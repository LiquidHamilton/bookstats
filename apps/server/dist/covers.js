import { createHash, randomBytes, randomUUID } from "node:crypto";
import { lookup } from "node:dns/promises";
import { mkdir, readFile, rename, stat, unlink, writeFile } from "node:fs/promises";
import { isIP } from "node:net";
import { dirname, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
const MAX_COVER_BYTES = 12 * 1024 * 1024;
const REMOTE_TIMEOUT_MS = 15_000;
const MAX_REDIRECTS = 4;
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const PROJECT_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "../../..");
export function coverStorageRoot() {
    return resolve(process.env.BOOKSTATS_COVER_DIR?.trim() || resolve(PROJECT_ROOT, "data", "covers"));
}
export function absoluteCoverAssetPath(storagePath) {
    const root = coverStorageRoot();
    const candidate = resolve(root, storagePath);
    if (candidate !== root && !candidate.startsWith(`${root}${sep}`))
        throw new Error("Invalid cover asset path.");
    return candidate;
}
export async function archiveCoverForUser(db, userId, coverValue) {
    const prepared = coverValue.startsWith("data:") ? prepareDataUrl(coverValue) : await fetchRemoteCover(coverValue);
    if (!prepared.bytes.length || prepared.bytes.length > MAX_COVER_BYTES)
        throw new Error("Cover image is too large.");
    const contentHash = createHash("sha256").update(prepared.bytes).digest("hex");
    const existing = await db.query(`SELECT id, access_token, mime_type, byte_size, storage_path, source_url
       FROM cover_assets WHERE user_id = $1 AND content_sha256 = $2 LIMIT 1`, [userId, contentHash]);
    if (existing.rows[0]) {
        const row = existing.rows[0];
        try {
            await stat(absoluteCoverAssetPath(row.storage_path));
        }
        catch {
            await mkdir(resolve(coverStorageRoot(), userId), { recursive: true });
            await writeCoverFile(row.storage_path, prepared.bytes);
        }
        if (!row.source_url && prepared.sourceUrl) {
            await db.query("UPDATE cover_assets SET source_url = $1, updated_at = now() WHERE id = $2 AND user_id = $3", [prepared.sourceUrl, row.id, userId]);
        }
        return rowToRecord({ ...row, source_url: row.source_url ?? prepared.sourceUrl ?? null });
    }
    const id = randomUUID();
    const accessToken = randomBytes(32).toString("base64url");
    const storagePath = `${userId}/${id}.${prepared.extension}`;
    await mkdir(resolve(coverStorageRoot(), userId), { recursive: true });
    await writeCoverFile(storagePath, prepared.bytes);
    try {
        const inserted = await db.query(`INSERT INTO cover_assets (id, user_id, content_sha256, access_token, mime_type, byte_size, storage_path, source_url)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
       RETURNING id, access_token, mime_type, byte_size, storage_path, source_url`, [id, userId, contentHash, accessToken, prepared.mimeType, prepared.bytes.length, storagePath, prepared.sourceUrl ?? null]);
        return rowToRecord(inserted.rows[0]);
    }
    catch (error) {
        // If another request archived the same image concurrently, reuse that row.
        const raced = await db.query(`SELECT id, access_token, mime_type, byte_size, storage_path, source_url
         FROM cover_assets WHERE user_id = $1 AND content_sha256 = $2 LIMIT 1`, [userId, contentHash]);
        if (raced.rows[0]) {
            if (raced.rows[0].storage_path !== storagePath)
                await unlink(absoluteCoverAssetPath(storagePath)).catch(() => undefined);
            return rowToRecord(raced.rows[0]);
        }
        // Do not leave an unreferenced file behind when the database insert fails.
        await unlink(absoluteCoverAssetPath(storagePath)).catch(() => undefined);
        throw error;
    }
}
export async function getCoverAssetByToken(db, id, accessToken) {
    if (!UUID_RE.test(id) || !accessToken)
        return undefined;
    const result = await db.query(`SELECT id, access_token, mime_type, byte_size, storage_path, source_url
       FROM cover_assets WHERE id = $1 AND access_token = $2 LIMIT 1`, [id, accessToken]);
    return result.rows[0] ? rowToRecord(result.rows[0]) : undefined;
}
export async function userOwnsCoverAsset(db, userId, id, accessToken) {
    if (!id || !accessToken || !UUID_RE.test(id))
        return false;
    const result = await db.query("SELECT 1 FROM cover_assets WHERE id = $1 AND user_id = $2 AND access_token = $3 LIMIT 1", [id, userId, accessToken]);
    return Boolean(result.rows[0]);
}
export async function listCoverAssetPathsForUser(db, userId) {
    const result = await db.query("SELECT storage_path FROM cover_assets WHERE user_id = $1", [userId]);
    return result.rows.map((row) => row.storage_path);
}
export async function removeStoredCoverFiles(storagePaths) {
    await Promise.all(storagePaths.map((storagePath) => unlink(absoluteCoverAssetPath(storagePath)).catch(() => undefined)));
}
async function writeCoverFile(storagePath, bytes) {
    const destination = absoluteCoverAssetPath(storagePath);
    const temp = `${destination}.${randomBytes(6).toString("hex")}.tmp`;
    await writeFile(temp, bytes, { mode: 0o600 });
    await rename(temp, destination);
}
function rowToRecord(row) {
    return {
        id: row.id,
        accessToken: row.access_token,
        mimeType: row.mime_type,
        byteSize: Number(row.byte_size),
        storagePath: row.storage_path,
        sourceUrl: row.source_url ?? undefined
    };
}
function prepareDataUrl(value) {
    const match = /^data:([^;,]+);base64,([A-Za-z0-9+/=\r\n]+)$/.exec(value);
    if (!match)
        throw new Error("Custom cover data is not a supported base64 image.");
    if (!match[1].toLowerCase().startsWith("image/"))
        throw new Error("Custom cover data is not an image.");
    const encoded = match[2].replace(/\s+/g, "");
    if (encoded.length > Math.ceil(MAX_COVER_BYTES * 4 / 3) + 16)
        throw new Error("Cover image is too large.");
    return identifyImage(Buffer.from(encoded, "base64"));
}
async function fetchRemoteCover(value) {
    let current = new URL(value);
    for (let redirect = 0; redirect <= MAX_REDIRECTS; redirect += 1) {
        await assertSafeRemoteUrl(current);
        const response = await fetch(current, {
            redirect: "manual",
            signal: AbortSignal.timeout(REMOTE_TIMEOUT_MS),
            headers: { "user-agent": process.env.BOOKSTATS_METADATA_USER_AGENT ?? "BookStats cover archiver" }
        });
        if (response.status >= 300 && response.status < 400) {
            const location = response.headers.get("location");
            if (!location || redirect === MAX_REDIRECTS)
                throw new Error("Cover URL redirected too many times.");
            current = new URL(location, current);
            continue;
        }
        if (!response.ok)
            throw new Error(`Cover URL returned HTTP ${response.status}.`);
        const declaredLength = Number(response.headers.get("content-length") ?? 0);
        if (declaredLength > MAX_COVER_BYTES)
            throw new Error("Cover image is too large.");
        if (!response.body)
            throw new Error("Cover URL returned no image data.");
        const reader = response.body.getReader();
        const chunks = [];
        let total = 0;
        while (true) {
            const { done, value: chunk } = await reader.read();
            if (done)
                break;
            if (!chunk)
                continue;
            total += chunk.byteLength;
            if (total > MAX_COVER_BYTES) {
                await reader.cancel();
                throw new Error("Cover image is too large.");
            }
            chunks.push(Buffer.from(chunk));
        }
        const identified = identifyImage(Buffer.concat(chunks, total));
        return { ...identified, sourceUrl: value };
    }
    throw new Error("Could not download cover image.");
}
async function assertSafeRemoteUrl(url) {
    if (url.protocol !== "https:" && url.protocol !== "http:")
        throw new Error("Cover URL must use HTTP or HTTPS.");
    if (url.username || url.password)
        throw new Error("Cover URL credentials are not allowed.");
    const hostname = url.hostname.replace(/^\[|\]$/g, "").toLowerCase();
    if (!hostname || hostname === "localhost" || hostname.endsWith(".localhost"))
        throw new Error("Local network cover URLs are not allowed.");
    const addresses = isIP(hostname) ? [{ address: hostname }] : await lookup(hostname, { all: true, verbatim: true });
    if (!addresses.length || addresses.some(({ address }) => !isPublicAddress(address)))
        throw new Error("Private or local network cover URLs are not allowed.");
}
function isPublicAddress(address) {
    if (address.includes(":")) {
        const value = address.toLowerCase();
        if (value === "::" || value === "::1" || value.startsWith("fe8") || value.startsWith("fe9") || value.startsWith("fea") || value.startsWith("feb") || value.startsWith("fc") || value.startsWith("fd"))
            return false;
        if (value.startsWith("::ffff:"))
            return false;
        return true;
    }
    return isPublicIpv4(address);
}
function isPublicIpv4(address) {
    const parts = address.split(".").map(Number);
    if (parts.length !== 4 || parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255))
        return false;
    const [a, b] = parts;
    if (a === 0 || a === 10 || a === 127 || a >= 224)
        return false;
    if (a === 100 && b >= 64 && b <= 127)
        return false;
    if (a === 169 && b === 254)
        return false;
    if (a === 172 && b >= 16 && b <= 31)
        return false;
    if (a === 192 && b === 168)
        return false;
    if (a === 198 && (b === 18 || b === 19))
        return false;
    return true;
}
function identifyImage(bytes) {
    if (bytes.length < 12)
        throw new Error("Cover URL did not return a recognizable image.");
    if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff)
        return { bytes, mimeType: "image/jpeg", extension: "jpg" };
    if (bytes.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])))
        return { bytes, mimeType: "image/png", extension: "png" };
    const prefix = bytes.subarray(0, 6).toString("ascii");
    if (prefix === "GIF87a" || prefix === "GIF89a")
        return { bytes, mimeType: "image/gif", extension: "gif" };
    if (bytes.subarray(0, 4).toString("ascii") === "RIFF" && bytes.subarray(8, 12).toString("ascii") === "WEBP")
        return { bytes, mimeType: "image/webp", extension: "webp" };
    throw new Error("Cover URL did not return a supported JPEG, PNG, GIF, or WebP image.");
}
// Exported only for release/migration validation.
export async function readStoredCover(asset) {
    return readFile(absoluteCoverAssetPath(asset.storagePath));
}
