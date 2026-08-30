import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import cors from "@fastify/cors";
import { createHash, randomBytes, scrypt as scryptCallback, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";
import { Pool, type PoolClient } from "pg";
import type { Book, MetadataCandidate, ReadingGoal, Shelf, SyncAcknowledgement, SyncMutation, SyncRecord, UserAccount } from "@bookstats/domain";
import { getMetadataDetails, metadataProviderStatuses, searchAllMetadata } from "./metadata.js";
import { emailConfigured, emailProvider, feedbackConfigured, sendFeedbackEmail, sendPasswordChangedEmail, sendPasswordResetEmail, sendVerificationEmail } from "./email.js";
import { registerAdminRoutes } from "./admin.js";
import { archiveCoverForUser, fetchRemoteCoverForInspection, getCoverAssetByToken, listCoverAssetPathsForUser, readStoredCover, removeStoredCoverFiles } from "./covers.js";

const APP_VERSION = "1.2.5";
const MIN_CLIENT_VERSION = "1.0.1";
const SESSION_DAYS = 30;
const MIN_PASSWORD_LENGTH = 10;
const VERIFICATION_HOURS = 24;
const RESET_HOURS = 1;
const TOKEN_THROTTLE_MINUTES = 2;
const scrypt = promisify(scryptCallback);

for (const candidate of [process.env.BOOKSTATS_ENV_FILE, ".env", "../../.env"].filter(Boolean) as string[]) {
  try { process.loadEnvFile?.(candidate); break; } catch { /* try the next optional location */ }
}

const app = Fastify({ logger: true, bodyLimit: 25 * 1024 * 1024, trustProxy: "127.0.0.1" });
const configuredOrigins = process.env.BOOKSTATS_CORS_ORIGIN?.split(",").map((item: string) => item.trim()).filter(Boolean);
const allowedOrigins = configuredOrigins?.length ? expandWebOriginAliases(configuredOrigins) : undefined;
await app.register(cors, {
  origin: allowedOrigins?.length ? allowedOrigins : true,
  credentials: true,
  allowedHeaders: ["content-type", "authorization", "x-bookstats-client-version"]
});

app.addHook("onRequest", async (request, reply) => {
  if (request.method === "OPTIONS") return;
  const path = request.url.split("?", 1)[0];
  if (!path.startsWith("/api/v1/")) return;
  // Cover images are loaded by normal <img> requests and cannot attach the client-version header.
  if (request.method === "GET" && /^\/api\/v1\/covers\/[^/]+\/[^/]+$/.test(path)) return;
  const header = request.headers["x-bookstats-client-version"];
  const clientVersion = Array.isArray(header) ? header[0] : header;
  if (!clientVersion || compareAppVersions(clientVersion, MIN_CLIENT_VERSION) < 0) {
    return reply.code(426).send({
      error: "BookStats has been updated. Refresh the page to continue.",
      serverVersion: APP_VERSION,
      minimumClientVersion: MIN_CLIENT_VERSION
    });
  }
});

function expandWebOriginAliases(origins: string[]): string[] {
  const expanded = new Set(origins);
  for (const origin of origins) {
    try {
      const url = new URL(origin);
      if (url.protocol !== "https:" && url.protocol !== "http:") continue;
      if (url.hostname === "localhost" || url.hostname.endsWith(".localhost") || /^[0-9.]+$/.test(url.hostname)) continue;
      const alias = new URL(url.toString());
      alias.hostname = url.hostname.startsWith("www.") ? url.hostname.slice(4) : `www.${url.hostname}`;
      expanded.add(alias.origin);
    } catch { /* non-web origins such as tauri://localhost are left unchanged */ }
  }
  return [...expanded];
}

function compareAppVersions(left: string, right: string): number {
  const parse = (value: string) => {
    const match = value.trim().match(/^(\d+)\.(\d+)\.(\d+)/);
    return match ? [Number(match[1]), Number(match[2]), Number(match[3])] : [0, 0, 0];
  };
  const a = parse(left); const b = parse(right);
  for (let i = 0; i < 3; i += 1) { if (a[i] !== b[i]) return a[i] - b[i]; }
  return 0;
}

const databaseUrl = process.env.DATABASE_URL;
const pool = databaseUrl ? new Pool({ connectionString: databaseUrl }) : undefined;
registerAdminRoutes(app, pool, APP_VERSION);

app.get("/health", async () => ({
  ok: true,
  service: "bookstats-api",
  version: APP_VERSION,
  databaseConfigured: Boolean(pool),
  emailConfigured: emailConfigured(),
  emailProvider: emailProvider(),
  feedbackConfigured: feedbackConfigured(),
  metadataProvider: metadataProviderStatuses().filter((provider) => provider.configured).map((provider) => provider.label).join(" + "),
  metadataProviders: metadataProviderStatuses()
}));

app.get("/api/v1", async () => ({ name: "BookStats API", version: 1, appVersion: APP_VERSION, minimumClientVersion: MIN_CLIENT_VERSION, status: "accounts-sync-metadata-email-feedback-cover-assets" }));


app.get<{ Querystring: { url?: string } }>("/api/v1/covers/inspect", async (request, reply) => {
  const coverUrl = request.query.url?.trim();
  if (!coverUrl) return reply.code(400).send({ error: "A cover URL is required." });
  try {
    const image = await fetchRemoteCoverForInspection(coverUrl);
    reply.header("Cache-Control", "private, max-age=86400");
    reply.header("Content-Length", String(image.bytes.length));
    reply.type(image.mimeType);
    return reply.send(image.bytes);
  } catch (error) {
    request.log.debug({ error, coverUrl }, "Could not inspect catalog cover");
    return reply.code(422).send({ error: "Catalog cover could not be inspected." });
  }
});

app.get<{ Params: { assetId: string; accessToken: string } }>("/api/v1/covers/:assetId/:accessToken", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const asset = await getCoverAssetByToken(db, request.params.assetId, request.params.accessToken);
  if (!asset) return reply.code(404).send({ error: "Cover not found." });
  try {
    const bytes = await readStoredCover(asset);
    reply.header("Cache-Control", "private, max-age=31536000, immutable");
    reply.type(asset.mimeType);
    return reply.send(bytes);
  } catch {
    return reply.code(404).send({ error: "Cover file is unavailable." });
  }
});

app.post<{ Body: { coverUrl?: string } }>("/api/v1/covers/archive", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  if (!request.bookstatsUser.emailVerified) return reply.code(403).send({ error: "Verify your email address before storing cloud cover assets." });
  const coverUrl = request.body?.coverUrl?.trim();
  if (!coverUrl) return reply.code(400).send({ error: "A selected cover is required." });
  try {
    const asset = await archiveCoverForUser(db, request.bookstatsUser.id, coverUrl);
    return { assetId: asset.id, assetToken: asset.accessToken, sourceUrl: asset.sourceUrl };
  } catch (error) {
    request.log.warn({ error }, "Could not archive selected cover");
    return reply.code(422).send({ error: error instanceof Error ? error.message : "BookStats could not archive that cover." });
  }
});

const feedbackWindows = new Map<string, number[]>();
app.post<{ Body: { kind?: string; message?: string; contactEmail?: string; diagnostics?: Record<string, unknown> } }>("/api/v1/feedback", async (request, reply) => {
  if (!feedbackConfigured()) return reply.code(503).send({ error: "Feedback delivery is not configured on this BookStats server." });
  const kind = request.body?.kind === "bug" ? "bug" : request.body?.kind === "feature" ? "feature" : undefined;
  const message = request.body?.message?.trim();
  if (!kind || !message) return reply.code(400).send({ error: "Choose a feedback type and enter a message." });
  if (message.length > 6000) return reply.code(400).send({ error: "Feedback messages can be up to 6,000 characters." });
  const contactEmail = request.body?.contactEmail?.trim();
  if (contactEmail && !normalizeEmail(contactEmail)) return reply.code(400).send({ error: "Enter a valid contact email or leave it blank." });

  const now = Date.now();
  const cutoff = now - 15 * 60 * 1000;
  const attempts = (feedbackWindows.get(request.ip) ?? []).filter((timestamp) => timestamp > cutoff);
  if (attempts.length >= 5) return reply.code(429).send({ error: "Too many feedback messages were sent recently. Please try again later." });
  attempts.push(now); feedbackWindows.set(request.ip, attempts);

  const rawDiagnostics = request.body?.diagnostics ?? {};
  const allowedKeys = ["version", "platform", "storage", "signedIn", "emailVerified", "bookCount", "shelfCount"] as const;
  const diagnostics: Record<string, string | number | boolean | undefined> = {};
  for (const key of allowedKeys) {
    const value = rawDiagnostics[key];
    if (typeof value === "string") diagnostics[key] = value.slice(0, 160);
    else if (typeof value === "number" || typeof value === "boolean") diagnostics[key] = value;
  }
  try {
    await sendFeedbackEmail({ kind, message, contactEmail: contactEmail || undefined, diagnostics });
    return { ok: true };
  } catch (error) {
    request.log.error(error, "Could not send BookStats feedback");
    return reply.code(502).send({ error: "BookStats could not send that feedback right now." });
  }
});

app.post<{ Body: { email?: string; password?: string; confirmPassword?: string; displayName?: string } }>("/api/v1/auth/register", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const email = normalizeEmail(request.body?.email);
  const displayName = request.body?.displayName?.trim();
  const password = request.body?.password ?? "";
  const confirmPassword = request.body?.confirmPassword ?? "";
  const passwordError = validateNewPassword(password, confirmPassword);
  if (!email || !displayName || passwordError) {
    return reply.code(400).send({ error: passwordError ?? `Email, display name, and a password of at least ${MIN_PASSWORD_LENGTH} characters are required.` });
  }

  const client = await db.connect();
  try {
    const passwordHash = await hashPassword(password);
    const result = await client.query<AccountRow>(
      `INSERT INTO users (email, display_name, password_hash) VALUES ($1, $2, $3)
       RETURNING id, email, display_name, created_at, email_verified_at`, [email, displayName, passwordHash]
    );
    const user = accountFromRow(result.rows[0]);
    const token = await createSession(client, user.id);
    let emailVerificationSent = false;
    if (emailConfigured()) {
      try {
        const verificationToken = await issueToken(client, "email_verification_tokens", user.id, VERIFICATION_HOURS);
        await sendVerificationEmail(user.email, user.displayName, verificationToken);
        emailVerificationSent = true;
      } catch (error) { request.log.error(error, "Could not send verification email"); }
    }
    return reply.code(201).send({ token, user, emailVerificationSent });
  } catch (error) {
    if (isUniqueViolation(error)) return reply.code(409).send({ error: "An account already exists for that email address." });
    request.log.error(error);
    return reply.code(500).send({ error: "Could not create the account." });
  } finally { client.release(); }
});

app.post<{ Body: { email?: string; password?: string } }>("/api/v1/auth/login", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const email = normalizeEmail(request.body?.email);
  const password = request.body?.password ?? "";
  if (!email || !password) return reply.code(400).send({ error: "Email and password are required." });
  const client = await db.connect();
  try {
    const result = await client.query<AccountRow & { password_hash: string; disabled_at: Date | null }>(
      "SELECT id, email, display_name, created_at, email_verified_at, password_hash, disabled_at FROM users WHERE email = $1", [email]
    );
    const row = result.rows[0];
    if (!row || !(await verifyPassword(password, row.password_hash))) return reply.code(401).send({ error: "Incorrect email or password." });
    if (row.disabled_at) return reply.code(403).send({ error: "This BookStats account has been disabled. Contact the administrator for assistance." });
    const user = accountFromRow(row);
    const token = await createSession(client, user.id);
    return { token, user };
  } finally { client.release(); }
});

app.get("/api/v1/auth/me", { preHandler: authenticate }, async (request) => ({ user: request.bookstatsUser }));

app.post("/api/v1/auth/logout", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  if (request.bookstatsTokenHash) await db.query("DELETE FROM sessions WHERE token_hash = $1", [request.bookstatsTokenHash]);
  return { ok: true };
});

app.post<{ Body: { currentPassword?: string; password?: string; confirmPassword?: string } }>("/api/v1/auth/change-password", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  const currentPassword = request.body?.currentPassword ?? "";
  const password = request.body?.password ?? "";
  const confirmPassword = request.body?.confirmPassword ?? "";
  const passwordError = validateNewPassword(password, confirmPassword);
  if (!currentPassword || passwordError) return reply.code(400).send({ error: passwordError ?? "Current password is required." });
  const result = await db.query<{ password_hash: string }>("SELECT password_hash FROM users WHERE id = $1", [request.bookstatsUser.id]);
  if (!result.rows[0] || !(await verifyPassword(currentPassword, result.rows[0].password_hash))) return reply.code(401).send({ error: "Current password is incorrect." });
  const passwordHash = await hashPassword(password);
  await db.query("UPDATE users SET password_hash = $1, updated_at = now() WHERE id = $2", [passwordHash, request.bookstatsUser.id]);
  if (request.bookstatsTokenHash) await db.query("DELETE FROM sessions WHERE user_id = $1 AND token_hash <> $2", [request.bookstatsUser.id, request.bookstatsTokenHash]);
  if (emailConfigured()) void sendPasswordChangedEmail(request.bookstatsUser.email, request.bookstatsUser.displayName).catch((error) => request.log.error(error, "Could not send password changed email"));
  return { ok: true };
});

app.post<{ Body: { currentPassword?: string; email?: string } }>("/api/v1/auth/change-email", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  const currentPassword = request.body?.currentPassword ?? "";
  const email = normalizeEmail(request.body?.email);
  if (!currentPassword || !email) return reply.code(400).send({ error: "Current password and a valid new email are required." });
  const client = await db.connect();
  try {
    const result = await client.query<{ password_hash: string }>("SELECT password_hash FROM users WHERE id = $1", [request.bookstatsUser.id]);
    if (!result.rows[0] || !(await verifyPassword(currentPassword, result.rows[0].password_hash))) return reply.code(401).send({ error: "Current password is incorrect." });
    await client.query("BEGIN");
    const changed = await client.query<AccountRow>(`UPDATE users SET email = $1, email_verified_at = NULL, updated_at = now() WHERE id = $2 RETURNING id, email, display_name, created_at, email_verified_at`, [email, request.bookstatsUser.id]);
    let verificationSent = false;
    if (emailConfigured()) {
      const token = await issueToken(client, "email_verification_tokens", request.bookstatsUser.id, VERIFICATION_HOURS);
      try { await sendVerificationEmail(email, request.bookstatsUser.displayName, token); verificationSent = true; }
      catch (error) { request.log.error(error, "Could not send email-change verification message"); }
    }
    await client.query("COMMIT");
    return { ok: true, user: accountFromRow(changed.rows[0]), verificationSent };
  } catch (error) {
    await client.query("ROLLBACK");
    if (isUniqueViolation(error)) return reply.code(409).send({ error: "An account already exists for that email address." });
    request.log.error(error); return reply.code(500).send({ error: "Could not change the email address." });
  } finally { client.release(); }
});

app.delete("/api/v1/account/cloud-library", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  const result = await db.query("DELETE FROM library_records WHERE user_id = $1", [request.bookstatsUser.id]);
  return { ok: true, deleted: result.rowCount ?? 0 };
});

app.post<{ Body: { password?: string } }>("/api/v1/account/delete", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  const password = request.body?.password ?? "";
  const result = await db.query<{ password_hash: string }>("SELECT password_hash FROM users WHERE id = $1", [request.bookstatsUser.id]);
  if (!result.rows[0] || !(await verifyPassword(password, result.rows[0].password_hash))) return reply.code(401).send({ error: "Password is incorrect." });
  const coverPaths = await listCoverAssetPathsForUser(db, request.bookstatsUser.id);
  await db.query("DELETE FROM users WHERE id = $1", [request.bookstatsUser.id]);
  await removeStoredCoverFiles(coverPaths);
  return { ok: true };
});

app.post("/api/v1/auth/resend-verification", { preHandler: authenticate }, async (request, reply) => {
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  if (request.bookstatsUser.emailVerified) return { ok: true, alreadyVerified: true };
  if (!emailConfigured()) return reply.code(503).send({ error: "Email delivery is not configured on this BookStats server." });
  const client = await db.connect();
  try {
    const recent = await hasRecentToken(client, "email_verification_tokens", request.bookstatsUser.id);
    if (recent) return { ok: true, throttled: true };
    const token = await issueToken(client, "email_verification_tokens", request.bookstatsUser.id, VERIFICATION_HOURS);
    await sendVerificationEmail(request.bookstatsUser.email, request.bookstatsUser.displayName, token);
    return { ok: true };
  } catch (error) {
    request.log.error(error);
    return reply.code(502).send({ error: "Could not send the verification email." });
  } finally { client.release(); }
});

app.post<{ Body: { token?: string } }>("/api/v1/auth/verify-email", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const token = request.body?.token?.trim();
  if (!token) return reply.code(400).send({ error: "Verification token is required." });
  const client = await db.connect();
  try {
    await client.query("BEGIN");
    const result = await client.query<AccountRow & { token_id: string }>(
      `SELECT u.id, u.email, u.display_name, u.created_at, u.email_verified_at, t.id AS token_id
       FROM email_verification_tokens t JOIN users u ON u.id = t.user_id
       WHERE t.token_hash = $1 AND t.used_at IS NULL AND t.expires_at > now()
       FOR UPDATE OF t`, [sha256(token)]
    );
    const row = result.rows[0];
    if (!row) { await client.query("ROLLBACK"); return reply.code(400).send({ error: "This verification link is invalid or has expired." }); }
    await client.query("UPDATE users SET email_verified_at = COALESCE(email_verified_at, now()), updated_at = now() WHERE id = $1", [row.id]);
    await client.query("UPDATE email_verification_tokens SET used_at = now() WHERE user_id = $1 AND used_at IS NULL", [row.id]);
    const refreshed = await client.query<AccountRow>("SELECT id, email, display_name, created_at, email_verified_at FROM users WHERE id = $1", [row.id]);
    await client.query("COMMIT");
    return { ok: true, user: accountFromRow(refreshed.rows[0]) };
  } catch (error) {
    await client.query("ROLLBACK"); request.log.error(error);
    return reply.code(500).send({ error: "Could not verify the email address." });
  } finally { client.release(); }
});

app.post<{ Body: { email?: string } }>("/api/v1/auth/forgot-password", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const email = normalizeEmail(request.body?.email);
  const generic = { ok: true, message: "If that email belongs to a BookStats account, a password reset link will be sent." };
  if (!email) return generic;
  const client = await db.connect();
  try {
    const result = await client.query<AccountRow>("SELECT id, email, display_name, created_at, email_verified_at FROM users WHERE email = $1", [email]);
    const user = result.rows[0] ? accountFromRow(result.rows[0]) : undefined;
    if (!user || !emailConfigured()) return generic;
    if (await hasRecentToken(client, "password_reset_tokens", user.id)) return generic;
    const token = await issueToken(client, "password_reset_tokens", user.id, RESET_HOURS);
    try { await sendPasswordResetEmail(user.email, user.displayName, token); }
    catch (error) { request.log.error(error, "Could not send password reset email"); }
    return generic;
  } finally { client.release(); }
});

app.post<{ Body: { token?: string; password?: string; confirmPassword?: string } }>("/api/v1/auth/reset-password", async (request, reply) => {
  const db = requireDb(reply); if (!db) return;
  const token = request.body?.token?.trim();
  const password = request.body?.password ?? "";
  const confirmPassword = request.body?.confirmPassword ?? "";
  const passwordError = validateNewPassword(password, confirmPassword);
  if (!token || passwordError) return reply.code(400).send({ error: passwordError ?? "Reset token is required." });
  const client = await db.connect();
  let changedUser: UserAccount | undefined;
  try {
    await client.query("BEGIN");
    const result = await client.query<AccountRow & { token_id: string }>(
      `SELECT u.id, u.email, u.display_name, u.created_at, u.email_verified_at, t.id AS token_id
       FROM password_reset_tokens t JOIN users u ON u.id = t.user_id
       WHERE t.token_hash = $1 AND t.used_at IS NULL AND t.expires_at > now()
       FOR UPDATE OF t`, [sha256(token)]
    );
    const row = result.rows[0];
    if (!row) { await client.query("ROLLBACK"); return reply.code(400).send({ error: "This password reset link is invalid or has expired." }); }
    const passwordHash = await hashPassword(password);
    await client.query("UPDATE users SET password_hash = $1, updated_at = now() WHERE id = $2", [passwordHash, row.id]);
    await client.query("UPDATE password_reset_tokens SET used_at = now() WHERE user_id = $1 AND used_at IS NULL", [row.id]);
    await client.query("DELETE FROM sessions WHERE user_id = $1", [row.id]);
    await client.query("COMMIT");
    changedUser = accountFromRow(row);
  } catch (error) {
    await client.query("ROLLBACK"); request.log.error(error);
    return reply.code(500).send({ error: "Could not reset the password." });
  } finally { client.release(); }
  if (changedUser && emailConfigured()) {
    void sendPasswordChangedEmail(changedUser.email, changedUser.displayName).catch((error) => request.log.error(error, "Could not send password changed email"));
  }
  return { ok: true };
});

app.post<{ Body: { cursor?: string; changes?: SyncMutation[] } }>("/api/v1/sync", { preHandler: authenticate }, async (request, reply) => {
  const syncStartedAt = performance.now();
  const requestBytes = Number(request.headers["content-length"] ?? 0) || undefined;
  const db = requireDb(reply); if (!db || !request.bookstatsUser) return;
  if (!request.bookstatsUser.emailVerified) return reply.code(403).send({ error: "Verify your email address before enabling cloud synchronization." });
  const rawChanges = Array.isArray(request.body?.changes) ? request.body.changes.slice(0, 10_000) : [];
  const assetIds = [...new Set(rawChanges.flatMap((change) => change?.book?.coverAssetId && /^[0-9a-f-]{36}$/i.test(change.book.coverAssetId) ? [change.book.coverAssetId] : []))];
  const validCoverAssets = new Set<string>();
  if (assetIds.length) {
    const ownedAssets = await db.query<{ id: string; access_token: string }>(
      "SELECT id, access_token FROM cover_assets WHERE user_id = $1 AND id::text = ANY($2::text[])", [request.bookstatsUser.id, assetIds]
    );
    for (const row of ownedAssets.rows) validCoverAssets.add(`${row.id}:${row.access_token}`);
  }
  const changes: SyncMutation[] = [];
  for (const change of rawChanges) {
    if (change?.entityType === "shelf" || change?.entityType === "goal" || change?.deleted || !change?.book) { changes.push(change); continue; }
    changes.push({ ...change, book: await prepareBookCoverForCloud(db, request.bookstatsUser.id, change.book, validCoverAssets, request.log) });
  }
  const cursor = parseCursor(request.body?.cursor);
  const client = await db.connect();
  let accepted = 0;
  const acknowledged: SyncAcknowledgement[] = [];
  try {
    await client.query("BEGIN");
    for (const change of changes) {
      if (!change?.id || !change.clientUpdatedAt) continue;
      const entityType = change.entityType === "shelf" ? "shelf" : change.entityType === "goal" ? "goal" : "book";
      const payload = entityType === "shelf" ? change.shelf : entityType === "goal" ? change.goal : change.book;
      if (!change.deleted && (!payload || payload.id !== change.id)) continue;
      const incomingTime = new Date(change.clientUpdatedAt);
      if (Number.isNaN(incomingTime.getTime())) continue;
      const existing = await client.query<{ client_updated_at: Date }>(
        "SELECT client_updated_at FROM library_records WHERE user_id = $1 AND id = $2 FOR UPDATE", [request.bookstatsUser.id, change.id]
      );
      const acknowledgement: SyncAcknowledgement = { id: change.id, entityType, deleted: Boolean(change.deleted), clientUpdatedAt: incomingTime.toISOString() };
      // Incremental retries are idempotent. If the server already has the same or a
      // newer client timestamp, this mutation no longer needs to remain in the device
      // outbox even though it should not create another server revision.
      if (existing.rows[0] && existing.rows[0].client_updated_at.getTime() >= incomingTime.getTime()) {
        acknowledged.push(acknowledgement);
        continue;
      }
      await client.query(
        `INSERT INTO library_records (id, user_id, record_type, book_data, client_updated_at, deleted_at)
         VALUES ($1, $2, $3, $4::jsonb, $5, $6)
         ON CONFLICT (user_id, id) DO UPDATE SET
           record_type = EXCLUDED.record_type,
           book_data = EXCLUDED.book_data,
           client_updated_at = EXCLUDED.client_updated_at,
           deleted_at = EXCLUDED.deleted_at,
           revision = library_records.revision + 1,
           updated_at = now()`,
        [change.id, request.bookstatsUser.id, entityType, change.deleted ? null : JSON.stringify(payload ?? null), incomingTime.toISOString(), change.deleted ? incomingTime.toISOString() : null]
      );
      accepted += 1;
      acknowledged.push(acknowledgement);
    }
    const serverCursorResult = await client.query<{ cursor: Date }>("SELECT now() AS cursor");
    const serverCursor = serverCursorResult.rows[0].cursor.toISOString();
    const pulled = await client.query<{
      id: string; record_type: "book" | "shelf" | "goal"; book_data: Book | Shelf | ReadingGoal | null; client_updated_at: Date; updated_at: Date; revision: string | number; deleted_at: Date | null;
    }>(
      `SELECT id, record_type, book_data, client_updated_at, updated_at, revision, deleted_at
       FROM library_records WHERE user_id = $1 AND updated_at > $2 ORDER BY updated_at ASC`,
      [request.bookstatsUser.id, cursor]
    );
    await client.query("COMMIT");
    const records: SyncRecord[] = pulled.rows.map((row) => {
      const entityType = row.record_type === "shelf" ? "shelf" : row.record_type === "goal" ? "goal" : "book";
      return {
        id: row.id,
        entityType,
        deleted: Boolean(row.deleted_at),
        book: entityType === "book" ? (row.book_data as Book | null) ?? undefined : undefined,
        shelf: entityType === "shelf" ? (row.book_data as Shelf | null) ?? undefined : undefined,
        goal: entityType === "goal" ? (row.book_data as ReadingGoal | null) ?? undefined : undefined,
        clientUpdatedAt: row.client_updated_at.toISOString(),
        serverUpdatedAt: row.updated_at.toISOString(),
        revision: Number(row.revision)
      };
    });
    request.log.info({
      sync: {
        incomingRecords: rawChanges.length,
        requestBytes,
        accepted,
        acknowledged: acknowledged.length,
        pulled: records.length,
        durationMs: Math.round(performance.now() - syncStartedAt)
      }
    }, "BookStats sync batch complete");
    return { cursor: serverCursor, changes: records, accepted, acknowledged };
  } catch (error) {
    await client.query("ROLLBACK");
    request.log.error({ error, sync: { incomingRecords: rawChanges.length, requestBytes, durationMs: Math.round(performance.now() - syncStartedAt) } }, "BookStats sync batch failed");
    return reply.code(500).send({ error: "Synchronization failed." });
  } finally { client.release(); }
});

app.get<{ Querystring: { q?: string; isbn?: string } }>("/api/v1/metadata/search", async (request, reply) => {
  const q = request.query.q?.trim();
  const isbn = request.query.isbn?.replace(/[^0-9Xx]/g, "").toUpperCase();
  if (!q && !isbn) return reply.code(400).send({ error: "Provide q or isbn." });
  const providerKey = metadataProviderStatuses().filter((provider) => provider.configured).map((provider) => provider.id).sort().join(",");
  const cacheKey = `metadata-search-v6:${providerKey}:${isbn ? `isbn:${isbn}` : q!.toLowerCase()}`;
  const cached = await getCached<MetadataCandidate[]>(cacheKey);
  if (cached) return { results: cached, cached: true, providers: metadataProviderStatuses() };
  try {
    const results = await searchAllMetadata(q ?? isbn ?? "", isbn);
    await putCached(cacheKey, results, isbn ? 7 * 24 * 60 * 60 : 24 * 60 * 60);
    return { results, cached: false, providers: metadataProviderStatuses() };
  } catch (error) {
    request.log.error(error);
    return reply.code(502).send({ error: "Book lookup is temporarily unavailable." });
  }
});

app.post<{ Body: { candidate?: MetadataCandidate } }>("/api/v1/metadata/details", async (request, reply) => {
  const candidate = request.body?.candidate;
  if (!candidate?.workId || !candidate.title) return reply.code(400).send({ error: "A metadata candidate is required." });
  const refs = candidate.sourceRefs?.map((ref) => `${ref.provider}:${ref.workId}:${ref.editionId ?? ""}`).sort().join("|") ?? `${candidate.source}:${candidate.workId}:${candidate.editionId ?? ""}`;
  const providerKey = metadataProviderStatuses().filter((provider) => provider.configured).map((provider) => provider.id).sort().join(",");
  const cacheKey = `metadata-details-v6:${providerKey}:${candidate.exactEdition ? `exact:${candidate.isbn ?? ""}:` : ""}${refs}`;
  const cached = await getCached<MetadataCandidate>(cacheKey);
  if (cached) return { details: cached, cached: true };
  try {
    const details = await getMetadataDetails(candidate);
    await putCached(cacheKey, details, 7 * 24 * 60 * 60);
    return { details, cached: false };
  } catch (error) {
    request.log.error(error);
    return reply.code(502).send({ error: "Could not load additional book details." });
  }
});

async function prepareBookCoverForCloud(db: Pool, userId: string, book: Book, validCoverAssets: Set<string>, log: FastifyRequest["log"]): Promise<Book> {
  const selected = book.coverUrl?.trim();
  if (!book.coverAssetId && !book.coverArchivePending && !selected?.startsWith("data:")) return book;
  const next: Book = { ...book };
  const hasOwnedAsset = Boolean(next.coverAssetId && next.coverAssetToken && validCoverAssets.has(`${next.coverAssetId}:${next.coverAssetToken}`));
  if (!hasOwnedAsset) {
    next.coverAssetId = undefined;
    next.coverAssetToken = undefined;
    // A portable/imported record may contain an asset reference from another account.
    // Preserve a usable remote source when one exists instead of leaving the book coverless.
    if (!selected && next.coverSourceUrl && /^https?:\/\//i.test(next.coverSourceUrl)) next.coverUrl = next.coverSourceUrl;
  }

  const legacyCustomCover = Boolean(selected?.startsWith("data:"));
  if ((!hasOwnedAsset && legacyCustomCover) || (next.coverArchivePending && selected)) {
    try {
      const asset = await archiveCoverForUser(db, userId, selected!);
      next.coverAssetId = asset.id;
      next.coverAssetToken = asset.accessToken;
      next.coverSourceUrl = selected && /^https?:\/\//i.test(selected) ? selected : (next.coverSourceUrl ?? asset.sourceUrl);
      next.coverArchivePending = undefined;
      // Custom images are no longer embedded in synchronized JSON after the server owns a durable copy.
      if (legacyCustomCover) next.coverUrl = undefined;
    } catch (error) {
      log.warn({ error, bookId: book.id }, "Selected cover remains pending because it could not be archived");
      next.coverArchivePending = true;
    }
  } else if (hasOwnedAsset) {
    next.coverArchivePending = undefined;
  }
  return next;
}

async function authenticate(request: FastifyRequest, reply: FastifyReply) {
  const db = requireDb(reply); if (!db) return;
  const header = request.headers.authorization;
  const token = header?.startsWith("Bearer ") ? header.slice(7).trim() : "";
  if (!token) return reply.code(401).send({ error: "Authentication required." });
  const tokenHash = sha256(token);
  const result = await db.query<AccountRow>(
    `SELECT u.id, u.email, u.display_name, u.created_at, u.email_verified_at
     FROM sessions s JOIN users u ON u.id = s.user_id
     WHERE s.token_hash = $1 AND s.expires_at > now() AND u.disabled_at IS NULL`, [tokenHash]
  );
  if (!result.rows[0]) return reply.code(401).send({ error: "Session expired or invalid." });
  request.bookstatsUser = accountFromRow(result.rows[0]);
  request.bookstatsTokenHash = tokenHash;
  void db.query("UPDATE sessions SET last_used_at = now() WHERE token_hash = $1", [tokenHash]);
}

function requireDb(reply: FastifyReply): Pool | undefined {
  if (pool) return pool;
  void reply.code(503).send({ error: "Cloud features require DATABASE_URL and the BookStats PostgreSQL migrations." });
  return undefined;
}

function normalizeEmail(value?: string): string | undefined {
  const email = value?.trim().toLowerCase();
  return email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : undefined;
}

function validateNewPassword(password: string, confirmPassword: string): string | undefined {
  if (password.length < MIN_PASSWORD_LENGTH) return `Password must be at least ${MIN_PASSWORD_LENGTH} characters.`;
  if (password.length > 256) return "Password is too long.";
  if (password !== confirmPassword) return "Passwords do not match.";
  return undefined;
}

async function hashPassword(password: string): Promise<string> {
  const salt = randomBytes(16);
  const derived = await scrypt(password, salt, 64) as Buffer;
  return `scrypt:${salt.toString("hex")}:${derived.toString("hex")}`;
}

async function verifyPassword(password: string, encoded: string): Promise<boolean> {
  const [algorithm, saltHex, hashHex] = encoded.split(":");
  if (algorithm !== "scrypt" || !saltHex || !hashHex) return false;
  const expected = Buffer.from(hashHex, "hex");
  const actual = await scrypt(password, Buffer.from(saltHex, "hex"), expected.length) as Buffer;
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}

async function createSession(client: PoolClient, userId: string): Promise<string> {
  const token = randomBytes(32).toString("base64url");
  await client.query(
    `INSERT INTO sessions (user_id, token_hash, expires_at) VALUES ($1, $2, now() + ($3 || ' days')::interval)`,
    [userId, sha256(token), SESSION_DAYS]
  );
  return token;
}

type TokenTable = "email_verification_tokens" | "password_reset_tokens";
async function issueToken(client: PoolClient, table: TokenTable, userId: string, expiresHours: number): Promise<string> {
  const token = randomBytes(32).toString("base64url");
  await client.query(`UPDATE ${table} SET used_at = now() WHERE user_id = $1 AND used_at IS NULL`, [userId]);
  await client.query(
    `INSERT INTO ${table} (user_id, token_hash, expires_at) VALUES ($1, $2, now() + ($3 || ' hours')::interval)`,
    [userId, sha256(token), expiresHours]
  );
  return token;
}

async function hasRecentToken(client: PoolClient, table: TokenTable, userId: string): Promise<boolean> {
  const result = await client.query(
    `SELECT 1 FROM ${table} WHERE user_id = $1 AND used_at IS NULL AND created_at > now() - ($2 || ' minutes')::interval LIMIT 1`,
    [userId, TOKEN_THROTTLE_MINUTES]
  );
  return Boolean(result.rows[0]);
}

function sha256(value: string): string { return createHash("sha256").update(value).digest("hex"); }
function isUniqueViolation(error: unknown): boolean { return Boolean(error && typeof error === "object" && "code" in error && (error as { code?: string }).code === "23505"); }
function parseCursor(value?: string): string {
  if (!value) return "1970-01-01T00:00:00.000Z";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "1970-01-01T00:00:00.000Z" : date.toISOString();
}
interface AccountRow { id: string; email: string; display_name: string; created_at: Date; email_verified_at: Date | null; }
function accountFromRow(row: AccountRow): UserAccount {
  return { id: row.id, email: row.email, displayName: row.display_name, createdAt: row.created_at.toISOString(), emailVerified: Boolean(row.email_verified_at) };
}
async function getCached<T>(key: string): Promise<T | undefined> {
  if (!pool) return undefined;
  try {
    const result = await pool.query<{ payload: T }>("SELECT payload FROM metadata_cache WHERE cache_key = $1 AND expires_at > now()", [key]);
    return result.rows[0]?.payload;
  } catch { return undefined; }
}
async function putCached(key: string, payload: unknown, ttlSeconds: number): Promise<void> {
  if (!pool) return;
  try {
    await pool.query(
      `INSERT INTO metadata_cache (cache_key, payload, expires_at) VALUES ($1, $2::jsonb, now() + ($3 || ' seconds')::interval)
       ON CONFLICT (cache_key) DO UPDATE SET payload = EXCLUDED.payload, expires_at = EXCLUDED.expires_at, updated_at = now()`,
      [key, JSON.stringify(payload), ttlSeconds]
    );
  } catch { /* cache failure must not break lookup */ }
}

declare module "fastify" {
  interface FastifyRequest {
    bookstatsUser?: UserAccount;
    bookstatsTokenHash?: string;
  }
}

const host = process.env.BOOKSTATS_HOST ?? "127.0.0.1";
const port = Number(process.env.BOOKSTATS_PORT ?? 8787);
try { await app.listen({ host, port }); }
catch (error) { app.log.error(error); process.exit(1); }
