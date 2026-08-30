import { cloudCoverUrl, fetchCatalogCoverForInspection } from "./api";

const MAX_WIDTH = 700;
const MAX_HEIGHT = 1050;
const JPEG_QUALITY = 0.84;
const COVER_VALIDATION_CONCURRENCY = 6;
const INSPECTION_CACHE_VERSION = 3;
const INSPECTION_CACHE_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const PLACEHOLDER_URL_PATTERN = /(?:image[-_ ]?not[-_ ]?available|no[-_ ]?(?:image|cover)|missing[-_ ]?(?:image|cover)|default[-_ ]?(?:image|cover)|placeholder|cover[-_ ]?coming[-_ ]?soon|spacer|transparent[-_ ]?pixel)/i;

export interface CoverDisplayRecord {
  coverUrl?: string;
  coverAssetId?: string;
  coverAssetToken?: string;
  coverSourceUrl?: string;
  cachedCoverDataUrl?: string;
  updatedAt?: string;
}

export type CoverQualityReason =
  | "placeholder-url"
  | "unreachable"
  | "not-image"
  | "too-small"
  | "bad-aspect"
  | "blank"
  | "placeholder-image";

export interface CoverInspection {
  usable: boolean;
  reason?: CoverQualityReason;
  width?: number;
  height?: number;
  /** Higher is better. Used only to rank already-usable catalog candidates. */
  qualityScore: number;
  /** 64-bit average hash of normalized pixels. Useful for near-duplicate artwork removal. */
  perceptualHash?: string;
}

interface StoredInspection extends CoverInspection {
  checkedAt: number;
}

const inspectionPromises = new Map<string, Promise<CoverInspection>>();

export async function prepareUploadedCover(file: File): Promise<string> {
  if (!file.type.startsWith("image/")) throw new Error("Choose an image file for the book cover.");
  if (file.size > 12 * 1024 * 1024) throw new Error("Cover images must be smaller than 12 MB.");
  return resizeImage(file);
}

export async function cacheRemoteCover(url: string): Promise<string | undefined> {
  if (!/^https?:\/\//i.test(url)) return undefined;
  try {
    const response = await fetch(url, { cache: "force-cache" });
    if (!response.ok) return undefined;
    const blob = await response.blob();
    if (!blob.type.startsWith("image/")) return undefined;
    return await resizeImage(blob);
  } catch {
    // Some third-party cover hosts disallow browser CORS. The original URL remains usable.
    return undefined;
  }
}

async function resizeImage(blob: Blob): Promise<string> {
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    const scale = Math.min(1, MAX_WIDTH / image.naturalWidth, MAX_HEIGHT / image.naturalHeight);
    const width = Math.max(1, Math.round(image.naturalWidth * scale));
    const height = Math.max(1, Math.round(image.naturalHeight * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("This browser could not prepare the cover image.");
    context.drawImage(image, 0, 0, width, height);
    return canvas.toDataURL("image/jpeg", JPEG_QUALITY);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("The selected cover image could not be read."));
    image.src = src;
  });
}

/** Fast, network-free screening used by both the picker and Library Health. */
export function coverUrlLooksSuspiciousByName(url?: string): boolean {
  if (!url?.trim()) return false;
  return PLACEHOLDER_URL_PATTERN.test(safeDecodedUrl(url));
}

/**
 * Validate and rank catalog images. Results are cached by URL for the browser session and
 * persisted as small verdict/metric records so reopening the cover picker does not download
 * and analyze the same image again. Near-identical artwork is collapsed using a perceptual hash.
 */
export async function rankUsableCoverUrls(values: string[]): Promise<string[]> {
  const urls = [...new Set(values.map((value) => value.trim()).filter(Boolean))];
  if (!urls.length) return [];
  const inspections = await inspectMany(urls);
  const ranked = urls
    .map((url, index) => ({ url, index, inspection: inspections[index] }))
    .filter(({ inspection }) => inspection.usable)
    .sort((left, right) => right.inspection.qualityScore - left.inspection.qualityScore || left.index - right.index);

  const accepted: typeof ranked = [];
  for (const candidate of ranked) {
    const hash = candidate.inspection.perceptualHash;
    if (hash && accepted.some((existing) => existing.inspection.perceptualHash && perceptualHashDistance(hash, existing.inspection.perceptualHash!) <= 3)) continue;
    accepted.push(candidate);
  }
  return accepted.map(({ url }) => url);
}

/** Backward-compatible name; v1.2 also quality-ranks and near-deduplicates results. */
export async function filterUsableCoverUrls(values: string[]): Promise<string[]> {
  return rankUsableCoverUrls(values);
}

export async function inspectCoverUrl(url: string, forceRefresh = false): Promise<CoverInspection> {
  const normalized = url.trim();
  if (!normalized) return { usable: false, reason: "unreachable", qualityScore: 0 };
  const existing = inspectionPromises.get(normalized);
  if (existing) return existing;

  if (!forceRefresh) {
    const cached = readStoredInspection(normalized);
    if (cached) return cached;
  }

  const promise = inspectCoverUrlUncached(normalized)
    .then((inspection) => {
      writeStoredInspection(normalized, inspection);
      return inspection;
    })
    .catch(() => ({ usable: false, reason: "unreachable" as const, qualityScore: 0 }))
    .finally(() => inspectionPromises.delete(normalized));
  // Keep only in-flight promises. Once inspection settles, the compact persistent verdict
  // becomes the cache, avoiding an ever-growing in-memory map on very large libraries.
  inspectionPromises.set(normalized, promise);
  return promise;
}

/**
 * Inspect a selected cover using the cheapest representative first (normally the local cache).
 * A healthy fallback makes the record healthy even if one transient display source is unavailable.
 */
export async function inspectCoverRecord(book: CoverDisplayRecord, forceRefresh = false): Promise<CoverInspection | undefined> {
  const sources = coverSources(book);
  if (!sources.length) return undefined;
  let last: CoverInspection | undefined;
  for (const source of sources) {
    const inspection = await inspectCoverUrl(source, forceRefresh);
    last = inspection;
    if (inspection.usable) return inspection;
  }
  return last;
}

async function inspectMany(urls: string[]): Promise<CoverInspection[]> {
  const results = new Array<CoverInspection>(urls.length);
  let nextIndex = 0;
  async function worker(): Promise<void> {
    while (nextIndex < urls.length) {
      const index = nextIndex++;
      results[index] = await inspectCoverUrl(urls[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(COVER_VALIDATION_CONCURRENCY, urls.length) }, () => worker()));
  return results;
}

async function inspectCoverUrlUncached(url: string): Promise<CoverInspection> {
  if (coverUrlLooksSuspiciousByName(url)) return { usable: false, reason: "placeholder-url", qualityScore: 0 };

  let blob: Blob | undefined;
  if (/^data:image\//i.test(url)) {
    try { blob = await (await fetch(url)).blob(); } catch { return { usable: false, reason: "unreachable", qualityScore: 0 }; }
  } else if (/^https?:\/\//i.test(url)) {
    // Third-party hosts commonly allow <img> display while blocking browser fetch() with CORS.
    // Prefer the hardened BookStats proxy, then direct-fetch as a dev/backward-compatible fallback.
    blob = await fetchCatalogCoverForInspection(url);
    if (!blob) {
      try {
        const response = await fetch(url, { cache: "force-cache" });
        if (!response.ok) return { usable: false, reason: "unreachable", qualityScore: 0 };
        blob = await response.blob();
      } catch {
        return { usable: false, reason: "unreachable", qualityScore: 0 };
      }
    }
  } else {
    return { usable: false, reason: "not-image", qualityScore: 0 };
  }

  if (!blob.type.startsWith("image/")) return { usable: false, reason: "not-image", qualityScore: 0 };
  return imageBlobInspection(blob, url);
}

function safeDecodedUrl(url: string): string {
  try { return decodeURIComponent(url); } catch { return url; }
}

async function imageBlobInspection(blob: Blob, sourceUrl: string): Promise<CoverInspection> {
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    const naturalWidth = image.naturalWidth;
    const naturalHeight = image.naturalHeight;
    if (naturalWidth < 48 || naturalHeight < 64) return { usable: false, reason: "too-small", width: naturalWidth, height: naturalHeight, qualityScore: 0 };
    const aspect = naturalWidth / naturalHeight;
    if (aspect < 0.35 || aspect > 1.15) return { usable: false, reason: "bad-aspect", width: naturalWidth, height: naturalHeight, qualityScore: 0 };

    const width = 40;
    const height = 56;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return { usable: true, width: naturalWidth, height: naturalHeight, qualityScore: coverQualityScore(naturalWidth, naturalHeight, sourceUrl) };
    context.drawImage(image, 0, 0, width, height);
    const pixels = context.getImageData(0, 0, width, height).data;
    let nearWhite = 0;
    let nearGray = 0;
    let colorful = 0;
    let dark = 0;
    let luminanceSum = 0;
    let luminanceSquaredSum = 0;
    const total = width * height;
    for (let index = 0; index < pixels.length; index += 4) {
      const red = pixels[index];
      const green = pixels[index + 1];
      const blue = pixels[index + 2];
      const maximum = Math.max(red, green, blue);
      const minimum = Math.min(red, green, blue);
      const luminance = 0.2126 * red + 0.7152 * green + 0.0722 * blue;
      luminanceSum += luminance;
      luminanceSquaredSum += luminance * luminance;
      if (red >= 242 && green >= 242 && blue >= 242) nearWhite += 1;
      if (maximum - minimum <= 8) nearGray += 1;
      if (maximum - minimum >= 24) colorful += 1;
      if (luminance <= 155) dark += 1;
    }
    const whiteRatio = nearWhite / total;
    const grayRatio = nearGray / total;
    const colorfulRatio = colorful / total;
    const darkRatio = dark / total;
    const meanLuminance = luminanceSum / total;
    const variance = Math.max(0, luminanceSquaredSum / total - meanLuminance * meanLuminance);
    const luminanceDeviation = Math.sqrt(variance);

    if (whiteRatio >= 0.965) return { usable: false, reason: "blank", width: naturalWidth, height: naturalHeight, qualityScore: 0 };
    if (whiteRatio >= 0.78 && grayRatio >= 0.94 && colorfulRatio <= 0.012 && darkRatio <= 0.12 && meanLuminance >= 220) {
      return { usable: false, reason: "placeholder-image", width: naturalWidth, height: naturalHeight, qualityScore: 0 };
    }
    if (grayRatio >= 0.985 && colorfulRatio <= 0.004 && luminanceDeviation <= 22 && meanLuminance >= 205) {
      return { usable: false, reason: "placeholder-image", width: naturalWidth, height: naturalHeight, qualityScore: 0 };
    }

    return {
      usable: true,
      width: naturalWidth,
      height: naturalHeight,
      qualityScore: coverQualityScore(naturalWidth, naturalHeight, sourceUrl),
      perceptualHash: imagePerceptualHash(image)
    };
  } catch {
    return { usable: false, reason: "not-image", qualityScore: 0 };
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function coverQualityScore(width: number, height: number, url: string): number {
  const pixels = width * height;
  const resolution = Math.min(55, Math.max(0, Math.log2(Math.max(1, pixels) / (120 * 180)) * 10 + 25));
  const aspect = width / height;
  const aspectFit = Math.max(0, 32 - Math.abs(aspect - 0.67) * 65);
  const urlBonus = /(?:zoom=4|zoom=5|[-_]L\.(?:jpe?g|png|webp)|large|xlarge)/i.test(url) ? 5 : 0;
  return Math.round(resolution + aspectFit + urlBonus);
}

function imagePerceptualHash(image: HTMLImageElement): string | undefined {
  const size = 8;
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  if (!context) return undefined;
  context.drawImage(image, 0, 0, size, size);
  const data = context.getImageData(0, 0, size, size).data;
  const luminance: number[] = [];
  for (let index = 0; index < data.length; index += 4) luminance.push(0.2126 * data[index] + 0.7152 * data[index + 1] + 0.0722 * data[index + 2]);
  const average = luminance.reduce((sum, value) => sum + value, 0) / luminance.length;
  let hex = "";
  for (let index = 0; index < luminance.length; index += 4) {
    let nibble = 0;
    for (let offset = 0; offset < 4; offset += 1) if (luminance[index + offset] >= average) nibble |= 1 << (3 - offset);
    hex += nibble.toString(16);
  }
  return hex;
}

function perceptualHashDistance(left: string, right: string): number {
  if (left.length !== right.length) return Number.MAX_SAFE_INTEGER;
  let distance = 0;
  for (let index = 0; index < left.length; index += 1) {
    const xor = Number.parseInt(left[index], 16) ^ Number.parseInt(right[index], 16);
    distance += BIT_COUNT[xor] ?? 4;
  }
  return distance;
}

const BIT_COUNT = [0, 1, 1, 2, 1, 2, 2, 3, 1, 2, 2, 3, 2, 3, 3, 4];

function inspectionCacheKey(url: string): string {
  // Local cached covers are data URLs and can be hundreds of kilobytes long. Hashing every
  // character synchronously for every Library Health pass can monopolize the UI thread.
  // A cache key is not an integrity/security hash, so fingerprint large values from multiple
  // representative slices plus the exact length. Remote URLs remain hashed in full.
  const fingerprint = url.length <= 4096
    ? url
    : `${url.slice(0, 1024)}|${url.slice(Math.max(0, Math.floor(url.length / 2) - 512), Math.floor(url.length / 2) + 512)}|${url.slice(-1024)}`;
  let hash = 2166136261;
  for (let index = 0; index < fingerprint.length; index += 1) {
    hash ^= fingerprint.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return `bookstats.coverInspection.v${INSPECTION_CACHE_VERSION}.${(hash >>> 0).toString(16)}.${url.length}`;
}

function readStoredInspection(url: string): CoverInspection | undefined {
  if (typeof localStorage === "undefined") return undefined;
  try {
    const raw = localStorage.getItem(inspectionCacheKey(url));
    if (!raw) return undefined;
    const stored = JSON.parse(raw) as StoredInspection;
    if (!stored || Date.now() - stored.checkedAt > INSPECTION_CACHE_TTL_MS || typeof stored.usable !== "boolean") return undefined;
    const { checkedAt: _checkedAt, ...inspection } = stored;
    return inspection;
  } catch { return undefined; }
}

function writeStoredInspection(url: string, inspection: CoverInspection): void {
  if (typeof localStorage === "undefined") return;
  try { localStorage.setItem(inspectionCacheKey(url), JSON.stringify({ ...inspection, checkedAt: Date.now() } satisfies StoredInspection)); }
  catch { /* Storage pressure must never break cover lookup. */ }
}

/**
 * Stable fingerprint for the cover-bearing fields on a book. Unlike updatedAt, this changes
 * only when the selected cover itself changes, so unrelated metadata edits do not invalidate
 * Library Health's persisted cover verdict. Large data URLs use only short representative
 * slices and their exact length, avoiding full-image hashing when Health opens.
 */
export function coverRecordRevision(book: CoverDisplayRecord): string {
  const parts: string[] = [`v${INSPECTION_CACHE_VERSION}`];
  if (book.cachedCoverDataUrl?.trim()) parts.push(`cache:${coverRevisionPart(book.cachedCoverDataUrl)}`);
  if (book.coverAssetId?.trim()) parts.push(`asset:${book.coverAssetId.trim()}:${book.coverAssetToken?.trim() ?? ""}`);
  if (book.coverUrl?.trim()) parts.push(`url:${coverRevisionPart(book.coverUrl)}`);
  if (book.coverSourceUrl?.trim()) parts.push(`source:${coverRevisionPart(book.coverSourceUrl)}`);
  return parts.length > 1 ? parts.join("|") : "none";
}

function coverRevisionPart(value: string): string {
  const normalized = value.trim();
  if (normalized.length <= 512) return normalized;
  const middle = Math.floor(normalized.length / 2);
  return `${normalized.length}:${normalized.slice(0, 96)}:${normalized.slice(Math.max(0, middle - 48), middle + 48)}:${normalized.slice(-96)}`;
}

/**
 * Return cover sources in display priority order. A selected cover is user data, so the UI
 * should gracefully fall through to the original source if a local cache or archived copy
 * is temporarily unavailable rather than rendering a broken image.
 */
export function coverSources(book: CoverDisplayRecord): string[] {
  const sources: string[] = [];
  const add = (value?: string) => {
    const normalized = value?.trim();
    if (normalized && !sources.includes(normalized)) sources.push(normalized);
  };
  add(book.cachedCoverDataUrl);
  if (book.coverAssetId && book.coverAssetToken) {
    add(cloudCoverUrl(book.coverAssetId, book.coverAssetToken, book.updatedAt));
  }
  add(book.coverUrl);
  add(book.coverSourceUrl);
  return sources;
}

/** Backward-compatible helper for callers that only need the first preferred source. */
export function displayCover(book: CoverDisplayRecord): string | undefined {
  return coverSources(book)[0];
}
