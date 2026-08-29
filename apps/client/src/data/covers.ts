import { cloudCoverUrl } from "./api";

const MAX_WIDTH = 700;
const MAX_HEIGHT = 1050;
const JPEG_QUALITY = 0.84;

export interface CoverDisplayRecord {
  coverUrl?: string;
  coverAssetId?: string;
  coverAssetToken?: string;
  coverSourceUrl?: string;
  cachedCoverDataUrl?: string;
  updatedAt?: string;
}

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

const PLACEHOLDER_URL_PATTERN = /(?:image[-_ ]?not[-_ ]?available|no[-_ ]?(?:image|cover)|missing[-_ ]?(?:image|cover)|default[-_ ]?(?:image|cover)|placeholder|cover[-_ ]?coming[-_ ]?soon)/i;
const COVER_VALIDATION_CONCURRENCY = 4;

/**
 * Remove catalog results that are clearly placeholders or effectively blank images.
 * Validation is deliberately conservative: if a third-party host blocks browser CORS,
 * BookStats keeps the URL unless its name itself identifies it as a placeholder.
 */
export async function filterUsableCoverUrls(values: string[]): Promise<string[]> {
  const urls = [...new Set(values.map((value) => value.trim()).filter(Boolean))];
  if (!urls.length) return [];
  const results = new Array<boolean>(urls.length);
  let nextIndex = 0;

  async function worker(): Promise<void> {
    while (nextIndex < urls.length) {
      const index = nextIndex++;
      results[index] = await coverUrlLooksUsable(urls[index]);
    }
  }

  await Promise.all(Array.from({ length: Math.min(COVER_VALIDATION_CONCURRENCY, urls.length) }, () => worker()));
  return urls.filter((_url, index) => results[index]);
}

async function coverUrlLooksUsable(url: string): Promise<boolean> {
  if (PLACEHOLDER_URL_PATTERN.test(safeDecodedUrl(url))) return false;
  if (!/^https?:\/\//i.test(url)) return true;
  try {
    const response = await fetch(url, { cache: "force-cache" });
    if (!response.ok) return false;
    const blob = await response.blob();
    if (!blob.type.startsWith("image/")) return false;
    return imageBlobLooksUsable(blob);
  } catch {
    // Some cover hosts allow <img> display while blocking fetch() with CORS. Do not
    // discard a potentially valid cover solely because it cannot be inspected.
    return true;
  }
}

function safeDecodedUrl(url: string): string {
  try { return decodeURIComponent(url); } catch { return url; }
}

async function imageBlobLooksUsable(blob: Blob): Promise<boolean> {
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    if (image.naturalWidth < 2 || image.naturalHeight < 2) return false;

    const width = 32;
    const height = 48;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return true;
    context.drawImage(image, 0, 0, width, height);
    const pixels = context.getImageData(0, 0, width, height).data;
    let nearWhite = 0;
    let nonWhite = 0;
    let dark = 0;
    const total = width * height;
    for (let index = 0; index < pixels.length; index += 4) {
      const red = pixels[index];
      const green = pixels[index + 1];
      const blue = pixels[index + 2];
      if (red >= 242 && green >= 242 && blue >= 242) nearWhite += 1;
      else nonWhite += 1;
      if (red <= 205 && green <= 205 && blue <= 205) dark += 1;
    }
    const whiteRatio = nearWhite / total;
    const nonWhiteRatio = nonWhite / total;
    const darkRatio = dark / total;
    // Blank files are almost entirely white. Common “Image Not Available” cards are
    // similarly white with only a small amount of dark text/icon artwork.
    return !(whiteRatio >= 0.93 && nonWhiteRatio <= 0.07 && darkRatio <= 0.05);
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
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
    // updatedAt intentionally participates in the URL. If an archived file is repaired in
    // place, a newly saved book gets a fresh browser request instead of retaining a failed
    // image request for the otherwise immutable asset URL until the whole page is reloaded.
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
