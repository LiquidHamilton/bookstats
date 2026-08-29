import { cloudCoverUrl, fetchCatalogCoverForInspection } from "./api";

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

  // Third-party cover hosts commonly allow <img> display while blocking browser fetch()
  // with CORS. In that case the old validator had to accept the image without looking at
  // it, which is why catalog placeholders still leaked into the picker. Ask the BookStats
  // server to safely fetch the same public image so the browser can inspect its pixels.
  let blob = await fetchCatalogCoverForInspection(url);
  if (!blob) {
    // Keep a direct fallback for local/dev servers that have not yet been upgraded.
    try {
      const response = await fetch(url, { cache: "force-cache" });
      if (!response.ok) return false;
      const direct = await response.blob();
      if (!direct.type.startsWith("image/")) return false;
      blob = direct;
    } catch {
      // Cover choices are optional. If BookStats cannot validate one, omit it rather than
      // showing a broken/placeholder candidate to the user.
      return false;
    }
  }
  return imageBlobLooksUsable(blob);
}

function safeDecodedUrl(url: string): string {
  try { return decodeURIComponent(url); } catch { return url; }
}

async function imageBlobLooksUsable(blob: Blob): Promise<boolean> {
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImage(objectUrl);
    // Catalog thumbnails smaller than this are almost always missing-image shims, tracking
    // pixels, or provider placeholders rather than usable cover artwork.
    if (image.naturalWidth < 48 || image.naturalHeight < 64) return false;
    const aspect = image.naturalWidth / image.naturalHeight;
    if (aspect < 0.35 || aspect > 1.15) return false;

    const width = 40;
    const height = 56;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return true;
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

    // Completely blank/near-blank files.
    if (whiteRatio >= 0.965) return false;

    // Provider "Image Not Available" cards are generally bright, almost entirely grayscale
    // and contain only a small amount of dark text/iconography. This deliberately requires
    // all of those signals together so real light-colored covers are retained.
    if (whiteRatio >= 0.78 && grayRatio >= 0.94 && colorfulRatio <= 0.012 && darkRatio <= 0.12 && meanLuminance >= 220) return false;

    // Flat gray/white placeholder tiles that do not contain enough visual information to be
    // useful as cover art.
    if (grayRatio >= 0.985 && colorfulRatio <= 0.004 && luminanceDeviation <= 22 && meanLuminance >= 205) return false;

    return true;
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
