import type { Book, BookFormat, ReadingStatus } from "@bookstats/domain";
import { normalizeIsbn, normalizedReadingSessions } from "@bookstats/domain";

export interface ImportedBook {
  book: Book;
  shelfNames: string[];
}

export interface ExternalImportResult {
  source: "goodreads" | "librarything";
  items: ImportedBook[];
  warnings: string[];
}

export async function importGoodreadsCsv(file: File): Promise<ExternalImportResult> {
  const rows = parseCsv(await file.text());
  if (!rows.length) throw new Error("The Goodreads CSV did not contain any rows.");
  const required = ["Book Id", "Title", "Author", "Exclusive Shelf"];
  const headers = new Set(Object.keys(rows[0] ?? {}));
  if (required.some((header) => !headers.has(header))) throw new Error("This does not look like a Goodreads library export.");

  const now = new Date().toISOString();
  const items = rows.filter((row) => row.Title?.trim()).map((row): ImportedBook => {
    const parsedTitle = parseGoodreadsTitle(row.Title);
    const status = goodreadsStatus(row["Exclusive Shelf"]);
    const readDate = normalizeSlashDate(row["Date Read"]);
    const ratingValue = Number(row["My Rating"] || 0);
    const shelves = splitShelfList(row.Bookshelves).filter((name) => !isGoodreadsSystemShelf(name));
    const isbn13 = stripGoodreadsIsbn(row.ISBN13);
    const isbn10 = stripGoodreadsIsbn(row.ISBN);
    const book: Book = {
      id: crypto.randomUUID(),
      title: parsedTitle.title,
      author: normalizeWhitespace(row.Author) || "Unknown author",
      additionalAuthors: splitAdditionalAuthors(row["Additional Authors"]),
      isbn: normalizeIsbn(isbn13 || isbn10 || "") || undefined,
      series: parsedTitle.series,
      seriesVolume: parsedTitle.volume,
      publicationYear: numberOrUndefined(row["Year Published"] || row["Original Publication Year"]),
      publisher: valueOrUndefined(row.Publisher),
      pages: numberOrUndefined(row["Number of Pages"]),
      format: normalizeFormat(row.Binding),
      status,
      owned: Number(row["Owned Copies"] || 0) > 0,
      shelfIds: [],
      rating: ratingValue > 0 ? ratingValue : undefined,
      tags: [],
      review: valueOrUndefined(row["My Review"]),
      notes: valueOrUndefined(row["Private Notes"]),
      sourceIds: { goodreads: row["Book Id"] },
      metadataSource: "goodreads",
      dateAdded: normalizeAddedDate(row["Date Added"]) ?? now,
      readDates: readDate ? [readDate] : [],
      createdAt: now,
      updatedAt: now
    };
    return { book, shelfNames: shelves };
  });

  const warnings: string[] = [];
  const unreadableRereads = rows.filter((row) => Number(row["Read Count"] || 0) > 1).length;
  if (unreadableRereads) warnings.push(`${unreadableRereads} Goodreads record(s) report more than one read, but the CSV provides only one Date Read field; BookStats preserves the available completion date without inventing missing reread dates.`);
  return { source: "goodreads", items, warnings };
}

export async function importLibraryThingJson(file: File): Promise<ExternalImportResult> {
  const raw = JSON.parse(await file.text()) as Record<string, LibraryThingRecord>;
  if (!raw || Array.isArray(raw) || typeof raw !== "object") throw new Error("This does not look like a LibraryThing JSON export.");
  const records = Object.values(raw).filter((record) => record && typeof record === "object" && record.title);
  if (!records.length) throw new Error("The LibraryThing JSON did not contain any books.");

  const seriesFrequency = new Map<string, number>();
  for (const record of records) for (const series of record.series ?? []) {
    const key = series.trim().toLocaleLowerCase();
    if (key) seriesFrequency.set(key, (seriesFrequency.get(key) ?? 0) + 1);
  }

  const now = new Date().toISOString();
  const items = records.map((record): ImportedBook => {
    const collections = (record.collections ?? []).map((item) => item.trim()).filter(Boolean);
    const authorEntries = (record.authors ?? []).filter((entry) => !entry.role || entry.role.toLowerCase() === "author");
    const primaryAuthor = authorEntries[0]?.fl || flipLibraryThingName(record.primaryauthor) || "Unknown author";
    const readDate = normalizeDashDate(record.dateread);
    const selectedSeries = chooseLibraryThingSeries(record.series ?? [], seriesFrequency, record.title);
    const status = libraryThingStatus(collections, Boolean(readDate));
    const owned = collections.some((name) => name.toLowerCase() === "owned") && !collections.some((name) => name.toLowerCase() === "read but unowned");
    const shelfNames = collections.filter((name) => !isLibraryThingSystemCollection(name));
    const genres = record.genre?.map((item) => item.trim()).filter(Boolean) ?? [];
    const book: Book = {
      id: crypto.randomUUID(),
      title: record.title.trim(),
      author: primaryAuthor,
      additionalAuthors: authorEntries.slice(1).map((entry) => entry.fl || flipLibraryThingName(entry.lf)).filter((name): name is string => Boolean(name)),
      isbn: pickLibraryThingIsbn(record.isbn, record.originalisbn),
      series: selectedSeries,
      publicationYear: yearFromLibraryThingDate(record.date),
      publisher: parseLibraryThingPublisher(record.publication),
      language: record.language_codeA?.[0] || record.language?.[0],
      pages: numberOrUndefined(record.pages),
      format: normalizeFormat(record.format?.[0]?.text),
      status,
      owned,
      shelfIds: [],
      rating: typeof record.rating === "number" && record.rating > 0 ? record.rating : undefined,
      tags: (record.tags ?? []).map((item) => item.trim()).filter(Boolean),
      genre: genres[0],
      review: valueOrUndefined(record.review),
      metadataSource: "librarything",
      sourceIds: { librarything: record.books_id },
      dateAdded: record.entrydate ? `${record.entrydate}T00:00:00.000Z` : now,
      readDates: readDate ? [readDate] : [],
      createdAt: now,
      updatedAt: now
    };
    return { book, shelfNames };
  });

  return {
    source: "librarything",
    items,
    warnings: [
      "LibraryThing exports can list the same work under several translated series names. BookStats now prefers an English-looking series name when one is available, then uses collection-wide consistency to choose among the remaining candidates; review series/volume fields after import if needed.",
      "LibraryThing catalog entry IDs are treated as individual owned entries. Re-importing the same entry updates it, while separate LibraryThing entries with the same title or ISBN remain separate BookStats books."
    ]
  };
}

export function mergeImportedBook(existing: Book, incoming: Book): Book {
  const incomingHasLaterUpdate = incoming.updatedAt > existing.updatedAt;
  return {
    ...incoming,
    ...existing,
    id: existing.id,
    title: existing.title || incoming.title,
    author: existing.author || incoming.author,
    additionalAuthors: union(existing.additionalAuthors, incoming.additionalAuthors),
    isbn: existing.isbn || incoming.isbn,
    series: existing.series || incoming.series,
    seriesVolume: existing.seriesVolume || incoming.seriesVolume,
    publicationYear: existing.publicationYear ?? incoming.publicationYear,
    publisher: existing.publisher || incoming.publisher,
    language: existing.language || incoming.language,
    pages: existing.pages ?? incoming.pages,
    format: existing.format || incoming.format,
    condition: existing.condition ?? incoming.condition,
    status: existing.status === "not_started" && incoming.status !== "not_started" ? incoming.status : existing.status,
    owned: existing.owned || incoming.owned,
    shelfIds: union(existing.shelfIds ?? [], incoming.shelfIds ?? []),
    rating: existing.rating ?? incoming.rating,
    tags: union(existing.tags ?? [], incoming.tags ?? []),
    genre: existing.genre || incoming.genre,
    description: existing.description || incoming.description,
    review: existing.review || incoming.review,
    notes: existing.notes || incoming.notes,
    coverUrl: existing.coverUrl || incoming.coverUrl,
    coverAssetId: existing.coverAssetId ?? incoming.coverAssetId,
    coverAssetToken: existing.coverAssetToken ?? incoming.coverAssetToken,
    coverSourceUrl: existing.coverSourceUrl ?? incoming.coverSourceUrl,
    coverArchivePending: existing.coverAssetId ? undefined : (existing.coverArchivePending ?? incoming.coverArchivePending),
    cachedCoverDataUrl: existing.cachedCoverDataUrl,
    metadataSource: existing.metadataSource || incoming.metadataSource,
    metadataWorkId: existing.metadataWorkId || incoming.metadataWorkId,
    metadataEditionId: existing.metadataEditionId || incoming.metadataEditionId,
    metadataMatchType: existing.metadataMatchType || incoming.metadataMatchType,
    metadataSourceRefs: mergeMetadataSourceRefs(existing.metadataSourceRefs ?? [], incoming.metadataSourceRefs ?? []),
    metadataSources: { ...(incoming.metadataSources ?? {}), ...(existing.metadataSources ?? {}) },
    metadataConfidence: { ...(incoming.metadataConfidence ?? {}), ...(existing.metadataConfidence ?? {}) },
    seriesMetadata: existing.seriesMetadata ?? incoming.seriesMetadata,
    seriesCompletionOverride: existing.seriesCompletionOverride ?? incoming.seriesCompletionOverride,
    loans: mergeLoans(existing.loans, incoming.loans),
    duplicateIgnoreIds: union(existing.duplicateIgnoreIds ?? [], incoming.duplicateIgnoreIds ?? []),
    healthExceptions: union(existing.healthExceptions ?? [], incoming.healthExceptions ?? []),
    sourceIds: { ...(incoming.sourceIds ?? {}), ...(existing.sourceIds ?? {}) },
    dateAdded: existing.dateAdded < incoming.dateAdded ? existing.dateAdded : incoming.dateAdded,
    readingSessions: mergeReadingSessions(existing, incoming),
    readDates: union(existing.readDates ?? [], incoming.readDates ?? []).sort(),
    createdAt: existing.createdAt,
    updatedAt: incomingHasLaterUpdate ? incoming.updatedAt : new Date().toISOString()
  };
}

function mergeLoans(left: Book["loans"], right: Book["loans"]): Book["loans"] {
  const byId = new Map<string, NonNullable<Book["loans"]>[number]>();
  for (const loan of [...(right ?? []), ...(left ?? [])]) {
    const current = byId.get(loan.id);
    if (!current || loan.updatedAt >= current.updatedAt) byId.set(loan.id, { ...loan });
  }
  const values = [...byId.values()].sort((a, b) => a.loanedAt.localeCompare(b.loanedAt));
  return values.length ? values : undefined;
}

export function bookMatchKeys(book: Book): string[] {
  const keys = Object.entries(book.sourceIds ?? {}).map(([source, id]) => `source:${source}:${id}`);
  const isbn = book.isbn ? normalizeIsbn(book.isbn) : "";
  if (isbn) keys.push(`isbn:${isbn}`);
  keys.push(`title:${normalizeKey(book.title)}|${normalizeKey(book.author)}`);
  return keys;
}

export function findExistingBook(incoming: Book, existingBooks: Book[]): Book | undefined {
  const wanted = new Set(bookMatchKeys(incoming));
  return existingBooks.find((book) => bookMatchKeys(book).some((key) => wanted.has(key)));
}

function mergeReadingSessions(existing: Book, incoming: Book) {
  const sessions = [...normalizedReadingSessions(incoming), ...normalizedReadingSessions(existing)];
  const byKey = new Map<string, typeof sessions[number]>();
  for (const session of sessions) {
    const key = session.id.startsWith("legacy-") ? `${session.startedAt ?? ""}|${session.finishedAt ?? ""}|${session.progressPages ?? ""}` : session.id;
    const current = byKey.get(key);
    if (!current || session.updatedAt > current.updatedAt) byKey.set(key, session);
  }
  return [...byKey.values()].sort((a, b) => (a.finishedAt ?? a.startedAt ?? "9999").localeCompare(b.finishedAt ?? b.startedAt ?? "9999"));
}

type CsvRow = Record<string, string>;
function parseCsv(text: string): CsvRow[] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') { field += '"'; index += 1; }
      else if (char === '"') quoted = false;
      else field += char;
      continue;
    }
    if (char === '"') quoted = true;
    else if (char === ",") { row.push(field); field = ""; }
    else if (char === "\n") { row.push(field.replace(/\r$/, "")); rows.push(row); row = []; field = ""; }
    else field += char;
  }
  if (field.length || row.length) { row.push(field.replace(/\r$/, "")); rows.push(row); }
  const header = rows.shift()?.map((item) => item.replace(/^\uFEFF/, "").trim()) ?? [];
  return rows.filter((values) => values.some(Boolean)).map((values) => Object.fromEntries(header.map((name, index) => [name, values[index] ?? ""])));
}

function trailingParenthetical(value: string): { base: string; content: string } | undefined {
  const text = value.trim();
  if (!text.endsWith(")")) return undefined;
  let depth = 0;
  for (let index = text.length - 1; index >= 0; index -= 1) {
    if (text[index] === ")") depth += 1;
    else if (text[index] === "(") {
      depth -= 1;
      if (depth === 0) {
        const base = text.slice(0, index).trim();
        const content = text.slice(index + 1, -1).trim();
        return base && content ? { base, content } : undefined;
      }
    }
  }
  return undefined;
}

function parseGoodreadsTitle(value: string): { title: string; series?: string; volume?: string } {
  const title = value.trim();
  const trailing = trailingParenthetical(title);
  if (!trailing) return { title };
  const match = trailing.content.match(/^(.+),\s*#([0-9]+(?:\.[0-9]+)?(?:\s*(?:&|and|,|[-–—])\s*[0-9]+(?:\.[0-9]+)?)*)$/i);
  if (!match) return { title };
  return { title: trailing.base, series: match[1].trim(), volume: match[2] };
}

function goodreadsStatus(value: string): ReadingStatus {
  switch (value.trim().toLowerCase()) {
    case "read": return "read";
    case "currently-reading": return "currently_reading";
    case "to-read": return "want_to_read";
    default: return "not_started";
  }
}
function libraryThingStatus(collections: string[], hasReadDate: boolean): ReadingStatus {
  const normalized = new Set(collections.map((name) => name.toLowerCase()));
  if (normalized.has("dnf") || normalized.has("did not finish (dnf)")) return "did_not_finish";
  if (normalized.has("to read")) return "want_to_read";
  if (normalized.has("read") || normalized.has("read but unowned") || hasReadDate) return "read";
  return "not_started";
}
function isGoodreadsSystemShelf(name: string): boolean { return ["read", "currently-reading", "to-read"].includes(name.trim().toLowerCase()); }
function isLibraryThingSystemCollection(name: string): boolean { return ["owned", "unread", "read", "to read", "dnf", "did not finish (dnf)"].includes(name.trim().toLowerCase()); }
function splitShelfList(value: string): string[] { return value.split(",").map((item) => item.trim()).filter(Boolean); }
function splitAdditionalAuthors(value: string): string[] { return value.split(",").map(normalizeWhitespace).filter(Boolean); }
function stripGoodreadsIsbn(value: string): string { return value.replace(/^="?/, "").replace(/"?$/, "").trim(); }
function normalizeSlashDate(value: string): string | undefined { const match = value.trim().match(/^(\d{4})\/(\d{2})\/(\d{2})$/); return match ? `${match[1]}-${match[2]}-${match[3]}` : undefined; }
function normalizeDashDate(value?: string): string | undefined { const match = value?.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/); return match ? match[0] : undefined; }
function normalizeAddedDate(value: string): string | undefined { const date = normalizeSlashDate(value); return date ? `${date}T00:00:00.000Z` : undefined; }
function numberOrUndefined(value?: string): number | undefined { const number = Number(String(value ?? "").trim()); return Number.isFinite(number) && number > 0 ? number : undefined; }
function valueOrUndefined(value?: string): string | undefined { const trimmed = value?.trim(); return trimmed || undefined; }
function normalizeWhitespace(value?: string): string { return (value ?? "").replace(/\s+/g, " ").trim(); }
function normalizeKey(value: string): string { return normalizeWhitespace(value).toLocaleLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim(); }
function union<T>(left: T[], right: T[]): T[] { return [...new Set([...left, ...right])]; }

function mergeMetadataSourceRefs(left: NonNullable<Book["metadataSourceRefs"]>, right: NonNullable<Book["metadataSourceRefs"]>): NonNullable<Book["metadataSourceRefs"]> {
  return [...new Map([...left, ...right].map((ref) => [`${ref.provider}:${ref.workId}:${ref.editionId ?? ""}`, ref] as const)).values()];
}
function normalizeFormat(value?: string): BookFormat | undefined {
  const text = value?.trim().toLowerCase();
  if (!text) return undefined;
  if (text.includes("mass market")) return "Mass Market Paperback";
  if (text.includes("hardcover") || text.includes("hardback")) return "Hardcover";
  if (text.includes("paperback") || text.includes("trade paperback")) return "Paperback";
  if (text.includes("ebook") || text.includes("kindle") || text.includes("electronic")) return "eBook";
  if (text.includes("audio")) return "Audiobook";
  if (text.includes("graphic") || text.includes("manga") || text.includes("comic")) return "Graphic Novel";
  if (text.includes("omnibus")) return "Omnibus";
  return "Other";
}

interface LibraryThingAuthor { lf?: string; fl?: string; role?: string; }
interface LibraryThingRecord {
  books_id: string;
  title: string;
  primaryauthor?: string;
  authors?: LibraryThingAuthor[];
  tags?: string[];
  collections?: string[];
  isbn?: string | string[] | Record<string, string>;
  originalisbn?: string;
  date?: string;
  publication?: string;
  language?: string[];
  language_codeA?: string[];
  series?: string[];
  genre?: string[];
  workcode?: string;
  rating?: number;
  entrydate?: string;
  format?: Array<{ text?: string }>;
  pages?: string;
  dateread?: string;
  review?: string;
}

function flipLibraryThingName(value?: string): string | undefined {
  const text = value?.trim();
  if (!text) return undefined;
  const match = text.match(/^([^,]+),\s*(.+)$/);
  return match ? `${match[2]} ${match[1]}` : text;
}
function pickLibraryThingIsbn(value: LibraryThingRecord["isbn"], original?: string): string | undefined {
  const candidates: string[] = [];
  if (typeof value === "string") candidates.push(value);
  else if (Array.isArray(value)) candidates.push(...value);
  else if (value && typeof value === "object") candidates.push(...Object.values(value));
  if (original) candidates.push(original);
  const normalized = candidates.map(normalizeIsbn).filter(Boolean);
  return normalized.find((isbn) => isbn.length === 13) ?? normalized[0];
}

function parseLibraryThingPublisher(value?: string): string | undefined {
  const text = value?.trim();
  if (!text) return undefined;
  // Typical exports look like "Penguin Classics (2002), Edition: Revised ed., 448 pages"
  // or "Tor, Hardcover, 168 pages". Keep only the leading publisher portion.
  const beforeParen = text.split("(", 1)[0]?.trim();
  const beforeComma = (beforeParen || text).split(",", 1)[0]?.trim();
  return beforeComma || undefined;
}

function yearFromLibraryThingDate(value?: string): number | undefined {
  const match = value?.match(/^(\d{4})/);
  return match ? Number(match[1]) : undefined;
}
const LIBRARYTHING_FOREIGN_SERIES_WORDS = new Set([
  // German
  "der", "die", "das", "den", "dem", "des", "und", "ein", "eine", "einer", "eines", "im", "auf", "mit", "ohne", "zur", "zum", "von", "romane", "chroniken",
  // Spanish / Catalan
  "de", "del", "la", "las", "el", "los", "y", "en", "para", "con", "sin", "una", "uno", "trilogia", "trilogía", "fundacion", "fundación", "cronologico", "cronológico", "publicacion", "publicación", "leyendas", "preludios",
  // French
  "le", "les", "du", "et", "un", "une", "dans", "avec", "sans", "trilogie", "fondation", "chroniques", "legendes", "légendes", "origines", "genese", "genèse",
  // Italian
  "il", "lo", "gli", "di", "della", "delle", "degli", "nel", "senza", "ciclo", "fondazioni", "ruota", "tempo",
  // Dutch
  "het", "van", "voor", "zonder", "oversteek", "duin",
  // Portuguese
  "da", "do", "dos", "em", "com", "fundacao", "fundação",
  // Polish / Czech / Slovak / South Slavic and other common LibraryThing translations
  "serijal", "zakladi", "fundacja", "chronologiczny", "publikacja", "nadace", "zakladna", "základňa", "romany", "romány", "robotech",
  // Additional translated-series words seen commonly in LibraryThing's multilingual series lists
  "kulttuuri", "kultur", "zyklus", "torre", "nera", "mundo", "rio", "río", "chair", "poule", "cabane", "magique",
  "operatiivinen", "keskus", "ringenes", "herre", "merkilliset", "matkat", "edice", "disgfyd", "verschollene", "flotte",
  "odyssee", "reihe", "epreuve", "épreuve", "arkiverne", "pisteur", "kirjat", "publiceringsordning", "opowiesci", "opowieści",
  "narnii", "letopisi", "narnije", "cronache", "kronieken", "fuerzas", "defensa", "coloniales", "materia", "oscura",
  "quinta", "ola", "nuit", "cronicas", "crônicas", "gelo", "fogo", "sormusten", "herrasta", "seigneur", "anneaux", "senhor", "aneis", "anéis"
]);

const LIBRARYTHING_ENGLISH_SERIES_WORDS = new Set([
  "the", "of", "and", "to", "for", "from", "with", "without", "chronicles", "chronicle", "diaries", "diary",
  "universe", "collection", "collections", "story", "stories", "complete", "adventures", "adventure", "tales", "tale",
  "world", "worlds", "lord", "lords", "ring", "rings", "house", "houses", "magic", "tree", "wheel", "time", "dark",
  "tower", "realm", "lost", "fleet", "odyssey", "pathfinder", "center", "foundation", "files", "war", "wars", "children",
  "king", "kings", "queen", "queens", "dragon", "dragons", "empire", "earth", "moon", "sun", "stars", "star"
]);

function chooseLibraryThingSeries(series: string[], frequency: Map<string, number>, bookTitle: string): string | undefined {
  const candidates = [...new Set(series.map((item) => item.trim()).filter(Boolean))];
  if (!candidates.length) return undefined;
  return candidates.sort((a, b) => seriesScore(b, frequency, bookTitle) - seriesScore(a, frequency, bookTitle) || a.length - b.length)[0];
}
function seriesScore(value: string, frequency: Map<string, number>, bookTitle: string): number {
  const count = frequency.get(value.toLocaleLowerCase()) ?? 0;
  const tokens = normalizeSeriesTokens(value);
  const titleTokens = new Set(normalizeSeriesTokens(bookTitle));
  const ascii = /^[\x00-\x7F]+$/.test(value) ? 5 : -120;
  const foreign = tokens.some((token) => LIBRARYTHING_FOREIGN_SERIES_WORDS.has(token)) ? -500 : 0;
  const english = tokens.filter((token) => LIBRARYTHING_ENGLISH_SERIES_WORDS.has(token)).length * 20
    + tokens.filter((token) => ["world", "land", "house", "verse", "fleet"].some((suffix) => token.length > suffix.length && token.endsWith(suffix))).length * 20;
  const titleOverlap = tokens.filter((token) => token.length >= 4 && titleTokens.has(token)).length * 8;
  const exactTitle = normalizeSeriesKey(value) === normalizeSeriesKey(bookTitle) ? 150 : 0;
  const clean = /[{}\[\]]/.test(value) ? -15 : 0;
  const concise = value.length <= 40 ? 2 : -2;
  return count * 10 + ascii + foreign + english + titleOverlap + exactTitle + clean + concise;
}
function normalizeSeriesTokens(value: string): string[] {
  return normalizeSeriesKey(value).split(" ").filter(Boolean);
}
function normalizeSeriesKey(value: string): string {
  return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
