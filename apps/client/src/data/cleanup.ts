import type { Book } from "@bookstats/domain";
import { normalizeIsbn, normalizedReadDates, normalizedReadingSessions } from "@bookstats/domain";
import type { CoverInspection } from "./covers";
import { coverUrlLooksSuspiciousByName } from "./covers";

export interface DuplicateGroup {
  id: string;
  books: Book[];
  reason: string;
  signals: string[];
  confidence: "high" | "medium";
  /** True when the candidate records carry different valid ISBNs and may be intentional editions. */
  editionConflict: boolean;
}

export function findDuplicateGroups(books: Book[]): DuplicateGroup[] {
  const parent = new Map<string, string>();
  const reasons = new Map<string, Set<string>>();
  const byKey = new Map<string, string>();
  const bookById = new Map(books.map((book) => [book.id, book] as const));
  for (const book of books) parent.set(book.id, book.id);

  const find = (id: string): string => {
    const p = parent.get(id) ?? id;
    if (p === id) return id;
    const root = find(p); parent.set(id, root); return root;
  };
  const ignoredPair = (a: string, b: string) => {
    const left = bookById.get(a); const right = bookById.get(b);
    return Boolean(left?.duplicateIgnoreIds?.includes(b) || right?.duplicateIgnoreIds?.includes(a));
  };
  const union = (a: string, b: string, reason: string) => {
    if (ignoredPair(a, b)) return;
    const ra = find(a); const rb = find(b);
    if (ra !== rb) parent.set(rb, ra);
    const root = find(a);
    const set = reasons.get(root) ?? new Set<string>(); set.add(reason); reasons.set(root, set);
    if (rb !== root && reasons.has(rb)) { for (const item of reasons.get(rb)!) set.add(item); reasons.delete(rb); }
  };

  for (const book of books) {
    const keys: Array<[string, string]> = [];
    for (const [source, id] of Object.entries(book.sourceIds ?? {})) if (id) keys.push([`source:${source}:${id}`, `same ${source} record`]);
    const isbn = canonicalDuplicateIsbn(book.isbn);
    if (isbn) keys.push([`isbn:${isbn}`, "same ISBN edition"]);
    const title = normalizeText(book.title); const author = normalizeText(book.author);
    if (title && author) keys.push([`title-author:${title}|${author}`, "same title and author"]);
    for (const [key, reason] of keys) {
      const previous = byKey.get(key);
      if (previous) union(previous, book.id, reason); else byKey.set(key, book.id);
    }
  }

  const groups = new Map<string, Book[]>();
  for (const book of books) {
    const root = find(book.id);
    const items = groups.get(root) ?? []; items.push(book); groups.set(root, items);
  }
  return [...groups.entries()].filter(([, items]) => items.length > 1).map(([root, items]) => {
    const ordered = items.sort((a, b) => completenessScore(b) - completenessScore(a) || a.title.localeCompare(b.title));
    const signals = [...(reasons.get(root) ?? new Set(["matching library data"]))];
    const isbns = [...new Set(ordered.map((book) => canonicalDuplicateIsbn(book.isbn)).filter((isbn): isbn is string => Boolean(isbn)))];
    const strong = signals.some((reason) => reason.startsWith("same ") && reason !== "same title and author");
    return {
      id: root,
      books: ordered,
      reason: signals.join(" and "),
      signals,
      confidence: strong ? "high" as const : "medium" as const,
      editionConflict: isbns.length > 1
    };
  }).sort((a, b) => (a.confidence === b.confidence ? 0 : a.confidence === "high" ? -1 : 1) || b.books.length - a.books.length || a.books[0].title.localeCompare(b.books[0].title));
}

export type MetadataIssue =
  | "Cover"
  | "Cover quality"
  | "Duplicate cover"
  | "Description"
  | "ISBN"
  | "Invalid ISBN"
  | "Pages"
  | "Invalid pages"
  | "Publication year"
  | "Invalid publication year"
  | "Series position"
  | "Series consistency"
  | "Author"
  | "Reading history"
  | "Source ID conflict";

export const METADATA_ISSUE_ORDER: MetadataIssue[] = [
  "Cover", "Cover quality", "Duplicate cover", "Description", "ISBN", "Invalid ISBN", "Pages", "Invalid pages",
  "Publication year", "Invalid publication year", "Series position", "Series consistency", "Author", "Reading history", "Source ID conflict"
];

const HEALTH_EXEMPTABLE_ISSUES = new Set<MetadataIssue>([
  "Cover", "Duplicate cover", "Description", "ISBN", "Pages", "Publication year", "Series position", "Source ID conflict"
]);

/** Cleanup findings a user may legitimately classify as intentional/not applicable. */
export function isMetadataIssueExemptable(issue: MetadataIssue): boolean {
  return HEALTH_EXEMPTABLE_ISSUES.has(issue);
}

function hasHealthException(book: Book, issue: MetadataIssue): boolean {
  return isMetadataIssueExemptable(issue) && Boolean(book.healthExceptions?.includes(issue));
}

/** Record-local cleanup findings. Library-wide/contextual checks are added below. */
export function metadataIssues(book: Book, expectSeriesPosition = false): MetadataIssue[] {
  const issues: MetadataIssue[] = [];
  const hasCover = Boolean(book.coverAssetId || book.coverUrl || book.cachedCoverDataUrl || book.coverSourceUrl);
  if (!hasCover) issues.push("Cover");
  else if (coverUrlLooksSuspiciousByName(book.coverUrl) || coverUrlLooksSuspiciousByName(book.coverSourceUrl) || (book.coverAssetId && !book.coverAssetToken)) issues.push("Cover quality");
  if (!book.description?.trim()) issues.push("Description");

  const isbn = book.isbn?.trim();
  if (!isbn) issues.push("ISBN");
  else if (!validIsbn(isbn)) issues.push("Invalid ISBN");

  if (book.pages === undefined || book.pages === null) issues.push("Pages");
  else if (!Number.isFinite(book.pages) || book.pages <= 0 || book.pages > 100_000) issues.push("Invalid pages");

  if (book.publicationYear === undefined || book.publicationYear === null) issues.push("Publication year");
  else if (!Number.isInteger(book.publicationYear) || book.publicationYear < 1000 || book.publicationYear > new Date().getFullYear() + 1) issues.push("Invalid publication year");

  if (expectSeriesPosition && book.series?.trim() && !book.seriesVolume?.trim()) issues.push("Series position");
  if (!book.series?.trim() && book.seriesVolume?.trim()) issues.push("Series consistency");
  if (!book.author?.trim()) issues.push("Author");
  if (readingHistoryInconsistent(book)) issues.push("Reading history");
  return issues;
}

/**
 * Decide whether this particular series membership actually expects an ordered position.
 * Generic franchise/umbrella memberships are intentionally left unnumbered unless BookStats
 * has concrete catalog or library evidence that this is an ordered sequence.
 */
export function seriesPositionExpected(book: Book, books: Book[]): boolean {
  const series = normalizeText(book.series ?? "");
  if (!series || book.seriesVolume?.trim()) return false;

  const catalog = book.seriesMetadata;
  if (catalog && normalizeText(catalog.name) === series) {
    const catalogPositions = catalog.books
      .map((entry) => numericSeriesPosition(entry.position))
      .filter((value): value is number => value !== undefined);
    const uniqueCatalogPositions = new Set(catalogPositions);

    // Strongest evidence: the matching catalog row for this exact/likely edition has a position.
    const isbn = canonicalDuplicateIsbn(book.isbn);
    const title = normalizeText(book.title);
    const author = normalizeText(book.author);
    const direct = catalog.books.find((entry) => {
      const entryIsbn = canonicalDuplicateIsbn(entry.isbn);
      if (isbn && entryIsbn && isbn === entryIsbn) return true;
      return Boolean(title && normalizeText(entry.title) === title && (!entry.author || !author || normalizeText(entry.author) === author));
    });
    if (numericSeriesPosition(direct?.position) !== undefined) return true;

    // Two or more distinct catalog positions are strong evidence of an ordered series.
    if (uniqueCatalogPositions.size >= 2) return true;
  }

  // Library evidence is deliberately conservative. A couple of numbered entries in a huge
  // umbrella collection should not force every other membership to invent a volume number.
  const sameSeries = books.filter((candidate) => normalizeText(candidate.series ?? "") === series);
  const positioned = sameSeries.filter((candidate) => numericSeriesPosition(candidate.seriesVolume) !== undefined).length;
  return sameSeries.length >= 3 && positioned >= 2 && positioned / sameSeries.length >= 0.5;
}

/** Build library-wide issues that require comparing records or asynchronously inspected covers. */
export function libraryMetadataIssueMap(books: Book[], coverInspections: ReadonlyMap<string, CoverInspection> = new Map()): Map<string, MetadataIssue[]> {
  const bookById = new Map(books.map((book) => [book.id, book] as const));
  const result = new Map(books.map((book) => [book.id, metadataIssues(book, seriesPositionExpected(book, books))] as const));
  const add = (id: string, issue: MetadataIssue) => {
    const current = result.get(id) ?? [];
    if (!current.includes(issue)) current.push(issue);
    result.set(id, current);
  };

  for (const book of books) {
    const inspection = coverInspections.get(book.id);
    if (inspection && !inspection.usable && (book.coverAssetId || book.coverUrl || book.cachedCoverDataUrl || book.coverSourceUrl)) add(book.id, "Cover quality");
  }

  // Imported source IDs normally identify one source record. A user-confirmed "keep separate"
  // decision is authoritative here too: only unresolved pairs sharing an ID remain conflicts.
  const sourceOwners = new Map<string, string[]>();
  for (const book of books) for (const [source, id] of Object.entries(book.sourceIds ?? {})) {
    const normalized = id.trim();
    if (!normalized) continue;
    const key = `${source.toLocaleLowerCase()}:${normalized}`;
    sourceOwners.set(key, [...(sourceOwners.get(key) ?? []), book.id]);
  }
  const ignoredPair = (a: string, b: string) => {
    const left = bookById.get(a); const right = bookById.get(b);
    return Boolean(left?.duplicateIgnoreIds?.includes(b) || right?.duplicateIgnoreIds?.includes(a));
  };
  for (const ids of sourceOwners.values()) {
    if (ids.length < 2) continue;
    for (const id of ids) if (ids.some((other) => other !== id && !ignoredPair(id, other))) add(id, "Source ID conflict");
  }

  // Exact reuse of a selected cloud/source cover across unrelated titles is usually a bad
  // import/provider placeholder rather than intentional artwork reuse. Copies/editions with
  // the same normalized title+author are not flagged.
  const coverOwners = new Map<string, Book[]>();
  for (const book of books) {
    const key = selectedCoverIdentity(book);
    if (!key) continue;
    coverOwners.set(key, [...(coverOwners.get(key) ?? []), book]);
  }
  for (const group of coverOwners.values()) {
    if (group.length < 2) continue;
    const works = new Set(group.map((book) => `${normalizeText(book.title)}|${normalizeText(book.author)}`));
    if (works.size < 2) continue;
    for (const book of group) add(book.id, "Duplicate cover");
  }

  // Saved exceptions make eligible checks not applicable. Objective integrity failures remain.
  for (const book of books) {
    const current = result.get(book.id) ?? [];
    result.set(book.id, current.filter((issue) => !hasHealthException(book, issue)));
  }
  return result;
}

export interface LibraryHealthSummary {
  score: number;
  totalChecks: number;
  passedChecks: number;
  booksToReview: number;
  duplicateGroups: number;
  issueCounts: Array<{ issue: MetadataIssue; count: number }>;
}

/**
 * Health scores objective integrity only. Missing optional metadata can still appear as a
 * cleanup opportunity, but it does not count as a failed check or prevent a 100% score.
 */
export function libraryHealth(books: Book[], duplicateGroupsOverride?: number, issueMap?: ReadonlyMap<string, MetadataIssue[]>): LibraryHealthSummary {
  const effectiveIssueMap = issueMap ?? libraryMetadataIssueMap(books);
  const counts = new Map<MetadataIssue, number>(METADATA_ISSUE_ORDER.map((issue) => [issue, 0]));
  let totalChecks = 0; let failedChecks = 0; let booksToReview = 0;

  for (const book of books) {
    const issues = effectiveIssueMap.get(book.id) ?? [];
    const exceptions = new Set((book.healthExceptions ?? []).filter((issue): issue is MetadataIssue => METADATA_ISSUE_ORDER.includes(issue as MetadataIssue)));
    const hasCover = Boolean(book.coverAssetId || book.coverUrl || book.cachedCoverDataUrl || book.coverSourceUrl);
    const hasSourceId = Object.values(book.sourceIds ?? {}).some((id) => Boolean(id?.trim()));
    const hasSeriesData = Boolean(book.series?.trim() || book.seriesVolume?.trim());
    const expectSeriesPosition = seriesPositionExpected(book, books);

    // Only checks with something objective to validate belong in the denominator.
    // User-confirmed exceptions make eligible checks explicitly not applicable.
    const applicableChecks = new Set<string>();
    if (hasCover) applicableChecks.add("cover-quality");
    if (selectedCoverIdentity(book) && !exceptions.has("Duplicate cover")) applicableChecks.add("duplicate-cover");
    if (book.isbn?.trim()) applicableChecks.add("isbn-validity");
    if (book.pages !== undefined && book.pages !== null) applicableChecks.add("pages-validity");
    if (book.publicationYear !== undefined && book.publicationYear !== null) applicableChecks.add("publication-year-validity");
    if (hasSeriesData) applicableChecks.add("series-consistency");
    if (expectSeriesPosition && !exceptions.has("Series position")) applicableChecks.add("series-position");
    applicableChecks.add("author");
    applicableChecks.add("reading-history");
    if (hasSourceId && !exceptions.has("Source ID conflict")) applicableChecks.add("source-id-conflict");

    const failures = new Set<string>();
    for (const issue of issues) {
      if (issue === "Cover quality") failures.add("cover-quality");
      else if (issue === "Duplicate cover") failures.add("duplicate-cover");
      else if (issue === "Invalid ISBN") failures.add("isbn-validity");
      else if (issue === "Invalid pages") failures.add("pages-validity");
      else if (issue === "Invalid publication year") failures.add("publication-year-validity");
      else if (issue === "Series consistency") failures.add("series-consistency");
      else if (issue === "Series position") failures.add("series-position");
      else if (issue === "Author") failures.add("author");
      else if (issue === "Reading history") failures.add("reading-history");
      else if (issue === "Source ID conflict") failures.add("source-id-conflict");
    }

    totalChecks += applicableChecks.size;
    failedChecks += [...failures].filter((failure) => applicableChecks.has(failure)).length;
    if (issues.length) booksToReview += 1;
    for (const issue of issues) counts.set(issue, (counts.get(issue) ?? 0) + 1);
  }

  const passedChecks = Math.max(0, totalChecks - failedChecks);
  return {
    score: totalChecks ? Math.round((passedChecks / totalChecks) * 100) : 100,
    totalChecks,
    passedChecks,
    booksToReview,
    duplicateGroups: duplicateGroupsOverride ?? findDuplicateGroups(books).length,
    issueCounts: METADATA_ISSUE_ORDER.map((issue) => ({ issue, count: counts.get(issue) ?? 0 })).filter((item) => item.count > 0)
  };
}

function numericSeriesPosition(value?: string): number | undefined {
  if (!value?.trim()) return undefined;
  const clean = value.trim().replace(/^#\s*/, "");
  if (!/^\d+(?:\.\d+)?$/.test(clean)) return undefined;
  const number = Number(clean);
  return Number.isFinite(number) && number > 0 && number <= 500 ? number : undefined;
}

function validIsbn(value: string): boolean {
  const isbn = normalizeIsbn(value);
  if (isbn.length === 10 && /^\d{9}[\dX]$/.test(isbn)) {
    const total = isbn.split("").reduce((sum, char, index) => sum + (char === "X" ? 10 : Number(char)) * (10 - index), 0);
    return total % 11 === 0;
  }
  if (isbn.length === 13 && /^\d{13}$/.test(isbn)) {
    const total = isbn.slice(0, 12).split("").reduce((sum, char, index) => sum + Number(char) * (index % 2 === 0 ? 1 : 3), 0);
    return (10 - (total % 10)) % 10 === Number(isbn[12]);
  }
  return false;
}

function readingHistoryInconsistent(book: Book): boolean {
  const sessions = normalizedReadingSessions(book);
  const finished = sessions.filter((session) => Boolean(session.finishedAt));
  if (book.status === "read" && finished.length === 0 && normalizedReadDates(book).length === 0) return true;
  const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
  for (const session of sessions) {
    if (session.startedAt && !validCalendarDate(session.startedAt)) return true;
    if (session.finishedAt && !validCalendarDate(session.finishedAt)) return true;
    if (session.startedAt && session.finishedAt && session.startedAt > session.finishedAt) return true;
    if ((session.startedAt && session.startedAt > tomorrow) || (session.finishedAt && session.finishedAt > tomorrow)) return true;
    if (session.progressPages !== undefined && (!Number.isFinite(session.progressPages) || session.progressPages < 0)) return true;
  }
  return false;
}

function validCalendarDate(value: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00Z`);
  return Number.isFinite(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function selectedCoverIdentity(book: Book): string | undefined {
  if (book.coverAssetId) return `asset:${book.coverAssetId}`;
  const url = book.coverSourceUrl?.trim() || book.coverUrl?.trim();
  if (!url || url.startsWith("data:")) return undefined;
  try {
    const parsed = new URL(url);
    parsed.hash = "";
    // Cache-busting params should not make the same selected artwork appear unique.
    for (const key of ["v", "cache", "cachebust", "cb"]) parsed.searchParams.delete(key);
    return `url:${parsed.toString()}`;
  } catch { return `url:${url}`; }
}

export function mergeBooks(preferred: Book, others: Book[]): Book {
  const all = [preferred, ...others];
  const firstDefined = <K extends keyof Book>(key: K): Book[K] => all.map((book) => book[key]).find((value) => value !== undefined && value !== "" && value !== null) as Book[K];
  const unionStrings = (values: string[][]) => [...new Set(values.flat().map((value) => value.trim()).filter(Boolean))];
  const sourceIds = Object.assign({}, ...all.map((book) => book.sourceIds ?? {}));
  const metadataSources = Object.assign({}, ...[...all].reverse().map((book) => book.metadataSources ?? {}));
  const metadataConfidence = Object.assign({}, ...[...all].reverse().map((book) => book.metadataConfidence ?? {}));
  const metadataSourceRefs = [...new Map(all.flatMap((book) => book.metadataSourceRefs ?? []).map((ref) => [`${ref.provider}:${ref.workId}:${ref.editionId ?? ""}`, ref] as const)).values()];
  const readDates = [...new Set(all.flatMap((book) => normalizedReadDates(book)))].sort();
  const sessionMap = new Map<string, ReturnType<typeof normalizedReadingSessions>[number]>();
  for (const book of all) for (const session of normalizedReadingSessions(book)) {
    const key = session.id.startsWith("legacy-") ? `${session.startedAt ?? ""}|${session.finishedAt ?? ""}|${session.progressPages ?? ""}` : session.id;
    const existing = sessionMap.get(key);
    if (!existing || session.updatedAt > existing.updatedAt) sessionMap.set(key, session);
  }
  const readingSessions = [...sessionMap.values()].sort((a, b) => (a.finishedAt ?? a.startedAt ?? "9999").localeCompare(b.finishedAt ?? b.startedAt ?? "9999"));
  const createdAt = all.map((book) => book.createdAt).filter(Boolean).sort()[0] ?? preferred.createdAt;
  // Keep a cover URL and its local cache paired so a merge cannot display a cache
  // that belongs to a different edition's remote cover.
  const coverSource = all.find((book) => Boolean(book.coverAssetId || book.coverUrl || book.cachedCoverDataUrl));
  return {
    ...preferred,
    additionalAuthors: unionStrings(all.map((book) => book.additionalAuthors ?? [])),
    isbn: firstDefined("isbn"),
    series: firstDefined("series"),
    seriesVolume: firstDefined("seriesVolume"),
    publicationYear: firstDefined("publicationYear"),
    publisher: firstDefined("publisher"),
    language: firstDefined("language"),
    pages: firstDefined("pages"),
    format: firstDefined("format"),
    condition: firstDefined("condition"),
    owned: all.some((book) => book.owned),
    shelfIds: unionStrings(all.map((book) => book.shelfIds ?? [])),
    rating: firstDefined("rating"),
    tags: unionStrings(all.map((book) => book.tags ?? [])),
    genre: firstDefined("genre"),
    description: firstDefined("description"),
    review: firstDefined("review"),
    notes: firstDefined("notes"),
    coverUrl: coverSource?.coverUrl,
    coverAssetId: coverSource?.coverAssetId,
    coverAssetToken: coverSource?.coverAssetToken,
    coverSourceUrl: coverSource?.coverSourceUrl,
    coverArchivePending: coverSource?.coverArchivePending,
    cachedCoverDataUrl: coverSource?.cachedCoverDataUrl,
    metadataSource: firstDefined("metadataSource"),
    metadataWorkId: firstDefined("metadataWorkId"),
    metadataEditionId: firstDefined("metadataEditionId"),
    metadataMatchType: firstDefined("metadataMatchType"),
    metadataSourceRefs: metadataSourceRefs.length ? metadataSourceRefs : undefined,
    metadataSources: Object.keys(metadataSources).length ? metadataSources : undefined,
    metadataConfidence: Object.keys(metadataConfidence).length ? metadataConfidence : undefined,
    seriesMetadata: firstDefined("seriesMetadata"),
    seriesCompletionOverride: firstDefined("seriesCompletionOverride"),
    loans: mergeLoanRecords(all),
    duplicateIgnoreIds: unionStrings(all.map((book) => book.duplicateIgnoreIds ?? [])).filter((id) => !all.some((book) => book.id === id)),
    healthExceptions: unionStrings(all.map((book) => book.healthExceptions ?? [])),
    sourceIds: Object.keys(sourceIds).length ? sourceIds : undefined,
    readingSessions,
    readDates,
    dateRead: undefined,
    dateAdded: all.map((book) => book.dateAdded).filter(Boolean).sort()[0] ?? preferred.dateAdded,
    createdAt,
    updatedAt: new Date().toISOString()
  };
}

function mergeLoanRecords(books: Book[]): Book["loans"] {
  const map = new Map<string, NonNullable<Book["loans"]>[number]>();
  for (const book of books) for (const loan of book.loans ?? []) {
    const existing = map.get(loan.id);
    if (!existing || loan.updatedAt > existing.updatedAt) map.set(loan.id, { ...loan });
  }
  const loans = [...map.values()].sort((a, b) => a.loanedAt.localeCompare(b.loanedAt));
  return loans.length ? loans : undefined;
}

function completenessScore(book: Book): number {
  let score = 0;
  for (const value of [book.isbn, book.series, book.publisher, book.language, book.format, book.genre, book.description, book.review, book.coverAssetId || book.coverUrl]) if (value) score += 1;
  if (book.pages) score += 1; if (book.publicationYear) score += 1; if (typeof book.rating === "number") score += 1;
  score += Math.min(3, normalizedReadDates(book).length); score += Math.min(3, book.tags?.length ?? 0); score += Math.min(3, book.shelfIds?.length ?? 0);
  return score;
}

function normalizeText(value: string): string {
  return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function canonicalDuplicateIsbn(value?: string): string | undefined {
  const isbn = normalizeIsbn(value ?? "");
  if (isbn.length === 13 && /^\d{13}$/.test(isbn)) return isbn;
  if (isbn.length !== 10 || !/^\d{9}[\dX]$/.test(isbn)) return undefined;
  const body = `978${isbn.slice(0, 9)}`;
  let sum = 0;
  for (let index = 0; index < 12; index += 1) sum += Number(body[index]) * (index % 2 === 0 ? 1 : 3);
  return `${body}${(10 - (sum % 10)) % 10}`;
}
