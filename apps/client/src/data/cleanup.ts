import type { Book } from "@bookstats/domain";
import { normalizeIsbn, normalizedReadDates, normalizedReadingSessions } from "@bookstats/domain";

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

export type MetadataIssue = "Cover" | "Description" | "ISBN" | "Pages" | "Publication year" | "Series position";
export function metadataIssues(book: Book): MetadataIssue[] {
  const issues: MetadataIssue[] = [];
  if (!book.coverAssetId && !book.coverUrl && !book.cachedCoverDataUrl) issues.push("Cover");
  if (!book.description?.trim()) issues.push("Description");
  if (!book.isbn?.trim()) issues.push("ISBN");
  if (!book.pages) issues.push("Pages");
  if (!book.publicationYear) issues.push("Publication year");
  if (book.series?.trim() && !book.seriesVolume?.trim()) issues.push("Series position");
  return issues;
}

export interface LibraryHealthSummary {
  score: number;
  totalChecks: number;
  passedChecks: number;
  incompleteBooks: number;
  duplicateGroups: number;
  issueCounts: Array<{ issue: MetadataIssue; count: number }>;
}

export function libraryHealth(books: Book[], duplicateGroupsOverride?: number): LibraryHealthSummary {
  const issueOrder: MetadataIssue[] = ["Cover", "Description", "ISBN", "Pages", "Publication year", "Series position"];
  const counts = new Map<MetadataIssue, number>(issueOrder.map((issue) => [issue, 0]));
  let totalChecks = 0; let missingChecks = 0; let incompleteBooks = 0;
  for (const book of books) {
    const missing = metadataIssues(book);
    totalChecks += 5 + (book.series?.trim() ? 1 : 0);
    missingChecks += missing.length;
    if (missing.length) incompleteBooks += 1;
    for (const issue of missing) counts.set(issue, (counts.get(issue) ?? 0) + 1);
  }
  const passedChecks = Math.max(0, totalChecks - missingChecks);
  return {
    score: totalChecks ? Math.round((passedChecks / totalChecks) * 100) : 100,
    totalChecks,
    passedChecks,
    incompleteBooks,
    duplicateGroups: duplicateGroupsOverride ?? findDuplicateGroups(books).length,
    issueCounts: issueOrder.map((issue) => ({ issue, count: counts.get(issue) ?? 0 })).filter((item) => item.count > 0)
  };
}

export function mergeBooks(preferred: Book, others: Book[]): Book {
  const all = [preferred, ...others];
  const firstDefined = <K extends keyof Book>(key: K): Book[K] => all.map((book) => book[key]).find((value) => value !== undefined && value !== "" && value !== null) as Book[K];
  const unionStrings = (values: string[][]) => [...new Set(values.flat().map((value) => value.trim()).filter(Boolean))];
  const sourceIds = Object.assign({}, ...all.map((book) => book.sourceIds ?? {}));
  const metadataSources = Object.assign({}, ...[...all].reverse().map((book) => book.metadataSources ?? {}));
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
    seriesMetadata: firstDefined("seriesMetadata"),
    seriesCompletionOverride: firstDefined("seriesCompletionOverride"),
    loans: mergeLoanRecords(all),
    duplicateIgnoreIds: unionStrings(all.map((book) => book.duplicateIgnoreIds ?? [])).filter((id) => !all.some((book) => book.id === id)),
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
