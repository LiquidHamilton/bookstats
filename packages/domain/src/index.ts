export type ReadingStatus =
  | "not_started"
  | "want_to_read"
  | "currently_reading"
  | "read"
  | "did_not_finish"
  | "on_hold";

export type BookFormat =
  | "Hardcover"
  | "Paperback"
  | "Mass Market Paperback"
  | "eBook"
  | "Audiobook"
  | "Graphic Novel"
  | "Omnibus"
  | "Other";

export const BOOK_CONDITIONS = ["New", "Like New", "Very Good", "Good", "Acceptable", "Poor"] as const;
export type BookCondition = (typeof BOOK_CONDITIONS)[number];

export interface ReadingSession {
  id: string;
  /** Inclusive calendar date when this reading began. */
  startedAt?: string;
  /** Completion date. A session without this date is an active/incomplete reading. */
  finishedAt?: string;
  /** Current page within this reading, mainly used by active sessions. */
  progressPages?: number;
  /** Optional session-specific note. */
  notes?: string;
  createdAt: string;
  updatedAt: string;
}

export type ReadingGoalMetric = "books" | "pages" | "rereads" | "new_authors" | "owned_books";

export interface ReadingGoal {
  id: string;
  name: string;
  metric: ReadingGoalMetric;
  target: number;
  startDate: string;
  endDate: string;
  createdAt: string;
  updatedAt: string;
}

export type ShelfKind = "manual" | "smart";
export type ShelfMatchMode = "all" | "any";
export type ShelfRuleField =
  | "status"
  | "condition"
  | "owned"
  | "rating"
  | "title"
  | "author"
  | "series"
  | "format"
  | "genre"
  | "tag"
  | "pages"
  | "publicationYear"
  | "readCount"
  | "lastRead"
  | "dateAdded";
export type ShelfRuleOperator =
  | "equals"
  | "not_equals"
  | "contains"
  | "not_contains"
  | "gte"
  | "lte"
  | "is_true"
  | "is_false";

export interface ShelfRule {
  id: string;
  field: ShelfRuleField;
  operator: ShelfRuleOperator;
  value?: string;
}

export interface ShelfRuleGroup {
  id: string;
  /** Rules inside this group are combined with AND (all) or OR (any). */
  match: ShelfMatchMode;
  rules: ShelfRule[];
}

export interface Shelf {
  id: string;
  name: string;
  /** User-controlled display position. Older shelves without this field are ordered by name until rearranged. */
  order?: number;
  /** Omitted on pre-v0.6 shelves; omitted is treated as a normal/manual shelf. */
  kind?: ShelfKind;
  /** Legacy pre-v0.9.4 flat smart-shelf rules. New shelves use ruleGroups. */
  rules?: ShelfRule[];
  /** Legacy flat-rule match mode, or the between-group match mode for grouped smart shelves. */
  match?: ShelfMatchMode;
  /** Grouped Boolean rules: each group has its own AND/OR mode, then groups are combined by match. */
  ruleGroups?: ShelfRuleGroup[];
  createdAt: string;
  updatedAt: string;
}


export interface LoanRecord {
  id: string;
  borrower: string;
  loanedAt: string;
  dueAt?: string;
  returnedAt?: string;
  notes?: string;
  createdAt: string;
  updatedAt: string;
}

export interface SeriesCompletionOverride {
  /** Hide this series from collection-completion tracking without changing book metadata. */
  ignoredFromTracking?: boolean;
  /** Optional manual expected size for the mainline series. */
  expectedCount?: number;
  /** Catalog rows explicitly excluded from completion calculations. */
  excludedProviderIds?: string[];
  /** When present, only these provider rows count toward completion. */
  includedProviderIds?: string[];
  /** User-authored entries for books a provider catalog omitted. */
  manualBooks?: SeriesCatalogBook[];
  updatedAt: string;
}

export interface Book {
  id: string;
  title: string;
  author: string;
  additionalAuthors: string[];
  isbn?: string;
  series?: string;
  seriesVolume?: string;
  publicationYear?: number;
  publisher?: string;
  language?: string;
  pages?: number;
  format?: BookFormat;
  /** User-entered physical condition. Catalog metadata never overwrites this field. */
  condition?: BookCondition;
  /** A book has exactly one lifecycle/reading status. */
  status: ReadingStatus;
  owned: boolean;
  /** User-defined many-to-many shelves, separate from reading status. */
  shelfIds: string[];
  /** 0.5 through 5.0 in half-star increments. */
  rating?: number;
  tags: string[];
  genre?: string;
  description?: string;
  review?: string;
  notes?: string;
  /** Original catalog/source URL, or a legacy data URL for a local/custom cover. */
  coverUrl?: string;
  /** Definitive BookStats cloud asset selected by the user. */
  coverAssetId?: string;
  /** Opaque per-asset access token used only to render the selected cloud cover. */
  coverAssetToken?: string;
  /** Original external URL retained as provenance after BookStats archives the image. */
  coverSourceUrl?: string;
  /** True when a selected cover still needs to be archived by the cloud service. */
  coverArchivePending?: boolean;
  /** Local-only copy of the selected cover. Never intentionally uploaded during normal sync. */
  cachedCoverDataUrl?: string;
  /** Legacy primary metadata reference retained for backward compatibility. */
  metadataSource?: string;
  metadataWorkId?: string;
  metadataEditionId?: string;
  /** How the most recently applied catalog record was matched. */
  metadataMatchType?: MetadataMatchType;
  /** Provider work/edition references retained for provenance and future refreshes. */
  metadataSourceRefs?: MetadataSourceRef[];
  /** Provider provenance for catalog-managed fields. */
  metadataSources?: Partial<Record<MetadataField, MetadataProvider>>;
  /** Internal 0-100 confidence assigned to catalog-managed fields. Not shown in the normal UI. */
  metadataConfidence?: Partial<Record<MetadataField, number>>;
  /** @deprecated Legacy pre-v1.0.2 field. Retained only so older exports remain readable; current metadata lookup ignores it. */
  metadataManualFields?: MetadataField[];
  /** Known series catalog information from a provider such as Hardcover or Google Books. */
  seriesMetadata?: SeriesMetadata;
  /** User-controlled completion rules for this series. Applied consistently across books in the same series. */
  seriesCompletionOverride?: SeriesCompletionOverride;
  /** Lending history. The newest record without returnedAt is the active loan. */
  loans?: LoanRecord[];
  /** Record IDs that the user explicitly confirmed are separate editions/copies, not duplicates. */
  duplicateIgnoreIds?: string[];
  /** Library Health findings the user explicitly confirmed are intentional/not applicable for this record. */
  healthExceptions?: string[];
  /** Stable IDs from imported services, used to make repeated imports idempotent. */
  sourceIds?: Record<string, string>;
  dateAdded: string;
  /** Reading attempts, including active progress and historical start/finish dates. */
  readingSessions?: ReadingSession[];
  /** Completion dates in chronological order. Kept for backward compatibility/imports. */
  readDates: string[];
  /** Legacy v0.1 field retained only for import/backward compatibility. */
  dateRead?: string;
  createdAt: string;
  updatedAt: string;
}

export interface LibrarySummary {
  totalBooks: number;
  ownedBooks: number;
  readBooks: number;
  unreadBooks: number;
  wantToRead: number;
  pagesKnown: number;
  pagesRead: number;
  averageRating: number | null;
  uniqueAuthors: number;
  totalReadings: number;
  rereads: number;
  uniqueGenres: number;
  uniqueSeries: number;
}

export type MetadataProvider = "openlibrary" | "googlebooks" | "hardcover" | "aggregate";
export type MetadataMatchType = "exact_isbn" | "edition" | "work" | "search";
export type MetadataField =
  | "title"
  | "author"
  | "additionalAuthors"
  | "isbn"
  | "series"
  | "seriesVolume"
  | "publicationYear"
  | "publisher"
  | "language"
  | "pages"
  | "format"
  | "genre"
  | "description"
  | "coverUrl";

export interface MetadataSourceRef {
  provider: Exclude<MetadataProvider, "aggregate">;
  workId: string;
  editionId?: string;
  exactIsbn?: string;
  sourceUrl?: string;
}

export interface MetadataCoverCandidate {
  url: string;
  provider: MetadataProvider;
  confidence: number;
  exactEdition?: boolean;
}

export interface SeriesCatalogBook {
  providerId: string;
  title: string;
  position?: string;
  isbn?: string;
  coverUrl?: string;
  author?: string;
}

export interface SeriesMetadata {
  provider: "googlebooks" | "hardcover" | "openlibrary";
  id: string;
  name: string;
  totalBooks?: number;
  primaryBooksCount?: number;
  isCompleted?: boolean;
  books: SeriesCatalogBook[];
  updatedAt?: string;
}

export interface MetadataSeriesMembership {
  /** Catalog/provider that supplied this series membership. */
  provider?: Exclude<MetadataProvider, "aggregate">;
  /** Provider series ID when one is available. */
  seriesId?: string;
  /** Series name and its matching position must travel together. */
  name: string;
  volume?: string;
  /** Full catalog metadata when details for this membership were loaded. */
  metadata?: SeriesMetadata;
}

export interface MetadataCandidate {
  source: MetadataProvider;
  /** Primary work/provider ID retained for compatibility with the pre-v0.9 client. */
  workId: string;
  editionId?: string;
  sourceRefs?: MetadataSourceRef[];
  matchType?: MetadataMatchType;
  confidence?: number;
  exactEdition?: boolean;
  title: string;
  author: string;
  additionalAuthors: string[];
  isbn?: string;
  publicationYear?: number;
  publisher?: string;
  language?: string;
  pages?: number;
  format?: BookFormat;
  series?: string;
  seriesVolume?: string;
  seriesMetadata?: SeriesMetadata;
  /** All known series memberships, preserving each name/position pairing. */
  seriesMemberships?: MetadataSeriesMembership[];
  subjects: string[];
  description?: string;
  coverUrl?: string;
  /** Alternate cover choices aggregated across metadata providers. */
  coverUrls?: string[];
  /** Per-cover provenance retained even when client-side quality ranking changes display order. */
  coverCandidates?: MetadataCoverCandidate[];
  sourceUrl?: string;
  /** Provider selected for each normalized field after merge. */
  fieldSources?: Partial<Record<MetadataField, MetadataProvider>>;
  /** Internal 0-100 confidence for each normalized field after provider merge. */
  fieldConfidence?: Partial<Record<MetadataField, number>>;
}

export interface UserAccount {
  id: string;
  email: string;
  displayName: string;
  createdAt: string;
  emailVerified: boolean;
}

export interface AuthResponse {
  token: string;
  user: UserAccount;
  emailVerificationSent?: boolean;
}

export type SyncEntityType = "book" | "shelf" | "goal";

export interface SyncMutation {
  id: string;
  entityType?: SyncEntityType;
  deleted: boolean;
  book?: Book;
  shelf?: Shelf;
  goal?: ReadingGoal;
  clientUpdatedAt: string;
}

export interface SyncRecord {
  id: string;
  entityType?: SyncEntityType;
  deleted: boolean;
  book?: Book;
  shelf?: Shelf;
  goal?: ReadingGoal;
  clientUpdatedAt: string;
  serverUpdatedAt: string;
  revision: number;
}

export interface SyncAcknowledgement {
  id: string;
  entityType?: SyncEntityType;
  deleted: boolean;
  clientUpdatedAt: string;
}

export interface SyncResponse {
  cursor: string;
  changes: SyncRecord[];
  accepted: number;
  /** Mutations the server has durably accepted or already superseded. */
  acknowledged?: SyncAcknowledgement[];
}

export const READING_STATUS_LABELS: Record<ReadingStatus, string> = {
  not_started: "Not Started",
  want_to_read: "Want to Read",
  currently_reading: "Currently Reading",
  read: "Read",
  did_not_finish: "Did Not Finish",
  on_hold: "On Hold"
};

export function normalizedReadingSessions(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): ReadingSession[] {
  const sessions = Array.isArray(book.readingSessions) ? book.readingSessions.filter(Boolean).map((session) => ({ ...session })) : [];
  if (sessions.length > 0) return sessions.sort((a, b) => (a.finishedAt ?? a.startedAt ?? "9999").localeCompare(b.finishedAt ?? b.startedAt ?? "9999"));
  const legacyDates = Array.isArray(book.readDates) ? book.readDates.filter(Boolean) : [];
  if (book.dateRead) legacyDates.push(book.dateRead);
  return [...new Set(legacyDates)].sort().map((date, index) => ({
    id: `legacy-${index}-${date}`,
    finishedAt: date,
    createdAt: `${date}T12:00:00.000Z`,
    updatedAt: `${date}T12:00:00.000Z`
  }));
}

export function normalizedReadDates(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): string[] {
  const dates = [
    ...(Array.isArray(book.readDates) ? book.readDates.filter(Boolean) : []),
    ...(Array.isArray(book.readingSessions) ? book.readingSessions.map((session) => session.finishedAt).filter((value): value is string => Boolean(value)) : []),
    ...(book.dateRead ? [book.dateRead] : [])
  ];
  return [...new Set(dates)].sort();
}

export function activeReadingSession(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): ReadingSession | undefined {
  return [...normalizedReadingSessions(book)].reverse().find((session) => !session.finishedAt);
}


export function normalizedLoans(book: Pick<Book, "loans">): LoanRecord[] {
  return (Array.isArray(book.loans) ? book.loans : [])
    .filter((loan): loan is LoanRecord => Boolean(loan?.id && loan.borrower?.trim() && loan.loanedAt))
    .map((loan) => ({ ...loan, borrower: loan.borrower.trim() }))
    .sort((a, b) => (a.loanedAt || a.createdAt).localeCompare(b.loanedAt || b.createdAt));
}

export function activeLoan(book: Pick<Book, "loans">): LoanRecord | undefined {
  return [...normalizedLoans(book)].reverse().find((loan) => !loan.returnedAt);
}

export function loanIsOverdue(loan: Pick<LoanRecord, "dueAt" | "returnedAt">, today = new Date().toISOString().slice(0, 10)): boolean {
  return Boolean(!loan.returnedAt && loan.dueAt && loan.dueAt < today);
}

export function normalizeIsbn(value: string): string {
  return value.replace(/[^0-9Xx]/g, "").toUpperCase();
}

export function isSmartShelf(shelf: Shelf): boolean {
  return shelf.kind === "smart";
}

export function sortShelves(shelves: Shelf[]): Shelf[] {
  const explicitOrders = shelves.map((shelf) => shelf.order).filter((value): value is number => Number.isFinite(value));
  let fallbackOrder = explicitOrders.length ? Math.max(...explicitOrders) + 1 : 0;
  return [...shelves]
    .sort((a, b) => {
      const aOrder = typeof a.order === "number" && Number.isFinite(a.order) ? a.order : Number.MAX_SAFE_INTEGER;
      const bOrder = typeof b.order === "number" && Number.isFinite(b.order) ? b.order : Number.MAX_SAFE_INTEGER;
      return aOrder - bOrder || a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    })
    .map((shelf) => typeof shelf.order === "number" && Number.isFinite(shelf.order) ? shelf : { ...shelf, order: fallbackOrder++ });
}

export function normalizedShelfRuleGroups(shelf: Pick<Shelf, "rules" | "match" | "ruleGroups">): ShelfRuleGroup[] {
  if (Array.isArray(shelf.ruleGroups) && shelf.ruleGroups.length > 0) {
    return shelf.ruleGroups
      .filter((group) => Array.isArray(group.rules) && group.rules.length > 0)
      .map((group, index) => ({
        id: group.id || `group-${index + 1}`,
        match: group.match === "any" ? "any" : "all",
        rules: group.rules
      }));
  }
  const legacyRules = Array.isArray(shelf.rules) ? shelf.rules : [];
  if (legacyRules.length === 0) return [];
  return [{ id: "legacy", match: shelf.match === "any" ? "any" : "all", rules: legacyRules }];
}

export function shelfRuleGroupMatchesBook(group: ShelfRuleGroup, book: Book): boolean {
  if (!group.rules.length) return false;
  const checks = group.rules.map((rule) => shelfRuleMatchesBook(rule, book));
  return group.match === "any" ? checks.some(Boolean) : checks.every(Boolean);
}

export function shelfMatchesBook(shelf: Shelf, book: Book): boolean {
  if (!isSmartShelf(shelf)) return (book.shelfIds ?? []).includes(shelf.id);
  const grouped = Array.isArray(shelf.ruleGroups) && shelf.ruleGroups.length > 0;
  const groups = normalizedShelfRuleGroups(shelf);
  if (groups.length === 0) return false;
  const checks = groups.map((group) => shelfRuleGroupMatchesBook(group, book));
  // On legacy flat shelves, match belongs to the rules inside the one generated group.
  // On v0.9.4+ grouped shelves, match combines the groups themselves.
  return grouped && shelf.match === "all" ? checks.every(Boolean) : grouped ? checks.some(Boolean) : checks[0];
}

export function shelfRuleMatchesBook(rule: ShelfRule, book: Book): boolean {
  const text = (value: unknown) => String(value ?? "").trim().toLocaleLowerCase();
  const number = (value: unknown) => {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  };
  const compareText = (actual: unknown) => {
    const left = text(actual);
    const right = text(rule.value);
    if (rule.operator === "contains") return Boolean(right) && left.includes(right);
    if (rule.operator === "not_contains") return Boolean(right) && !left.includes(right);
    if (rule.operator === "not_equals") return left !== right;
    return left === right;
  };
  const compareNumber = (actual: unknown) => {
    const left = number(actual);
    const right = number(rule.value);
    if (left === undefined || right === undefined) return false;
    if (rule.operator === "gte") return left >= right;
    if (rule.operator === "lte") return left <= right;
    if (rule.operator === "not_equals") return left !== right;
    return left === right;
  };
  const compareDate = (actual: unknown) => {
    const left = String(actual ?? "").slice(0, 10);
    const right = String(rule.value ?? "").slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(left) || !/^\d{4}-\d{2}-\d{2}$/.test(right)) return false;
    if (rule.operator === "gte") return left >= right;
    if (rule.operator === "lte") return left <= right;
    if (rule.operator === "not_equals") return left !== right;
    return left === right;
  };

  switch (rule.field) {
    case "owned": return rule.operator === "is_false" ? !book.owned : book.owned;
    case "status": return compareText(book.status);
    case "condition": return compareText(book.condition);
    case "rating": return compareNumber(book.rating);
    case "pages": return compareNumber(book.pages);
    case "publicationYear": return compareNumber(book.publicationYear);
    case "readCount": return compareNumber(normalizedReadDates(book).length);
    case "lastRead": return compareDate(normalizedReadDates(book).at(-1));
    case "dateAdded": return compareDate(book.dateAdded);
    case "tag": {
      const value = text(rule.value);
      const tags = (book.tags ?? []).map(text);
      if (rule.operator === "not_contains" || rule.operator === "not_equals") return !tags.some((tag) => rule.operator === "not_contains" ? tag.includes(value) : tag === value);
      return tags.some((tag) => rule.operator === "contains" ? tag.includes(value) : tag === value);
    }
    case "title": return compareText(book.title);
    case "author": {
      const value = text(rule.value);
      const authors = [book.author, ...(book.additionalAuthors ?? [])].map(text);
      if (rule.operator === "not_contains") return authors.every((author) => !author.includes(value));
      if (rule.operator === "not_equals") return authors.every((author) => author !== value);
      if (rule.operator === "contains") return Boolean(value) && authors.some((author) => author.includes(value));
      return authors.some((author) => author === value);
    }
    case "series": return compareText(book.series);
    case "format": return compareText(book.format);
    case "genre": return compareText(book.genre);
  }
}
