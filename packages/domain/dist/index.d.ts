export type ReadingStatus = "not_started" | "want_to_read" | "currently_reading" | "read" | "did_not_finish" | "on_hold";
export type BookFormat = "Hardcover" | "Paperback" | "Mass Market Paperback" | "eBook" | "Audiobook" | "Graphic Novel" | "Omnibus" | "Other";
export declare const BOOK_CONDITIONS: readonly ["New", "Like New", "Very Good", "Good", "Acceptable", "Poor"];
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
export type ShelfRuleField = "status" | "condition" | "owned" | "rating" | "title" | "author" | "series" | "format" | "genre" | "tag" | "pages" | "publicationYear" | "readCount" | "lastRead" | "dateAdded";
export type ShelfRuleOperator = "equals" | "not_equals" | "contains" | "not_contains" | "gte" | "lte" | "is_true" | "is_false";
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
export type MetadataField = "title" | "author" | "additionalAuthors" | "isbn" | "series" | "seriesVolume" | "publicationYear" | "publisher" | "language" | "pages" | "format" | "genre" | "description" | "coverUrl";
export interface MetadataSourceRef {
    provider: Exclude<MetadataProvider, "aggregate">;
    workId: string;
    editionId?: string;
    exactIsbn?: string;
    sourceUrl?: string;
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
    sourceUrl?: string;
    /** Provider selected for each normalized field after merge. */
    fieldSources?: Partial<Record<MetadataField, MetadataProvider>>;
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
export interface SyncResponse {
    cursor: string;
    changes: SyncRecord[];
    accepted: number;
}
export declare const READING_STATUS_LABELS: Record<ReadingStatus, string>;
export declare function normalizedReadingSessions(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): ReadingSession[];
export declare function normalizedReadDates(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): string[];
export declare function activeReadingSession(book: Pick<Book, "readingSessions" | "readDates" | "dateRead">): ReadingSession | undefined;
export declare function normalizedLoans(book: Pick<Book, "loans">): LoanRecord[];
export declare function activeLoan(book: Pick<Book, "loans">): LoanRecord | undefined;
export declare function loanIsOverdue(loan: Pick<LoanRecord, "dueAt" | "returnedAt">, today?: string): boolean;
export declare function normalizeIsbn(value: string): string;
export declare function isSmartShelf(shelf: Shelf): boolean;
export declare function sortShelves(shelves: Shelf[]): Shelf[];
export declare function normalizedShelfRuleGroups(shelf: Pick<Shelf, "rules" | "match" | "ruleGroups">): ShelfRuleGroup[];
export declare function shelfRuleGroupMatchesBook(group: ShelfRuleGroup, book: Book): boolean;
export declare function shelfMatchesBook(shelf: Shelf, book: Book): boolean;
export declare function shelfRuleMatchesBook(rule: ShelfRule, book: Book): boolean;
