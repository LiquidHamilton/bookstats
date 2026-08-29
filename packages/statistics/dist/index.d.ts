import { type Book, type LibrarySummary, type ReadingGoal, type SeriesCatalogBook, type SeriesCompletionOverride, type SeriesMetadata } from "@bookstats/domain";
export interface DistributionDatum {
    name: string;
    value: number;
}
export interface ReadEvent {
    bookId: string;
    title: string;
    author: string;
    date: string;
    startedAt?: string;
    year: number;
    month: number;
    pages: number;
    isReread: boolean;
    durationDays?: number;
}
export declare function readEvents(books: Book[]): ReadEvent[];
export declare function summarizeLibrary(books: Book[]): LibrarySummary;
export declare function statusDistribution(books: Book[]): DistributionDatum[];
export declare function ratingDistribution(books: Book[]): DistributionDatum[];
export declare function formatDistribution(books: Book[]): DistributionDatum[];
export declare function genreDistribution(books: Book[]): DistributionDatum[];
export declare function seriesDistribution(books: Book[]): DistributionDatum[];
export declare function authorLibraryDistribution(books: Book[]): DistributionDatum[];
export declare function authorReadDistribution(books: Book[]): DistributionDatum[];
export declare function booksReadByYear(books: Book[]): DistributionDatum[];
export declare function pagesReadByYear(books: Book[]): DistributionDatum[];
export declare function booksAddedByYear(books: Book[]): DistributionDatum[];
export declare function publicationDecadeDistribution(books: Book[]): DistributionDatum[];
export declare function pageLengthDistribution(books: Book[]): DistributionDatum[];
export declare function averageRatingByAuthor(books: Book[]): Array<DistributionDatum & {
    count: number;
}>;
export declare function readingMonthDistribution(books: Book[]): DistributionDatum[];
export interface MonthlyReadingDatum {
    name: string;
    books: number;
    pages: number;
}
export declare function monthlyReadingForYear(books: Book[], year: number): MonthlyReadingDatum[];
export declare function newAuthorsByYear(books: Book[]): DistributionDatum[];
export declare function ratingTrendByYear(books: Book[]): DistributionDatum[];
export declare function readFormatDistribution(books: Book[]): DistributionDatum[];
export declare function readGenreDistribution(books: Book[]): DistributionDatum[];
export declare function readOwnershipDistribution(books: Book[]): DistributionDatum[];
export interface ReadingExtremes {
    longest?: {
        title: string;
        author: string;
        pages: number;
    };
    shortest?: {
        title: string;
        author: string;
        pages: number;
    };
}
export declare function readingExtremes(books: Book[]): ReadingExtremes;
export interface GoalProgress {
    current: number;
    target: number;
    percent: number;
    remaining: number;
    complete: boolean;
    elapsedPercent: number;
    daysRemaining: number;
    onPace: boolean;
    projected: number | null;
}
export declare function readingGoalProgress(goal: ReadingGoal, books: Book[], today?: string): GoalProgress;
export interface ReadingPaceSummary {
    timedReads: number;
    averageDaysToFinish: number | null;
    fastestDays: number | null;
    slowestDays: number | null;
    averagePagesPerDay: number | null;
    averagePagesPerRead: number | null;
    activeReadings: number;
}
export declare function readingPaceSummary(books: Book[]): ReadingPaceSummary;
export interface SeriesCatalogStatus extends SeriesCatalogBook {
    inLibrary: boolean;
    owned: boolean;
    read: boolean;
}
export interface SeriesProgress {
    name: string;
    total: number;
    read: number;
    owned: number;
    /** Reading completion among the series books already in the user's library. */
    completionPercent: number;
    knownVolumes: number[];
    /** Numeric gaps inferred only from volume numbers already stored in the user's library. */
    missingVolumeGaps: number[];
    /** Provider-backed catalog information when a selected edition supplied it. */
    catalogProvider?: SeriesMetadata["provider"];
    catalogTotal?: number;
    catalogPrimaryBooksCount?: number;
    catalogIsCompleted?: boolean;
    /** Provider rows before mainline cleanup; useful for explaining noisy catalogs without exposing provider branding. */
    rawCatalogCount?: number;
    /** User-authored completion rules copied from the series books. */
    completionOverride?: SeriesCompletionOverride;
    /** How many mainline series positions are represented by owned library records. */
    collectionPercent?: number;
    catalogBooks: SeriesCatalogStatus[];
    missingCatalogBooks: SeriesCatalogStatus[];
}
export declare function seriesProgress(books: Book[]): SeriesProgress[];
export interface LibraryFlowDatum {
    name: string;
    added: number;
    read: number;
    pages: number;
}
export declare function libraryFlowByYear(books: Book[]): LibraryFlowDatum[];
