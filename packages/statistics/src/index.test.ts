import { describe, expect, it } from "vitest";
import type { Book, ReadingGoal } from "@bookstats/domain";
import {
  authorReadDistribution,
  booksReadByYear,
  monthlyReadingForYear,
  ratingDistribution,
  ratingTrendByYear,
  readingGoalProgress,
  readingPaceSummary,
  seriesProgress,
  summarizeLibrary
} from "./index.js";

const now = "2026-08-24T12:00:00.000Z";
const books: Book[] = [
  {
    id: "1", title: "A", author: "Author One", additionalAuthors: [], status: "read",
    owned: true, rating: 4.5, pages: 300, tags: [], shelfIds: [], genre: "Science Fiction", readDates: ["2024-05-01", "2026-03-01"], dateAdded: now, createdAt: now, updatedAt: now,
    series: "Example Series", seriesVolume: "1"
  },
  {
    id: "2", title: "B", author: "Author Two", additionalAuthors: [], status: "want_to_read",
    owned: true, pages: 200, tags: [], shelfIds: [], readDates: [], dateAdded: now, createdAt: now, updatedAt: now,
    series: "Example Series", seriesVolume: "3"
  }
];

const sessionBook: Book = {
  id: "3", title: "C", author: "Author Three", additionalAuthors: [], status: "read",
  owned: false, rating: 3.5, pages: 240, tags: [], shelfIds: [], dateAdded: now, createdAt: now, updatedAt: now,
  readingSessions: [
    { id: "s1", startedAt: "2026-01-01", finishedAt: "2026-01-04", createdAt: now, updatedAt: now },
    { id: "s2", startedAt: "2026-02-10", finishedAt: "2026-02-12", createdAt: now, updatedAt: now }
  ]
};

function goal(overrides: Partial<ReadingGoal> = {}): ReadingGoal {
  return {
    id: "goal-1",
    name: "2026 goal",
    metric: "books",
    target: 6,
    startDate: "2026-01-01",
    endDate: "2026-12-31",
    createdAt: now,
    updatedAt: now,
    ...overrides
  };
}

describe("statistics", () => {
  it("summarizes the core library and reading metrics", () => {
    expect(summarizeLibrary(books)).toEqual({
      totalBooks: 2,
      ownedBooks: 2,
      readBooks: 1,
      unreadBooks: 1,
      wantToRead: 1,
      pagesKnown: 500,
      pagesRead: 600,
      averageRating: 4.5,
      uniqueAuthors: 2,
      totalReadings: 2,
      rereads: 1,
      uniqueGenres: 1,
      uniqueSeries: 1
    });
  });

  it("counts rereads as separate reading events by year and author", () => {
    expect(booksReadByYear(books)).toEqual([{ name: "2024", value: 1 }, { name: "2026", value: 1 }]);
    expect(authorReadDistribution(books)[0]).toEqual({ name: "Author One", value: 2 });
  });

  it("keeps half-star ratings as their own buckets", () => {
    const distribution = ratingDistribution(books);
    expect(distribution.find((datum) => datum.name === "4.5★")?.value).toBe(1);
    expect(distribution.find((datum) => datum.name === "5★")?.value).toBe(0);
  });

  it("derives reading pace from explicit reading sessions", () => {
    const pace = readingPaceSummary([sessionBook]);
    expect(pace.timedReads).toBe(2);
    expect(pace.averageDaysToFinish).toBe(3.5);
    expect(pace.fastestDays).toBe(3);
    expect(pace.slowestDays).toBe(4);
    expect(pace.averagePagesPerDay).toBe(70);
    expect(pace.averagePagesPerRead).toBe(240);
    expect(pace.activeReadings).toBe(0);
  });

  it("tracks goal progress, pace, projections, rereads, and new authors", () => {
    const allBooks = [...books, sessionBook];
    const bookGoal = readingGoalProgress(goal(), allBooks, "2026-06-30");
    expect(bookGoal.current).toBe(3);
    expect(bookGoal.percent).toBe(50);
    expect(bookGoal.complete).toBe(false);
    expect(bookGoal.daysRemaining).toBeGreaterThan(0);
    expect(bookGoal.onPace).toBe(true);
    expect(bookGoal.projected).not.toBeNull();
    expect(bookGoal.projected!).toBeGreaterThan(bookGoal.current);

    expect(readingGoalProgress(goal({ metric: "rereads", target: 3 }), allBooks, "2026-06-30").current).toBe(2);
    expect(readingGoalProgress(goal({ metric: "new_authors", target: 2 }), allBooks, "2026-06-30").current).toBe(1);
  });

  it("builds monthly reading and rating trends from completion dates", () => {
    const allBooks = [...books, sessionBook];
    const monthly = monthlyReadingForYear(allBooks, 2026);
    expect(monthly[0]).toEqual({ name: "Jan", books: 1, pages: 240 });
    expect(monthly[1]).toEqual({ name: "Feb", books: 1, pages: 240 });
    expect(monthly[2]).toEqual({ name: "Mar", books: 1, pages: 300 });

    expect(ratingTrendByYear(allBooks)).toEqual([
      { name: "2024", value: 4.5 },
      { name: "2026", value: (4.5 + 3.5 + 3.5) / 3 }
    ]);
  });

  it("detects only gaps between known series volume numbers", () => {
    const progress = seriesProgress(books);
    expect(progress).toHaveLength(1);
    expect(progress[0].name).toBe("Example Series");
    expect(progress[0].knownVolumes).toEqual([1, 3]);
    expect(progress[0].missingVolumeGaps).toEqual([2]);
  });

  it("uses provider series catalogs to identify books missing from the library", () => {
    const catalogBooks: Book[] = books.map((book, index) => ({
      ...book,
      seriesMetadata: index === 0 ? {
        provider: "hardcover",
        id: "42",
        name: "Example Series",
        totalBooks: 3,
        primaryBooksCount: 3,
        isCompleted: true,
        books: [
          { providerId: "a", title: "A", position: "1", isbn: "9780306406157", author: "Author One" },
          { providerId: "middle", title: "Middle Book", position: "2", isbn: "9780140328721", author: "Author One" },
          { providerId: "b", title: "B", position: "3", author: "Author Two" }
        ]
      } : undefined
    }));
    catalogBooks[0].isbn = "0306406152"; // ISBN-10 equivalent of the catalog ISBN-13.
    const progress = seriesProgress(catalogBooks)[0];
    expect(progress.catalogProvider).toBe("hardcover");
    expect(progress.catalogTotal).toBe(3);
    expect(progress.collectionPercent).toBeCloseTo((2 / 3) * 100);
    expect(progress.missingCatalogBooks.map((book) => book.title)).toEqual(["Middle Book"]);
    expect(progress.catalogBooks.find((book) => book.title === "A")?.inLibrary).toBe(true);
  });
  it("collapses noisy translated series catalogs to the mainline numbered sequence", () => {
    const catalogEntries = Array.from({ length: 131 }, (_, index) => {
      const position = String((index % 5) + 1);
      const variant = Math.floor(index / 5);
      return {
        providerId: `catalog-${index + 1}`,
        title: variant === 0 ? `Mainline ${position}` : `Translated variant ${variant}-${position}`,
        position,
        author: "Series Author"
      };
    });
    const seriesBooks: Book[] = Array.from({ length: 5 }, (_, index) => ({
      id: `owned-${index + 1}`, title: `Mainline ${index + 1}`, author: "Series Author", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: [`2026-0${index + 1}-01`], dateAdded: now, createdAt: now, updatedAt: now,
      series: "Mainline Series", seriesVolume: String(index + 1),
      seriesMetadata: index === 0 ? { provider: "hardcover", id: "mainline", name: "Mainline Series", totalBooks: 131, primaryBooksCount: 131, isCompleted: true, books: catalogEntries } : undefined
    }));
    const progress = seriesProgress(seriesBooks)[0];
    expect(progress.rawCatalogCount).toBe(131);
    expect(progress.catalogBooks).toHaveLength(5);
    expect(progress.catalogTotal).toBe(5);
    expect(progress.missingCatalogBooks).toHaveLength(0);
    expect(progress.collectionPercent).toBe(100);
    expect(progress.catalogBooks.map((entry) => entry.position)).toEqual(["1", "2", "3", "4", "5"]);
  });

  it("fills provider catalog gaps from clearly numbered books already in the library", () => {
    const providerPositions = [1, 2, 3, 4, 5, 6, 8, 9, 10, 12];
    const source: Book[] = Array.from({ length: 12 }, (_, index) => ({
      id: `ever-${index + 1}`, title: `Everworld ${index + 1}`, author: "K.A. Applegate", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: ["2026-01-01"], dateAdded: now, createdAt: now, updatedAt: now,
      series: "Everworld", seriesVolume: String(index + 1),
      seriesMetadata: index === 0 ? { provider: "hardcover", id: "everworld", name: "Everworld", totalBooks: 13, primaryBooksCount: 12, books: providerPositions.map((position) => ({ providerId: `provider-${position}`, title: `Everworld ${position}`, position: String(position) })) } : undefined
    }));
    const progress = seriesProgress(source)[0];
    expect(progress.catalogTotal).toBe(12);
    expect(progress.catalogBooks).toHaveLength(12);
    expect(progress.catalogBooks.map((entry) => entry.position)).toEqual(Array.from({ length: 12 }, (_, index) => String(index + 1)));
    expect(progress.catalogBooks.find((entry) => entry.position === "7")?.providerId).toContain("library:ever-7");
    expect(progress.catalogBooks.find((entry) => entry.position === "11")?.providerId).toContain("library:ever-11");
    expect(progress.missingCatalogBooks).toHaveLength(0);
    expect(progress.collectionPercent).toBe(100);
  });

  it("lets one omnibus satisfy multiple mainline series positions", () => {
    const catalog = Array.from({ length: 6 }, (_, index) => ({ providerId: `issue-${index + 1}`, title: `Issue ${index + 1}`, position: String(index + 1) }));
    const source: Book[] = [
      { id: "omni-1", title: "Omnibus Vol. 1", author: "Author", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: ["2026-01-01"], dateAdded: now, createdAt: now, updatedAt: now, series: "Comic Series", seriesVolume: "1 & 2", seriesMetadata: { provider: "hardcover", id: "comic", name: "Comic Series", totalBooks: 6, primaryBooksCount: 6, books: catalog } },
      { id: "omni-2", title: "Omnibus Vol. 2", author: "Author", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: ["2026-01-01"], dateAdded: now, createdAt: now, updatedAt: now, series: "Comic Series", seriesVolume: "3-4" },
      { id: "omni-3", title: "Omnibus Vol. 3", author: "Author", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: ["2026-01-01"], dateAdded: now, createdAt: now, updatedAt: now, series: "Comic Series", seriesVolume: "5, 6" }
    ];
    const progress = seriesProgress(source)[0];
    expect(progress.total).toBe(3);
    expect(progress.knownVolumes).toEqual([1, 2, 3, 4, 5, 6]);
    expect(progress.missingVolumeGaps).toEqual([]);
    expect(progress.catalogBooks).toHaveLength(6);
    expect(progress.catalogBooks.every((entry) => entry.owned)).toBe(true);
    expect(progress.missingCatalogBooks).toHaveLength(0);
    expect(progress.collectionPercent).toBe(100);
  });

  it("honors manual series completion rules when automatic catalog data is wrong", () => {
    const source: Book[] = [{
      id: "override-1", title: "First", author: "Author", additionalAuthors: [], status: "read", owned: true, tags: [], shelfIds: [], readDates: ["2026-01-01"], dateAdded: now, createdAt: now, updatedAt: now, series: "Override Series", seriesVolume: "1",
      seriesMetadata: { provider: "hardcover", id: "override", name: "Override Series", totalBooks: 4, primaryBooksCount: 4, books: [
        { providerId: "one", title: "First", position: "1" }, { providerId: "two", title: "Second", position: "2" }, { providerId: "side", title: "Side Story", position: "3" }, { providerId: "four", title: "Fourth", position: "4" }
      ] },
      seriesCompletionOverride: { expectedCount: 3, includedProviderIds: ["one", "two"], manualBooks: [{ providerId: "manual:three", title: "Real Third", position: "3" }], updatedAt: now }
    }];
    const progress = seriesProgress(source)[0];
    expect(progress.catalogTotal).toBe(3);
    expect(progress.catalogBooks.map((entry) => entry.title)).toEqual(["First", "Second", "Real Third"]);
    expect(progress.missingCatalogBooks.map((entry) => entry.title)).toEqual(["Second", "Real Third"]);
  });

});
