import type { Book, ReadingGoal, Shelf } from "@bookstats/domain";
import { sortShelves } from "@bookstats/domain";
import { db, type Tombstone } from "./db";

export interface LibraryRepository {
  readonly kind: "indexeddb" | "sqlite";
  listBooks(): Promise<Book[]>;
  getBook(id: string): Promise<Book | undefined>;
  putBook(book: Book): Promise<void>;
  bulkPutBooks(books: Book[]): Promise<void>;
  deleteBook(id: string, trackTombstone?: boolean): Promise<void>;
  listShelves(): Promise<Shelf[]>;
  putShelf(shelf: Shelf): Promise<void>;
  bulkPutShelves(shelves: Shelf[]): Promise<void>;
  deleteShelf(id: string, trackTombstone?: boolean): Promise<void>;
  listGoals(): Promise<ReadingGoal[]>;
  putGoal(goal: ReadingGoal): Promise<void>;
  bulkPutGoals(goals: ReadingGoal[]): Promise<void>;
  deleteGoal(id: string, trackTombstone?: boolean): Promise<void>;
  listTombstones(): Promise<Tombstone[]>;
  clearTombstones(ids: string[]): Promise<void>;
  listShelfTombstones(): Promise<Tombstone[]>;
  clearShelfTombstones(ids: string[]): Promise<void>;
  listGoalTombstones(): Promise<Tombstone[]>;
  clearGoalTombstones(ids: string[]): Promise<void>;
  getMeta(key: string): Promise<string | undefined>;
  setMeta(key: string, value: string): Promise<void>;
  clearLibraryData(): Promise<void>;
}

let repositoryPromise: Promise<LibraryRepository> | undefined;

export function isDesktopRuntime(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export function getLibraryRepository(): Promise<LibraryRepository> {
  repositoryPromise ??= createRepository();
  return repositoryPromise;
}

async function createRepository(): Promise<LibraryRepository> {
  if (isDesktopRuntime()) {
    try {
      const repository = await createSqliteRepository();
      await seedIfEmpty(repository);
      return repository;
    } catch (error) {
      console.error("Could not initialize desktop SQLite; falling back to IndexedDB.", error);
    }
  }
  const repository = createIndexedDbRepository();
  await seedIfEmpty(repository);
  return repository;
}

function normalizeBook(book: Book): Book {
  const { metadataManualFields: _legacyMetadataManualFields, ...rest } = book;
  return {
    ...rest,
    additionalAuthors: book.additionalAuthors ?? [],
    tags: book.tags ?? [],
    shelfIds: book.shelfIds ?? [],
    loans: Array.isArray(book.loans) ? book.loans.map((loan) => ({ ...loan })) : undefined,
    duplicateIgnoreIds: Array.isArray(book.duplicateIgnoreIds) ? [...new Set(book.duplicateIgnoreIds.filter(Boolean))] : undefined,
    seriesCompletionOverride: book.seriesCompletionOverride ? { ...book.seriesCompletionOverride, excludedProviderIds: [...(book.seriesCompletionOverride.excludedProviderIds ?? [])], includedProviderIds: [...(book.seriesCompletionOverride.includedProviderIds ?? [])], manualBooks: book.seriesCompletionOverride.manualBooks?.map((entry) => ({ ...entry })) } : undefined,
    readingSessions: Array.isArray(book.readingSessions) ? book.readingSessions.map((session) => ({ ...session })) : undefined,
    readDates: Array.isArray(book.readDates) ? book.readDates : (book.dateRead ? [book.dateRead] : [])
  };
}

function normalizeGoal(goal: ReadingGoal): ReadingGoal {
  return { ...goal, target: Math.max(1, Number(goal.target) || 1) };
}

function createIndexedDbRepository(): LibraryRepository {
  return {
    kind: "indexeddb",
    async listBooks() { return (await db.books.toArray()).map(normalizeBook); },
    async getBook(id) { const book = await db.books.get(id); return book ? normalizeBook(book) : undefined; },
    async putBook(book) {
      const normalized = normalizeBook(book);
      await db.transaction("rw", db.books, db.tombstones, async () => {
        await db.books.put(normalized);
        await db.tombstones.delete(normalized.id);
      });
    },
    async bulkPutBooks(books) {
      const normalized = books.map(normalizeBook);
      await db.transaction("rw", db.books, db.tombstones, async () => {
        await db.books.bulkPut(normalized);
        await db.tombstones.bulkDelete(normalized.map((book) => book.id));
      });
    },
    async deleteBook(id, trackTombstone = true) {
      await db.transaction("rw", db.books, db.tombstones, async () => {
        await db.books.delete(id);
        if (trackTombstone) await db.tombstones.put({ id, deletedAt: new Date().toISOString() });
      });
    },
    async listShelves() {
      const shelves = await db.shelves.toArray();
      const ordered = sortShelves(shelves);
      // v0.9.2 and earlier shelves have no explicit order. Persist the stable
      // fallback once so later edits to a single shelf cannot reshuffle the rest.
      if (shelves.some((shelf) => typeof shelf.order !== "number" || !Number.isFinite(shelf.order))) await db.shelves.bulkPut(ordered);
      return ordered;
    },
    async putShelf(shelf) { await db.shelves.put(shelf); await db.shelfTombstones.delete(shelf.id); },
    async bulkPutShelves(shelves) { await db.shelves.bulkPut(shelves); await db.shelfTombstones.bulkDelete(shelves.map((shelf) => shelf.id)); },
    async deleteShelf(id, trackTombstone = true) {
      const now = new Date().toISOString();
      await db.transaction("rw", db.books, db.shelves, db.shelfTombstones, async () => {
        await db.books.filter((book) => (book.shelfIds ?? []).includes(id)).modify((book) => {
          book.shelfIds = (book.shelfIds ?? []).filter((shelfId) => shelfId !== id);
          book.updatedAt = now;
        });
        await db.shelves.delete(id);
        if (trackTombstone) await db.shelfTombstones.put({ id, deletedAt: now });
      });
    },
    async listGoals() { return (await db.goals.toArray()).map(normalizeGoal).sort((a, b) => b.startDate.localeCompare(a.startDate) || a.name.localeCompare(b.name)); },
    async putGoal(goal) { await db.goals.put(normalizeGoal(goal)); await db.goalTombstones.delete(goal.id); },
    async bulkPutGoals(goals) { await db.goals.bulkPut(goals.map(normalizeGoal)); await db.goalTombstones.bulkDelete(goals.map((goal) => goal.id)); },
    async deleteGoal(id, trackTombstone = true) {
      await db.transaction("rw", db.goals, db.goalTombstones, async () => {
        await db.goals.delete(id);
        if (trackTombstone) await db.goalTombstones.put({ id, deletedAt: new Date().toISOString() });
      });
    },
    listTombstones: () => db.tombstones.toArray(),
    clearTombstones: (ids) => db.tombstones.bulkDelete(ids),
    listShelfTombstones: () => db.shelfTombstones.toArray(),
    clearShelfTombstones: (ids) => db.shelfTombstones.bulkDelete(ids),
    listGoalTombstones: () => db.goalTombstones.toArray(),
    clearGoalTombstones: (ids) => db.goalTombstones.bulkDelete(ids),
    async getMeta(key) { return (await db.meta.get(key))?.value; },
    async setMeta(key, value) { await db.meta.put({ key, value }); },
    async clearLibraryData() {
      // On the web this is also a privacy boundary between signed-in users.
      // Clear metadata because it can contain import history and local safety
      // backups with complete library records, not just harmless preferences.
      await db.transaction(
        "rw",
        [db.books, db.shelves, db.goals, db.tombstones, db.shelfTombstones, db.goalTombstones, db.meta],
        async () => {
          await db.books.clear();
          await db.shelves.clear();
          await db.goals.clear();
          await db.tombstones.clear();
          await db.shelfTombstones.clear();
          await db.goalTombstones.clear();
          await db.meta.clear();
        }
      );
    }
  };
}

async function createSqliteRepository(): Promise<LibraryRepository> {
  const { default: Database } = await import("@tauri-apps/plugin-sql");
  const sqlite = await Database.load("sqlite:bookstats.db");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS books (id TEXT PRIMARY KEY NOT NULL, data TEXT NOT NULL, updated_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS shelves (id TEXT PRIMARY KEY NOT NULL, data TEXT NOT NULL, updated_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS goals (id TEXT PRIMARY KEY NOT NULL, data TEXT NOT NULL, updated_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS tombstones (id TEXT PRIMARY KEY NOT NULL, deleted_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS shelf_tombstones (id TEXT PRIMARY KEY NOT NULL, deleted_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS goal_tombstones (id TEXT PRIMARY KEY NOT NULL, deleted_at TEXT NOT NULL)");
  await sqlite.execute("CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY NOT NULL, value TEXT NOT NULL)");

  const countRows = await sqlite.select<Array<{ count: number }>>("SELECT COUNT(*) AS count FROM books");
  if (Number(countRows[0]?.count ?? 0) === 0) {
    // Older desktop builds used the WebView's IndexedDB. Migrate those records once.
    const legacy = await db.books.toArray();
    for (const book of legacy) {
      const normalized = normalizeBook(book);
      await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
    }
    const legacyShelves = await db.shelves.toArray();
    for (const shelf of legacyShelves) {
      await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
    }
    const legacyGoals = await db.goals.toArray().catch(() => [] as ReadingGoal[]);
    for (const goal of legacyGoals) {
      const normalized = normalizeGoal(goal);
      await sqlite.execute("INSERT OR REPLACE INTO goals (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
    }
  }

  return {
    kind: "sqlite",
    async listBooks() {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM books ORDER BY updated_at DESC");
      return rows.map((row) => normalizeBook(JSON.parse(row.data) as Book));
    },
    async getBook(id) {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM books WHERE id = $1 LIMIT 1", [id]);
      return rows[0] ? normalizeBook(JSON.parse(rows[0].data) as Book) : undefined;
    },
    async putBook(book) {
      const normalized = normalizeBook(book);
      await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
      await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [normalized.id]);
    },
    async bulkPutBooks(books) {
      for (const book of books) {
        const normalized = normalizeBook(book);
        await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
        await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [normalized.id]);
      }
    },
    async deleteBook(id, trackTombstone = true) {
      await sqlite.execute("DELETE FROM books WHERE id = $1", [id]);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO tombstones (id, deleted_at) VALUES ($1, $2)", [id, new Date().toISOString()]);
    },
    async listShelves() {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM shelves");
      const shelves = rows.map((row) => JSON.parse(row.data) as Shelf);
      const ordered = sortShelves(shelves);
      if (shelves.some((shelf) => typeof shelf.order !== "number" || !Number.isFinite(shelf.order))) {
        for (const shelf of ordered) await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
      }
      return ordered;
    },
    async putShelf(shelf) {
      await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
      await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [shelf.id]);
    },
    async bulkPutShelves(shelves) {
      for (const shelf of shelves) {
        await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
        await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [shelf.id]);
      }
    },
    async deleteShelf(id, trackTombstone = true) {
      const now = new Date().toISOString();
      const affected = await sqlite.select<Array<{ id: string; data: string }>>("SELECT id, data FROM books");
      for (const row of affected) {
        const book = normalizeBook(JSON.parse(row.data) as Book);
        if (!(book.shelfIds ?? []).includes(id)) continue;
        const updated = { ...book, shelfIds: book.shelfIds.filter((shelfId) => shelfId !== id), updatedAt: now };
        await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [updated.id, JSON.stringify(updated), updated.updatedAt]);
      }
      await sqlite.execute("DELETE FROM shelves WHERE id = $1", [id]);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO shelf_tombstones (id, deleted_at) VALUES ($1, $2)", [id, now]);
    },
    async listGoals() {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM goals ORDER BY updated_at DESC");
      return rows.map((row) => normalizeGoal(JSON.parse(row.data) as ReadingGoal)).sort((a, b) => b.startDate.localeCompare(a.startDate) || a.name.localeCompare(b.name));
    },
    async putGoal(goal) {
      const normalized = normalizeGoal(goal);
      await sqlite.execute("INSERT OR REPLACE INTO goals (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
      await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [normalized.id]);
    },
    async bulkPutGoals(goals) {
      for (const goal of goals) {
        const normalized = normalizeGoal(goal);
        await sqlite.execute("INSERT OR REPLACE INTO goals (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
        await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [normalized.id]);
      }
    },
    async deleteGoal(id, trackTombstone = true) {
      await sqlite.execute("DELETE FROM goals WHERE id = $1", [id]);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO goal_tombstones (id, deleted_at) VALUES ($1, $2)", [id, new Date().toISOString()]);
    },
    async listTombstones() {
      const rows = await sqlite.select<Array<{ id: string; deleted_at: string }>>("SELECT id, deleted_at FROM tombstones");
      return rows.map((row) => ({ id: row.id, deletedAt: row.deleted_at }));
    },
    async clearTombstones(ids) { for (const id of ids) await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [id]); },
    async listShelfTombstones() {
      const rows = await sqlite.select<Array<{ id: string; deleted_at: string }>>("SELECT id, deleted_at FROM shelf_tombstones");
      return rows.map((row) => ({ id: row.id, deletedAt: row.deleted_at }));
    },
    async clearShelfTombstones(ids) { for (const id of ids) await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [id]); },
    async listGoalTombstones() {
      const rows = await sqlite.select<Array<{ id: string; deleted_at: string }>>("SELECT id, deleted_at FROM goal_tombstones");
      return rows.map((row) => ({ id: row.id, deletedAt: row.deleted_at }));
    },
    async clearGoalTombstones(ids) { for (const id of ids) await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [id]); },
    async getMeta(key) {
      const rows = await sqlite.select<Array<{ value: string }>>("SELECT value FROM metadata WHERE key = $1", [key]);
      return rows[0]?.value;
    },
    async setMeta(key, value) { await sqlite.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES ($1, $2)", [key, value]); },
    async clearLibraryData() {
      await sqlite.execute("DELETE FROM books");
      await sqlite.execute("DELETE FROM shelves");
      await sqlite.execute("DELETE FROM goals");
      await sqlite.execute("DELETE FROM tombstones");
      await sqlite.execute("DELETE FROM shelf_tombstones");
      await sqlite.execute("DELETE FROM goal_tombstones");
    }
  };
}

async function seedIfEmpty(repository: LibraryRepository): Promise<void> {
  if (!(await repository.getMeta("initialized"))) await repository.setMeta("initialized", "1");
}
