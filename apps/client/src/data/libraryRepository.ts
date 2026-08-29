import type { Book, ReadingGoal, Shelf, SyncAcknowledgement, SyncEntityType } from "@bookstats/domain";
import { sortShelves } from "@bookstats/domain";
import { db, type SyncOutboxEntry, type Tombstone } from "./db";

export interface LibraryRepository {
  readonly kind: "indexeddb" | "sqlite";
  listBooks(): Promise<Book[]>;
  getBook(id: string): Promise<Book | undefined>;
  putBook(book: Book, trackSync?: boolean): Promise<void>;
  bulkPutBooks(books: Book[], trackSync?: boolean): Promise<void>;
  deleteBook(id: string, trackTombstone?: boolean): Promise<void>;
  listShelves(): Promise<Shelf[]>;
  getShelf(id: string): Promise<Shelf | undefined>;
  putShelf(shelf: Shelf, trackSync?: boolean): Promise<void>;
  bulkPutShelves(shelves: Shelf[], trackSync?: boolean): Promise<void>;
  deleteShelf(id: string, trackTombstone?: boolean): Promise<void>;
  listGoals(): Promise<ReadingGoal[]>;
  getGoal(id: string): Promise<ReadingGoal | undefined>;
  putGoal(goal: ReadingGoal, trackSync?: boolean): Promise<void>;
  bulkPutGoals(goals: ReadingGoal[], trackSync?: boolean): Promise<void>;
  deleteGoal(id: string, trackTombstone?: boolean): Promise<void>;
  listTombstones(): Promise<Tombstone[]>;
  clearTombstones(ids: string[]): Promise<void>;
  listShelfTombstones(): Promise<Tombstone[]>;
  clearShelfTombstones(ids: string[]): Promise<void>;
  listGoalTombstones(): Promise<Tombstone[]>;
  clearGoalTombstones(ids: string[]): Promise<void>;
  listSyncOutbox(): Promise<SyncOutboxEntry[]>;
  markSyncOutbox(entries: SyncOutboxEntry[]): Promise<void>;
  acknowledgeSyncChanges(acknowledged: SyncAcknowledgement[]): Promise<void>;
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

export function syncOutboxKey(entityType: SyncEntityType, id: string): string {
  return `${entityType}:${id}`;
}

export function syncOutboxEntry(entityType: SyncEntityType, id: string, clientUpdatedAt: string): SyncOutboxEntry {
  return { key: syncOutboxKey(entityType, id), entityType, id, clientUpdatedAt };
}

function createIndexedDbRepository(): LibraryRepository {
  return {
    kind: "indexeddb",
    async listBooks() { return (await db.books.toArray()).map(normalizeBook); },
    async getBook(id) { const book = await db.books.get(id); return book ? normalizeBook(book) : undefined; },
    async putBook(book, trackSync = true) {
      const normalized = normalizeBook(book);
      await db.transaction("rw", db.books, db.tombstones, db.syncOutbox, async () => {
        await db.books.put(normalized);
        await db.tombstones.delete(normalized.id);
        if (trackSync) await db.syncOutbox.put(syncOutboxEntry("book", normalized.id, normalized.updatedAt));
        else await db.syncOutbox.delete(syncOutboxKey("book", normalized.id));
      });
    },
    async bulkPutBooks(books, trackSync = true) {
      const normalized = books.map(normalizeBook);
      await db.transaction("rw", db.books, db.tombstones, db.syncOutbox, async () => {
        await db.books.bulkPut(normalized);
        await db.tombstones.bulkDelete(normalized.map((book) => book.id));
        const keys = normalized.map((book) => syncOutboxKey("book", book.id));
        if (trackSync) await db.syncOutbox.bulkPut(normalized.map((book) => syncOutboxEntry("book", book.id, book.updatedAt)));
        else await db.syncOutbox.bulkDelete(keys);
      });
    },
    async deleteBook(id, trackTombstone = true) {
      await db.transaction("rw", db.books, db.tombstones, db.syncOutbox, async () => {
        await db.books.delete(id);
        await db.syncOutbox.delete(syncOutboxKey("book", id));
        if (trackTombstone) await db.tombstones.put({ id, deletedAt: new Date().toISOString() });
        else await db.tombstones.delete(id);
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
    async getShelf(id) { return db.shelves.get(id); },
    async putShelf(shelf, trackSync = true) {
      await db.transaction("rw", db.shelves, db.shelfTombstones, db.syncOutbox, async () => {
        await db.shelves.put(shelf);
        await db.shelfTombstones.delete(shelf.id);
        if (trackSync) await db.syncOutbox.put(syncOutboxEntry("shelf", shelf.id, shelf.updatedAt));
        else await db.syncOutbox.delete(syncOutboxKey("shelf", shelf.id));
      });
    },
    async bulkPutShelves(shelves, trackSync = true) {
      await db.transaction("rw", db.shelves, db.shelfTombstones, db.syncOutbox, async () => {
        await db.shelves.bulkPut(shelves);
        await db.shelfTombstones.bulkDelete(shelves.map((shelf) => shelf.id));
        const keys = shelves.map((shelf) => syncOutboxKey("shelf", shelf.id));
        if (trackSync) await db.syncOutbox.bulkPut(shelves.map((shelf) => syncOutboxEntry("shelf", shelf.id, shelf.updatedAt)));
        else await db.syncOutbox.bulkDelete(keys);
      });
    },
    async deleteShelf(id, trackTombstone = true) {
      const now = new Date().toISOString();
      await db.transaction("rw", db.books, db.shelves, db.shelfTombstones, db.syncOutbox, async () => {
        const affected = await db.books.filter((book) => (book.shelfIds ?? []).includes(id)).toArray();
        if (affected.length) {
          const updated = affected.map((book) => ({
            ...book,
            shelfIds: (book.shelfIds ?? []).filter((shelfId) => shelfId !== id),
            updatedAt: trackTombstone ? now : book.updatedAt
          }));
          await db.books.bulkPut(updated);
          if (trackTombstone) await db.syncOutbox.bulkPut(updated.map((book) => syncOutboxEntry("book", book.id, book.updatedAt)));
        }
        await db.shelves.delete(id);
        await db.syncOutbox.delete(syncOutboxKey("shelf", id));
        if (trackTombstone) await db.shelfTombstones.put({ id, deletedAt: now });
        else await db.shelfTombstones.delete(id);
      });
    },
    async listGoals() { return (await db.goals.toArray()).map(normalizeGoal).sort((a, b) => b.startDate.localeCompare(a.startDate) || a.name.localeCompare(b.name)); },
    async getGoal(id) { const goal = await db.goals.get(id); return goal ? normalizeGoal(goal) : undefined; },
    async putGoal(goal, trackSync = true) {
      const normalized = normalizeGoal(goal);
      await db.transaction("rw", db.goals, db.goalTombstones, db.syncOutbox, async () => {
        await db.goals.put(normalized);
        await db.goalTombstones.delete(normalized.id);
        if (trackSync) await db.syncOutbox.put(syncOutboxEntry("goal", normalized.id, normalized.updatedAt));
        else await db.syncOutbox.delete(syncOutboxKey("goal", normalized.id));
      });
    },
    async bulkPutGoals(goals, trackSync = true) {
      const normalized = goals.map(normalizeGoal);
      await db.transaction("rw", db.goals, db.goalTombstones, db.syncOutbox, async () => {
        await db.goals.bulkPut(normalized);
        await db.goalTombstones.bulkDelete(normalized.map((goal) => goal.id));
        const keys = normalized.map((goal) => syncOutboxKey("goal", goal.id));
        if (trackSync) await db.syncOutbox.bulkPut(normalized.map((goal) => syncOutboxEntry("goal", goal.id, goal.updatedAt)));
        else await db.syncOutbox.bulkDelete(keys);
      });
    },
    async deleteGoal(id, trackTombstone = true) {
      await db.transaction("rw", db.goals, db.goalTombstones, db.syncOutbox, async () => {
        await db.goals.delete(id);
        await db.syncOutbox.delete(syncOutboxKey("goal", id));
        if (trackTombstone) await db.goalTombstones.put({ id, deletedAt: new Date().toISOString() });
        else await db.goalTombstones.delete(id);
      });
    },
    listTombstones: () => db.tombstones.toArray(),
    clearTombstones: (ids) => db.tombstones.bulkDelete(ids),
    listShelfTombstones: () => db.shelfTombstones.toArray(),
    clearShelfTombstones: (ids) => db.shelfTombstones.bulkDelete(ids),
    listGoalTombstones: () => db.goalTombstones.toArray(),
    clearGoalTombstones: (ids) => db.goalTombstones.bulkDelete(ids),
    async listSyncOutbox() { return db.syncOutbox.orderBy("clientUpdatedAt").toArray(); },
    async markSyncOutbox(entries) { if (entries.length) await db.syncOutbox.bulkPut(entries); },
    async acknowledgeSyncChanges(acknowledged) {
      if (!acknowledged.length) return;
      await db.transaction("rw", db.syncOutbox, db.tombstones, db.shelfTombstones, db.goalTombstones, async () => {
        for (const ack of acknowledged) {
          const entityType = ack.entityType ?? "book";
          if (!ack.deleted) {
            const key = syncOutboxKey(entityType, ack.id);
            const entry = await db.syncOutbox.get(key);
            if (entry?.clientUpdatedAt === ack.clientUpdatedAt) await db.syncOutbox.delete(key);
            continue;
          }
          if (entityType === "shelf") {
            const tombstone = await db.shelfTombstones.get(ack.id);
            if (tombstone?.deletedAt === ack.clientUpdatedAt) await db.shelfTombstones.delete(ack.id);
          } else if (entityType === "goal") {
            const tombstone = await db.goalTombstones.get(ack.id);
            if (tombstone?.deletedAt === ack.clientUpdatedAt) await db.goalTombstones.delete(ack.id);
          } else {
            const tombstone = await db.tombstones.get(ack.id);
            if (tombstone?.deletedAt === ack.clientUpdatedAt) await db.tombstones.delete(ack.id);
          }
        }
      });
    },
    async getMeta(key) { return (await db.meta.get(key))?.value; },
    async setMeta(key, value) { await db.meta.put({ key, value }); },
    async clearLibraryData() {
      // On the web this is also a privacy boundary between signed-in users.
      // Clear metadata because it can contain import history and local safety
      // backups with complete library records, not just harmless preferences.
      await db.transaction(
        "rw",
        [db.books, db.shelves, db.goals, db.tombstones, db.shelfTombstones, db.goalTombstones, db.syncOutbox, db.meta],
        async () => {
          await db.books.clear();
          await db.shelves.clear();
          await db.goals.clear();
          await db.tombstones.clear();
          await db.shelfTombstones.clear();
          await db.goalTombstones.clear();
          await db.syncOutbox.clear();
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
  await sqlite.execute("CREATE TABLE IF NOT EXISTS sync_outbox (key TEXT PRIMARY KEY NOT NULL, entity_type TEXT NOT NULL, record_id TEXT NOT NULL, client_updated_at TEXT NOT NULL)");
  await sqlite.execute("CREATE INDEX IF NOT EXISTS sync_outbox_updated_at_idx ON sync_outbox (client_updated_at)");
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

  const putOutbox = async (entityType: SyncEntityType, id: string, clientUpdatedAt: string) => {
    const entry = syncOutboxEntry(entityType, id, clientUpdatedAt);
    await sqlite.execute("INSERT OR REPLACE INTO sync_outbox (key, entity_type, record_id, client_updated_at) VALUES ($1, $2, $3, $4)", [entry.key, entry.entityType, entry.id, entry.clientUpdatedAt]);
  };
  const deleteOutbox = (entityType: SyncEntityType, id: string) => sqlite.execute("DELETE FROM sync_outbox WHERE key = $1", [syncOutboxKey(entityType, id)]);

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
    async putBook(book, trackSync = true) {
      const normalized = normalizeBook(book);
      await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
      await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [normalized.id]);
      if (trackSync) await putOutbox("book", normalized.id, normalized.updatedAt); else await deleteOutbox("book", normalized.id);
    },
    async bulkPutBooks(books, trackSync = true) {
      for (const book of books) {
        const normalized = normalizeBook(book);
        await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
        await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [normalized.id]);
        if (trackSync) await putOutbox("book", normalized.id, normalized.updatedAt); else await deleteOutbox("book", normalized.id);
      }
    },
    async deleteBook(id, trackTombstone = true) {
      await sqlite.execute("DELETE FROM books WHERE id = $1", [id]);
      await deleteOutbox("book", id);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO tombstones (id, deleted_at) VALUES ($1, $2)", [id, new Date().toISOString()]);
      else await sqlite.execute("DELETE FROM tombstones WHERE id = $1", [id]);
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
    async getShelf(id) {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM shelves WHERE id = $1 LIMIT 1", [id]);
      return rows[0] ? JSON.parse(rows[0].data) as Shelf : undefined;
    },
    async putShelf(shelf, trackSync = true) {
      await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
      await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [shelf.id]);
      if (trackSync) await putOutbox("shelf", shelf.id, shelf.updatedAt); else await deleteOutbox("shelf", shelf.id);
    },
    async bulkPutShelves(shelves, trackSync = true) {
      for (const shelf of shelves) {
        await sqlite.execute("INSERT OR REPLACE INTO shelves (id, data, updated_at) VALUES ($1, $2, $3)", [shelf.id, JSON.stringify(shelf), shelf.updatedAt]);
        await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [shelf.id]);
        if (trackSync) await putOutbox("shelf", shelf.id, shelf.updatedAt); else await deleteOutbox("shelf", shelf.id);
      }
    },
    async deleteShelf(id, trackTombstone = true) {
      const now = new Date().toISOString();
      const affected = await sqlite.select<Array<{ id: string; data: string }>>("SELECT id, data FROM books");
      for (const row of affected) {
        const book = normalizeBook(JSON.parse(row.data) as Book);
        if (!(book.shelfIds ?? []).includes(id)) continue;
        const updated = { ...book, shelfIds: book.shelfIds.filter((shelfId) => shelfId !== id), updatedAt: trackTombstone ? now : book.updatedAt };
        await sqlite.execute("INSERT OR REPLACE INTO books (id, data, updated_at) VALUES ($1, $2, $3)", [updated.id, JSON.stringify(updated), updated.updatedAt]);
        if (trackTombstone) await putOutbox("book", updated.id, updated.updatedAt);
      }
      await sqlite.execute("DELETE FROM shelves WHERE id = $1", [id]);
      await deleteOutbox("shelf", id);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO shelf_tombstones (id, deleted_at) VALUES ($1, $2)", [id, now]);
      else await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1", [id]);
    },
    async listGoals() {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM goals ORDER BY updated_at DESC");
      return rows.map((row) => normalizeGoal(JSON.parse(row.data) as ReadingGoal)).sort((a, b) => b.startDate.localeCompare(a.startDate) || a.name.localeCompare(b.name));
    },
    async getGoal(id) {
      const rows = await sqlite.select<Array<{ data: string }>>("SELECT data FROM goals WHERE id = $1 LIMIT 1", [id]);
      return rows[0] ? normalizeGoal(JSON.parse(rows[0].data) as ReadingGoal) : undefined;
    },
    async putGoal(goal, trackSync = true) {
      const normalized = normalizeGoal(goal);
      await sqlite.execute("INSERT OR REPLACE INTO goals (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
      await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [normalized.id]);
      if (trackSync) await putOutbox("goal", normalized.id, normalized.updatedAt); else await deleteOutbox("goal", normalized.id);
    },
    async bulkPutGoals(goals, trackSync = true) {
      for (const goal of goals) {
        const normalized = normalizeGoal(goal);
        await sqlite.execute("INSERT OR REPLACE INTO goals (id, data, updated_at) VALUES ($1, $2, $3)", [normalized.id, JSON.stringify(normalized), normalized.updatedAt]);
        await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [normalized.id]);
        if (trackSync) await putOutbox("goal", normalized.id, normalized.updatedAt); else await deleteOutbox("goal", normalized.id);
      }
    },
    async deleteGoal(id, trackTombstone = true) {
      await sqlite.execute("DELETE FROM goals WHERE id = $1", [id]);
      await deleteOutbox("goal", id);
      if (trackTombstone) await sqlite.execute("INSERT OR REPLACE INTO goal_tombstones (id, deleted_at) VALUES ($1, $2)", [id, new Date().toISOString()]);
      else await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1", [id]);
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
    async listSyncOutbox() {
      const rows = await sqlite.select<Array<{ key: string; entity_type: SyncEntityType; record_id: string; client_updated_at: string }>>("SELECT key, entity_type, record_id, client_updated_at FROM sync_outbox ORDER BY client_updated_at ASC");
      return rows.map((row) => ({ key: row.key, entityType: row.entity_type, id: row.record_id, clientUpdatedAt: row.client_updated_at }));
    },
    async markSyncOutbox(entries) {
      for (const entry of entries) {
        await sqlite.execute("INSERT OR REPLACE INTO sync_outbox (key, entity_type, record_id, client_updated_at) VALUES ($1, $2, $3, $4)", [entry.key, entry.entityType, entry.id, entry.clientUpdatedAt]);
      }
    },
    async acknowledgeSyncChanges(acknowledged) {
      for (const ack of acknowledged) {
        const entityType = ack.entityType ?? "book";
        if (!ack.deleted) {
          await sqlite.execute("DELETE FROM sync_outbox WHERE key = $1 AND client_updated_at = $2", [syncOutboxKey(entityType, ack.id), ack.clientUpdatedAt]);
        } else if (entityType === "shelf") {
          await sqlite.execute("DELETE FROM shelf_tombstones WHERE id = $1 AND deleted_at = $2", [ack.id, ack.clientUpdatedAt]);
        } else if (entityType === "goal") {
          await sqlite.execute("DELETE FROM goal_tombstones WHERE id = $1 AND deleted_at = $2", [ack.id, ack.clientUpdatedAt]);
        } else {
          await sqlite.execute("DELETE FROM tombstones WHERE id = $1 AND deleted_at = $2", [ack.id, ack.clientUpdatedAt]);
        }
      }
    },
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
      await sqlite.execute("DELETE FROM sync_outbox");
    }
  };
}

async function seedIfEmpty(repository: LibraryRepository): Promise<void> {
  if (!(await repository.getMeta("initialized"))) await repository.setMeta("initialized", "1");
}
