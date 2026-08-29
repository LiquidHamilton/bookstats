import Dexie, { type EntityTable } from "dexie";
import type { Book, ReadingGoal, Shelf, SyncEntityType } from "@bookstats/domain";

export interface Tombstone {
  id: string;
  deletedAt: string;
}

export interface LocalMeta {
  key: string;
  value: string;
}

export interface SyncOutboxEntry {
  key: string;
  entityType: SyncEntityType;
  id: string;
  clientUpdatedAt: string;
}

export const db = new Dexie("bookstats") as Dexie & {
  books: EntityTable<Book, "id">;
  shelves: EntityTable<Shelf, "id">;
  goals: EntityTable<ReadingGoal, "id">;
  tombstones: EntityTable<Tombstone, "id">;
  shelfTombstones: EntityTable<Tombstone, "id">;
  goalTombstones: EntityTable<Tombstone, "id">;
  syncOutbox: EntityTable<SyncOutboxEntry, "key">;
  meta: EntityTable<LocalMeta, "key">;
};

db.version(1).stores({
  books: "id, title, author, status, owned, rating, dateAdded, updatedAt, *tags"
});

db.version(2).stores({
  books: "id, title, author, series, status, owned, rating, dateAdded, updatedAt, *tags, *readDates"
}).upgrade(async (transaction) => {
  await transaction.table("books").toCollection().modify((book: Partial<Book>) => {
    book.additionalAuthors ??= [];
    book.tags ??= [];
    book.readDates = Array.isArray(book.readDates) ? book.readDates : (book.dateRead ? [book.dateRead] : []);
  });
});

db.version(3).stores({
  books: "id, title, author, series, status, owned, rating, dateAdded, updatedAt, *tags, *readDates",
  tombstones: "id, deletedAt",
  meta: "key"
});

db.version(4).stores({
  books: "id, title, author, series, status, owned, rating, dateAdded, updatedAt, *tags, *readDates, *shelfIds",
  shelves: "id, &name, updatedAt",
  tombstones: "id, deletedAt",
  shelfTombstones: "id, deletedAt",
  meta: "key"
}).upgrade(async (transaction) => {
  await transaction.table("books").toCollection().modify((book: Partial<Book>) => {
    book.additionalAuthors ??= [];
    book.tags ??= [];
    book.shelfIds ??= [];
    book.readDates = Array.isArray(book.readDates) ? book.readDates : (book.dateRead ? [book.dateRead] : []);
  });
});

// v0.7 adds reading goals. Reading sessions live inside each Book record so old
// libraries require no destructive data migration.
db.version(5).stores({
  books: "id, title, author, series, status, owned, rating, dateAdded, updatedAt, *tags, *readDates, *shelfIds",
  shelves: "id, &name, updatedAt",
  goals: "id, metric, startDate, endDate, updatedAt",
  tombstones: "id, deletedAt",
  shelfTombstones: "id, deletedAt",
  goalTombstones: "id, deletedAt",
  meta: "key"
});

// v1.1 adds a persistent per-record sync outbox. Local writes update one compact
// outbox entry, so repeated edits coalesce and a normal sync no longer has to scan
// or upload the complete library.
db.version(6).stores({
  books: "id, title, author, series, status, owned, rating, dateAdded, updatedAt, *tags, *readDates, *shelfIds",
  shelves: "id, &name, updatedAt",
  goals: "id, metric, startDate, endDate, updatedAt",
  tombstones: "id, deletedAt",
  shelfTombstones: "id, deletedAt",
  goalTombstones: "id, deletedAt",
  syncOutbox: "key, entityType, id, clientUpdatedAt",
  meta: "key"
});
