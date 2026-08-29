import type { Book, ReadingGoal, Shelf, SyncMutation } from "@bookstats/domain";
import { syncAccountLibrary } from "./api";
import type { LibraryRepository } from "./libraryRepository";

export interface SyncResult {
  pushed: number;
  pulled: number;
  cursor: string;
}

export async function synchronizeLibrary(repository: LibraryRepository, books: Book[], shelves: Shelf[], goals: ReadingGoal[], accountId: string): Promise<SyncResult> {
  const [tombstones, shelfTombstones, goalTombstones] = await Promise.all([
    repository.listTombstones(), repository.listShelfTombstones(), repository.listGoalTombstones()
  ]);
  const changes: SyncMutation[] = [
    ...books.map((book) => ({ id: book.id, entityType: "book" as const, deleted: false, book: syncableBook(book), clientUpdatedAt: book.updatedAt })),
    ...shelves.map((shelf) => ({ id: shelf.id, entityType: "shelf" as const, deleted: false, shelf, clientUpdatedAt: shelf.updatedAt })),
    ...goals.map((goal) => ({ id: goal.id, entityType: "goal" as const, deleted: false, goal, clientUpdatedAt: goal.updatedAt })),
    ...tombstones.map((item) => ({ id: item.id, entityType: "book" as const, deleted: true, clientUpdatedAt: item.deletedAt })),
    ...shelfTombstones.map((item) => ({ id: item.id, entityType: "shelf" as const, deleted: true, clientUpdatedAt: item.deletedAt })),
    ...goalTombstones.map((item) => ({ id: item.id, entityType: "goal" as const, deleted: true, clientUpdatedAt: item.deletedAt }))
  ];
  const cursorKey = `cloudSyncCursor:${accountId}`;
  const cursor = await repository.getMeta(cursorKey);
  const response = await syncAccountLibrary(cursor, changes);
  for (const record of response.changes) {
    const entityType = record.entityType ?? "book";
    if (entityType === "shelf") {
      if (record.deleted) await repository.deleteShelf(record.id, false);
      else if (record.shelf) await repository.putShelf(record.shelf);
      continue;
    }
    if (entityType === "goal") {
      if (record.deleted) await repository.deleteGoal(record.id, false);
      else if (record.goal) await repository.putGoal(record.goal);
      continue;
    }
    // A user can save while this network request is in flight. Never let an older
    // server response overwrite that newer local save; the queued follow-up sync will
    // send the local edit back to the server.
    const local = await repository.getBook(record.id);
    if (local && isNewerThan(local.updatedAt, record.clientUpdatedAt)) continue;
    if (record.deleted) {
      await repository.deleteBook(record.id, false);
      continue;
    }
    if (record.book) {
      const sameSelectedCover = Boolean(local && (
        (local.coverAssetId && local.coverAssetId === record.book.coverAssetId) ||
        (local.coverUrl && local.coverUrl === record.book.coverUrl) ||
        (local.coverUrl && local.coverUrl === record.book.coverSourceUrl) ||
        (record.book.coverAssetId && local.coverArchivePending) ||
        // v1.0.1 server migration moves custom data-URL covers out of cloud JSON. Keep
        // the device-local resized copy when that same selection gains an asset reference.
        (record.book.coverAssetId && !local.coverAssetId && Boolean(local.cachedCoverDataUrl) && Boolean(local.coverUrl?.startsWith("data:")))
      ));
      const cachedCoverDataUrl = sameSelectedCover ? local?.cachedCoverDataUrl : undefined;
      await repository.putBook({ ...record.book, shelfIds: record.book.shelfIds ?? [], cachedCoverDataUrl });
    }
  }
  if (tombstones.length) await repository.clearTombstones(tombstones.map((item) => item.id));
  if (shelfTombstones.length) await repository.clearShelfTombstones(shelfTombstones.map((item) => item.id));
  if (goalTombstones.length) await repository.clearGoalTombstones(goalTombstones.map((item) => item.id));
  await repository.setMeta(cursorKey, response.cursor);
  return { pushed: response.accepted, pulled: response.changes.length, cursor: response.cursor };
}

function syncableBook(book: Book): Book {
  const { cachedCoverDataUrl: _localCache, ...rest } = book;
  return rest as Book;
}

export async function resetSyncCursor(repository: LibraryRepository, accountId: string): Promise<void> {
  await repository.setMeta(`cloudSyncCursor:${accountId}`, "1970-01-01T00:00:00.000Z");
}

function isNewerThan(localUpdatedAt?: string, remoteUpdatedAt?: string): boolean {
  if (!localUpdatedAt || !remoteUpdatedAt) return false;
  const local = Date.parse(localUpdatedAt);
  const remote = Date.parse(remoteUpdatedAt);
  return Number.isFinite(local) && Number.isFinite(remote) && local > remote;
}
