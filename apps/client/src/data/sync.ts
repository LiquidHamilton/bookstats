import type { Book, ReadingGoal, Shelf, SyncAcknowledgement, SyncMutation, SyncRecord, SyncResponse } from "@bookstats/domain";
import { ApiError, syncAccountLibrary } from "./api";
import { syncOutboxEntry, type LibraryRepository } from "./libraryRepository";

const EPOCH = "1970-01-01T00:00:00.000Z";
const MAX_BATCH_RECORDS = 100;
// Normal sync traffic should stay comfortably below both Fastify's and NGINX's
// emergency body limits. A single unusually large record is still sent alone.
const MAX_BATCH_BYTES = 900 * 1024;
const MAX_SYNC_ATTEMPTS = 4;
const RETRY_BASE_MS = 350;

export interface SyncResult {
  pushed: number;
  pulled: number;
  cursor: string;
  batches: number;
}

export async function synchronizeLibrary(repository: LibraryRepository, accountId: string): Promise<SyncResult> {
  await ensureOutboxInitialized(repository, accountId);

  const cursorKey = `cloudSyncCursor:${accountId}`;
  const [cursor, outbox, tombstones, shelfTombstones, goalTombstones] = await Promise.all([
    repository.getMeta(cursorKey),
    repository.listSyncOutbox(),
    repository.listTombstones(),
    repository.listShelfTombstones(),
    repository.listGoalTombstones()
  ]);

  const changes = await materializePendingChanges(repository, outbox);
  changes.push(
    ...tombstones.map((item) => ({ id: item.id, entityType: "book" as const, deleted: true, clientUpdatedAt: item.deletedAt })),
    ...shelfTombstones.map((item) => ({ id: item.id, entityType: "shelf" as const, deleted: true, clientUpdatedAt: item.deletedAt })),
    ...goalTombstones.map((item) => ({ id: item.id, entityType: "goal" as const, deleted: true, clientUpdatedAt: item.deletedAt }))
  );
  changes.sort((left, right) => left.clientUpdatedAt.localeCompare(right.clientUpdatedAt) || mutationKey(left).localeCompare(mutationKey(right)));

  const batches = splitIntoBatches(changes, cursor);
  // Pull-only syncs are intentional. They let a device receive cloud changes without
  // uploading the local library when nothing on this device changed.
  if (batches.length === 0) batches.push([]);

  let activeCursor = cursor;
  let pushed = 0;
  let pulled = 0;
  const startedAt = performance.now();

  for (let index = 0; index < batches.length; index += 1) {
    const batch = batches[index];
    const bytes = requestBytes(activeCursor, batch);
    console.debug(`[BookStats sync] batch ${index + 1}/${batches.length}: ${batch.length} local changes, ${bytes} bytes`);
    const response = await syncBatchWithRetry(activeCursor, batch, index + 1, batches.length);
    pushed += response.accepted;
    pulled += response.changes.length;

    await applySyncResponse(repository, response);
    // v1.1 servers explicitly acknowledge both newly accepted mutations and safe
    // retries/stale mutations that no longer need to be sent. If a deployment briefly
    // hits an older server, leaving the outbox intact is safer than guessing.
    if (response.acknowledged?.length) await repository.acknowledgeSyncChanges(response.acknowledged);

    activeCursor = response.cursor;
    // Persist the server cursor after every successful batch. If a later batch fails,
    // already-applied cloud changes do not need to be downloaded again, while the
    // persistent outbox keeps every unacknowledged local mutation queued.
    await repository.setMeta(cursorKey, activeCursor);
  }

  console.debug(`[BookStats sync] complete: ${pushed} accepted, ${pulled} pulled, ${batches.length} batches, ${Math.round(performance.now() - startedAt)} ms`);
  return { pushed, pulled, cursor: activeCursor ?? EPOCH, batches: batches.length };
}

async function syncBatchWithRetry(cursor: string | undefined, batch: SyncMutation[], batchNumber: number, batchCount: number): Promise<SyncResponse> {
  let lastError: unknown;
  for (let attempt = 1; attempt <= MAX_SYNC_ATTEMPTS; attempt += 1) {
    const startedAt = performance.now();
    try {
      const response = await syncAccountLibrary(cursor, batch);
      console.debug(`[BookStats sync] batch ${batchNumber}/${batchCount} attempt ${attempt}: ok in ${Math.round(performance.now() - startedAt)} ms; ${response.accepted} accepted, ${response.acknowledged?.length ?? 0} acknowledged, ${response.changes.length} pulled`);
      return response;
    } catch (error) {
      lastError = error;
      if (!retryableSyncError(error) || attempt === MAX_SYNC_ATTEMPTS) throw error;
      const exponential = RETRY_BASE_MS * 2 ** (attempt - 1);
      const jitter = Math.round(exponential * (0.15 + Math.random() * 0.2));
      const delay = exponential + jitter;
      console.debug(`[BookStats sync] batch ${batchNumber}/${batchCount} attempt ${attempt} failed transiently; retrying in ${delay} ms`);
      await sleep(delay);
    }
  }
  throw lastError;
}

function retryableSyncError(error: unknown): boolean {
  if (!(error instanceof ApiError)) return false;
  return error.status === 0 || error.status === 408 || error.status === 425 || error.status === 429 || error.status >= 500;
}

function sleep(ms: number): Promise<void> { return new Promise((resolve) => setTimeout(resolve, ms)); }

async function ensureOutboxInitialized(repository: LibraryRepository, accountId: string): Promise<void> {
  const initializedKey = `cloudSyncOutboxInitialized:${accountId}`;
  if (await repository.getMeta(initializedKey) === "1") return;

  // v1.0.x had no persistent outbox. Seed it once from records newer than the last
  // successful sync so an upgrade preserves unsynced offline edits without re-uploading
  // an already-synchronized multi-thousand-book library.
  const baseline = validTimestamp(await repository.getMeta(`lastSuccessfulSync:${accountId}`)) ?? EPOCH;
  const [books, shelves, goals] = await Promise.all([repository.listBooks(), repository.listShelves(), repository.listGoals()]);
  const entries = [
    ...books.filter((book) => isAfter(book.updatedAt, baseline)).map((book) => syncOutboxEntry("book", book.id, book.updatedAt)),
    ...shelves.filter((shelf) => isAfter(shelf.updatedAt, baseline)).map((shelf) => syncOutboxEntry("shelf", shelf.id, shelf.updatedAt)),
    ...goals.filter((goal) => isAfter(goal.updatedAt, baseline)).map((goal) => syncOutboxEntry("goal", goal.id, goal.updatedAt))
  ];
  await repository.markSyncOutbox(entries);
  await repository.setMeta(initializedKey, "1");
}

async function materializePendingChanges(repository: LibraryRepository, outbox: Awaited<ReturnType<LibraryRepository["listSyncOutbox"]>>): Promise<SyncMutation[]> {
  const changes: SyncMutation[] = [];
  const staleAcknowledgements: SyncAcknowledgement[] = [];

  for (const entry of outbox) {
    if (entry.entityType === "shelf") {
      const shelf = await repository.getShelf(entry.id);
      if (!shelf) { staleAcknowledgements.push({ id: entry.id, entityType: "shelf", deleted: false, clientUpdatedAt: entry.clientUpdatedAt }); continue; }
      if (shelf.updatedAt !== entry.clientUpdatedAt) await repository.markSyncOutbox([syncOutboxEntry("shelf", shelf.id, shelf.updatedAt)]);
      changes.push({ id: shelf.id, entityType: "shelf", deleted: false, shelf, clientUpdatedAt: shelf.updatedAt });
      continue;
    }
    if (entry.entityType === "goal") {
      const goal = await repository.getGoal(entry.id);
      if (!goal) { staleAcknowledgements.push({ id: entry.id, entityType: "goal", deleted: false, clientUpdatedAt: entry.clientUpdatedAt }); continue; }
      if (goal.updatedAt !== entry.clientUpdatedAt) await repository.markSyncOutbox([syncOutboxEntry("goal", goal.id, goal.updatedAt)]);
      changes.push({ id: goal.id, entityType: "goal", deleted: false, goal, clientUpdatedAt: goal.updatedAt });
      continue;
    }
    const book = await repository.getBook(entry.id);
    if (!book) { staleAcknowledgements.push({ id: entry.id, entityType: "book", deleted: false, clientUpdatedAt: entry.clientUpdatedAt }); continue; }
    if (book.updatedAt !== entry.clientUpdatedAt) await repository.markSyncOutbox([syncOutboxEntry("book", book.id, book.updatedAt)]);
    changes.push({ id: book.id, entityType: "book", deleted: false, book: syncableBook(book), clientUpdatedAt: book.updatedAt });
  }

  // These entries point at records that no longer exist and have no tombstone. They can
  // only be stale local bookkeeping, so remove them without involving the network.
  if (staleAcknowledgements.length) await repository.acknowledgeSyncChanges(staleAcknowledgements);
  return changes;
}

async function applySyncResponse(repository: LibraryRepository, response: SyncResponse): Promise<void> {
  const [bookTombstones, shelfTombstones, goalTombstones] = await Promise.all([
    repository.listTombstones(), repository.listShelfTombstones(), repository.listGoalTombstones()
  ]);
  const deletedBooks = new Map(bookTombstones.map((item) => [item.id, item.deletedAt]));
  const deletedShelves = new Map(shelfTombstones.map((item) => [item.id, item.deletedAt]));
  const deletedGoals = new Map(goalTombstones.map((item) => [item.id, item.deletedAt]));

  for (const record of response.changes) {
    const entityType = record.entityType ?? "book";
    if (entityType === "shelf") {
      if (isNewerThan(deletedShelves.get(record.id), record.clientUpdatedAt)) continue;
      const local = await repository.getShelf(record.id);
      if (local && isNewerThan(local.updatedAt, record.clientUpdatedAt)) continue;
      if (record.deleted) await repository.deleteShelf(record.id, false);
      else if (record.shelf) await repository.putShelf(record.shelf, false);
      continue;
    }
    if (entityType === "goal") {
      if (isNewerThan(deletedGoals.get(record.id), record.clientUpdatedAt)) continue;
      const local = await repository.getGoal(record.id);
      if (local && isNewerThan(local.updatedAt, record.clientUpdatedAt)) continue;
      if (record.deleted) await repository.deleteGoal(record.id, false);
      else if (record.goal) await repository.putGoal(record.goal, false);
      continue;
    }

    // A user can save or delete while this network request is in flight. Never let an
    // older server response overwrite that newer local action; the outbox/tombstone
    // remains queued for the next batch.
    if (isNewerThan(deletedBooks.get(record.id), record.clientUpdatedAt)) continue;
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
      await repository.putBook({ ...record.book, shelfIds: record.book.shelfIds ?? [], cachedCoverDataUrl }, false);
    }
  }
}

function splitIntoBatches(changes: SyncMutation[], cursor?: string): SyncMutation[][] {
  const batches: SyncMutation[][] = [];
  let current: SyncMutation[] = [];

  for (const change of changes) {
    const candidate = [...current, change];
    const exceedsCount = candidate.length > MAX_BATCH_RECORDS;
    const exceedsBytes = current.length > 0 && requestBytes(cursor, candidate) > MAX_BATCH_BYTES;
    if (exceedsCount || exceedsBytes) {
      batches.push(current);
      current = [change];
    } else {
      current = candidate;
    }
  }
  if (current.length) batches.push(current);
  return batches;
}

function requestBytes(cursor: string | undefined, changes: SyncMutation[]): number {
  return new TextEncoder().encode(JSON.stringify({ cursor, changes })).byteLength;
}

function mutationKey(change: SyncMutation): string {
  return `${change.entityType ?? "book"}:${change.id}:${change.deleted ? "delete" : "put"}`;
}

function syncableBook(book: Book): Book {
  const { cachedCoverDataUrl: _localCache, ...rest } = book;
  return rest as Book;
}

export async function resetSyncCursor(repository: LibraryRepository, accountId: string): Promise<void> {
  await Promise.all([
    repository.setMeta(`cloudSyncCursor:${accountId}`, EPOCH),
    repository.setMeta(`cloudSyncOutboxInitialized:${accountId}`, "0"),
    repository.setMeta(`lastSuccessfulSync:${accountId}`, EPOCH)
  ]);
}

function validTimestamp(value?: string): string | undefined {
  return value && Number.isFinite(Date.parse(value)) ? value : undefined;
}

function isAfter(value: string | undefined, baseline: string): boolean {
  if (!value) return false;
  const parsed = Date.parse(value);
  const cutoff = Date.parse(baseline);
  return Number.isFinite(parsed) && Number.isFinite(cutoff) && parsed > cutoff;
}

function isNewerThan(localUpdatedAt?: string, remoteUpdatedAt?: string): boolean {
  if (!localUpdatedAt || !remoteUpdatedAt) return false;
  const local = Date.parse(localUpdatedAt);
  const remote = Date.parse(remoteUpdatedAt);
  return Number.isFinite(local) && Number.isFinite(remote) && local > remote;
}

// Kept separate for future targeted tests and to make the conflict rule explicit.
export function syncRecordKey(record: Pick<SyncRecord, "entityType" | "id">): string {
  return `${record.entityType ?? "book"}:${record.id}`;
}
