import type { Book, ReadingGoal, Shelf } from "@bookstats/domain";
import { isSmartShelf } from "@bookstats/domain";
import type { LibraryRepository } from "./libraryRepository";
import type { ExternalImportResult, ImportedBook } from "./importers";
import { bookMatchKeys, mergeImportedBook } from "./importers";

const HISTORY_KEY = "importHistory:v1";
const MAX_BATCHES = 8;

export type ImportSource = "goodreads" | "librarything" | "bookstats";
export interface ImportChange<T> {
  id: string;
  before?: T;
  after: T;
}
export interface ImportBatch {
  id: string;
  source: ImportSource;
  sourceName: string;
  createdAt: string;
  bookChanges: ImportChange<Book>[];
  shelfChanges: ImportChange<Shelf>[];
  goalChanges: ImportChange<ReadingGoal>[];
  createdBooks: number;
  updatedBooks: number;
  ambiguousBooks: number;
  createdShelves: number;
  warnings: string[];
}

export interface ExternalImportPlan {
  source: ImportSource;
  sourceName: string;
  total: number;
  createdBooks: number;
  updatedBooks: number;
  ambiguousBooks: number;
  createdShelves: number;
  warnings: string[];
  bookChanges: ImportChange<Book>[];
  shelfChanges: ImportChange<Shelf>[];
  goalChanges: ImportChange<ReadingGoal>[];
  preview: Array<{ title: string; author: string; action: "new" | "update" | "ambiguous-new"; matchedTitle?: string }>;
}

export function buildExternalImportPlan(result: ExternalImportResult, currentBooks: Book[], currentShelves: Shelf[]): ExternalImportPlan {
  const shelfByName = new Map(currentShelves.map((shelf) => [shelf.name.toLocaleLowerCase(), shelf]));
  const createdShelves: Shelf[] = [];
  const finalBooks = new Map(currentBooks.map((book) => [book.id, book]));
  const identityIndex = buildIdentityIndex(currentBooks);
  const now = new Date().toISOString();
  const changesById = new Map<string, ImportChange<Book>>();
  let createdBooks = 0;
  let updatedBooks = 0;
  let ambiguousBooks = 0;
  const preview: ExternalImportPlan["preview"] = [];

  for (const item of result.items) {
    const shelfIds = resolveShelfIds(item, shelfByName, createdShelves, now);
    const incoming = { ...item.book, shelfIds, updatedAt: now };
    const candidateIds = candidateIdsForBook(incoming, identityIndex, result.source);
    let existing: Book | undefined;
    let action: "new" | "update" | "ambiguous-new" = "new";
    if (candidateIds.size === 1) {
      const existingId = [...candidateIds][0];
      existing = finalBooks.get(existingId);
      action = existing ? "update" : "new";
    } else if (candidateIds.size > 1) {
      action = "ambiguous-new";
      ambiguousBooks += 1;
    }

    let finalBook: Book;
    if (existing) {
      finalBook = mergeImportedBook(existing, incoming);
      finalBooks.set(finalBook.id, finalBook);
      const previousChange = changesById.get(existing.id);
      changesById.set(existing.id, { id: existing.id, before: previousChange?.before ?? existing, after: finalBook });
      updatedBooks += previousChange ? 0 : 1;
      preview.push({ title: incoming.title, author: incoming.author, action, matchedTitle: existing.title });
    } else {
      finalBook = incoming;
      finalBooks.set(finalBook.id, finalBook);
      changesById.set(finalBook.id, { id: finalBook.id, after: finalBook });
      createdBooks += 1;
      preview.push({ title: incoming.title, author: incoming.author, action });
    }
    for (const key of bookMatchKeys(finalBook)) addIdentity(identityIndex, key, finalBook.id);
  }

  return {
    source: result.source,
    sourceName: result.source === "goodreads" ? "Goodreads" : "LibraryThing",
    total: result.items.length,
    createdBooks,
    updatedBooks,
    ambiguousBooks,
    createdShelves: createdShelves.length,
    warnings: result.warnings,
    bookChanges: [...changesById.values()],
    shelfChanges: createdShelves.map((shelf) => ({ id: shelf.id, after: shelf })),
    goalChanges: [],
    preview
  };
}

function resolveShelfIds(item: ImportedBook, shelfByName: Map<string, Shelf>, createdShelves: Shelf[], now: string): string[] {
  const ids: string[] = [];
  for (const originalName of item.shelfNames) {
    const shelfName = originalName.trim();
    if (!shelfName) continue;
    let key = shelfName.toLocaleLowerCase();
    let shelf = shelfByName.get(key);
    if (shelf && isSmartShelf(shelf)) {
      const importedName = `${shelfName} (Imported)`;
      key = importedName.toLocaleLowerCase();
      shelf = shelfByName.get(key);
      if (!shelf) {
        shelf = { id: crypto.randomUUID(), name: importedName, kind: "manual", createdAt: now, updatedAt: now };
        shelfByName.set(key, shelf); createdShelves.push(shelf);
      }
    }
    if (!shelf) {
      shelf = { id: crypto.randomUUID(), name: shelfName, kind: "manual", createdAt: now, updatedAt: now };
      shelfByName.set(key, shelf); createdShelves.push(shelf);
    }
    ids.push(shelf.id);
  }
  return ids;
}

function buildIdentityIndex(books: Book[]): Map<string, Set<string>> {
  const result = new Map<string, Set<string>>();
  for (const book of books) for (const key of bookMatchKeys(book)) addIdentity(result, key, book.id);
  return result;
}
function addIdentity(index: Map<string, Set<string>>, key: string, id: string): void {
  const ids = index.get(key) ?? new Set<string>(); ids.add(id); index.set(key, ids);
}
function candidateIdsForBook(book: Book, index: Map<string, Set<string>>, source: ExternalImportResult["source"]): Set<string> {
  const sourceKeys = bookMatchKeys(book).filter((key) => key.startsWith("source:"));
  for (const key of sourceKeys) {
    const ids = index.get(key);
    if (ids?.size) return new Set(ids);
  }
  // A LibraryThing books_id identifies one catalog entry/copy, not merely the work.
  // If that stable entry ID has never been imported before, keep it as a new BookStats
  // record even when another owned copy has the same ISBN or title + author. This keeps
  // collectors' separate editions/copies intact while repeated imports remain idempotent.
  if (source === "librarything") return new Set();
  const isbnKey = bookMatchKeys(book).find((key) => key.startsWith("isbn:"));
  if (isbnKey && index.get(isbnKey)?.size) return new Set(index.get(isbnKey));
  const titleKey = bookMatchKeys(book).find((key) => key.startsWith("title:"));
  return titleKey && index.get(titleKey)?.size ? new Set(index.get(titleKey)) : new Set();
}

export async function applyExternalImportPlan(repository: LibraryRepository, plan: ExternalImportPlan): Promise<ImportBatch> {
  if (plan.shelfChanges.length) await repository.bulkPutShelves(plan.shelfChanges.map((change) => change.after));
  if (plan.goalChanges.length) await repository.bulkPutGoals(plan.goalChanges.map((change) => change.after));
  if (plan.bookChanges.length) await repository.bulkPutBooks(plan.bookChanges.map((change) => change.after));
  const batch: ImportBatch = {
    id: crypto.randomUUID(), source: plan.source, sourceName: plan.sourceName, createdAt: new Date().toISOString(),
    bookChanges: plan.bookChanges, shelfChanges: plan.shelfChanges, goalChanges: plan.goalChanges,
    createdBooks: plan.createdBooks, updatedBooks: plan.updatedBooks, ambiguousBooks: plan.ambiguousBooks,
    createdShelves: plan.createdShelves, warnings: plan.warnings
  };
  await saveImportBatch(repository, batch);
  return batch;
}

export async function getImportHistory(repository: LibraryRepository): Promise<ImportBatch[]> {
  try { return JSON.parse((await repository.getMeta(HISTORY_KEY)) ?? "[]") as ImportBatch[]; } catch { return []; }
}

export async function saveImportBatch(repository: LibraryRepository, batch: ImportBatch): Promise<void> {
  const history = await getImportHistory(repository);
  await repository.setMeta(HISTORY_KEY, JSON.stringify([batch, ...history.filter((item) => item.id !== batch.id)].slice(0, MAX_BATCHES)));
}

export async function undoImportBatch(repository: LibraryRepository, batch: ImportBatch): Promise<{ restored: number; removed: number; skipped: number }> {
  let restored = 0; let removed = 0; let skipped = 0;
  const currentBooks = new Map((await repository.listBooks()).map((book) => [book.id, book]));
  const currentShelves = new Map((await repository.listShelves()).map((shelf) => [shelf.id, shelf]));
  const currentGoals = new Map((await repository.listGoals()).map((goal) => [goal.id, goal]));

  for (const change of [...batch.bookChanges].reverse()) {
    const current = currentBooks.get(change.id);
    if (!current || current.updatedAt !== change.after.updatedAt) { skipped += 1; continue; }
    if (change.before) { await repository.putBook({ ...change.before, updatedAt: new Date().toISOString() }); restored += 1; }
    else { await repository.deleteBook(change.id); removed += 1; }
  }
  for (const change of [...batch.shelfChanges].reverse()) {
    const current = currentShelves.get(change.id);
    if (!current || current.updatedAt !== change.after.updatedAt) { skipped += 1; continue; }
    if (change.before) { await repository.putShelf({ ...change.before, updatedAt: new Date().toISOString() }); restored += 1; }
    else { await repository.deleteShelf(change.id); removed += 1; }
  }
  for (const change of [...(batch.goalChanges ?? [])].reverse()) {
    const current = currentGoals.get(change.id);
    if (!current || current.updatedAt !== change.after.updatedAt) { skipped += 1; continue; }
    if (change.before) { await repository.putGoal({ ...change.before, updatedAt: new Date().toISOString() }); restored += 1; }
    else { await repository.deleteGoal(change.id); removed += 1; }
  }
  const history = await getImportHistory(repository);
  await repository.setMeta(HISTORY_KEY, JSON.stringify(history.filter((item) => item.id !== batch.id)));
  return { restored, removed, skipped };
}
