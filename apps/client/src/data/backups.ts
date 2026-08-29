import type { Book, ReadingGoal, Shelf } from "@bookstats/domain";
import type { LibraryRepository } from "./libraryRepository";

const LOCAL_BACKUPS_KEY = "localBackups:v1";
const DAILY_BACKUP_KEY = "lastAutomaticBackupDate:v1";
const MAX_LOCAL_BACKUPS = 5;

export interface BackupPreferences {
  layout?: "grid" | "table";
  visibleColumns?: string[];
}

export interface BookStatsBackup {
  format: "bookstats-backup";
  backupVersion: 1;
  appVersion: string;
  createdAt: string;
  reason: "manual" | "automatic" | "before-import" | "before-import-undo" | "before-merge" | "before-bulk-delete" | "before-restore";
  books: Book[];
  shelves: Shelf[];
  goals: ReadingGoal[];
  preferences?: BackupPreferences;
}

export function stripLocalCoverCache(book: Book): Book {
  const { cachedCoverDataUrl: _cache, ...portable } = book;
  return portable;
}

export async function createBackupSnapshot(
  repository: LibraryRepository,
  appVersion: string,
  reason: BookStatsBackup["reason"],
  preferences?: BackupPreferences,
  keepLocally = true
): Promise<BookStatsBackup> {
  const [books, shelves, goals] = await Promise.all([repository.listBooks(), repository.listShelves(), repository.listGoals()]);
  const backup: BookStatsBackup = {
    format: "bookstats-backup", backupVersion: 1, appVersion, createdAt: new Date().toISOString(), reason,
    books: books.map(stripLocalCoverCache), shelves, goals, preferences
  };
  if (keepLocally) {
    const backups = await listLocalBackups(repository);
    await repository.setMeta(LOCAL_BACKUPS_KEY, JSON.stringify([backup, ...backups].slice(0, MAX_LOCAL_BACKUPS)));
  }
  return backup;
}

export async function createDailyBackupIfNeeded(repository: LibraryRepository, appVersion: string, preferences?: BackupPreferences): Promise<void> {
  const today = new Date().toISOString().slice(0, 10);
  if ((await repository.getMeta(DAILY_BACKUP_KEY)) === today) return;
  const books = await repository.listBooks();
  if (books.length === 0) return;
  await createBackupSnapshot(repository, appVersion, "automatic", preferences, true);
  await repository.setMeta(DAILY_BACKUP_KEY, today);
}

export async function listLocalBackups(repository: LibraryRepository): Promise<BookStatsBackup[]> {
  try {
    const parsed = JSON.parse((await repository.getMeta(LOCAL_BACKUPS_KEY)) ?? "[]") as BookStatsBackup[];
    return Array.isArray(parsed) ? parsed.filter(isBackup) : [];
  } catch { return []; }
}

export async function removeLocalBackup(repository: LibraryRepository, createdAt: string): Promise<void> {
  const backups = await listLocalBackups(repository);
  await repository.setMeta(LOCAL_BACKUPS_KEY, JSON.stringify(backups.filter((backup) => backup.createdAt !== createdAt)));
}

export function parseBackup(text: string): BookStatsBackup {
  const parsed = JSON.parse(text) as unknown;
  if (!isBackup(parsed)) throw new Error("This file is not a BookStats recovery backup.");
  return parsed;
}

export async function restoreBackup(repository: LibraryRepository, backup: BookStatsBackup): Promise<void> {
  const [currentBooks, currentShelves, currentGoals] = await Promise.all([repository.listBooks(), repository.listShelves(), repository.listGoals()]);
  const bookIds = new Set(backup.books.map((book) => book.id));
  const shelfIds = new Set(backup.shelves.map((shelf) => shelf.id));
  const goalIds = new Set(backup.goals.map((goal) => goal.id));
  for (const book of currentBooks) if (!bookIds.has(book.id)) await repository.deleteBook(book.id);
  for (const goal of currentGoals) if (!goalIds.has(goal.id)) await repository.deleteGoal(goal.id);
  for (const shelf of currentShelves) if (!shelfIds.has(shelf.id)) await repository.deleteShelf(shelf.id);
  const now = new Date().toISOString();
  if (backup.shelves.length) await repository.bulkPutShelves(backup.shelves.map((shelf) => ({ ...shelf, updatedAt: now })));
  if (backup.goals.length) await repository.bulkPutGoals(backup.goals.map((goal) => ({ ...goal, updatedAt: now })));
  if (backup.books.length) await repository.bulkPutBooks(backup.books.map((book) => ({ ...book, cachedCoverDataUrl: undefined, updatedAt: now })));
}

export function downloadBackup(backup: BookStatsBackup): void {
  const url = URL.createObjectURL(new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" }));
  const anchor = document.createElement("a"); anchor.href = url; anchor.download = `bookstats-backup-${backup.createdAt.slice(0, 10)}.json`; anchor.click(); URL.revokeObjectURL(url);
}

function isBackup(value: unknown): value is BookStatsBackup {
  if (!value || typeof value !== "object") return false;
  const backup = value as Partial<BookStatsBackup>;
  return backup.format === "bookstats-backup" && backup.backupVersion === 1 && Array.isArray(backup.books) && Array.isArray(backup.shelves) && Array.isArray(backup.goals);
}
