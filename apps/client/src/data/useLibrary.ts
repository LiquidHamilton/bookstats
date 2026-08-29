import { useCallback, useEffect, useRef, useState } from "react";
import type { Book, ReadingGoal, Shelf } from "@bookstats/domain";
import { sortShelves } from "@bookstats/domain";
import { getLibraryRepository, type LibraryRepository } from "./libraryRepository";

export function useLibrary() {
  const [repository, setRepository] = useState<LibraryRepository>();
  const [books, setBooks] = useState<Book[]>([]);
  const [shelves, setShelves] = useState<Shelf[]>([]);
  const [goals, setGoals] = useState<ReadingGoal[]>([]);
  const [loading, setLoading] = useState(true);
  // A background sync can trigger a full refresh at the same time a user saves.
  // Track local writes on both sides of their async repository operation so a
  // refresh never commits a snapshot that was captured mid-save.
  const localWriteRevision = useRef(0);

  const refresh = useCallback(async (repo?: LibraryRepository) => {
    const active = repo ?? repository ?? await getLibraryRepository();
    setRepository(active);
    while (true) {
      const revision = localWriteRevision.current;
      const [nextBooks, nextShelves, nextGoals] = await Promise.all([active.listBooks(), active.listShelves(), active.listGoals()]);
      if (revision !== localWriteRevision.current) continue;
      setBooks(nextBooks);
      setShelves(sortShelves(nextShelves));
      setGoals(nextGoals);
      setLoading(false);
      return;
    }
  }, [repository]);

  useEffect(() => { void refresh(); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const putBook = useCallback(async (book: Book) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.putBook(book);
      // Do not reread a multi-thousand-record browser database after every edit.
      // The repository write is already complete, so update the in-memory snapshot directly.
      setBooks((current) => {
        const index = current.findIndex((item) => item.id === book.id);
        if (index < 0) return [...current, book];
        const next = current.slice(); next[index] = book; return next;
      });
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const bulkPutBooks = useCallback(async (items: Book[]) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.bulkPutBooks(items);
      const updates = new Map(items.map((book) => [book.id, book]));
      setBooks((current) => {
        const seen = new Set<string>();
        const next = current.map((book) => { const replacement = updates.get(book.id); if (replacement) { seen.add(book.id); return replacement; } return book; });
        for (const book of items) if (!seen.has(book.id)) next.push(book);
        return next;
      });
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const deleteBook = useCallback(async (id: string, trackTombstone = true) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.deleteBook(id, trackTombstone);
      setBooks((current) => current.filter((book) => book.id !== id));
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const putShelf = useCallback(async (shelf: Shelf) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.putShelf(shelf);
      setShelves((current) => {
        const index = current.findIndex((item) => item.id === shelf.id);
        const next = index < 0 ? [...current, shelf] : current.map((item) => item.id === shelf.id ? shelf : item);
        return sortShelves(next);
      });
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const bulkPutShelves = useCallback(async (items: Shelf[]) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.bulkPutShelves(items);
      const updates = new Map(items.map((shelf) => [shelf.id, shelf]));
      setShelves((current) => {
        const seen = new Set<string>();
        const next = current.map((shelf) => { const replacement = updates.get(shelf.id); if (replacement) { seen.add(shelf.id); return replacement; } return shelf; });
        for (const shelf of items) if (!seen.has(shelf.id)) next.push(shelf);
        return sortShelves(next);
      });
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const deleteShelf = useCallback(async (id: string, trackTombstone = true) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.deleteShelf(id, trackTombstone);
      setShelves((current) => current.filter((shelf) => shelf.id !== id));
      setBooks((current) => current.map((book) => (book.shelfIds ?? []).includes(id) ? { ...book, shelfIds: book.shelfIds.filter((shelfId) => shelfId !== id) } : book));
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const putGoal = useCallback(async (goal: ReadingGoal) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.putGoal(goal);
      setGoals((current) => { const index = current.findIndex((item) => item.id === goal.id); return index < 0 ? [...current, goal] : current.map((item) => item.id === goal.id ? goal : item); });
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  const deleteGoal = useCallback(async (id: string, trackTombstone = true) => {
    const active = repository ?? await getLibraryRepository();
    localWriteRevision.current += 1;
    try {
      await active.deleteGoal(id, trackTombstone);
      setGoals((current) => current.filter((goal) => goal.id !== id));
    } finally { localWriteRevision.current += 1; }
  }, [repository]);

  return {
    books, shelves, goals, loading, repository, storageKind: repository?.kind, refresh,
    putBook, bulkPutBooks, deleteBook, putShelf, bulkPutShelves, deleteShelf, putGoal, deleteGoal
  };
}
