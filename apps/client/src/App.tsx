import { useDeferredValue, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import type { Book, ReadingGoal, ReadingStatus, SeriesCompletionOverride, Shelf, UserAccount } from "@bookstats/domain";
import { isSmartShelf, normalizedReadDates, READING_STATUS_LABELS, shelfMatchesBook } from "@bookstats/domain";
import { ArrowDown, ArrowUp, ArrowUpDown, BarChart3, BookOpen, CheckSquare2, Cloud, Columns3, Filter, FolderOpen, Grid2X2, Handshake, Library, List, LogIn, MessageSquareText, Plus, RefreshCw, ScanLine, Search, Shuffle, SlidersHorizontal, Sparkles, UserPlus, UserRound, Wrench, X } from "lucide-react";
import { checkServerCompatibility, currentAccount, deleteAccount as deleteCloudAccount, deleteCloudLibrary, logoutAccount, updateRequiredEventName } from "./data/api";
import { useLibrary } from "./data/useLibrary";
import { resetSyncCursor, synchronizeLibrary } from "./data/sync";
import { importGoodreadsCsv, importLibraryThingJson, mergeImportedBook } from "./data/importers";
import { mergeBooks } from "./data/cleanup";
import { applyExternalImportPlan, buildExternalImportPlan, getImportHistory, type ExternalImportPlan, type ImportBatch, undoImportBatch } from "./data/importHistory";
import { createBackupSnapshot, createDailyBackupIfNeeded, downloadBackup, listLocalBackups, parseBackup, removeLocalBackup, restoreBackup, type BookStatsBackup } from "./data/backups";
import { BookForm } from "./components/BookForm";
import { BookCard } from "./components/BookCard";
import { BookDetail } from "./components/BookDetail";
import { StatisticsView } from "./components/StatisticsView";
import { ToolsView } from "./components/ToolsView";
import { AccountView } from "./components/AccountView";
import { FeedbackView } from "./components/FeedbackView";
import { ShelfManager } from "./components/ShelfManager";
import { LibraryCleanup } from "./components/LibraryCleanup";
import { AdvancedFilterPanel, filterMatchesBook, type AdvancedFilterState } from "./components/AdvancedFilterPanel";
import { BulkEditPanel } from "./components/BulkEditPanel";
import { ImportPreview } from "./components/ImportPreview";
import { LendingView } from "./components/LendingView";
import { BarcodeScanner } from "./components/BarcodeScanner";
import { InstallAppPrompt } from "./components/InstallAppPrompt";
import { applyDesktopUpdate, isTauriRuntime, type DesktopUpdateProgress } from "./data/desktopUpdater";
import "./styles/app.css";

type View = "library" | "lending" | "statistics" | "tools" | "feedback" | "account";
type StatusFilter = "all" | ReadingStatus;
type SortKey = "title" | "author" | "series" | "status" | "shelves" | "rating" | "format" | "pages" | "reads" | "lastRead" | "owned";
type SortDirection = "asc" | "desc";
const TABLE_COLUMNS: Array<{ key: SortKey; label: string }> = [
  { key: "title", label: "Title" }, { key: "author", label: "Author" }, { key: "series", label: "Series" }, { key: "status", label: "Status" }, { key: "shelves", label: "Shelves" },
  { key: "rating", label: "Rating" }, { key: "format", label: "Format" }, { key: "pages", label: "Pages" }, { key: "reads", label: "Reads" }, { key: "lastRead", label: "Last read" }, { key: "owned", label: "Owned" }
];
const DEFAULT_VISIBLE_COLUMNS = TABLE_COLUMNS.map((column) => column.key);
const COLUMN_STORAGE_KEY = "bookstats.visibleTableColumns";
const LAYOUT_STORAGE_KEY = "bookstats.libraryLayout";
const LIBRARY_RENDER_BATCH = 120;
const ACCOUNT_WELCOME_STORAGE_KEY = "bookstats.accountWelcomeDismissed.v1";

export default function App() {
  const {
    books, shelves, goals, loading, repository, storageKind, refresh, putBook, bulkPutBooks, deleteBook: deleteLocalBook,
    putShelf, bulkPutShelves, deleteShelf: deleteLocalShelf, putGoal, deleteGoal: deleteLocalGoal
  } = useLibrary();
  const [view, setView] = useState<View>(() => {
    if (typeof window !== "undefined") { const params = new URLSearchParams(window.location.search); if (params.has("verify") || params.has("reset")) return "account"; }
    return "library";
  });
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState<StatusFilter>("all");
  const [shelfFilter, setShelfFilter] = useState<string>("all");
  const [advancedFilter, setAdvancedFilter] = useState<AdvancedFilterState>();
  const [filterOpen, setFilterOpen] = useState(false);
  const [editing, setEditing] = useState<Book | null | undefined>(undefined);
  const [selectedBookId, setSelectedBookId] = useState<string>();
  const [managingShelves, setManagingShelves] = useState(false);
  const [cleanupOpen, setCleanupOpen] = useState(false);
  const [randomBook, setRandomBook] = useState<Book | null>(null);
  const [layout, setLayout] = useState<"grid" | "table">(loadLayout);
  const [visibleColumns, setVisibleColumns] = useState<SortKey[]>(loadVisibleColumns);
  const [sortKey, setSortKey] = useState<SortKey>("author");
  const [sortDirection, setSortDirection] = useState<SortDirection>("asc");
  const [selectionMode, setSelectionMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkOpen, setBulkOpen] = useState(false);
  const [importPlan, setImportPlan] = useState<ExternalImportPlan>();
  const [backups, setBackups] = useState<BookStatsBackup[]>([]);
  const [importHistory, setImportHistory] = useState<ImportBatch[]>([]);
  const [account, setAccount] = useState<UserAccount | null>(null);
  const [accountReady, setAccountReady] = useState(false);
  const [accountWelcomeOpen, setAccountWelcomeOpen] = useState(false);
  const [accountEntryMode, setAccountEntryMode] = useState<"login" | "register">("login");
  const [syncing, setSyncing] = useState(false);
  const [lastSync, setLastSync] = useState<string>();
  const [syncError, setSyncError] = useState<string>();
  const [renderLimit, setRenderLimit] = useState(LIBRARY_RENDER_BATCH);
  const [scannerOpen, setScannerOpen] = useState(false);
  const [newBookIsbn, setNewBookIsbn] = useState<string>();
  const [serverUpdateVersion, setServerUpdateVersion] = useState<string>();
  const [desktopUpdateProgress, setDesktopUpdateProgress] = useState<DesktopUpdateProgress>({ phase: "idle" });
  const [desktopUpdateError, setDesktopUpdateError] = useState<string>();
  const syncInFlight = useRef(false);
  const syncQueued = useRef(false);

  useEffect(() => {
    const eventName = updateRequiredEventName();
    const onUpdateRequired = (event: Event) => {
      const detail = (event as CustomEvent<{ serverVersion?: string }>).detail;
      setServerUpdateVersion(detail?.serverVersion || "newer");
    };
    const check = () => { void checkServerCompatibility(); };
    window.addEventListener(eventName, onUpdateRequired);
    window.addEventListener("focus", check);
    const interval = window.setInterval(check, 5 * 60 * 1000);
    check();
    return () => {
      window.removeEventListener(eventName, onUpdateRequired);
      window.removeEventListener("focus", check);
      window.clearInterval(interval);
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    void currentAccount().then((current) => {
      if (cancelled) return;
      setAccount(current);
      if (current) return;
      const params = new URLSearchParams(window.location.search);
      const handlingAccountLink = params.has("verify") || params.has("reset");
      if (!handlingAccountLink && localStorage.getItem(ACCOUNT_WELCOME_STORAGE_KEY) !== "1") setAccountWelcomeOpen(true);
    }).finally(() => { if (!cancelled) setAccountReady(true); });
    return () => { cancelled = true; };
  }, []);
  useEffect(() => {
    if (!repository || !account) { setLastSync(undefined); return; }
    void repository.getMeta(`lastSuccessfulSync:${account.id}`).then(setLastSync);
  }, [repository, account]);
  useEffect(() => { if (repository && account?.emailVerified) void performSync().catch(() => undefined); }, [repository, account?.id, account?.emailVerified]);
  useEffect(() => { localStorage.setItem(COLUMN_STORAGE_KEY, JSON.stringify(visibleColumns)); }, [visibleColumns]);
  useEffect(() => { localStorage.setItem(LAYOUT_STORAGE_KEY, layout); }, [layout]);
  useEffect(() => {
    if (!repository || loading) return;
    void createDailyBackupIfNeeded(repository, __BOOKSTATS_VERSION__, backupPreferences()).then(() => reloadSafetyData(repository)).catch(() => undefined);
  }, [repository, loading, books.length]); // one lightweight daily safety snapshot per device
  useEffect(() => { if (repository && view === "tools") void reloadSafetyData(repository); }, [repository, view]);

  async function installDesktopUpdate(): Promise<void> {
    setDesktopUpdateError(undefined);
    try {
      await applyDesktopUpdate(setDesktopUpdateProgress);
    } catch (error) {
      console.error("BookStats desktop update failed", error);
      setDesktopUpdateProgress({ phase: "error" });
      setDesktopUpdateError(error instanceof Error ? error.message : "BookStats could not apply the desktop update.");
    }
  }

  const selectedBook = selectedBookId ? books.find((book) => book.id === selectedBookId) : undefined;
  const selectedShelf = shelfFilter === "all" ? undefined : shelves.find((shelf) => shelf.id === shelfFilter);
  const selectedBooks = useMemo(() => books.filter((book) => selectedIds.has(book.id)), [books, selectedIds]);
  const deferredQuery = useDeferredValue(query);
  const shelfIndex = useMemo(() => {
    const namesByBook = new Map<string, string[]>();
    const counts = new Map<string, number>(shelves.map((shelf) => [shelf.id, 0]));
    for (const book of books) {
      const names: string[] = [];
      for (const shelf of shelves) {
        if (!shelfMatchesBook(shelf, book)) continue;
        names.push(shelf.name);
        counts.set(shelf.id, (counts.get(shelf.id) ?? 0) + 1);
      }
      namesByBook.set(book.id, names);
    }
    return { namesByBook, counts };
  }, [books, shelves]);
  const shelfNamesByBook = shelfIndex.namesByBook;
  const shelfCounts = shelfIndex.counts;
  const filtered = useMemo(() => {
    const needle = deferredQuery.trim().toLowerCase();
    const matches = books.filter((book) => {
      if (statusFilter !== "all" && book.status !== statusFilter) return false;
      if (selectedShelf && !shelfMatchesBook(selectedShelf, book)) return false;
      if (!filterMatchesBook(advancedFilter, book)) return false;
      if (!needle) return true;
      const searchable = [book.title, book.author, book.isbn ?? "", book.series ?? "", book.seriesVolume ?? "", book.publisher ?? "", book.genre ?? "", book.condition ?? "", ...book.tags, ...(shelfNamesByBook.get(book.id) ?? [])].join(" ").toLowerCase();
      return searchable.includes(needle);
    });
    return [...matches].sort((a, b) => compareBooks(a, b, sortKey, sortDirection, shelfNamesByBook));
  }, [books, deferredQuery, statusFilter, selectedShelf, sortKey, sortDirection, advancedFilter, shelfNamesByBook]);
  const renderedBooks = useMemo(() => filtered.slice(0, renderLimit), [filtered, renderLimit]);
  const detailNavigation = useMemo(() => {
    if (!selectedBook) return { previous: undefined as Book | undefined, next: undefined as Book | undefined };
    const filteredIndex = filtered.findIndex((item) => item.id === selectedBook.id);
    const ordered = filteredIndex >= 0 ? filtered : [...books].sort((a, b) => compareBooks(a, b, sortKey, sortDirection, shelfNamesByBook));
    const index = filteredIndex >= 0 ? filteredIndex : ordered.findIndex((item) => item.id === selectedBook.id);
    return {
      previous: index > 0 ? ordered[index - 1] : undefined,
      next: index >= 0 && index < ordered.length - 1 ? ordered[index + 1] : undefined
    };
  }, [selectedBook, filtered, books, sortKey, sortDirection, shelfNamesByBook]);

  useEffect(() => { setRenderLimit(LIBRARY_RENDER_BATCH); }, [deferredQuery, statusFilter, shelfFilter, advancedFilter, sortKey, sortDirection, layout]);

  async function reloadSafetyData(repo = repository) {
    if (!repo) return;
    const [nextBackups, nextImports] = await Promise.all([listLocalBackups(repo), getImportHistory(repo)]);
    setBackups(nextBackups); setImportHistory(nextImports);
  }
  function backupPreferences() { return { layout, visibleColumns }; }
  async function safetyBackup(reason: BookStatsBackup["reason"]) {
    if (!repository || books.length === 0) return;
    await createBackupSnapshot(repository, __BOOKSTATS_VERSION__, reason, backupPreferences(), true);
    await reloadSafetyData(repository);
  }

  async function performSync() {
    if (!repository || !account || !account.emailVerified) return;
    if (syncInFlight.current) { syncQueued.current = true; return; }
    syncInFlight.current = true; setSyncing(true); setSyncError(undefined);
    try {
      do {
        syncQueued.current = false;
        const result = await synchronizeLibrary(repository, account.id);
        const completedAt = new Date().toISOString();
        await repository.setMeta(`lastSuccessfulSync:${account.id}`, completedAt);
        setLastSync(completedAt);
        // Local saves already updated React state. Re-read storage only when cloud
        // records were reconciled, keeping the common one-book edit path lightweight.
        if (result.pulled > 0) await refresh(repository);
      } while (syncQueued.current);
    } catch (error) { setSyncError(error instanceof Error ? error.message : "Synchronization failed."); throw error; }
    finally { syncInFlight.current = false; setSyncing(false); }
  }

  async function saveBook(book: Book) { await putBook({ ...book, shelfIds: book.shelfIds ?? [] }); if (account?.emailVerified) void performSync().catch(() => undefined); }
  async function deleteBook(book: Book) {
    if (!window.confirm(`Remove “${book.title}” from your library?`)) return;
    await deleteLocalBook(book.id); setSelectedIds((ids) => { const next = new Set(ids); next.delete(book.id); return next; });
    if (selectedBookId === book.id) setSelectedBookId(undefined); if (account?.emailVerified) void performSync().catch(() => undefined);
  }
  async function saveGoal(goal: ReadingGoal) { await putGoal(goal); if (account?.emailVerified) void performSync().catch(() => undefined); }
  async function deleteGoal(goal: ReadingGoal) { await deleteLocalGoal(goal.id); if (account?.emailVerified) void performSync().catch(() => undefined); }

  async function createShelf(name: string, options: Partial<Pick<Shelf, "kind" | "match" | "rules" | "ruleGroups">> = {}): Promise<Shelf> {
    const normalized = name.trim();
    const existing = shelves.find((shelf) => shelf.name.localeCompare(normalized, undefined, { sensitivity: "base" }) === 0);
    if (existing) { const requestedKind = options.kind ?? "manual"; if (requestedKind === "manual" && !isSmartShelf(existing)) return existing; throw new Error(`A shelf named “${existing.name}” already exists.`); }
    const now = new Date().toISOString();
    const nextOrder = shelves.reduce((highest, shelf) => typeof shelf.order === "number" && Number.isFinite(shelf.order) ? Math.max(highest, shelf.order) : highest, -1) + 1;
    const shelf: Shelf = { id: crypto.randomUUID(), name: normalized, order: nextOrder, kind: options.kind ?? "manual", match: options.match, rules: options.rules, ruleGroups: options.ruleGroups, createdAt: now, updatedAt: now };
    await putShelf(shelf); if (account?.emailVerified) void performSync().catch(() => undefined); return shelf;
  }
  async function updateShelf(shelf: Shelf) { const duplicate = shelves.find((item) => item.id !== shelf.id && item.name.localeCompare(shelf.name, undefined, { sensitivity: "base" }) === 0); if (duplicate) throw new Error(`A shelf named “${duplicate.name}” already exists.`); await putShelf(shelf); if (account?.emailVerified) void performSync().catch(() => undefined); }
  async function reorderShelves(ordered: Shelf[]) {
    const now = new Date().toISOString();
    await bulkPutShelves(ordered.map((shelf, index) => ({ ...shelf, order: index, updatedAt: now })));
    if (account?.emailVerified) void performSync().catch(() => undefined);
  }
  async function removeShelf(shelf: Shelf) {
    const count = books.filter((book) => shelfMatchesBook(shelf, book)).length;
    const assignmentText = isSmartShelf(shelf) ? " Its rules will be removed; no books will be changed." : `${count ? ` It will be removed from ${count} ${count === 1 ? "book" : "books"}.` : ""} The books themselves will not be deleted.`;
    if (!window.confirm(`Delete the “${shelf.name}” ${isSmartShelf(shelf) ? "smart " : ""}shelf?${assignmentText}`)) return;
    await deleteLocalShelf(shelf.id); if (shelfFilter === shelf.id) setShelfFilter("all"); if (account?.emailVerified) void performSync().catch(() => undefined);
  }
  async function saveFilterAsShelf(name: string, filter: AdvancedFilterState) { await createShelf(name, { kind: "smart", match: filter.match, ruleGroups: filter.ruleGroups.map((group) => ({ ...group, rules: group.rules.map((rule) => ({ ...rule })) })) }); }

  async function mergeDuplicateRecords(keep: Book, remove: Book[]) {
    if (!repository || remove.length === 0) return;
    await safetyBackup("before-merge");
    const merged = mergeBooks(keep, remove); await repository.putBook(merged); for (const book of remove) await repository.deleteBook(book.id); await refresh(repository);
    if (selectedBookId && remove.some((book) => book.id === selectedBookId)) setSelectedBookId(merged.id); if (account?.emailVerified) void performSync().catch(() => undefined);
  }

  async function signOutAccount() {
    if (!account) return; const signedOutAccount = account;
    if (repository?.kind === "indexeddb") {
      if (signedOutAccount.emailVerified) { try { await performSync(); } catch { if (!window.confirm("BookStats could not complete a final cloud sync. Signing out will clear this browser's local cache. Sign out anyway?")) return; } }
      else if (books.length > 0 && !window.confirm("This account is not verified, so these browser books may not be backed up to the cloud. Signing out clears the browser cache. Sign out anyway?")) return;
    }
    await logoutAccount(); setAccount(null); setSyncError(undefined); setLastSync(undefined);
    if (repository?.kind === "indexeddb") { await repository.clearLibraryData(); await resetSyncCursor(repository, signedOutAccount.id); await refresh(repository); }
  }
  async function removeCloudCopyAndDisconnect() {
    if (!account || !repository) return; const old = account;
    await deleteCloudLibrary(); await logoutAccount(); setAccount(null); setLastSync(undefined); setSyncError(undefined); await resetSyncCursor(repository, old.id);
    if (repository.kind === "indexeddb") { await repository.clearLibraryData(); await refresh(repository); }
  }
  async function removeAccount(password: string) {
    if (!account || !repository) return; const old = account;
    await deleteCloudAccount(password); setAccount(null); setLastSync(undefined); setSyncError(undefined); await resetSyncCursor(repository, old.id);
    if (repository.kind === "indexeddb") { await repository.clearLibraryData(); await refresh(repository); }
  }

  function pickRandom() { const candidates = filtered.filter((book) => book.status !== "read"); setRandomBook(candidates.length ? candidates[Math.floor(Math.random() * candidates.length)] : null); }
  function changeSort(key: SortKey) { if (key === sortKey) setSortDirection((direction) => direction === "asc" ? "desc" : "asc"); else { setSortKey(key); setSortDirection(["title", "author", "series", "shelves"].includes(key) ? "asc" : "desc"); } }
  function toggleColumn(key: SortKey) { if (key === "title") return; setVisibleColumns((columns) => columns.includes(key) ? columns.filter((column) => column !== key) : DEFAULT_VISIBLE_COLUMNS.filter((column) => column === "title" || columns.includes(column) || column === key)); }
  function editFromDetail(book: Book) { setSelectedBookId(undefined); setEditing(book); }
  function toggleSelected(book: Book, selected: boolean) { setSelectedIds((ids) => { const next = new Set(ids); if (selected) next.add(book.id); else next.delete(book.id); return next; }); }
  function exitSelectionMode() { setSelectionMode(false); setSelectedIds(new Set()); setBulkOpen(false); }

  function exportBooks(items: Book[], includeGoals = false, filename = "bookstats-library") {
    const portableBooks = items.map(({ cachedCoverDataUrl: _cache, ...book }) => book);
    const referencedShelfIds = new Set(items.flatMap((book) => book.shelfIds ?? []));
    const portableShelves = includeGoals ? shelves : shelves.filter((shelf) => isSmartShelf(shelf) || referencedShelfIds.has(shelf.id));
    const payload = JSON.stringify({ format: "bookstats", version: 12, exportedAt: new Date().toISOString(), shelves: portableShelves, goals: includeGoals ? goals : [], books: portableBooks }, null, 2);
    const url = URL.createObjectURL(new Blob([payload], { type: "application/json" })); const anchor = document.createElement("a"); anchor.href = url; anchor.download = `${filename}-${new Date().toISOString().slice(0, 10)}.json`; anchor.click(); URL.revokeObjectURL(url);
  }
  function exportLibrary() { exportBooks(books, true, "bookstats-library"); }

  async function importLibrary(file: File) {
    try {
      if (!repository) throw new Error("Your local library is still opening. Try again in a moment.");
      const parsed = JSON.parse(await file.text()) as { format?: string; shelves?: Partial<Shelf>[]; goals?: Partial<ReadingGoal>[]; books?: Partial<Book>[] } | Partial<Book>[];
      const imported = Array.isArray(parsed) ? parsed : parsed.books; if (!Array.isArray(imported)) throw new Error("No books array found");
      const now = new Date().toISOString();
      const currentBooks = new Map(books.map((book) => [book.id, book])); const currentShelves = new Map(shelves.map((shelf) => [shelf.id, shelf])); const currentGoals = new Map(goals.map((goal) => [goal.id, goal]));
      const shelfPlan = planBookStatsShelves(Array.isArray((parsed as { shelves?: Partial<Shelf>[] }).shelves) ? (parsed as { shelves: Partial<Shelf>[] }).shelves : [], shelves, now);
      const bookChanges = imported.map((book) => {
        const normalizedShelfIds = [...new Set((book.shelfIds ?? []).map((id) => shelfPlan.idMap.get(id) ?? id).filter((id) => shelfPlan.validIds.has(id)))];
        const normalized: Book = { ...book, id: book.id || crypto.randomUUID(), title: book.title?.trim() || "Untitled", author: book.author?.trim() || "Unknown author", additionalAuthors: book.additionalAuthors ?? [], status: book.status ?? "not_started", owned: book.owned ?? false, shelfIds: normalizedShelfIds, tags: book.tags ?? [], readingSessions: Array.isArray(book.readingSessions) ? book.readingSessions : undefined, readDates: Array.isArray(book.readDates) ? [...new Set(book.readDates.filter(Boolean))].sort() : (book.dateRead ? [book.dateRead] : []), dateRead: undefined, dateAdded: book.dateAdded ?? now, createdAt: book.createdAt ?? now, updatedAt: now };
        const before = currentBooks.get(normalized.id); return { id: normalized.id, before, after: before ? mergeImportedBook(before, normalized) : normalized };
      });
      const shelfChanges = shelfPlan.changes;
      const importedGoals = Array.isArray((parsed as { goals?: Partial<ReadingGoal>[] }).goals) ? (parsed as { goals: Partial<ReadingGoal>[] }).goals.filter((goal) => goal.name && goal.metric && goal.startDate && goal.endDate).map((goal) => ({ id: goal.id || crypto.randomUUID(), name: goal.name!.trim(), metric: goal.metric!, target: Math.max(1, Number(goal.target) || 1), startDate: goal.startDate!, endDate: goal.endDate!, createdAt: goal.createdAt ?? now, updatedAt: now } as ReadingGoal)) : [];
      const goalChanges = importedGoals.map((goal) => ({ id: goal.id, before: currentGoals.get(goal.id), after: goal }));
      const createdBooks = bookChanges.filter((change) => !change.before).length;
      const warnings = ["BookStats exports match their own stable record IDs. Records already in this library are merged; unrelated existing records are left alone.", ...shelfPlan.warnings];
      setImportPlan({ source: "bookstats", sourceName: "BookStats", total: bookChanges.length, createdBooks, updatedBooks: bookChanges.length - createdBooks, ambiguousBooks: 0, createdShelves: shelfChanges.filter((change) => !change.before).length, warnings, bookChanges, shelfChanges, goalChanges, preview: bookChanges.map((change) => ({ title: change.after.title, author: change.after.author, action: change.before ? "update" : "new", matchedTitle: change.before?.title })) });
    } catch (error) { window.alert(`Could not import this file: ${error instanceof Error ? error.message : "unknown error"}`); }
  }

  async function importExternal(file: File, source: "goodreads" | "librarything") {
    try {
      if (!repository) throw new Error("Your local library is still opening. Try again in a moment.");
      const result = source === "goodreads" ? await importGoodreadsCsv(file) : await importLibraryThingJson(file);
      setImportPlan(buildExternalImportPlan(result, await repository.listBooks(), await repository.listShelves()));
    } catch (error) { window.alert(`Could not import this ${source === "goodreads" ? "Goodreads" : "LibraryThing"} file: ${error instanceof Error ? error.message : "unknown error"}`); }
  }
  async function commitImport() {
    if (!repository || !importPlan) return;
    await safetyBackup("before-import"); await applyExternalImportPlan(repository, importPlan); setImportPlan(undefined); await refresh(repository); await reloadSafetyData(repository);
    if (account?.emailVerified) void performSync().catch(() => undefined);
  }
  async function undoImport(batch: ImportBatch) {
    if (!repository || !window.confirm(`Undo the ${batch.sourceName} import from ${new Date(batch.createdAt).toLocaleString()}? Records you edited after the import will be skipped.`)) return;
    await safetyBackup("before-import-undo");
    const result = await undoImportBatch(repository, batch); await refresh(repository); await reloadSafetyData(repository); if (account?.emailVerified) void performSync().catch(() => undefined);
    window.alert(`Import undo finished.\n\n${result.removed} imported records removed\n${result.restored} previous records restored\n${result.skipped} newer records left untouched`);
  }

  async function createManualBackup() { if (!repository) return; const backup = await createBackupSnapshot(repository, __BOOKSTATS_VERSION__, "manual", backupPreferences(), true); downloadBackup(backup); await reloadSafetyData(repository); }
  async function restoreFromBackup(backup: BookStatsBackup) {
    if (!repository) return;
    if (!window.confirm(`Restore the backup from ${new Date(backup.createdAt).toLocaleString()}? Unlike Import, Restore replaces the current local library with the snapshot. A safety backup of the current library will be created first.`)) return;
    await safetyBackup("before-restore"); await restoreBackup(repository, backup); applyBackupPreferences(backup); await refresh(repository); await reloadSafetyData(repository); if (account?.emailVerified) void performSync().catch(() => undefined);
  }
  async function restoreBackupFile(file: File) { try { await restoreFromBackup(parseBackup(await file.text())); } catch (error) { window.alert(`Could not restore this backup: ${error instanceof Error ? error.message : "unknown error"}`); } }
  async function deleteBackup(backup: BookStatsBackup) { if (!repository) return; await removeLocalBackup(repository, backup.createdAt); await reloadSafetyData(repository); }
  function applyBackupPreferences(backup: BookStatsBackup) { if (backup.preferences?.layout) setLayout(backup.preferences.layout); if (Array.isArray(backup.preferences?.visibleColumns)) { const valid = backup.preferences.visibleColumns.filter((key): key is SortKey => DEFAULT_VISIBLE_COLUMNS.includes(key as SortKey)); setVisibleColumns(valid.includes("title") ? valid : ["title", ...valid]); } }

  async function applyBulkBooks(updated: Book[]) { await bulkPutBooks(updated); setSelectedIds(new Set(updated.map((book) => book.id))); if (account?.emailVerified) void performSync().catch(() => undefined); }
  async function deleteBulkBooks(items: Book[]) { if (!repository) return; await safetyBackup("before-bulk-delete"); for (const book of items) await repository.deleteBook(book.id); await refresh(repository); exitSelectionMode(); if (account?.emailVerified) void performSync().catch(() => undefined); }

  async function markDuplicateRecordsSeparate(items: Book[]) {
    if (items.length < 2) return;
    const ids = new Set(items.map((book) => book.id));
    const now = new Date().toISOString();
    const updated = items.map((book) => ({ ...book, duplicateIgnoreIds: [...new Set([...(book.duplicateIgnoreIds ?? []), ...[...ids].filter((id) => id !== book.id)])], updatedAt: now }));
    await bulkPutBooks(updated);
    if (account?.emailVerified) void performSync().catch(() => undefined);
  }

  async function saveSeriesCompletion(seriesName: string, override: SeriesCompletionOverride | undefined) {
    const key = normalizeSeriesName(seriesName);
    const now = new Date().toISOString();
    const matching = books.filter((book) => normalizeSeriesName(book.seriesMetadata?.name ?? book.series ?? "") === key);
    if (!matching.length) return;
    await bulkPutBooks(matching.map((book) => ({ ...book, seriesCompletionOverride: override ? { ...override, updatedAt: override.updatedAt || now } : undefined, updatedAt: now })));
    if (account?.emailVerified) void performSync().catch(() => undefined);
  }

  function openSeriesInLibrary(seriesName: string) {
    setSelectedBookId(undefined);
    setQuery("");
    setStatusFilter("all");
    setShelfFilter("all");
    setAdvancedFilter({ match: "all", ruleGroups: [{ id: crypto.randomUUID(), match: "all", rules: [{ id: crypto.randomUUID(), field: "series", operator: "equals", value: seriesName }] }] });
    setRandomBook(null);
    setView("library");
  }

  function dismissAccountWelcome() {
    localStorage.setItem(ACCOUNT_WELCOME_STORAGE_KEY, "1");
    setAccountWelcomeOpen(false);
  }

  function openAccountFromWelcome(mode: "login" | "register") {
    setAccountEntryMode(mode);
    dismissAccountWelcome();
    setView("account");
  }

  function startManualAdd() { setNewBookIsbn(undefined); setEditing(null); }
  function startScannedAdd(isbn: string) { setScannerOpen(false); setNewBookIsbn(isbn); setEditing(null); }

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><img className="brand-mark" src={`${import.meta.env.BASE_URL}bookstats-mark.svg`} alt="" /><div><strong>BookStats</strong><span>Personal library</span></div></div>
      <nav className="primary-nav"><button className={view === "library" && shelfFilter === "all" ? "active" : ""} onClick={() => { setView("library"); setShelfFilter("all"); }}><Library size={18} /><span>Library</span></button><button className={view === "lending" ? "active" : ""} onClick={() => setView("lending")}><Handshake size={18} /><span>Lending</span></button><button className={view === "statistics" ? "active" : ""} onClick={() => setView("statistics")}><BarChart3 size={18} /><span className="desktop-nav-label">Statistics</span><span className="mobile-nav-label">Stats</span></button><button className={view === "tools" ? "active" : ""} onClick={() => setView("tools")}><Wrench size={18} /><span>Tools</span></button><button className={view === "feedback" ? "active" : ""} onClick={() => setView("feedback")}><MessageSquareText size={18} /><span className="desktop-nav-label">Help / Feedback</span><span className="mobile-nav-label">Feedback</span></button><button className={view === "account" ? "active" : ""} onClick={() => { setAccountEntryMode("login"); setView("account"); }}><UserRound size={18} /><span>Account</span></button></nav>
      <div className="sidebar-shelves"><div className="sidebar-section-heading"><span>Shelves</span><button title="Add or manage shelves" aria-label="Add or manage shelves" onClick={() => setManagingShelves(true)}><Plus size={14} /></button></div><div className="sidebar-shelf-list">{shelves.length === 0 ? <button className="sidebar-empty-shelf" onClick={() => setManagingShelves(true)}>Create your first shelf</button> : shelves.map((shelf) => { const count = shelfCounts.get(shelf.id) ?? 0; return <button key={shelf.id} className={view === "library" && shelfFilter === shelf.id ? "active" : ""} onClick={() => { setView("library"); setShelfFilter(shelf.id); }}>{isSmartShelf(shelf) ? <Sparkles size={14} /> : <FolderOpen size={14} />}<span>{shelf.name}</span><small>{count}</small></button>; })}</div></div>
      <div className="sidebar-spacer" /><button className="local-mode account-status-button" onClick={() => { setAccountEntryMode("login"); setView("account"); }}><span className={`online-dot ${account ? "cloud-online" : ""}`} /><div><strong>{account ? account.displayName : "Local library"}</strong><span>{account ? (!account.emailVerified ? "Verify email to enable sync" : syncing ? "Synchronizing…" : lastSync ? `Synced ${formatRelativeSync(lastSync)}` : "Cloud sync enabled") : storageKind === "sqlite" ? "SQLite on this computer" : "Stored in this browser"}</span></div></button><div className="app-version" title={`BookStats ${__BOOKSTATS_VERSION__}`}>v{__BOOKSTATS_VERSION__}</div>
    </aside>

    <main className="main-content">
      {view === "library" && <>
        <header className="page-header"><div><p className="eyebrow">{shelfFilter === "all" ? "Your collection" : selectedShelf && isSmartShelf(selectedShelf) ? "Smart shelf" : "Shelf"}</p><h1>{shelfFilter === "all" ? "Library" : selectedShelf?.name ?? "Shelf"}</h1><p>{loading ? "Opening your library…" : books.length === 0 ? "Start building your library." : shelfFilter === "all" ? `${filtered.length} ${filtered.length === 1 ? "book" : "books"} in your collection.` : `${filtered.length} ${filtered.length === 1 ? "book" : "books"} ${selectedShelf && isSmartShelf(selectedShelf) ? "currently match this smart shelf." : "on this shelf."}`}</p></div><div className="page-header-actions"><button className="button secondary" onClick={() => setScannerOpen(true)}><ScanLine size={17} />Scan ISBN</button><button className="button primary" onClick={startManualAdd}><Plus size={18} />Add book</button></div></header>
        <section className="toolbar"><div className="search-box"><Search size={18} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search title, author, series, shelf, ISBN, publisher, genre, condition or tag…" /></div><div className="filter-group"><SlidersHorizontal size={17} /><select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value as StatusFilter)}><option value="all">All statuses</option>{Object.entries(READING_STATUS_LABELS).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></div><div className="filter-group"><FolderOpen size={16} /><select value={shelfFilter} onChange={(event) => setShelfFilter(event.target.value)}><option value="all">All shelves</option>{shelves.map((shelf) => <option key={shelf.id} value={shelf.id}>{shelf.name}{isSmartShelf(shelf) ? " (Smart)" : ""}</option>)}</select></div><button className={`button secondary compact ${advancedFilter ? "active-filter" : ""}`} onClick={() => setFilterOpen(true)}><Filter size={16} />{advancedFilter ? "Advanced filter active" : "Advanced filters"}</button><button className="button secondary compact mobile-shelf-button" onClick={() => setManagingShelves(true)}><FolderOpen size={16} />Manage shelves</button>{layout === "table" && <ColumnPicker visible={visibleColumns} onToggle={toggleColumn} onReset={() => setVisibleColumns(DEFAULT_VISIBLE_COLUMNS)} />}<div className="layout-toggle" aria-label="Library layout"><button className={layout === "grid" ? "active" : ""} onClick={() => setLayout("grid")} title="Cover grid"><Grid2X2 size={16} /></button><button className={layout === "table" ? "active" : ""} onClick={() => setLayout("table")} title="Table"><List size={17} /></button></div><button className={`button secondary compact ${selectionMode ? "active-filter" : ""}`} onClick={() => selectionMode ? exitSelectionMode() : setSelectionMode(true)}><CheckSquare2 size={16} />{selectionMode ? "Done selecting" : "Select"}</button><button className="button secondary compact" onClick={pickRandom}><Shuffle size={17} />Pick a book</button></section>
        {advancedFilter && <div className="active-filter-banner"><Filter size={14} /><span>Advanced rules are filtering this view.</span><button onClick={() => setFilterOpen(true)}>Edit</button><button onClick={() => setAdvancedFilter(undefined)}><X size={13} />Clear</button></div>}
        {selectionMode && <div className="selection-toolbar"><div><CheckSquare2 size={16} /><strong>{selectedIds.size.toLocaleString()} selected</strong><span>of {filtered.length.toLocaleString()} visible</span></div><div><button className="button secondary compact" onClick={() => setSelectedIds(new Set(filtered.map((book) => book.id)))}>Select visible</button><button className="button secondary compact" disabled={selectedIds.size === 0} onClick={() => setSelectedIds(new Set())}>Clear</button><button className="button primary compact" disabled={selectedIds.size === 0} onClick={() => setBulkOpen(true)}>Bulk edit</button></div></div>}
        {randomBook && <section className="random-pick"><div><p className="eyebrow">From the current unread view</p><strong>{randomBook.title}</strong><span>{randomBook.author}</span></div><div className="random-pick-actions"><button className="button secondary compact" onClick={() => setSelectedBookId(randomBook.id)}>Open</button><button className="icon-button random-pick-close" type="button" onClick={() => setRandomBook(null)} aria-label="Dismiss random book"><X size={16} /></button></div></section>}
        {!loading && (filtered.length === 0 ? <section className="empty-state"><BookOpen size={44} /><h2>{books.length === 0 ? "Your library is empty" : "No books match"}</h2><p>{books.length === 0 ? "Add your first book manually or use the catalog lookup inside Add Book." : "Try adjusting your search, shelf, status, or advanced filter rules."}</p>{books.length === 0 && <div className="empty-state-actions"><button className="button secondary" onClick={() => setScannerOpen(true)}><ScanLine size={17} />Scan ISBN</button><button className="button primary" onClick={startManualAdd}><Plus size={18} />Add your first book</button></div>}</section> : <>
          {layout === "grid" ? <section className="book-grid">{renderedBooks.map((book) => <BookCard key={book.id} book={book} shelfNames={shelfNamesByBook.get(book.id)} onOpen={(item) => setSelectedBookId(item.id)} onEdit={setEditing} onDelete={deleteBook} selectable={selectionMode} selected={selectedIds.has(book.id)} onSelect={toggleSelected} />)}</section> : <LibraryTable books={renderedBooks} shelfNamesByBook={shelfNamesByBook} visibleColumns={visibleColumns} sortKey={sortKey} sortDirection={sortDirection} onSort={changeSort} onOpen={(item) => setSelectedBookId(item.id)} onDelete={deleteBook} selectionMode={selectionMode} selectedIds={selectedIds} onSelect={toggleSelected} />}
          {renderedBooks.length < filtered.length && <div className="library-load-more"><span>Showing {renderedBooks.length.toLocaleString()} of {filtered.length.toLocaleString()} books</span><button className="button secondary" onClick={() => setRenderLimit((limit) => Math.min(filtered.length, limit + LIBRARY_RENDER_BATCH))}>Show {Math.min(LIBRARY_RENDER_BATCH, filtered.length - renderedBooks.length).toLocaleString()} more</button></div>}
        </>)}
      </>}
      {view === "lending" && <LendingView books={books} onSaveBook={saveBook} onOpenBook={(book) => setSelectedBookId(book.id)} />}
      {view === "statistics" && <StatisticsView books={books} shelves={shelves} goals={goals} onSaveGoal={saveGoal} onDeleteGoal={deleteGoal} onOpenSeries={openSeriesInLibrary} onSaveSeriesCompletion={saveSeriesCompletion} />}
      {view === "tools" && <ToolsView bookCount={books.length} onExport={exportLibrary} onImport={importLibrary} onImportGoodreads={(file) => importExternal(file, "goodreads")} onImportLibraryThing={(file) => importExternal(file, "librarything")} onOpenCleanup={() => setCleanupOpen(true)} backups={backups} imports={importHistory} onCreateBackup={createManualBackup} onRestoreBackupFile={restoreBackupFile} onRestoreLocalBackup={restoreFromBackup} onDownloadBackup={downloadBackup} onDeleteBackup={deleteBackup} onUndoImport={undoImport} />}
      {view === "feedback" && <FeedbackView account={account} storageKind={storageKind} bookCount={books.length} shelfCount={shelves.length} />}
      {view === "account" && <AccountView account={account} initialMode={accountEntryMode} storageKind={storageKind} syncing={syncing} lastSync={lastSync} syncError={syncError} onAccountChange={setAccount} onSync={performSync} onLogout={signOutAccount} onDeleteCloudData={removeCloudCopyAndDisconnect} onDeleteAccount={removeAccount} />}
    </main>

    {selectedBook && <BookDetail book={selectedBook} shelves={shelves} onEdit={editFromDetail} onOpenSeries={openSeriesInLibrary} onPrevious={detailNavigation.previous ? () => setSelectedBookId(detailNavigation.previous!.id) : undefined} onNext={detailNavigation.next ? () => setSelectedBookId(detailNavigation.next!.id) : undefined} onClose={() => setSelectedBookId(undefined)} />}
    {editing !== undefined && <BookForm book={editing ?? undefined} initialIsbn={editing === null ? newBookIsbn : undefined} autoLookupIsbn={Boolean(editing === null && newBookIsbn)} shelves={shelves} onCreateShelf={createShelf} onSave={saveBook} onClose={() => { setEditing(undefined); setNewBookIsbn(undefined); }} />}
    {managingShelves && <ShelfManager books={books} shelves={shelves} onCreate={createShelf} onUpdate={updateShelf} onReorder={reorderShelves} onDelete={removeShelf} onClose={() => setManagingShelves(false)} />}
    {cleanupOpen && <LibraryCleanup books={books} onMerge={mergeDuplicateRecords} onMarkSeparate={markDuplicateRecordsSeparate} onOpen={(book) => { setCleanupOpen(false); setSelectedBookId(book.id); }} onEdit={(book) => { setCleanupOpen(false); setEditing(book); }} onClose={() => setCleanupOpen(false)} />}
    {filterOpen && <AdvancedFilterPanel books={books} initial={advancedFilter} onApply={setAdvancedFilter} onSaveAsShelf={saveFilterAsShelf} onClose={() => setFilterOpen(false)} />}
    {bulkOpen && selectedBooks.length > 0 && <BulkEditPanel books={selectedBooks} shelves={shelves} onApply={applyBulkBooks} onDelete={deleteBulkBooks} onExport={(items) => exportBooks(items, false, "bookstats-selected")} onClose={() => setBulkOpen(false)} />}
    {importPlan && <ImportPreview plan={importPlan} onImport={commitImport} onCancel={() => setImportPlan(undefined)} />}
    {scannerOpen && <BarcodeScanner onDetected={startScannedAdd} onClose={() => setScannerOpen(false)} />}
    {accountWelcomeOpen && !account && !serverUpdateVersion && <div className="modal-backdrop account-welcome-backdrop" onMouseDown={(event) => event.target === event.currentTarget && dismissAccountWelcome()}>
      <section className="account-welcome-modal" role="dialog" aria-modal="true" aria-labelledby="bookstats-account-welcome-title">
        <button className="icon-button account-welcome-close" type="button" onClick={dismissAccountWelcome} aria-label="Continue without an account"><X size={18} /></button>
        <div className="account-welcome-icon"><Cloud size={28} /></div>
        <div className="account-welcome-copy"><p className="eyebrow">Optional BookStats account</p><h2 id="bookstats-account-welcome-title">Keep your library with you</h2><p>BookStats works without an account. Creating one adds cloud synchronization while your library continues to live locally on this device.</p></div>
        <div className="account-welcome-benefits">
          <div><CheckSquare2 size={17} /><span><strong>Sync across devices</strong><small>Keep edits, reading history, shelves, and covers available wherever you sign in.</small></span></div>
          <div><CheckSquare2 size={17} /><span><strong>Protect a browser library</strong><small>Your collection is not dependent on this one browser profile or device.</small></span></div>
          <div><CheckSquare2 size={17} /><span><strong>Still local-first</strong><small>You can continue using BookStats offline and keep independent exports and backups.</small></span></div>
        </div>
        <div className="account-welcome-actions"><button className="button primary" type="button" onClick={() => openAccountFromWelcome("register")}><UserPlus size={17} />Create account</button><button className="button secondary" type="button" onClick={() => openAccountFromWelcome("login")}><LogIn size={17} />Sign in</button><button className="text-button" type="button" onClick={dismissAccountWelcome}>Continue without an account</button></div>
      </section>
    </div>}
    {accountReady && !accountWelcomeOpen && !serverUpdateVersion && <InstallAppPrompt account={account} storageKind={storageKind} />}
    {serverUpdateVersion && <div className="modal-backdrop update-required-backdrop"><section className="update-required-modal" role="dialog" aria-modal="true" aria-labelledby="bookstats-update-title"><div className="update-required-icon"><RefreshCw size={28} /></div><div><p className="eyebrow">Update available</p><h2 id="bookstats-update-title">BookStats has been updated</h2>{isTauriRuntime() ? <><p>A newer desktop version is required. BookStats can download and install it for you.</p>{desktopUpdateProgress.phase === "checking" && <small>Checking signed update…</small>}{desktopUpdateProgress.phase === "downloading" && <small>{desktopUpdateProgress.total ? `Downloading update… ${Math.min(100, Math.round(((desktopUpdateProgress.downloaded ?? 0) / desktopUpdateProgress.total) * 100))}%` : "Downloading update…"}</small>}{desktopUpdateProgress.phase === "installing" && <small>Installing update…</small>}{desktopUpdateProgress.phase === "relaunching" && <small>Update installed. Relaunching BookStats…</small>}{desktopUpdateError && <small className="update-error">{desktopUpdateError}</small>}</> : <p>Refresh BookStats to load the current version before continuing.</p>}{serverUpdateVersion !== "newer" && <small>Current server version: {serverUpdateVersion}</small>}</div>{isTauriRuntime() ? <button className="button primary" disabled={["checking","downloading","installing","relaunching"].includes(desktopUpdateProgress.phase)} onClick={() => void installDesktopUpdate()}>{desktopUpdateProgress.phase === "error" ? "Retry update" : desktopUpdateProgress.phase === "idle" ? "Update BookStats" : "Updating…"}</button> : <button className="button primary" onClick={refreshBookStats}>Refresh BookStats</button>}</section></div>}
  </div>;
}

function refreshBookStats(): void {
  const url = new URL(window.location.href);
  url.searchParams.set("bookstats-refresh", String(Date.now()));
  window.location.replace(url.toString());
}

function normalizeSeriesName(value: string): string { return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim(); }

function planBookStatsShelves(imported: Partial<Shelf>[], current: Shelf[], now: string): { changes: Array<{ id: string; before?: Shelf; after: Shelf }>; idMap: Map<string, string>; validIds: Set<string>; warnings: string[] } {
  const byId = new Map(current.map((shelf) => [shelf.id, shelf]));
  const byName = new Map(current.map((shelf) => [shelf.name.trim().toLocaleLowerCase(), shelf]));
  const changes: Array<{ id: string; before?: Shelf; after: Shelf }> = [];
  const idMap = new Map<string, string>();
  const validIds = new Set(current.map((shelf) => shelf.id));
  const warnings: string[] = [];
  const plannedById = new Map<string, Shelf>();
  let nextOrder = current.reduce((highest, shelf) => typeof shelf.order === "number" && Number.isFinite(shelf.order) ? Math.max(highest, shelf.order) : highest, -1) + 1;

  const uniqueImportedName = (base: string) => {
    let candidate = `${base} (Imported)`; let index = 2;
    while (byName.has(candidate.toLocaleLowerCase())) candidate = `${base} (Imported ${index++})`;
    return candidate;
  };

  for (const raw of imported) {
    const sourceId = raw.id?.trim() || crypto.randomUUID();
    const name = raw.name?.trim() || "Imported shelf";
    const importedOrder = typeof raw.order === "number" && Number.isFinite(raw.order) ? raw.order : undefined;
    const normalized: Shelf = {
      id: sourceId,
      name,
      order: importedOrder,
      kind: raw.kind === "smart" ? "smart" : "manual",
      match: raw.kind === "smart" ? (raw.match ?? (Array.isArray(raw.ruleGroups) && raw.ruleGroups.length > 0 ? "any" : "all")) : undefined,
      rules: raw.kind === "smart" && Array.isArray(raw.rules) ? raw.rules : undefined,
      ruleGroups: raw.kind === "smart" && Array.isArray(raw.ruleGroups) ? raw.ruleGroups : undefined,
      createdAt: raw.createdAt ?? now,
      updatedAt: now
    };
    const sameId = byId.get(sourceId) ?? plannedById.get(sourceId);
    if (sameId) {
      const nameOwner = byName.get(name.toLocaleLowerCase());
      const safeName = nameOwner && nameOwner.id !== sameId.id ? uniqueImportedName(name) : name;
      const after = { ...normalized, id: sameId.id, name: safeName, order: importedOrder ?? sameId.order ?? nextOrder++, createdAt: sameId.createdAt || normalized.createdAt };
      if (safeName !== name) warnings.push(`Shelf “${name}” conflicts with another existing shelf name, so the imported record will be named “${safeName}”.`);
      changes.push({ id: sameId.id, before: byId.get(sameId.id), after });
      idMap.set(sourceId, sameId.id); validIds.add(sameId.id); byName.set(after.name.toLocaleLowerCase(), after); plannedById.set(after.id, after);
      continue;
    }

    const sameName = byName.get(name.toLocaleLowerCase());
    if (sameName && !isSmartShelf(sameName) && normalized.kind === "manual") {
      idMap.set(sourceId, sameName.id); validIds.add(sameName.id);
      continue;
    }
    if (sameName && isSmartShelf(sameName) && normalized.kind === "smart" && smartShelfDefinitionEqual(sameName, normalized)) {
      idMap.set(sourceId, sameName.id); validIds.add(sameName.id);
      continue;
    }

    const withOrder = { ...normalized, order: importedOrder ?? nextOrder++ };
    const after = sameName ? { ...withOrder, id: crypto.randomUUID(), name: uniqueImportedName(name) } : withOrder;
    if (sameName) warnings.push(`Shelf “${name}” already exists with a different type or smart-shelf rule set, so the imported shelf will be named “${after.name}”.`);
    changes.push({ id: after.id, after }); idMap.set(sourceId, after.id); validIds.add(after.id); byName.set(after.name.toLocaleLowerCase(), after); plannedById.set(after.id, after);
  }
  return { changes, idMap, validIds, warnings: [...new Set(warnings)] };
}

function smartShelfDefinitionEqual(a: Shelf, b: Shelf): boolean {
  if (!isSmartShelf(a) || !isSmartShelf(b)) return false;
  const normalizeRule = ({ field, operator, value }: NonNullable<Shelf["rules"]>[number]) => ({ field, operator, value: value ?? "" });
  const normalizeDefinition = (shelf: Shelf) => {
    if (shelf.ruleGroups?.length) return {
      match: shelf.match ?? "any",
      groups: shelf.ruleGroups.map((group) => ({ match: group.match ?? "all", rules: group.rules.map(normalizeRule) }))
    };
    return { match: "any", groups: [{ match: shelf.match ?? "all", rules: (shelf.rules ?? []).map(normalizeRule) }] };
  };
  return JSON.stringify(normalizeDefinition(a)) === JSON.stringify(normalizeDefinition(b));
}

function ColumnPicker({ visible, onToggle, onReset }: { visible: SortKey[]; onToggle: (key: SortKey) => void; onReset: () => void }) { return <details className="column-picker"><summary className="button secondary compact"><Columns3 size={16} />Hide / Show columns</summary><div className="column-picker-popover"><div><strong>Table columns</strong><button onClick={(event) => { event.preventDefault(); onReset(); }}>Show all</button></div>{TABLE_COLUMNS.map((column) => <label key={column.key} className={column.key === "title" ? "locked" : ""}><input type="checkbox" checked={visible.includes(column.key)} disabled={column.key === "title"} onChange={() => onToggle(column.key)} />{column.label}{column.key === "title" && <small>Always shown</small>}</label>)}</div></details>; }
function LibraryTable({ books, shelfNamesByBook, visibleColumns, sortKey, sortDirection, onSort, onOpen, onDelete, selectionMode, selectedIds, onSelect }: { books: Book[]; shelfNamesByBook: Map<string, string[]>; visibleColumns: SortKey[]; sortKey: SortKey; sortDirection: SortDirection; onSort: (key: SortKey) => void; onOpen: (book: Book) => void; onDelete: (book: Book) => void; selectionMode: boolean; selectedIds: Set<string>; onSelect: (book: Book, selected: boolean) => void }) {
  const columns = TABLE_COLUMNS.filter((column) => visibleColumns.includes(column.key));
  return <div className="table-wrap"><table className="library-table"><thead><tr>{selectionMode && <th className="select-column"><CheckSquare2 size={14} /></th>}{columns.map((column) => <th key={column.key}><button className={`sort-header ${sortKey === column.key ? "active" : ""}`} onClick={() => onSort(column.key)}>{column.label}<SortIcon active={sortKey === column.key} direction={sortDirection} /></button></th>)}<th></th></tr></thead><tbody>{books.map((book) => { const readDates = normalizedReadDates(book); const series = book.series ? `${book.series}${book.seriesVolume ? ` #${book.seriesVolume}` : ""}` : "—"; const bookShelves = (shelfNamesByBook.get(book.id) ?? []).join(", ") || "—"; const cells: Record<SortKey, ReactNode> = { title: <button className="title-link" onClick={() => onOpen(book)}>{book.title}</button>, author: book.author, series, status: READING_STATUS_LABELS[book.status], shelves: bookShelves, rating: typeof book.rating === "number" ? `${book.rating} ★` : "—", format: book.format ?? "—", pages: book.pages?.toLocaleString() ?? "—", reads: readDates.length || "—", lastRead: readDates.at(-1) ?? "—", owned: book.owned ? "Yes" : "No" }; return <tr key={book.id} className={selectedIds.has(book.id) ? "selected-row" : ""}>{selectionMode && <td className="select-column"><input type="checkbox" checked={selectedIds.has(book.id)} onChange={(event) => onSelect(book, event.target.checked)} aria-label={`Select ${book.title}`} /></td>}{columns.map((column) => <td key={column.key}>{cells[column.key]}</td>)}<td><button className="table-delete" onClick={() => void onDelete(book)}>Remove</button></td></tr>; })}</tbody></table></div>;
}
function SortIcon({ active, direction }: { active: boolean; direction: SortDirection }) { if (!active) return <ArrowUpDown size={12} />; return direction === "asc" ? <ArrowUp size={12} /> : <ArrowDown size={12} />; }
function compareBooks(a: Book, b: Book, key: SortKey, direction: SortDirection, shelfNamesByBook: Map<string, string[]>): number { const readA = normalizedReadDates(a); const readB = normalizedReadDates(b); const values: Record<SortKey, [string | number | boolean | undefined, string | number | boolean | undefined]> = { title: [a.title, b.title], author: [authorSortValue(a.author), authorSortValue(b.author)], series: [seriesSortValue(a), seriesSortValue(b)], status: [READING_STATUS_LABELS[a.status], READING_STATUS_LABELS[b.status]], shelves: [shelfSortValue(a, shelfNamesByBook), shelfSortValue(b, shelfNamesByBook)], rating: [a.rating, b.rating], format: [a.format, b.format], pages: [a.pages, b.pages], reads: [readA.length, readB.length], lastRead: [readA.at(-1), readB.at(-1)], owned: [a.owned, b.owned] }; const [left, right] = values[key]; const missingLeft = left === undefined || left === ""; const missingRight = right === undefined || right === ""; if (missingLeft && missingRight) return a.title.localeCompare(b.title); if (missingLeft) return 1; if (missingRight) return -1; let result: number; if (typeof left === "number" && typeof right === "number") result = left - right; else if (typeof left === "boolean" && typeof right === "boolean") result = Number(left) - Number(right); else result = String(left).localeCompare(String(right), undefined, { numeric: true, sensitivity: "base" }); return (direction === "asc" ? result : -result) || a.title.localeCompare(b.title); }
function authorSortValue(author: string): string | undefined {
  const clean = author.replace(/\s+/g, " ").trim();
  if (!clean) return undefined;
  const commaIndex = clean.indexOf(",");
  if (commaIndex > 0) return `${clean.slice(0, commaIndex).trim()}\u0000${clean.slice(commaIndex + 1).trim()}`;

  const parts = clean.split(" ");
  const suffixes = /^(?:jr\.?|sr\.?|ii|iii|iv|v)$/i;
  const suffix: string[] = [];
  while (parts.length > 1 && suffixes.test(parts.at(-1) ?? "")) suffix.unshift(parts.pop()!);
  if (parts.length === 1) return `${parts[0]}\u0000${suffix.join(" ")}`;

  // Keep common multi-word surname particles with the family name so authors such
  // as Ursula K. Le Guin sort under L rather than G. This remains intentionally
  // conservative because BookStats stores display names, not structured name parts.
  const surnameParticles = new Set(["da", "das", "de", "del", "della", "der", "di", "dos", "du", "la", "le", "van", "von"]);
  let surnameStart = parts.length - 1;
  while (surnameStart > 0 && surnameParticles.has(parts[surnameStart - 1].replace(/[.'’]/g, "").toLocaleLowerCase())) surnameStart -= 1;
  const surname = parts.slice(surnameStart).join(" ");
  const given = parts.slice(0, surnameStart).join(" ");
  return `${surname}\u0000${given}${suffix.length ? ` ${suffix.join(" ")}` : ""}`;
}
function seriesSortValue(book: Book): string | undefined { return book.series ? `${book.series}\u0000${book.seriesVolume ?? ""}` : undefined; }
function shelfSortValue(book: Book, shelfNamesByBook: Map<string, string[]>): string | undefined { const names = [...(shelfNamesByBook.get(book.id) ?? [])].sort(); return names.length ? names.join(" | ") : undefined; }
function loadVisibleColumns(): SortKey[] { try { const parsed = JSON.parse(localStorage.getItem(COLUMN_STORAGE_KEY) ?? "null") as unknown; if (Array.isArray(parsed)) { const valid = parsed.filter((key): key is SortKey => DEFAULT_VISIBLE_COLUMNS.includes(key as SortKey)); return valid.includes("title") ? valid : ["title", ...valid]; } } catch { /* use defaults */ } return DEFAULT_VISIBLE_COLUMNS; }
function loadLayout(): "grid" | "table" { try { return localStorage.getItem(LAYOUT_STORAGE_KEY) === "table" ? "table" : "grid"; } catch { return "grid"; } }
function formatRelativeSync(value: string): string { const delta = Date.now() - new Date(value).getTime(); if (!Number.isFinite(delta) || delta < 0) return "recently"; const minutes = Math.floor(delta / 60_000); if (minutes < 1) return "just now"; if (minutes < 60) return `${minutes}m ago`; const hours = Math.floor(minutes / 60); if (hours < 24) return `${hours}h ago`; return new Date(value).toLocaleDateString(); }
