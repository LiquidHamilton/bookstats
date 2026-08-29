import { BookCopy, FileJson, GitMerge, LibraryBig, RotateCcw, Sparkles, Upload } from "lucide-react";
import type { BookStatsBackup } from "../data/backups";
import type { ImportBatch } from "../data/importHistory";
import { BackupCenter, RecentBackups } from "./BackupCenter";

interface Props {
  bookCount: number;
  onExport: () => void;
  onImport: (file: File) => Promise<void>;
  onImportGoodreads: (file: File) => Promise<void>;
  onImportLibraryThing: (file: File) => Promise<void>;
  onOpenCleanup: () => void;
  backups: BookStatsBackup[];
  imports: ImportBatch[];
  onCreateBackup: () => Promise<void>;
  onRestoreBackupFile: (file: File) => Promise<void>;
  onRestoreLocalBackup: (backup: BookStatsBackup) => Promise<void>;
  onDownloadBackup: (backup: BookStatsBackup) => void;
  onDeleteBackup: (backup: BookStatsBackup) => Promise<void>;
  onUndoImport: (batch: ImportBatch) => Promise<void>;
}

export function ToolsView({ bookCount, onExport, onImport, onImportGoodreads, onImportLibraryThing, onOpenCleanup, backups, imports, onCreateBackup, onRestoreBackupFile, onRestoreLocalBackup, onDownloadBackup, onDeleteBackup, onUndoImport }: Props) {
  return <>
    <header className="page-header"><div><p className="eyebrow">Library utilities</p><h1>Tools</h1><p>Move your library, keep recovery snapshots, and clean up metadata when you need to.</p></div></header>
    <section className="tool-grid">
      <article className="tool-card"><div className="tool-icon"><FileJson size={21} /></div><div><h2>Export BookStats</h2><p>Create a portable, merge-friendly JSON copy of your books, shelves, reading sessions and goals. Importing an export adds or updates records instead of replacing the current library.</p></div><button className="button primary" disabled={bookCount === 0} onClick={onExport}><FileJson size={17} />Export library JSON</button></article>
      <article className="tool-card"><div className="tool-icon"><Upload size={21} /></div><div><h2>Import BookStats</h2><p>Bring in a BookStats export. Matching IDs are updated and existing records not present in the file are left alone.</p></div><label className="button secondary file-button"><Upload size={17} />Choose BookStats JSON<input type="file" accept="application/json,.json" onChange={(event) => { const file = event.target.files?.[0]; if (file) void onImport(file); event.currentTarget.value = ""; }} /></label></article>
      <article className="tool-card"><div className="tool-icon"><BookCopy size={21} /></div><div><h2>Import Goodreads</h2><p>Preview a Goodreads CSV before changing anything. BookStats shows new, matched and ambiguous records, then keeps the import available for safe undo.</p></div><label className="button secondary file-button"><Upload size={17} />Choose Goodreads CSV<input type="file" accept="text/csv,.csv" onChange={(event) => { const file = event.target.files?.[0]; if (file) void onImportGoodreads(file); event.currentTarget.value = ""; }} /></label></article>
      <article className="tool-card"><div className="tool-icon"><LibraryBig size={21} /></div><div><h2>Import LibraryThing</h2><p>Preview a LibraryThing JSON export with collections, tags, ratings, reading history, editions and series data before applying it.</p></div><label className="button secondary file-button"><Upload size={17} />Choose LibraryThing JSON<input type="file" accept="application/json,.json" onChange={(event) => { const file = event.target.files?.[0]; if (file) void onImportLibraryThing(file); event.currentTarget.value = ""; }} /></label></article>
      <article className="tool-card"><div className="tool-icon"><Sparkles size={21} /></div><div><h2>Library health & cleanup</h2><p>Review your library-health score, find missing metadata and series positions, detect possible duplicates, and work through cleanup one book at a time.</p></div><button className="button primary" disabled={bookCount === 0} onClick={onOpenCleanup}><GitMerge size={17} />Open health & cleanup</button></article>
      <BackupCenter backups={backups} onCreate={onCreateBackup} onRestoreFile={onRestoreBackupFile} onRestoreLocal={onRestoreLocalBackup} onDownloadLocal={onDownloadBackup} onDeleteLocal={onDeleteBackup} />
    </section>

    <RecentBackups backups={backups} onRestoreLocal={onRestoreLocalBackup} onDownloadLocal={onDownloadBackup} onDeleteLocal={onDeleteBackup} />

    <section className="recent-imports-section"><div className="section-heading"><div><p className="eyebrow">Import history</p><h2>Recent imports</h2><p>Undo only records that have not been edited since the import. BookStats skips anything you changed afterward instead of overwriting newer work.</p></div></div>{imports.length === 0 ? <p className="read-history-empty">No undoable imports yet.</p> : <div className="recent-import-list">{imports.map((batch) => <article key={batch.id}><div><strong>{batch.sourceName}</strong><span>{new Date(batch.createdAt).toLocaleString()} · {batch.createdBooks} new · {batch.updatedBooks} updated{batch.ambiguousBooks ? ` · ${batch.ambiguousBooks} ambiguous` : ""}</span></div><button className="button secondary compact" onClick={() => void onUndoImport(batch)}><RotateCcw size={14} />Undo import</button></article>)}</div>}</section>
  </>;
}
