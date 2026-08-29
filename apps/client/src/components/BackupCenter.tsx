import { ArchiveRestore, Download, History, ShieldCheck, Trash2, Upload } from "lucide-react";
import type { BookStatsBackup } from "../data/backups";

interface Props {
  backups: BookStatsBackup[];
  onCreate: () => Promise<void>;
  onRestoreFile: (file: File) => Promise<void>;
  onRestoreLocal: (backup: BookStatsBackup) => Promise<void>;
  onDownloadLocal: (backup: BookStatsBackup) => void;
  onDeleteLocal: (backup: BookStatsBackup) => Promise<void>;
}

export function BackupCenter({ onCreate, onRestoreFile }: Props) {
  return <article className="tool-card">
    <div className="tool-icon"><ShieldCheck size={21} /></div>
    <div>
      <h2>Backup & restore</h2>
      <p>Create a recovery snapshot of BookStats or restore one from a file. Restoring a backup replaces the current local library with the saved state, unlike an export, which merges records into the library you already have.</p>
    </div>
    <div className="tool-card-actions">
      <button className="button primary" onClick={() => void onCreate()}><Download size={16} />Create backup</button>
      <label className="button secondary file-button"><Upload size={16} />Restore from file<input type="file" accept="application/json,.json" onChange={(event) => { const file = event.target.files?.[0]; if (file) void onRestoreFile(file); event.currentTarget.value = ""; }} /></label>
    </div>
  </article>;
}

export function RecentBackups({ backups, onRestoreLocal, onDownloadLocal, onDeleteLocal }: Pick<Props, "backups" | "onRestoreLocal" | "onDownloadLocal" | "onDeleteLocal">) {
  return <section className="recent-backups-section">
    <div className="recent-backups-heading">
      <div><History size={16} /><strong>Recent local safety backups</strong></div>
      <span>BookStats keeps up to five automatic or pre-change snapshots on this device.</span>
    </div>
    <div className="backup-privacy-note"><ShieldCheck size={14} /><span>Backups include your library, shelves, goals, reading history, custom covers and basic view preferences. Passwords, session tokens, server credentials and reconstructable remote-cover cache are excluded.</span></div>
    {backups.length === 0 ? <p className="read-history-empty">No local safety backups yet. BookStats will create them automatically before major changes and during normal use.</p> : <div className="recent-backups-list">{backups.map((backup) => <div className="recent-backup-row" key={`${backup.createdAt}-${backup.reason}`}><div><strong>{formatReason(backup.reason)}</strong><span>{new Date(backup.createdAt).toLocaleString()} · {backup.books.length.toLocaleString()} books · {backup.goals.length.toLocaleString()} goals</span></div><div><button className="button secondary compact" onClick={() => onDownloadLocal(backup)}><Download size={14} />Download</button><button className="button secondary compact" onClick={() => void onRestoreLocal(backup)}><ArchiveRestore size={14} />Restore</button><button className="icon-button danger-icon" aria-label="Delete backup" title="Delete this local backup" onClick={() => void onDeleteLocal(backup)}><Trash2 size={15} /></button></div></div>)}</div>}
  </section>;
}

function formatReason(reason: BookStatsBackup["reason"]): string {
  switch (reason) {
    case "automatic": return "Daily safety backup";
    case "before-import": return "Before import";
    case "before-import-undo": return "Before import undo";
    case "before-merge": return "Before duplicate merge";
    case "before-bulk-delete": return "Before bulk delete";
    case "before-restore": return "Before restore";
    default: return "Manual backup";
  }
}
