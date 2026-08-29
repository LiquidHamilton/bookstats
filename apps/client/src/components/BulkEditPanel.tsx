import { useState } from "react";
import type { Book, BookCondition, ReadingStatus, Shelf } from "@bookstats/domain";
import { BOOK_CONDITIONS, isSmartShelf, READING_STATUS_LABELS } from "@bookstats/domain";
import { CheckSquare2, Download, Tags, Trash2, X } from "lucide-react";

type Action = "status" | "condition" | "ownership" | "add_shelf" | "remove_shelf" | "add_tags" | "remove_tags";

export function BulkEditPanel({
  books,
  shelves,
  onApply,
  onDelete,
  onExport,
  onClose
}: {
  books: Book[];
  shelves: Shelf[];
  onApply: (books: Book[]) => Promise<void>;
  onDelete: (books: Book[]) => Promise<void>;
  onExport: (books: Book[]) => void;
  onClose: () => void;
}) {
  const [action, setAction] = useState<Action>("status");
  const [status, setStatus] = useState<ReadingStatus>("want_to_read");
  const [condition, setCondition] = useState<BookCondition | "">("Good");
  const [owned, setOwned] = useState("true");
  const [shelfId, setShelfId] = useState("");
  const [tags, setTags] = useState("");
  const [working, setWorking] = useState(false);
  const manualShelves = shelves.filter((shelf) => !isSmartShelf(shelf));

  async function apply() {
    const now = new Date().toISOString();
    const tagValues = tags.split(",").map((tag) => tag.trim()).filter(Boolean);
    const updated = books.map((book) => {
      if (action === "status") return { ...book, status, updatedAt: now };
      if (action === "condition") return { ...book, condition: condition || undefined, updatedAt: now };
      if (action === "ownership") return { ...book, owned: owned === "true", updatedAt: now };
      if (action === "add_shelf" && shelfId) return { ...book, shelfIds: [...new Set([...(book.shelfIds ?? []), shelfId])], updatedAt: now };
      if (action === "remove_shelf" && shelfId) return { ...book, shelfIds: (book.shelfIds ?? []).filter((id) => id !== shelfId), updatedAt: now };
      if (action === "add_tags" && tagValues.length) return { ...book, tags: [...new Set([...(book.tags ?? []), ...tagValues])], updatedAt: now };
      if (action === "remove_tags" && tagValues.length) {
        const remove = new Set(tagValues.map((tag) => tag.toLocaleLowerCase()));
        return { ...book, tags: (book.tags ?? []).filter((tag) => !remove.has(tag.toLocaleLowerCase())), updatedAt: now };
      }
      return book;
    });
    setWorking(true);
    try { await onApply(updated); onClose(); } finally { setWorking(false); }
  }

  async function removeSelected() {
    if (!window.confirm(`Delete ${books.length.toLocaleString()} selected ${books.length === 1 ? "book" : "books"}? This can be recovered from a recent safety backup if necessary.`)) return;
    setWorking(true);
    try { await onDelete(books); onClose(); } finally { setWorking(false); }
  }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
    <section className="bulk-edit-modal">
      <div className="form-header"><div><p className="eyebrow">Library maintenance</p><h2>Bulk edit {books.length.toLocaleString()} {books.length === 1 ? "book" : "books"}</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div>
      <div className="bulk-edit-body">
        <label>Action<select value={action} onChange={(event) => setAction(event.target.value as Action)}><option value="status">Change status</option><option value="condition">Change condition</option><option value="ownership">Change ownership</option><option value="add_shelf">Add to shelf</option><option value="remove_shelf">Remove from shelf</option><option value="add_tags">Add tags</option><option value="remove_tags">Remove tags</option></select></label>
        {action === "status" && <label>New status<select value={status} onChange={(event) => setStatus(event.target.value as ReadingStatus)}>{(Object.entries(READING_STATUS_LABELS) as Array<[ReadingStatus, string]>).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>}
        {action === "condition" && <label>Condition<select value={condition} onChange={(event) => setCondition(event.target.value as BookCondition | "")}><option value="">Clear condition</option>{BOOK_CONDITIONS.map((value) => <option key={value} value={value}>{value}</option>)}</select></label>}
        {action === "ownership" && <label>Ownership<select value={owned} onChange={(event) => setOwned(event.target.value)}><option value="true">Owned</option><option value="false">Not owned</option></select></label>}
        {(action === "add_shelf" || action === "remove_shelf") && <label>Shelf<select value={shelfId} onChange={(event) => setShelfId(event.target.value)}><option value="">Choose a regular shelf…</option>{manualShelves.map((shelf) => <option key={shelf.id} value={shelf.id}>{shelf.name}</option>)}</select><small>Smart shelves are rule-based and cannot be assigned manually.</small></label>}
        {(action === "add_tags" || action === "remove_tags") && <label>Tags<input value={tags} onChange={(event) => setTags(event.target.value)} placeholder="science fiction, favorite" /><small>Separate multiple tags with commas.</small></label>}
      </div>
      <div className="bulk-edit-footer"><button className="button primary" disabled={working || ((action === "add_shelf" || action === "remove_shelf") && !shelfId)} onClick={() => void apply()}><CheckSquare2 size={16} />{working ? "Working…" : "Apply to selected"}</button><button className="button secondary" disabled={working} onClick={() => onExport(books)}><Download size={16} />Export selected</button><button className="button danger-button" disabled={working} onClick={() => void removeSelected()}><Trash2 size={16} />Delete selected</button></div>
      <p className="bulk-edit-note"><Tags size={14} />Only the fields you choose here are changed. Ratings, reviews, reading history and other metadata stay untouched.</p>
    </section>
  </div>;
}
