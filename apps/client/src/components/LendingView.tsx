import { useMemo, useState } from "react";
import type { Book, LoanRecord } from "@bookstats/domain";
import { activeLoan, loanIsOverdue, normalizedLoans } from "@bookstats/domain";
import { AlertTriangle, BookOpen, CalendarDays, CheckCircle2, Clock3, Handshake, Plus, Search, Undo2, X } from "lucide-react";
import { CoverImage } from "./CoverImage";

interface Props {
  books: Book[];
  onSaveBook: (book: Book) => Promise<void>;
  onOpenBook: (book: Book) => void;
}

export function LendingView({ books, onSaveBook, onOpenBook }: Props) {
  const [editing, setEditing] = useState<{ book?: Book; loan?: LoanRecord }>();
  const [working, setWorking] = useState<string>();
  const active = useMemo(() => books.map((book) => ({ book, loan: activeLoan(book) })).filter((row): row is { book: Book; loan: LoanRecord } => Boolean(row.loan)).sort((a, b) => (a.loan.dueAt ?? "9999").localeCompare(b.loan.dueAt ?? "9999") || a.book.title.localeCompare(b.book.title)), [books]);
  const today = new Date().toISOString().slice(0, 10);
  const overdue = active.filter(({ loan }) => loanIsOverdue(loan, today));
  const dueSoon = active.filter(({ loan }) => !loanIsOverdue(loan, today) && loan.dueAt && loan.dueAt >= today && loan.dueAt <= addDays(today, 7));
  const returned = useMemo(() => books.flatMap((book) => normalizedLoans(book).filter((loan) => loan.returnedAt).map((loan) => ({ book, loan }))).sort((a, b) => (b.loan.returnedAt ?? "").localeCompare(a.loan.returnedAt ?? "")).slice(0, 30), [books]);
  const lendable = useMemo(() => books.filter((book) => book.owned && !activeLoan(book)).sort((a, b) => a.title.localeCompare(b.title)), [books]);

  async function markReturned(book: Book, loan: LoanRecord) {
    setWorking(loan.id);
    try {
      const now = new Date().toISOString();
      const returnedAt = now.slice(0, 10);
      const loans = normalizedLoans(book).map((item) => item.id === loan.id ? { ...item, returnedAt, updatedAt: now } : item);
      await onSaveBook({ ...book, loans, updatedAt: now });
    } finally { setWorking(undefined); }
  }

  return <>
    <header className="page-header lending-header"><div><p className="eyebrow">Collection</p><h1>Lending</h1><p>Keep track of books that have left your shelves and who has them.</p></div><button className="button primary" disabled={lendable.length === 0} onClick={() => setEditing({})}><Plus size={17} />Loan a book</button></header>
    <section className="metric-grid lending-metrics">
      <Metric label="On loan" value={active.length} note={active.length === 1 ? "book currently lent out" : "books currently lent out"} />
      <Metric label="Overdue" value={overdue.length} note={overdue.length ? "past their due date" : "nothing overdue"} tone={overdue.length ? "warning" : undefined} />
      <Metric label="Due soon" value={dueSoon.length} note="within the next 7 days" />
      <Metric label="Returned" value={books.reduce((sum, book) => sum + normalizedLoans(book).filter((loan) => loan.returnedAt).length, 0)} note="loans in your history" />
    </section>

    <section className="lending-section">
      <div className="section-heading"><div><p className="eyebrow">Active</p><h2>Books away from home</h2></div></div>
      {active.length === 0 ? <div className="lending-empty"><Handshake size={38} /><h3>Nothing is currently on loan</h3><p>When you lend a book, BookStats will keep the borrower, dates and optional notes with that copy.</p></div> : <div className="loan-card-grid">{active.map(({ book, loan }) => <LoanCard key={loan.id} book={book} loan={loan} onOpen={() => onOpenBook(book)} onEdit={() => setEditing({ book, loan })} onReturn={() => void markReturned(book, loan)} working={working === loan.id} />)}</div>}
    </section>

    {returned.length > 0 && <section className="lending-section"><div className="section-heading"><div><p className="eyebrow">History</p><h2>Recently returned</h2></div></div><div className="loan-history-list">{returned.map(({ book, loan }) => <button key={`${book.id}:${loan.id}`} onClick={() => onOpenBook(book)}><span className="loan-history-icon"><CheckCircle2 size={15} /></span><span><strong>{book.title}</strong><small>{loan.borrower}</small></span><span>{loan.returnedAt ? formatDate(loan.returnedAt) : "Returned"}</span></button>)}</div></section>}

    {editing && <LoanEditor books={lendable} book={editing.book} loan={editing.loan} onClose={() => setEditing(undefined)} onSave={async (book, loan) => { const now = new Date().toISOString(); const previous = normalizedLoans(book).filter((item) => item.id !== loan.id); await onSaveBook({ ...book, loans: [...previous, loan], updatedAt: now }); setEditing(undefined); }} />}
  </>;
}

function LoanCard({ book, loan, onOpen, onEdit, onReturn, working }: { book: Book; loan: LoanRecord; onOpen: () => void; onEdit: () => void; onReturn: () => void; working: boolean }) {
  const overdue = loanIsOverdue(loan);
  return <article className={`loan-card ${overdue ? "overdue" : ""}`}><button className="loan-book" onClick={onOpen}><span className="loan-cover"><CoverImage book={book} alt="" fallback={<BookOpen size={24} />} /></span><span><strong>{book.title}</strong><small>{book.author}</small></span></button><div className="loan-person"><Handshake size={15} /><span><small>Borrower</small><strong>{loan.borrower}</strong></span></div><div className="loan-dates"><span><CalendarDays size={14} />Loaned {formatDate(loan.loanedAt)}</span>{loan.dueAt && <span className={overdue ? "overdue-text" : ""}>{overdue ? <AlertTriangle size={14} /> : <Clock3 size={14} />}{overdue ? "Overdue since" : "Due"} {formatDate(loan.dueAt)}</span>}</div>{loan.notes && <p>{loan.notes}</p>}<div className="loan-actions"><button className="button secondary compact" onClick={onEdit}>Edit</button><button className="button primary compact" disabled={working} onClick={onReturn}><Undo2 size={14} />{working ? "Returning…" : "Mark returned"}</button></div></article>;
}

function LoanEditor({ books, book: initialBook, loan, onClose, onSave }: { books: Book[]; book?: Book; loan?: LoanRecord; onClose: () => void; onSave: (book: Book, loan: LoanRecord) => Promise<void> }) {
  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState(initialBook?.id ?? "");
  const [borrower, setBorrower] = useState(loan?.borrower ?? "");
  const [loanedAt, setLoanedAt] = useState(loan?.loanedAt ?? new Date().toISOString().slice(0, 10));
  const [dueAt, setDueAt] = useState(loan?.dueAt ?? "");
  const [notes, setNotes] = useState(loan?.notes ?? "");
  const [saving, setSaving] = useState(false);
  const selected = initialBook ?? books.find((item) => item.id === selectedId);
  const visible = books.filter((item) => !query.trim() || `${item.title} ${item.author} ${item.isbn ?? ""}`.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase())).slice(0, 80);

  async function save() {
    if (!selected || !borrower.trim() || !loanedAt) return;
    setSaving(true);
    try {
      const now = new Date().toISOString();
      await onSave(selected, { id: loan?.id ?? crypto.randomUUID(), borrower: borrower.trim(), loanedAt, dueAt: dueAt || undefined, returnedAt: loan?.returnedAt, notes: notes.trim() || undefined, createdAt: loan?.createdAt ?? now, updatedAt: now });
    } finally { setSaving(false); }
  }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="loan-editor-modal"><div className="form-header"><div><p className="eyebrow">Lending</p><h2>{loan ? "Edit loan" : "Loan a book"}</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div>{!initialBook && <div className="loan-book-picker"><label><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search your owned books…" /></label><div>{visible.map((book) => <button key={book.id} className={selectedId === book.id ? "selected" : ""} onClick={() => setSelectedId(book.id)}><strong>{book.title}</strong><span>{book.author}</span></button>)}</div></div>}{selected && <div className="loan-selected-book"><BookOpen size={17} /><div><strong>{selected.title}</strong><span>{selected.author}</span></div></div>}<div className="loan-form-grid"><label>Borrower<input autoFocus={Boolean(initialBook)} value={borrower} onChange={(event) => setBorrower(event.target.value)} placeholder="Name" /></label><label>Loaned<input type="date" value={loanedAt} onChange={(event) => setLoanedAt(event.target.value)} /></label><label>Due date<input type="date" min={loanedAt || undefined} value={dueAt} onChange={(event) => setDueAt(event.target.value)} /></label><label className="wide">Notes<textarea rows={3} value={notes} onChange={(event) => setNotes(event.target.value)} placeholder="Optional reminder, contact info, etc." /></label></div><div className="form-actions"><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!selected || !borrower.trim() || !loanedAt || saving} onClick={() => void save()}><Handshake size={16} />{saving ? "Saving…" : "Save loan"}</button></div></section></div>;
}

function Metric({ label, value, note, tone }: { label: string; value: number; note: string; tone?: "warning" }) { return <article className={`metric ${tone ? `metric-${tone}` : ""}`}><span>{label}</span><strong>{value.toLocaleString()}</strong><small>{note}</small></article>; }
function formatDate(value: string): string { const date = new Date(`${value}T12:00:00`); return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }); }
function addDays(value: string, days: number): string { const date = new Date(`${value}T12:00:00`); date.setDate(date.getDate() + days); return date.toISOString().slice(0, 10); }
