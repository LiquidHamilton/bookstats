import { useEffect, useMemo, useState, type CSSProperties } from "react";
import type { Book } from "@bookstats/domain";
import { AlertTriangle, BookOpen, CheckCircle2, GitMerge, HeartPulse, Search, Sparkles, X } from "lucide-react";
import { CoverImage } from "./CoverImage";
import { findDuplicateGroups, libraryHealth, metadataIssues, type DuplicateGroup, type MetadataIssue } from "../data/cleanup";

interface Props {
  books: Book[];
  onMerge: (keep: Book, remove: Book[]) => Promise<void>;
  onMarkSeparate: (books: Book[]) => Promise<void>;
  onOpen: (book: Book) => void;
  onEdit: (book: Book) => void;
  onClose: () => void;
}

export function LibraryCleanup({ books, onMerge, onMarkSeparate, onOpen, onEdit, onClose }: Props) {
  const [tab, setTab] = useState<"health" | "duplicates" | "metadata">("health");
  const [query, setQuery] = useState("");
  const [issueFilter, setIssueFilter] = useState<MetadataIssue | "all">("all");
  const [working, setWorking] = useState<string>();
  const [metadataLimit, setMetadataLimit] = useState(100);
  const [reviewGroup, setReviewGroup] = useState<DuplicateGroup>();
  const duplicates = useMemo(() => findDuplicateGroups(books), [books]);
  const health = useMemo(() => libraryHealth(books, duplicates.length), [books, duplicates.length]);
  const incomplete = useMemo(() => books.map((book) => ({ book, issues: metadataIssues(book) })).filter((item) => item.issues.length > 0).sort((a, b) => b.issues.length - a.issues.length || a.book.title.localeCompare(b.book.title)), [books]);
  const needle = query.trim().toLocaleLowerCase();
  const visibleIncomplete = incomplete.filter(({ book, issues }) => (issueFilter === "all" || issues.includes(issueFilter)) && (!needle || `${book.title} ${book.author} ${issues.join(" ")}`.toLocaleLowerCase().includes(needle)));
  const renderedIncomplete = visibleIncomplete.slice(0, metadataLimit);

  useEffect(() => { setMetadataLimit(100); }, [query, issueFilter]);

  async function merge(keep: Book, remove: Book[]) {
    const list = remove.map((book) => `“${book.title}”`).join(", ");
    if (!window.confirm(`Merge ${list} into “${keep.title}”? Reading sessions, shelves, tags and source IDs will be combined. The other record${remove.length === 1 ? "" : "s"} will be deleted.`)) return;
    setWorking(keep.id);
    try { await onMerge(keep, remove); } finally { setWorking(undefined); }
  }

  function fixNext() {
    const next = visibleIncomplete[0] ?? incomplete[0];
    if (next) onEdit(next.book);
  }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
    <section className="cleanup-modal">
      <div className="form-header"><div><p className="eyebrow">Library intelligence</p><h2>Library health & cleanup</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div>
      <div className="cleanup-tabs"><button className={tab === "health" ? "active" : ""} onClick={() => setTab("health")}><HeartPulse size={16} />Library health <span>{health.score}%</span></button><button className={tab === "metadata" ? "active" : ""} onClick={() => setTab("metadata")}><Sparkles size={16} />Metadata <span>{incomplete.length}</span></button><button className={tab === "duplicates" ? "active" : ""} onClick={() => setTab("duplicates")}><GitMerge size={16} />Duplicates <span>{duplicates.length}</span></button></div>

      {tab === "health" && <div className="cleanup-content">
        <div className="health-hero"><div className="health-score-ring" style={{ "--health": `${health.score * 3.6}deg` } as CSSProperties}><div><strong>{health.score}%</strong><span>health</span></div></div><div><h3>{health.score >= 95 ? "Your library is in excellent shape" : health.score >= 80 ? "Your library is in good shape" : health.score >= 60 ? "There is useful cleanup to do" : "Your library has plenty of metadata to improve"}</h3><p>The score checks covers, descriptions, ISBNs, page counts and publication years for every book, plus series position when a book belongs to a series. It never judges ratings, reviews, notes or other personal choices.</p><button className="button primary" disabled={incomplete.length === 0} onClick={fixNext}><Sparkles size={16} />Repair next book</button></div></div>
        <div className="health-metric-grid"><div><span>Incomplete books</span><strong>{health.incompleteBooks.toLocaleString()}</strong></div><div><span>Possible duplicate groups</span><strong>{health.duplicateGroups.toLocaleString()}</strong></div><div><span>Checks passed</span><strong>{health.passedChecks.toLocaleString()}</strong><small>of {health.totalChecks.toLocaleString()}</small></div></div>
        <div className="health-issues"><div className="section-heading"><div><p className="eyebrow">Where to focus</p><h3>Missing metadata</h3></div></div>{health.issueCounts.length === 0 ? <div className="cleanup-empty"><CheckCircle2 size={34} /><h3>The Librarian Approves.</h3><p>Every applicable core metadata check currently passes. Your collection is suspiciously well organized.</p></div> : health.issueCounts.map(({ issue, count }) => { const percent = books.length ? Math.round((count / books.length) * 100) : 0; return <button key={issue} onClick={() => { setIssueFilter(issue); setTab("metadata"); }}><div><strong>{issue}</strong><span>{count.toLocaleString()} {count === 1 ? "book" : "books"}</span></div><div className="health-bar"><i style={{ width: `${Math.min(100, percent)}%` }} /></div><span>{percent}% missing</span></button>; })}</div>
      </div>}

      {tab === "duplicates" && <div className="cleanup-content">
        <div className="cleanup-intro"><AlertTriangle size={18} /><p>BookStats ranks duplicate candidates using source IDs, ISBNs and normalized title + author. Different ISBNs are highlighted as likely intentional editions. Nothing is merged or hidden until you decide.</p></div>
        {duplicates.length === 0 ? <div className="cleanup-empty"><CheckCircle2 size={34} /><h3>No unresolved duplicate candidates</h3><p>Your library doesn't currently have any duplicate groups that still need a decision.</p></div> : <div className="duplicate-groups">{duplicates.map((group) => <article className={`duplicate-group duplicate-${group.confidence}`} key={group.id}><div className="duplicate-heading"><div><strong>{group.books.length} possible copies <em>{group.confidence} confidence</em></strong><span>Matched by {group.reason}</span>{group.editionConflict && <small>Different ISBNs detected — these may be separate editions that should stay separate.</small>}</div><button className="button secondary compact" onClick={() => setReviewGroup(group)}>Compare / reconcile</button></div><div className="duplicate-books">{group.books.map((book, index) => <div className="duplicate-book" key={book.id}><MiniBook book={book} /><div className="duplicate-actions"><button className="button secondary compact" onClick={() => onOpen(book)}>View</button><button className="button primary compact" disabled={working === book.id} onClick={() => void merge(book, group.books.filter((item) => item.id !== book.id))}>{working === book.id ? "Merging…" : index === 0 ? "Merge into best record" : "Keep this & merge"}</button></div></div>)}</div></article>)}</div>}
      </div>}

      {tab === "metadata" && <div className="cleanup-content">
        <div className="cleanup-intro"><Sparkles size={18} /><p>Review missing catalog information one book at a time. Edit / repair now includes targeted “Fill missing metadata” and “Load catalog covers” actions, so you can repair blank fields without replacing metadata that is already correct. Exact ISBN matches remain edition-specific.</p></div>
        <div className="metadata-filter-row"><div className="cleanup-search"><Search size={16} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search books or missing fields…" /></div><select value={issueFilter} onChange={(event) => setIssueFilter(event.target.value as MetadataIssue | "all")}><option value="all">All missing fields</option>{["Cover", "Description", "ISBN", "Pages", "Publication year", "Series position"].map((issue) => <option key={issue} value={issue}>{issue}</option>)}</select><button className="button primary compact" disabled={visibleIncomplete.length === 0} onClick={fixNext}>Repair next</button></div>
        {incomplete.length === 0 ? <div className="cleanup-empty"><CheckCircle2 size={34} /><h3>Metadata looks complete</h3><p>Every book has the core catalog fields BookStats checks here.</p></div> : visibleIncomplete.length === 0 ? <div className="cleanup-empty"><Search size={30} /><h3>No matching cleanup items</h3><p>Try another field filter or search.</p></div> : <><div className="metadata-cleanup-list">{renderedIncomplete.map(({ book, issues }) => <div className="metadata-cleanup-row" key={book.id}><MiniBook book={book} /><div className="issue-chips">{issues.map((issue) => <span key={issue}>{issue}</span>)}</div><div className="cleanup-row-actions"><button className="button secondary compact" onClick={() => onOpen(book)}>View</button><button className="button primary compact" onClick={() => onEdit(book)}>Edit / repair</button></div></div>)}</div>{renderedIncomplete.length < visibleIncomplete.length && <div className="library-load-more"><span>Showing {renderedIncomplete.length.toLocaleString()} of {visibleIncomplete.length.toLocaleString()} cleanup items</span><button className="button secondary compact" onClick={() => setMetadataLimit((limit) => Math.min(visibleIncomplete.length, limit + 100))}>Show 100 more</button></div>}</>}
      </div>}
    </section>
    {reviewGroup && <DuplicateReconcile group={reviewGroup} onClose={() => setReviewGroup(undefined)} onOpen={onOpen} onMerge={async (keep) => { await merge(keep, reviewGroup.books.filter((book) => book.id !== keep.id)); setReviewGroup(undefined); }} onMarkSeparate={async () => { await onMarkSeparate(reviewGroup.books); setReviewGroup(undefined); }} />}
  </div>;
}

function DuplicateReconcile({ group, onClose, onOpen, onMerge, onMarkSeparate }: { group: DuplicateGroup; onClose: () => void; onOpen: (book: Book) => void; onMerge: (keep: Book) => Promise<void>; onMarkSeparate: () => Promise<void> }) {
  const [keepId, setKeepId] = useState(group.books[0]?.id ?? "");
  const [working, setWorking] = useState<"merge" | "separate">();
  const keep = group.books.find((book) => book.id === keepId) ?? group.books[0];
  const rows: Array<[string, (book: Book) => string]> = [
    ["ISBN", (book) => book.isbn ?? "—"], ["Publication", (book) => book.publicationYear?.toString() ?? "—"], ["Publisher", (book) => book.publisher ?? "—"], ["Format", (book) => book.format ?? "—"], ["Pages", (book) => book.pages?.toLocaleString() ?? "—"], ["Series", (book) => book.series ? `${book.series}${book.seriesVolume ? ` #${book.seriesVolume}` : ""}` : "—"], ["Condition", (book) => book.condition ?? "—"], ["Rating", (book) => typeof book.rating === "number" ? `${book.rating} ★` : "—"], ["Readings", (book) => String(book.readDates?.length ?? 0)], ["Tags", (book) => book.tags?.join(", ") || "—"]
  ];
  return <div className="modal-backdrop duplicate-review-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="duplicate-review-modal"><div className="form-header"><div><p className="eyebrow">Duplicate reconciliation</p><h2>Compare records</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div><div className={`duplicate-review-note ${group.editionConflict ? "edition-warning" : ""}`}><AlertTriangle size={17} /><div><strong>{group.editionConflict ? "Possible separate editions" : `${group.confidence === "high" ? "Strong" : "Possible"} duplicate match`}</strong><span>{group.reason}. Choose the record to keep if these are truly duplicates, or mark the whole group as separate so BookStats stops asking.</span></div></div><div className="duplicate-compare-wrap"><table className="duplicate-compare"><thead><tr><th>Field</th>{group.books.map((book) => <th key={book.id}><label><input type="radio" name="keeper" checked={keepId === book.id} onChange={() => setKeepId(book.id)} /><span>Keep this</span></label><button onClick={() => onOpen(book)}>{book.title}</button><small>{book.author}</small></th>)}</tr></thead><tbody>{rows.map(([label, get]) => { const values = group.books.map(get); const differs = new Set(values).size > 1; return <tr key={label} className={differs ? "differs" : ""}><th>{label}</th>{group.books.map((book) => <td key={book.id}>{get(book)}</td>)}</tr>; })}</tbody></table></div><p className="duplicate-merge-explain">Merging preserves reading sessions, shelves, tags, source IDs, lending history and the strongest available catalog fields. A safety backup is created before the merge.</p><div className="form-actions duplicate-review-actions"><button className="button secondary" disabled={Boolean(working)} onClick={async () => { setWorking("separate"); try { await onMarkSeparate(); } finally { setWorking(undefined); } }}>Keep as separate editions/copies</button><span /><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!keep || Boolean(working)} onClick={async () => { if (!keep) return; setWorking("merge"); try { await onMerge(keep); } finally { setWorking(undefined); } }}>{working === "merge" ? "Merging…" : "Merge into selected record"}</button></div></section></div>;
}

function MiniBook({ book }: { book: Book }) {
  return <div className="mini-book"><div className="mini-cover"><CoverImage book={book} alt="" fallback={<BookOpen size={18} />} /></div><div><strong>{book.title}</strong><span>{book.author}</span><small>{[book.isbn, book.publicationYear, book.format].filter(Boolean).join(" · ") || "No edition details"}</small></div></div>;
}
