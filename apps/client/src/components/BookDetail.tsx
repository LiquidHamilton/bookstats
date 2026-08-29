import { useEffect, useMemo, type ReactNode } from "react";
import DOMPurify from "dompurify";
import type { Book, Shelf } from "@bookstats/domain";
import { activeLoan, activeReadingSession, isSmartShelf, loanIsOverdue, normalizedReadingSessions, READING_STATUS_LABELS, shelfMatchesBook } from "@bookstats/domain";
import { AlertTriangle, BookOpen, CalendarDays, CheckCircle2, ChevronLeft, ChevronRight, Edit3, FolderOpen, Handshake, Hash, Sparkles, Star, X } from "lucide-react";
import { CoverImage } from "./CoverImage";

interface Props {
  book: Book;
  shelves: Shelf[];
  onEdit: (book: Book) => void;
  onOpenSeries: (seriesName: string) => void;
  onPrevious?: () => void;
  onNext?: () => void;
  onClose: () => void;
}

export function BookDetail({ book, shelves, onEdit, onOpenSeries, onPrevious, onNext, onClose }: Props) {
  useEffect(() => {
    const handle = (event: KeyboardEvent) => {
      if (event.key === "Escape") { onClose(); return; }
      if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
      if (event.key === "ArrowLeft" && onPrevious) { event.preventDefault(); onPrevious(); }
      if (event.key === "ArrowRight" && onNext) { event.preventDefault(); onNext(); }
    };
    window.addEventListener("keydown", handle);
    return () => window.removeEventListener("keydown", handle);
  }, [onClose, onPrevious, onNext]);

  const sessions = normalizedReadingSessions(book);
  const activeSession = activeReadingSession(book);
  const currentLoan = activeLoan(book);
  const matchingShelves = useMemo(() => shelves.filter((shelf) => shelfMatchesBook(shelf, book)), [book, shelves]);
  const regularShelves = matchingShelves.filter((shelf) => !isSmartShelf(shelf));
  const smartShelves = matchingShelves.filter(isSmartShelf);
  const series = book.series ? `${book.series}${book.seriesVolume ? ` · Book ${book.seriesVolume}` : ""}` : undefined;
  const details = [
    ["ISBN", book.isbn], ["Published", book.publicationYear?.toString()], ["Publisher", book.publisher], ["Format", book.format],
    ["Condition", book.condition], ["Pages", book.pages?.toLocaleString()], ["Genre", book.genre], ["Owned", book.owned ? "Yes" : "No"]
  ].filter(([, value]) => Boolean(value));
  const progressPercent = activeSession?.progressPages && book.pages ? Math.min(100, Math.round((activeSession.progressPages / book.pages) * 100)) : undefined;
  const catalogProviders = [...new Set((book.metadataSourceRefs ?? []).map((ref) => metadataProviderLabel(ref.provider)))];
  const descriptionHasHtml = Boolean(book.description && /<\/?(?:p|br|b|strong|i|em|u|s|strike|ul|ol|li|blockquote|a|small|sub|sup|h[2-4])\b/i.test(book.description));
  const descriptionHtml = useMemo(() => book.description ? DOMPurify.sanitize(book.description, {
    ALLOWED_TAGS: ["p", "br", "b", "strong", "i", "em", "u", "s", "strike", "ul", "ol", "li", "blockquote", "a", "small", "sub", "sup", "h2", "h3", "h4"],
    ALLOWED_ATTR: ["href", "title"]
  }) : "", [book.description]);

  return <div className="modal-backdrop detail-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
    <article className="book-detail">
      <div className="book-detail-topbar"><div><p className="eyebrow">Book details</p><span className={`status status-${book.status}`}>{READING_STATUS_LABELS[book.status]}</span></div><div className="detail-actions"><div className="detail-book-nav" aria-label="Book navigation"><button className="button secondary compact" type="button" disabled={!onPrevious} onClick={onPrevious} title="Previous book (Left arrow)"><ChevronLeft size={16} /><span>Previous</span></button><button className="button secondary compact" type="button" disabled={!onNext} onClick={onNext} title="Next book (Right arrow)"><span>Next</span><ChevronRight size={16} /></button></div><button className="button primary compact" onClick={() => onEdit(book)}><Edit3 size={15} />Edit</button><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div></div>

      <section className="book-detail-hero">
        <div className="book-detail-cover"><CoverImage book={book} alt={`Cover of ${book.title}`} fallback={<div className="cover placeholder"><BookOpen size={46} /><span>{book.title.slice(0, 1)}</span></div>} /></div>
        <div className="book-detail-heading">
          <h1>{book.title}</h1>
          <p className="book-detail-author">{book.author}</p>
          {book.additionalAuthors?.length > 0 && <p className="book-detail-additional">with {book.additionalAuthors.join(", ")}</p>}
          {series && <button type="button" className="book-detail-series" onClick={() => onOpenSeries(book.series!)} aria-label={`Open ${book.series} series in the library`}><BookOpen size={15} />{series}</button>}
          {(book.metadataMatchType === "exact_isbn" || catalogProviders.length > 0) && <div className="book-detail-catalog-source">{book.metadataMatchType === "exact_isbn" && <span className="exact-edition-chip"><CheckCircle2 size={12} />Exact ISBN edition</span>}{catalogProviders.length > 0 && <span>Catalog data: {catalogProviders.join(" + ")}</span>}</div>}
          {typeof book.rating === "number" && <div className="detail-rating"><Star size={17} fill="currentColor" /><strong>{book.rating}</strong><span>/ 5</span></div>}
          {activeSession && <div className="detail-progress-card"><div><span>Currently reading</span><strong>{activeSession.progressPages ? `Page ${activeSession.progressPages}${book.pages ? ` of ${book.pages}` : ""}` : "In progress"}</strong></div>{progressPercent !== undefined && <><div className="progress-track"><span style={{ width: `${progressPercent}%` }} /></div><small>{progressPercent}% complete</small></>}</div>}
          {currentLoan && <div className={`detail-loan-card ${loanIsOverdue(currentLoan) ? "overdue" : ""}`}><Handshake size={17} /><div><span>Currently on loan to <strong>{currentLoan.borrower}</strong></span><small>Loaned {formatDate(currentLoan.loanedAt)}{currentLoan.dueAt ? ` · ${loanIsOverdue(currentLoan) ? "Overdue since" : "Due"} ${formatDate(currentLoan.dueAt)}` : ""}</small></div>{loanIsOverdue(currentLoan) && <AlertTriangle size={16} />}</div>}
          <div className="detail-shelf-block">
            {regularShelves.length > 0 && <div><span className="detail-label"><FolderOpen size={13} />Shelves</span><div className="detail-chips">{regularShelves.map((shelf) => <span key={shelf.id}>{shelf.name}</span>)}</div></div>}
            {smartShelves.length > 0 && <div><span className="detail-label"><Sparkles size={13} />Smart shelves</span><div className="detail-chips smart">{smartShelves.map((shelf) => <span key={shelf.id}>{shelf.name}</span>)}</div></div>}
          </div>
        </div>
      </section>

      <section className="book-detail-body">
        {details.length > 0 && <div className="detail-facts">{details.map(([label, value]) => <div key={label}><span>{label}</span><strong>{value}</strong></div>)}</div>}
        {book.description && <DetailSection title="Description">{descriptionHasHtml ? <div className="detail-prose detail-rich-prose" dangerouslySetInnerHTML={{ __html: descriptionHtml }} /> : <p className="detail-prose">{book.description}</p>}</DetailSection>}
        {book.review && <DetailSection title="Your review"><p className="detail-prose review-prose">{book.review}</p></DetailSection>}
        {book.notes && <DetailSection title="Private notes"><p className="detail-prose">{book.notes}</p></DetailSection>}
        {book.tags?.length > 0 && <DetailSection title="Tags"><div className="detail-tags">{book.tags.map((tag) => <span key={tag}><Hash size={11} />{tag}</span>)}</div></DetailSection>}

        <DetailSection title="Reading history">
          {sessions.length === 0 ? <p className="detail-empty">No reading sessions recorded yet.</p> : <ol className="reading-timeline session-timeline">{sessions.map((session, index) => {
            const finished = Boolean(session.finishedAt);
            const days = durationDays(session.startedAt, session.finishedAt);
            return <li key={session.id}><span className={`timeline-dot ${finished ? "" : "active"}`}><CalendarDays size={13} /></span><div><strong>{finished ? (index === 0 ? "First read" : `Reading ${index + 1}`) : "Currently reading"}</strong><span>{session.startedAt ? `Started ${formatDate(session.startedAt)}` : "Start date not recorded"}{session.finishedAt ? ` · Finished ${formatDate(session.finishedAt)}` : ""}{days ? ` · ${days} ${days === 1 ? "day" : "days"}` : ""}</span>{session.progressPages && !finished && <span>Page {session.progressPages}{book.pages ? ` of ${book.pages}` : ""}</span>}{session.notes && <small>{session.notes}</small>}</div></li>;
          })}</ol>}
        </DetailSection>
      </section>
    </article>
  </div>;
}

function metadataProviderLabel(provider: "openlibrary" | "googlebooks" | "hardcover"): string { return provider === "googlebooks" ? "Google Books" : provider === "hardcover" ? "Hardcover" : "Open Library"; }

function DetailSection({ title, children }: { title: string; children: ReactNode }) {
  return <section className="detail-section"><h2>{title}</h2>{children}</section>;
}

function formatDate(value: string): string {
  const date = new Date(`${value}T12:00:00`);
  return Number.isNaN(date.getTime()) ? value : new Intl.DateTimeFormat(undefined, { year: "numeric", month: "long", day: "numeric" }).format(date);
}

function durationDays(start?: string, end?: string): number | undefined {
  if (!start || !end) return undefined;
  const a = new Date(`${start}T12:00:00`).getTime(); const b = new Date(`${end}T12:00:00`).getTime();
  if (Number.isNaN(a) || Number.isNaN(b) || b < a) return undefined;
  return Math.max(1, Math.round((b - a) / 86_400_000) + 1);
}
