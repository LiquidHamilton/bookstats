import type { Book } from "@bookstats/domain";
import { READING_STATUS_LABELS } from "@bookstats/domain";
import { BookOpen, MoreHorizontal, Star } from "lucide-react";
import { CoverImage } from "./CoverImage";

interface Props {
  book: Book;
  density?: "compact" | "standard" | "large";
  shelfNames?: string[];
  onOpen: (book: Book) => void;
  onEdit: (book: Book) => void;
  onDelete: (book: Book) => void;
  selectable?: boolean;
  selected?: boolean;
  onSelect?: (book: Book, selected: boolean) => void;
}

export function BookCard({ book, density = "standard", shelfNames = [], onOpen, onEdit, onDelete, selectable = false, selected = false, onSelect }: Props) {
  const seriesLabel = book.series ? `${book.series}${book.seriesVolume ? ` · #${book.seriesVolume}` : ""}` : null;
  const assignedShelves = shelfNames;
  return (
    <article className={`book-card book-card-${density} ${selected ? "selected" : ""}`}>
      {selectable && <label className="book-select-checkbox" title="Select book"><input type="checkbox" checked={selected} onChange={(event) => onSelect?.(book, event.target.checked)} /><span /></label>}
      <button className="cover-button" onClick={() => onOpen(book)} aria-label={`Open ${book.title}`} title={density === "compact" ? `${book.title} — ${book.author}` : undefined}>
        <CoverImage book={book} className="cover" alt="" fallback={<div className="cover placeholder"><BookOpen size={34} /><span>{book.title.slice(0, 1)}</span></div>} />
      </button>
      {density !== "compact" && <>
      <div className="book-card-body">
        <div className="book-card-title-row">
          <div className="min-width-zero">
            <button className="book-title-button" onClick={() => onOpen(book)} title={book.title}><h3>{book.title}</h3></button>
            <p>{book.author}</p>
            {seriesLabel && <p className="series-line" title={seriesLabel}>{seriesLabel}</p>}
          </div>
          <details className="book-menu">
            <summary aria-label="Book actions"><MoreHorizontal size={18} /></summary>
            <div className="menu-popover">
              <button onClick={() => onOpen(book)}>View details</button>
              <button onClick={() => onEdit(book)}>Edit</button>
              <button className="danger-text" onClick={() => onDelete(book)}>Delete</button>
            </div>
          </details>
        </div>
        <div className="book-meta-row"><span className={`status status-${book.status}`}>{READING_STATUS_LABELS[book.status]}</span>{typeof book.rating === "number" && <span className="rating"><Star size={14} fill="currentColor" />{book.rating}</span>}</div>
        {assignedShelves.length > 0 && <div className="shelf-chips">{assignedShelves.slice(0, 2).map((name) => <span key={name}>{name}</span>)}{assignedShelves.length > 2 && <span>+{assignedShelves.length - 2}</span>}</div>}
        {book.tags.length > 0 && <div className="tags">{book.tags.slice(0, 3).map((tag) => <span key={tag}>#{tag}</span>)}</div>}
      </div>
      </>}
    </article>
  );
}
