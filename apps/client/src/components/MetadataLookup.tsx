import { useEffect, useRef, useState } from "react";
import type { MetadataCandidate, MetadataProvider } from "@bookstats/domain";
import { BookOpen, CheckCircle2, LoaderCircle, Search } from "lucide-react";
import { metadataDetails, searchMetadata } from "../data/api";

interface Props {
  initialQuery?: string;
  autoSearch?: boolean;
  onApply: (candidate: MetadataCandidate) => void | Promise<void>;
}

const providerLabels: Record<MetadataProvider, string> = {
  aggregate: "Multiple sources",
  openlibrary: "Open Library",
  googlebooks: "Google Books",
  hardcover: "Hardcover"
};

export function MetadataLookup({ initialQuery = "", autoSearch = false, onApply }: Props) {
  const [query, setQuery] = useState(initialQuery);
  const [results, setResults] = useState<MetadataCandidate[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>();
  const [isbnSearch, setIsbnSearch] = useState(false);
  const autoSearched = useRef(false);

  async function search() {
    const value = query.trim();
    if (!value) return;
    setLoading(true); setError(undefined);
    try {
      const isbn = /^[\dXx\-\s]{9,20}$/.test(value) && value.replace(/[^0-9Xx]/g, "").length >= 10;
      setIsbnSearch(isbn);
      setResults(await searchMetadata(value, isbn));
    } catch (err) {
      setError(err instanceof Error ? err.message : "Lookup failed.");
    } finally { setLoading(false); }
  }

  useEffect(() => {
    if (!autoSearch || autoSearched.current || !initialQuery.trim()) return;
    autoSearched.current = true;
    void search();
  }, [autoSearch, initialQuery]); // search is intentionally triggered once for scanner-prefilled ISBNs

  async function choose(candidate: MetadataCandidate) {
    setLoading(true); setError(undefined);
    try {
      const details = await metadataDetails(candidate).catch(() => candidate);
      const coverUrl = details.coverUrl ?? candidate.coverUrl;
      const coverUrls = [...new Set([...(details.coverUrls ?? []), ...(candidate.coverUrls ?? []), ...(coverUrl ? [coverUrl] : [])])];
      await onApply({ ...candidate, ...details, coverUrl, coverUrls, subjects: details.subjects ?? candidate.subjects });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not load book details.");
    } finally { setLoading(false); }
  }

  return <section className="metadata-lookup">
    <div className="metadata-lookup-heading"><BookOpen size={18} /><div><strong>Look up book details</strong><span>BookStats combines Open Library, Google Books and Hardcover when configured. ISBN searches stay tied to that exact edition; choosing a result fills blank catalog fields and leaves existing values alone.</span></div></div>
    <div className="metadata-search-row">
      <input value={query} onChange={(event) => setQuery(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); void search(); } }} placeholder="ISBN, title, or title + author" />
      <button className="button secondary compact" type="button" disabled={loading || !query.trim()} onClick={() => void search()}>{loading ? <LoaderCircle className="spin" size={16} /> : <Search size={16} />}Search</button>
    </div>
    {isbnSearch && !loading && results.length > 0 && <p className="metadata-exact-note"><CheckCircle2 size={14} />Exact ISBN mode — edition-specific publisher, date, pages, format and identifiers are preserved.</p>}
    {error && <p className="inline-error">{error}</p>}
    {results.length > 0 && <div className="metadata-results">{results.map((candidate) => {
      const providers = [...new Set((candidate.sourceRefs ?? []).map((ref) => providerLabels[ref.provider]))];
      return <button type="button" className="metadata-result" key={`${candidate.source}-${candidate.workId}-${candidate.editionId ?? candidate.isbn ?? candidate.title}`} onClick={() => void choose(candidate)}>
        <div className="metadata-result-cover">{candidate.coverUrl ? <img src={candidate.coverUrl} alt="" /> : <BookOpen size={20} />}</div>
        <div><strong>{candidate.title}</strong><span>{candidate.author}</span><small>{[candidate.series ? `${candidate.series}${candidate.seriesVolume ? ` #${candidate.seriesVolume}` : ""}` : undefined, candidate.publicationYear, candidate.publisher, candidate.isbn].filter(Boolean).join(" · ")}</small><div className="metadata-source-line">{candidate.exactEdition && <em>Exact ISBN edition</em>}<span>{providers.length ? providers.join(" + ") : providerLabels[candidate.source]}</span></div></div>
      </button>;
    })}</div>}
    {!loading && isbnSearch && query.trim() && results.length === 0 && <p className="metadata-empty-note">No configured provider returned that exact ISBN. BookStats will not substitute a different edition for an ISBN lookup.</p>}
  </section>;
}
