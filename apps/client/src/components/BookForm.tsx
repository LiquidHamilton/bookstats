import { useEffect, useState } from "react";
import type { Book, BookCondition, BookFormat, MetadataCandidate, MetadataField, MetadataMatchType, MetadataProvider, MetadataSeriesMembership, MetadataSourceRef, ReadingSession, ReadingStatus, SeriesCompletionOverride, SeriesMetadata, Shelf } from "@bookstats/domain";
import { BOOK_CONDITIONS, isSmartShelf, normalizedReadingSessions, READING_STATUS_LABELS } from "@bookstats/domain";
import { CalendarPlus, ImagePlus, LoaderCircle, Plus, Sparkles, Trash2, X } from "lucide-react";
import { MetadataLookup } from "./MetadataLookup";
import { cacheRemoteCover, filterUsableCoverUrls, prepareUploadedCover } from "../data/covers";
import { CoverImage } from "./CoverImage";
import { archiveSelectedCover, getAuthToken, metadataDetails, searchMetadata } from "../data/api";

interface Props {
  book?: Book;
  initialIsbn?: string;
  autoLookupIsbn?: boolean;
  shelves: Shelf[];
  onCreateShelf: (name: string) => Promise<Shelf>;
  onSave: (book: Book) => Promise<void>;
  onClose: () => void;
}

const formats: BookFormat[] = [
  "Hardcover", "Paperback", "Mass Market Paperback", "eBook", "Audiobook",
  "Graphic Novel", "Omnibus", "Other"
];

const statuses = Object.keys(READING_STATUS_LABELS) as ReadingStatus[];
const ratings = Array.from({ length: 10 }, (_, index) => (index + 1) / 2);

function ratingLabel(value: number): string {
  return Number.isInteger(value) ? `${value} ★` : `${Math.floor(value)}½ ★`;
}

function hasMetadataValue(value: unknown): boolean {
  if (Array.isArray(value)) return value.some((item) => hasMetadataValue(item));
  if (value === undefined || value === null) return false;
  if (typeof value === "string") return value.trim().length > 0;
  if (typeof value === "number") return Number.isFinite(value);
  return true;
}

export function BookForm({ book, initialIsbn, autoLookupIsbn = false, shelves, onCreateShelf, onSave, onClose }: Props) {
  const [title, setTitle] = useState(book?.title ?? "");
  const [author, setAuthor] = useState(book?.author ?? "");
  const [additionalAuthors, setAdditionalAuthors] = useState(book?.additionalAuthors ?? []);
  const [isbn, setIsbn] = useState(book?.isbn ?? initialIsbn ?? "");
  const [series, setSeries] = useState(book?.series ?? "");
  const [seriesVolume, setSeriesVolume] = useState(book?.seriesVolume ?? "");
  const [pages, setPages] = useState(book?.pages?.toString() ?? "");
  const [publisher, setPublisher] = useState(book?.publisher ?? "");
  const [language, setLanguage] = useState(book?.language ?? "");
  const [year, setYear] = useState(book?.publicationYear?.toString() ?? "");
  const [format, setFormat] = useState<BookFormat | "">(book?.format ?? "");
  const [condition, setCondition] = useState<BookCondition | "">(book?.condition ?? "");
  const [status, setStatus] = useState<ReadingStatus>(book?.status ?? "not_started");
  const [owned, setOwned] = useState(book?.owned ?? true);
  const [selectedShelfIds, setSelectedShelfIds] = useState<string[]>(book?.shelfIds ?? []);
  const [newShelfName, setNewShelfName] = useState("");
  const [addingShelf, setAddingShelf] = useState(false);
  const [rating, setRating] = useState(book?.rating?.toString() ?? "");
  const [genre, setGenre] = useState(book?.genre ?? "");
  const [tags, setTags] = useState(book?.tags.join(", ") ?? "");
  const [description, setDescription] = useState(book?.description ?? "");
  const [review, setReview] = useState(book?.review ?? "");
  const [notes, setNotes] = useState(book?.notes ?? "");
  const [readingSessions, setReadingSessions] = useState<ReadingSession[]>(book ? normalizedReadingSessions(book) : []);
  const [saving, setSaving] = useState(false);
  const [coverUrl, setCoverUrl] = useState(book?.coverUrl ?? "");
  const [coverAssetId, setCoverAssetId] = useState(book?.coverAssetId);
  const [coverAssetToken, setCoverAssetToken] = useState(book?.coverAssetToken);
  const [coverSourceUrl, setCoverSourceUrl] = useState(book?.coverSourceUrl);
  const [coverArchivePending, setCoverArchivePending] = useState(Boolean(book?.coverArchivePending));
  const [cachedCoverDataUrl, setCachedCoverDataUrl] = useState(book?.cachedCoverDataUrl);
  const [coverError, setCoverError] = useState<string>();
  const [coverOptions, setCoverOptions] = useState<string[]>([]);
  const [metadataSource, setMetadataSource] = useState(book?.metadataSource);
  const [metadataWorkId, setMetadataWorkId] = useState(book?.metadataWorkId);
  const [metadataEditionId, setMetadataEditionId] = useState(book?.metadataEditionId);
  const [metadataMatchType, setMetadataMatchType] = useState<MetadataMatchType | undefined>(book?.metadataMatchType);
  const [metadataSourceRefs, setMetadataSourceRefs] = useState<MetadataSourceRef[]>(book?.metadataSourceRefs ?? []);
  const [metadataSources, setMetadataSources] = useState<Partial<Record<MetadataField, MetadataProvider>>>(book?.metadataSources ?? {});
  const [seriesMetadata, setSeriesMetadata] = useState<SeriesMetadata | undefined>(book?.seriesMetadata);
  const [seriesCompletionOverride, setSeriesCompletionOverride] = useState<SeriesCompletionOverride | undefined>(book?.seriesCompletionOverride);
  const [catalogAction, setCatalogAction] = useState<"covers" | "missing">();
  const [catalogMessage, setCatalogMessage] = useState<string>();
  const [catalogError, setCatalogError] = useState<string>();

  useEffect(() => {
    const handle = (event: KeyboardEvent) => event.key === "Escape" && onClose();
    window.addEventListener("keydown", handle);
    return () => window.removeEventListener("keydown", handle);
  }, [onClose]);

  function clearMetadataSource(field: MetadataField) {
    setMetadataSources((current) => {
      if (!current[field]) return current;
      const next = { ...current };
      delete next[field];
      return next;
    });
  }

  function updateAdditionalAuthor(index: number, value: string) {
    setAdditionalAuthors((current) => current.map((name, itemIndex) => itemIndex === index ? value : name));
    clearMetadataSource("additionalAuthors");
  }

  function addAdditionalAuthor() {
    setAdditionalAuthors((current) => [...current, ""]);
    clearMetadataSource("additionalAuthors");
  }

  function removeAdditionalAuthor(index: number) {
    setAdditionalAuthors((current) => current.filter((_, itemIndex) => itemIndex !== index));
    clearMetadataSource("additionalAuthors");
  }

  function selectCover(value: string, manual: boolean) {
    const next = value.trim();
    if (next === coverUrl.trim() && coverAssetId && coverAssetToken && !coverArchivePending) {
      if (manual) clearMetadataSource("coverUrl");
      return;
    }
    setCoverUrl(next);
    setCachedCoverDataUrl(undefined);
    setCoverAssetId(undefined);
    setCoverAssetToken(undefined);
    setCoverSourceUrl(/^https?:\/\//i.test(next) ? next : undefined);
    setCoverArchivePending(Boolean(next));
    if (manual) clearMetadataSource("coverUrl");
  }

  function clearCover() {
    setCoverUrl("");
    setCachedCoverDataUrl(undefined);
    setCoverAssetId(undefined);
    setCoverAssetToken(undefined);
    setCoverSourceUrl(undefined);
    setCoverArchivePending(false);
    setCoverOptions([]);
    clearMetadataSource("coverUrl");
  }

  function metadataSeriesMemberships(candidate: MetadataCandidate): MetadataSeriesMembership[] {
    const memberships = [...(candidate.seriesMemberships ?? [])];
    const fallbackName = candidate.series ?? candidate.seriesMetadata?.name;
    if (fallbackName?.trim()) memberships.unshift({
      provider: candidate.seriesMetadata?.provider ?? (candidate.source !== "aggregate" ? candidate.source : undefined),
      seriesId: candidate.seriesMetadata?.id,
      name: fallbackName,
      volume: candidate.seriesVolume,
      metadata: candidate.seriesMetadata
    });
    const seen = new Set<string>();
    return memberships.filter((membership) => {
      const name = normalizedCatalogText(membership.name);
      const volume = normalizedCatalogText(membership.volume ?? "");
      if (!name) return false;
      const key = `${name}|${volume}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function matchingSeriesMembership(candidate: MetadataCandidate, preferredSeries: string): MetadataSeriesMembership | undefined {
    const wanted = normalizedCatalogText(preferredSeries);
    if (!wanted) return undefined;
    const memberships = metadataSeriesMemberships(candidate);
    const exact = memberships.filter((membership) => normalizedCatalogText(membership.name) === wanted);
    if (exact.length) return exact.find((membership) => membership.volume) ?? exact[0];
    const compatible = memberships.filter((membership) => {
      const name = normalizedCatalogText(membership.name);
      return name.length >= 5 && wanted.length >= 5 && (name.includes(wanted) || wanted.includes(name));
    });
    return compatible.length === 1 ? compatible[0] : undefined;
  }

  function primarySeriesMembership(candidate: MetadataCandidate): MetadataSeriesMembership | undefined {
    const memberships = metadataSeriesMemberships(candidate);
    const preferred = candidate.series ?? candidate.seriesMetadata?.name;
    return (preferred ? matchingSeriesMembership(candidate, preferred) : undefined) ?? memberships[0];
  }

  function fullSeriesMetadata(membership: MetadataSeriesMembership | undefined, candidate: MetadataCandidate): SeriesMetadata | undefined {
    const direct = membership?.metadata;
    if (direct?.books?.length) return direct;
    if (candidate.seriesMetadata && membership && normalizedCatalogText(candidate.seriesMetadata.name) === normalizedCatalogText(membership.name)) return candidate.seriesMetadata;
    return undefined;
  }

  function applyMetadataSeriesPair(candidate: MetadataCandidate): Set<MetadataField> {
    const applied = new Set<MetadataField>();
    const currentSeries = series.trim();
    const currentVolume = seriesVolume.trim();

    // Series name + position are a pair. Metadata only fills missing pieces; it never
    // replaces a populated field. If only the series name exists, use the position
    // from that same membership. If only a position exists, only attach a series when
    // the provider explicitly reports the same position.
    if (currentSeries) {
      const membership = matchingSeriesMembership(candidate, currentSeries);
      if (!currentVolume && membership?.volume) {
        setSeriesVolume(membership.volume);
        applied.add("seriesVolume");
      }
      if (membership) setSeriesMetadata(fullSeriesMetadata(membership, candidate));
      return applied;
    }

    const membership = primarySeriesMembership(candidate);
    if (!membership) return applied;
    if (currentVolume && (!membership.volume || normalizedCatalogText(currentVolume) !== normalizedCatalogText(membership.volume))) return applied;

    setSeries(membership.name);
    applied.add("series");
    if (!currentVolume && membership.volume) {
      setSeriesVolume(membership.volume);
      applied.add("seriesVolume");
    }
    setSeriesMetadata(fullSeriesMetadata(membership, candidate));
    return applied;
  }

  function applyBlankMetadata(candidate: MetadataCandidate): MetadataField[] {
    const applied = new Set<MetadataField>();
    const apply = (field: MetadataField, currentValue: unknown, candidateValue: unknown, setter: (value: any) => void) => {
      if (hasMetadataValue(currentValue) || !hasMetadataValue(candidateValue)) return;
      setter(candidateValue);
      applied.add(field);
    };

    apply("title", title, candidate.title, (value) => setTitle(String(value)));
    apply("author", author, candidate.author, (value) => setAuthor(String(value)));
    if (!additionalAuthors.some((name) => name.trim()) && candidate.additionalAuthors.some((name) => name.trim())) {
      setAdditionalAuthors(candidate.additionalAuthors);
      applied.add("additionalAuthors");
    }
    apply("isbn", isbn, candidate.isbn, (value) => setIsbn(String(value)));
    for (const field of applyMetadataSeriesPair(candidate)) applied.add(field);
    apply("pages", pages, candidate.pages, (value) => setPages(String(value)));
    apply("publicationYear", year, candidate.publicationYear, (value) => setYear(String(value)));
    apply("publisher", publisher, candidate.publisher, (value) => setPublisher(String(value)));
    apply("language", language, candidate.language, (value) => setLanguage(String(value)));
    apply("format", format, candidate.format, (value) => setFormat(value as BookFormat));
    apply("description", description, candidate.description, (value) => setDescription(String(value)));
    const hasCurrentCover = hasMetadataValue(coverAssetId || coverUrl || coverSourceUrl || cachedCoverDataUrl);
    if (!genre.trim() && candidate.subjects.length) {
      setGenre(candidate.subjects[0]);
      applied.add("genre");
    }

    const coverChoices = [...new Set([...(candidate.coverUrls ?? []), ...(candidate.coverUrl ? [candidate.coverUrl] : [])])];
    if (coverChoices.length) {
      void filterUsableCoverUrls(coverChoices).then((usable) => {
        setCoverOptions(usable);
        if (!hasCurrentCover && usable[0]) {
          selectCover(usable[0], false);
          const provider = candidate.fieldSources?.coverUrl;
          if (provider) setMetadataSources((current) => ({ ...current, coverUrl: provider }));
        }
      });
    }
    setMetadataSources((current) => {
      const next = { ...current };
      for (const field of applied) {
        const provider = candidate.fieldSources?.[field];
        if (provider) next[field] = provider;
      }
      return next;
    });
    return [...applied];
  }

  function applyMetadata(candidate: MetadataCandidate) {
    applyBlankMetadata(candidate);
    setMetadataSource(candidate.source);
    setMetadataWorkId(candidate.workId);
    setMetadataEditionId(candidate.editionId);
    setMetadataMatchType(candidate.matchType);
    setMetadataSourceRefs(candidate.sourceRefs ?? []);
  }


  function normalizedCatalogText(value: string): string {
    return value.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  }

  function isbnLike(value: string): boolean {
    const clean = value.replace(/[^0-9Xx]/g, "");
    return clean.length === 10 || clean.length === 13;
  }

  async function searchForCurrentBook(allowTitleFallback: boolean): Promise<{ candidates: MetadataCandidate[]; exactIsbn: boolean; fellBack: boolean }> {
    const currentIsbn = isbn.trim();
    if (isbnLike(currentIsbn)) {
      const exact = await searchMetadata(currentIsbn, true);
      if (exact.length || !allowTitleFallback) return { candidates: exact, exactIsbn: true, fellBack: false };
    }
    const textQuery = [title.trim(), author.trim()].filter(Boolean).join(" ");
    if (!textQuery) return { candidates: [], exactIsbn: false, fellBack: Boolean(currentIsbn) };
    const results = await searchMetadata(textQuery, false);
    const wantedTitle = normalizedCatalogText(title);
    const wantedAuthor = normalizedCatalogText(author);
    const close = results.filter((candidate) => {
      const candidateTitle = normalizedCatalogText(candidate.title);
      const candidateAuthor = normalizedCatalogText(candidate.author);
      return candidateTitle === wantedTitle && (!wantedAuthor || candidateAuthor === wantedAuthor || candidateAuthor.includes(wantedAuthor) || wantedAuthor.includes(candidateAuthor));
    });
    return { candidates: close.length ? close : (allowTitleFallback ? results : []), exactIsbn: false, fellBack: Boolean(currentIsbn) };
  }

  async function coverUrlsFromCandidates(candidates: MetadataCandidate[], detailCount: number): Promise<string[]> {
    const detailResults = await Promise.allSettled(candidates.slice(0, detailCount).map((candidate) => metadataDetails(candidate)));
    const detailed = detailResults.flatMap((result) => result.status === "fulfilled" ? [result.value] : []);
    const urls = [...new Set([...candidates, ...detailed].flatMap((candidate) => [...(candidate.coverUrls ?? []), ...(candidate.coverUrl ? [candidate.coverUrl] : [])]).filter(Boolean))];
    return filterUsableCoverUrls(urls);
  }

  async function loadCatalogCovers() {
    setCatalogAction("covers"); setCatalogError(undefined); setCatalogMessage(undefined);
    try {
      const currentIsbn = isbnLike(isbn.trim()) ? isbn.trim() : "";
      let exactUrls: string[] = [];
      let alternateUrls: string[] = [];

      // Cover selection is deliberately broader than metadata repair. Keep exact-edition
      // ISBN artwork first, but also offer other likely editions of the same title/author.
      // Choosing one only changes the cover; it never changes the book's ISBN or metadata.
      if (currentIsbn) {
        const exactCandidates = await searchMetadata(currentIsbn, true);
        if (exactCandidates.length) exactUrls = await coverUrlsFromCandidates(exactCandidates, 1);
      }

      if (title.trim() && author.trim()) {
        const textResults = await searchMetadata(`${title.trim()} ${author.trim()}`, false);
        const wantedTitle = normalizedCatalogText(title);
        const wantedAuthor = normalizedCatalogText(author);
        const close = textResults.filter((candidate) => {
          const candidateTitle = normalizedCatalogText(candidate.title);
          const candidateAuthor = normalizedCatalogText(candidate.author);
          return candidateTitle === wantedTitle && (!wantedAuthor || candidateAuthor === wantedAuthor || candidateAuthor.includes(wantedAuthor) || wantedAuthor.includes(candidateAuthor));
        });
        const alternateCandidates = (close.length ? close : textResults).slice(0, 8);
        if (alternateCandidates.length) alternateUrls = await coverUrlsFromCandidates(alternateCandidates, alternateCandidates.length);
      }

      const exactSet = new Set(exactUrls);
      const urls = [...new Set([...exactUrls, ...alternateUrls.filter((url) => !exactSet.has(url))])];
      if (!urls.length) throw new Error("No catalog cover matches were found for this book.");
      setCoverOptions(urls);

      if (currentIsbn && exactUrls.length && alternateUrls.some((url) => !exactSet.has(url))) {
        setCatalogMessage(`${urls.length} catalog ${urls.length === 1 ? "cover" : "covers"} loaded. Exact-ISBN artwork is shown first, followed by alternate editions found by title and author.`);
      } else if (currentIsbn && exactUrls.length) {
        setCatalogMessage(`${urls.length} exact-ISBN ${urls.length === 1 ? "cover" : "covers"} loaded. No additional title/author artwork was found.`);
      } else if (currentIsbn) {
        setCatalogMessage(`${urls.length} alternate-edition ${urls.length === 1 ? "cover" : "covers"} loaded by title and author because the current ISBN did not return usable artwork.`);
      } else {
        setCatalogMessage(`${urls.length} catalog ${urls.length === 1 ? "cover" : "covers"} loaded by title and author.`);
      }
    } catch (error) { setCatalogError(error instanceof Error ? error.message : "Could not load catalog covers."); }
    finally { setCatalogAction(undefined); }
  }

  async function fillMissingCatalogMetadata() {
    setCatalogAction("missing"); setCatalogError(undefined); setCatalogMessage(undefined);
    try {
      const lookup = await searchForCurrentBook(false);
      if (!lookup.candidates.length) {
        if (isbnLike(isbn.trim())) throw new Error("No exact metadata match was found for this ISBN. BookStats will not fill edition fields from a different ISBN; use the full catalog search if you want to choose a title match manually.");
        throw new Error("No sufficiently close title/author match was found. Use the full catalog search above to choose the correct book manually.");
      }
      const candidate = await metadataDetails(lookup.candidates[0]).catch(() => lookup.candidates[0]);
      const applied = applyBlankMetadata(candidate);
      setMetadataSource(candidate.source);
      setMetadataWorkId(candidate.workId);
      setMetadataEditionId(candidate.editionId);
      setMetadataMatchType(candidate.matchType);
      setMetadataSourceRefs(candidate.sourceRefs ?? []);
      const urls = [...new Set([...(candidate.coverUrls ?? []), ...(candidate.coverUrl ? [candidate.coverUrl] : [])])];
      const labels: Partial<Record<MetadataField, string>> = { title: "title", author: "author", additionalAuthors: "additional authors", isbn: "ISBN", series: "series", seriesVolume: "series position", pages: "pages", publicationYear: "publication year", publisher: "publisher", language: "language", format: "format", description: "description", genre: "genre", coverUrl: "cover" };
      setCatalogMessage(applied.length ? `Filled ${applied.map((field) => labels[field] ?? field).join(", ")}. Existing values were left untouched.${urls.length ? " Cover choices were loaded too." : ""}` : urls.length ? "No blank catalog fields could be filled, but cover choices were loaded." : "No additional missing metadata was available from the matched catalog record.");
    } catch (error) { setCatalogError(error instanceof Error ? error.message : "Could not fill missing metadata."); }
    finally { setCatalogAction(undefined); }
  }


  function addReadingSession(completed = false) {
    const today = new Date().toISOString().slice(0, 10);
    const now = new Date().toISOString();
    const session: ReadingSession = {
      id: crypto.randomUUID(),
      startedAt: completed ? undefined : today,
      finishedAt: completed ? today : undefined,
      createdAt: now,
      updatedAt: now
    };
    setReadingSessions((sessions) => [...sessions, session]);
    setStatus(completed ? "read" : "currently_reading");
  }

  function updateReadingSession(id: string, patch: Partial<ReadingSession>) {
    setReadingSessions((sessions) => sessions.map((session) => session.id === id ? { ...session, ...patch, updatedAt: new Date().toISOString() } : session));
  }

  function removeReadingSession(id: string) {
    setReadingSessions((sessions) => sessions.filter((session) => session.id !== id));
  }

  function toggleShelf(id: string) {
    setSelectedShelfIds((ids) => ids.includes(id) ? ids.filter((item) => item !== id) : [...ids, id]);
  }

  async function addShelf() {
    const name = newShelfName.trim();
    if (!name) return;
    setAddingShelf(true);
    try {
      const shelf = await onCreateShelf(name);
      setSelectedShelfIds((ids) => ids.includes(shelf.id) ? ids : [...ids, shelf.id]);
      setNewShelfName("");
    } finally { setAddingShelf(false); }
  }

  async function uploadCover(file: File) {
    setCoverError(undefined);
    try {
      const dataUrl = await prepareUploadedCover(file);
      selectCover(dataUrl, true);
      // Keep a device-local copy even after an account archives the definitive cloud asset.
      setCachedCoverDataUrl(dataUrl);
      setCoverOptions([]);
    } catch (error) {
      setCoverError(error instanceof Error ? error.message : "Could not read that image.");
    }
  }

  async function submit(event: React.FormEvent) {
    event.preventDefault();
    if (!title.trim() || !author.trim()) return;
    setSaving(true);
    try {
      const now = new Date().toISOString();
      const normalizedSessions = readingSessions.map((session) => ({ ...session, progressPages: session.progressPages ? Math.max(0, Math.round(session.progressPages)) : undefined })).sort((a, b) => (a.finishedAt ?? a.startedAt ?? "9999").localeCompare(b.finishedAt ?? b.startedAt ?? "9999"));
      const normalizedDates = [...new Set(normalizedSessions.map((session) => session.finishedAt).filter((value): value is string => Boolean(value)))].sort();
      let finalCoverUrl = coverUrl.trim() || undefined;
      let finalCoverAssetId = coverAssetId;
      let finalCoverAssetToken = coverAssetToken;
      let finalCoverSourceUrl = coverSourceUrl;
      let finalCoverArchivePending = coverArchivePending;
      let localCoverCache = cachedCoverDataUrl;
      if (finalCoverUrl && /^https?:\/\//i.test(finalCoverUrl) && (!book || book.coverUrl !== finalCoverUrl || !localCoverCache)) {
        localCoverCache = await cacheRemoteCover(finalCoverUrl);
      }

      // Archiving is opportunistic so BookStats remains local-first/offline capable. If the
      // server cannot be reached, sync will retry because coverArchivePending remains true.
      if (finalCoverUrl && (!finalCoverAssetId || finalCoverArchivePending) && getAuthToken()) {
        const sourceBeforeArchive = finalCoverUrl;
        try {
          const archived = await archiveSelectedCover(sourceBeforeArchive);
          finalCoverAssetId = archived.assetId;
          finalCoverAssetToken = archived.assetToken;
          finalCoverSourceUrl = /^https?:\/\//i.test(sourceBeforeArchive) ? sourceBeforeArchive : archived.sourceUrl;
          finalCoverArchivePending = false;
          if (sourceBeforeArchive.startsWith("data:")) finalCoverUrl = undefined;
        } catch {
          // If a third-party URL has already disappeared but this device still has the
          // selected image cached, archive that local copy instead of losing the cover.
          if (/^https?:\/\//i.test(sourceBeforeArchive) && localCoverCache?.startsWith("data:")) {
            try {
              const archived = await archiveSelectedCover(localCoverCache);
              finalCoverAssetId = archived.assetId;
              finalCoverAssetToken = archived.assetToken;
              finalCoverSourceUrl = sourceBeforeArchive;
              finalCoverArchivePending = false;
            } catch {
              finalCoverArchivePending = true;
            }
          } else {
            finalCoverArchivePending = true;
          }
        }
      }
      const normalizedAdditionalAuthors = [...new Set(additionalAuthors.map((name) => name.trim()).filter(Boolean))]
        .filter((name) => normalizedCatalogText(name) !== normalizedCatalogText(author.trim()));
      await onSave({
        id: book?.id ?? crypto.randomUUID(),
        title: title.trim(),
        author: author.trim(),
        additionalAuthors: normalizedAdditionalAuthors,
        isbn: isbn.trim() || undefined,
        series: series.trim() || undefined,
        seriesVolume: seriesVolume.trim() || undefined,
        pages: pages ? Number(pages) : undefined,
        publicationYear: year ? Number(year) : undefined,
        publisher: publisher.trim() || undefined,
        language: language.trim() || undefined,
        format: format || undefined,
        condition: condition || undefined,
        status,
        owned,
        shelfIds: selectedShelfIds,
        rating: rating ? Number(rating) : undefined,
        genre: genre.trim() || undefined,
        tags: tags.split(",").map((tag) => tag.trim()).filter(Boolean),
        description: description.trim() || undefined,
        review: review.trim() || undefined,
        notes: notes.trim() || undefined,
        coverUrl: finalCoverUrl,
        coverAssetId: finalCoverAssetId,
        coverAssetToken: finalCoverAssetToken,
        coverSourceUrl: finalCoverSourceUrl,
        coverArchivePending: finalCoverArchivePending || undefined,
        cachedCoverDataUrl: localCoverCache,
        metadataSource,
        metadataWorkId,
        metadataEditionId,
        metadataMatchType,
        metadataSourceRefs,
        metadataSources,
        seriesMetadata,
        seriesCompletionOverride,
        loans: book?.loans,
        duplicateIgnoreIds: book?.duplicateIgnoreIds,
        sourceIds: book?.sourceIds,
        dateAdded: book?.dateAdded ?? now,
        readingSessions: normalizedSessions,
        readDates: normalizedDates,
        dateRead: undefined,
        createdAt: book?.createdAt ?? now,
        updatedAt: now
      });
      onClose();
    } finally { setSaving(false); }
  }

  const coverPreviewRecord = { coverUrl, coverAssetId, coverAssetToken, coverSourceUrl, cachedCoverDataUrl, updatedAt: book?.updatedAt };
  const manualShelves = shelves.filter((shelf) => !isSmartShelf(shelf));

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <form className="book-form" onSubmit={submit}>
        <div className="form-header">
          <div>
            <p className="eyebrow">{book ? "Edit library item" : "Add to library"}</p>
            <h2>{book ? book.title : "New book"}</h2>
          </div>
          <button className="icon-button" type="button" onClick={onClose} aria-label="Close"><X size={20} /></button>
        </div>

        <MetadataLookup initialQuery={book?.isbn || initialIsbn || [book?.title, book?.author].filter(Boolean).join(" ")} autoSearch={autoLookupIsbn} onApply={applyMetadata} />
        <div className="metadata-quick-repair"><div><Sparkles size={17} /><div><strong>Targeted metadata repair</strong><span>Fill only blank catalog fields. BookStats does not replace metadata you already have.</span></div></div><button className="button secondary compact" type="button" disabled={Boolean(catalogAction) || !title.trim() || !author.trim()} onClick={() => void fillMissingCatalogMetadata()}>{catalogAction === "missing" ? <LoaderCircle className="spin" size={15} /> : <Sparkles size={15} />}Fill missing metadata</button></div>
        {catalogMessage && <p className="metadata-repair-message">{catalogMessage}</p>}
        {catalogError && <p className="inline-error metadata-repair-error">{catalogError}</p>}

        <div className="form-section">
          <h3>Book details</h3>
          <div className="form-grid">
            <label className="wide">Title<input autoFocus required value={title} onChange={(e) => { setTitle(e.target.value); clearMetadataSource("title"); }} /></label>
            <label className="wide">Primary author<input required value={author} onChange={(e) => { setAuthor(e.target.value); clearMetadataSource("author"); }} /></label>
            <div className="wide additional-authors-editor">
              <div className="additional-authors-heading"><div><strong>Additional authors</strong><span>Add co-authors, editors, or other credited authors one at a time.</span></div><button className="button secondary compact" type="button" onClick={addAdditionalAuthor}><Plus size={14} />Add author</button></div>
              {additionalAuthors.length > 0 && <div className="additional-authors-list">{additionalAuthors.map((name, index) => <div className="additional-author-row" key={index}><input value={name} onChange={(event) => updateAdditionalAuthor(index, event.target.value)} placeholder={`Additional author ${index + 1}`} aria-label={`Additional author ${index + 1}`} /><button className="icon-button" type="button" onClick={() => removeAdditionalAuthor(index)} aria-label={`Remove additional author ${index + 1}`}><X size={15} /></button></div>)}</div>}
            </div>
            <label>Series<input value={series} onChange={(e) => {
              const nextSeries = e.target.value;
              if (normalizedCatalogText(nextSeries) !== normalizedCatalogText(series)) {
                setSeriesCompletionOverride(undefined);
                setSeriesMetadata(undefined);
                // A manually changed series name must not inherit a catalog position from
                // the previous series. Clear it so a later lookup can refill the matching pair.
                setSeriesVolume("");
                clearMetadataSource("seriesVolume");
              }
              setSeries(nextSeries);
              clearMetadataSource("series");
            }} placeholder="Series name" /></label>
            <label>Volume<input value={seriesVolume} onChange={(e) => { setSeriesVolume(e.target.value); setSeriesMetadata(undefined); clearMetadataSource("seriesVolume"); }} placeholder="Book or volume number" /></label>
            <label>ISBN<input value={isbn} onChange={(e) => { setIsbn(e.target.value); clearMetadataSource("isbn"); }} /></label>
            <label>Publication year<input type="number" min="0" max="9999" value={year} onChange={(e) => { setYear(e.target.value); clearMetadataSource("publicationYear"); }} /></label>
            <label>Pages<input type="number" min="0" value={pages} onChange={(e) => { setPages(e.target.value); clearMetadataSource("pages"); }} /></label>
            <label>Publisher<input value={publisher} onChange={(e) => { setPublisher(e.target.value); clearMetadataSource("publisher"); }} /></label>
            <label>Language<input value={language} onChange={(e) => { setLanguage(e.target.value); clearMetadataSource("language"); }} placeholder="eng" /></label>
            <label>Format<select value={format} onChange={(e) => { setFormat(e.target.value as BookFormat | ""); clearMetadataSource("format"); }}><option value="">Not set</option>{formats.map((item) => <option key={item}>{item}</option>)}</select></label>
            <label className="wide">Genre<input value={genre} onChange={(e) => { setGenre(e.target.value); clearMetadataSource("genre"); }} placeholder="Science Fiction" /></label>
            <label className="wide">Tags<input value={tags} onChange={(e) => setTags(e.target.value)} placeholder="signed, first edition, comfort read" /><small>Separate tags with commas. Shelves are managed separately below.</small></label>
            {(coverUrl || coverAssetId || cachedCoverDataUrl || coverSourceUrl) && <div className="wide selected-cover-preview"><CoverImage book={coverPreviewRecord} alt="Selected book cover" /><div><strong>Selected cover</strong><span>{coverAssetId ? "Stored in your BookStats cover archive and cached locally when available." : coverUrl.startsWith("data:") ? "Custom image stored locally; it will be archived when cloud sync is available." : cachedCoverDataUrl ? "Catalog cover cached locally; BookStats will archive the selected image when signed in." : "Catalog/URL cover."}</span></div></div>}
            {coverOptions.length > 0 && <div className="wide alternate-cover-picker"><div className="alternate-cover-heading"><strong>Catalog covers</strong><span>Choose the edition artwork you want BookStats to display.</span></div><div className="alternate-cover-grid">{coverOptions.map((url) => <button type="button" key={url} className={coverUrl === url ? "selected" : ""} onClick={() => selectCover(url, true)}><img src={url} alt="Alternate cover option" loading="lazy" onError={() => setCoverOptions((current) => current.filter((item) => item !== url))} /></button>)}</div></div>}
            <label className="wide">Cover URL<input value={coverUrl.startsWith("data:") ? "" : coverUrl} onChange={(e) => selectCover(e.target.value, true)} placeholder={coverUrl.startsWith("data:") ? "Custom cover selected" : "https://…"} /></label>
            <div className="wide cover-actions">
              <button className="button secondary" type="button" disabled={Boolean(catalogAction) || !title.trim() || !author.trim()} onClick={() => void loadCatalogCovers()}>{catalogAction === "covers" ? <LoaderCircle className="spin" size={16} /> : <ImagePlus size={16} />}Load catalog covers</button>
              <label className="button secondary file-button"><ImagePlus size={16} />Choose local cover<input type="file" accept="image/png,image/jpeg,image/webp,image/gif" onChange={(event) => { const file = event.target.files?.[0]; if (file) void uploadCover(file); event.currentTarget.value = ""; }} /></label>
              {(coverUrl || coverAssetId || cachedCoverDataUrl || coverSourceUrl) && <button className="button secondary" type="button" onClick={clearCover}>Clear cover</button>}
              <small>Load catalog covers shows exact-ISBN artwork first, then alternate editions found by title + author. Choosing a cover does not change the book's ISBN or other metadata.</small>
            </div>
            {coverError && <p className="wide inline-error">{coverError}</p>}
            <label className="wide">Book description<textarea rows={5} value={description} onChange={(e) => { setDescription(e.target.value); clearMetadataSource("description"); }} placeholder="Publisher or catalog description…" /></label>
          </div>
        </div>

        <div className="form-section">
          <h3>Your library & reading</h3>
          <div className="form-grid library-reading-grid">
            <label>Status<select value={status} onChange={(e) => setStatus(e.target.value as ReadingStatus)}>{statuses.map((item) => <option key={item} value={item}>{READING_STATUS_LABELS[item]}</option>)}</select></label>
            <label>Condition<select value={condition} onChange={(e) => setCondition(e.target.value as BookCondition | "")}><option value="">Not set</option>{BOOK_CONDITIONS.map((item) => <option key={item} value={item}>{item}</option>)}</select><small>Your own assessment; catalog lookups never change it.</small></label>
            <label>Rating<select value={rating} onChange={(e) => setRating(e.target.value)}><option value="">Unrated</option>{ratings.map((value) => <option key={value} value={value}>{ratingLabel(value)}</option>)}</select></label>
            <label className="checkbox-row wide"><input type="checkbox" checked={owned} onChange={(e) => setOwned(e.target.checked)} /> I own this edition</label>
          </div>

          <div className="shelf-picker">
            <div className="shelf-picker-heading"><div><strong>Shelves</strong><span>A book can be on as many shelves as you like while keeping one main status.</span></div></div>
            {manualShelves.length > 0 ? <div className="shelf-checkboxes">{manualShelves.map((shelf) => <label key={shelf.id}><input type="checkbox" checked={selectedShelfIds.includes(shelf.id)} onChange={() => toggleShelf(shelf.id)} /><span>{shelf.name}</span></label>)}</div> : <p className="read-history-empty">No regular shelves yet. Smart shelves are filled automatically and do not need to be assigned here.</p>}
            <div className="new-shelf-row"><input value={newShelfName} onChange={(event) => setNewShelfName(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); void addShelf(); } }} placeholder="New regular shelf" /><button className="button secondary compact" disabled={!newShelfName.trim() || addingShelf} type="button" onClick={() => void addShelf()}><Plus size={15} />Add shelf</button></div>
          </div>

          <div className="read-history-editor session-editor">
            <div className="read-history-heading">
              <div><strong>Reading history & progress</strong><span>Track starts, finishes, rereads, and your current page.</span></div>
              <div className="reading-session-actions"><button className="button secondary compact" type="button" onClick={() => addReadingSession(false)}><CalendarPlus size={15} />Start reading</button><button className="button secondary compact" type="button" onClick={() => addReadingSession(true)}>Add past read</button></div>
            </div>
            {readingSessions.length === 0 ? <p className="read-history-empty">No reading sessions recorded yet.</p> : (
              <div className="reading-session-list">
                {readingSessions.map((session, index) => {
                  const percent = pages && session.progressPages ? Math.min(100, Math.round((session.progressPages / Number(pages)) * 100)) : undefined;
                  const sessionTitle = session.finishedAt ? (index === 0 ? "First read" : `Reading ${index + 1}`) : "Currently reading";
                  return <div className={`reading-session-row ${session.finishedAt ? "completed" : "active"}`} key={session.id}>
                    <div className="reading-session-label">
                      <div className="reading-session-title"><strong>{sessionTitle}</strong><span>{session.finishedAt ? "Completed reading session" : "Reading in progress"}</span></div>
                      <div className="reading-session-row-actions">{percent !== undefined && !session.finishedAt && <span className="reading-progress-chip">{percent}%</span>}<button type="button" className="icon-button danger-icon reading-session-delete" onClick={() => removeReadingSession(session.id)} aria-label={`Remove ${sessionTitle.toLowerCase()}`}><Trash2 size={15} /></button></div>
                    </div>
                    <div className="reading-session-fields">
                      <label>Started<input type="date" value={session.startedAt ?? ""} onChange={(event) => updateReadingSession(session.id, { startedAt: event.target.value || undefined })} /></label>
                      <label>Finished<input type="date" value={session.finishedAt ?? ""} onChange={(event) => { const finishedAt = event.target.value || undefined; updateReadingSession(session.id, { finishedAt, progressPages: finishedAt && pages ? Number(pages) : session.progressPages }); if (finishedAt) setStatus("read"); }} /></label>
                      <label>Current page<input type="number" min="0" max={pages || undefined} value={session.progressPages ?? ""} onChange={(event) => updateReadingSession(session.id, { progressPages: event.target.value ? Number(event.target.value) : undefined })} placeholder={pages ? `of ${pages}` : "Page"} /></label>
                      <label className="session-note">Session note<input value={session.notes ?? ""} onChange={(event) => updateReadingSession(session.id, { notes: event.target.value || undefined })} placeholder="Optional note about this read" /></label>
                    </div>
                    {percent !== undefined && !session.finishedAt && <div className="reading-session-progress"><div className="progress-track"><span style={{ width: `${percent}%` }} /></div><small>{session.progressPages?.toLocaleString()} {pages ? `of ${Number(pages).toLocaleString()} pages` : "pages"}</small></div>}
                  </div>;
                })}
              </div>
            )}
          </div>

          <div className="form-grid text-fields">
            <label className="wide">Your review<textarea rows={5} value={review} onChange={(e) => setReview(e.target.value)} placeholder="Your thoughts on the book…" /></label>
            <label className="wide">Private notes<textarea rows={4} value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Edition notes, reminders, quotes to revisit…" /></label>
          </div>
        </div>

        <div className="form-actions">
          <button className="button secondary" type="button" onClick={onClose}>Cancel</button>
          <button className="button primary" disabled={saving || !title.trim() || !author.trim()}>{saving ? "Saving…" : "Save book"}</button>
        </div>
      </form>
    </div>
  );
}
