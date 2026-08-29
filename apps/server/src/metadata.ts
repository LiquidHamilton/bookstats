import type {
  BookFormat,
  MetadataCandidate,
  MetadataField,
  MetadataProvider,
  MetadataSourceRef,
  MetadataSeriesMembership,
  SeriesMetadata
} from "@bookstats/domain";

export interface MetadataProviderStatus {
  id: Exclude<MetadataProvider, "aggregate">;
  label: string;
  configured: boolean;
  role: string;
}

const GOOGLE_BOOKS_API = "https://www.googleapis.com/books/v1";
const HARDCOVER_API = "https://api.hardcover.app/v1/graphql";
const OPEN_LIBRARY_API = "https://openlibrary.org";

const metadataUserAgent = () => process.env.BOOKSTATS_METADATA_USER_AGENT ?? "BookStats/1.1.0 (local-development)";
const googleKey = () => process.env.BOOKSTATS_GOOGLE_BOOKS_API_KEY?.trim();
const hardcoverToken = () => process.env.BOOKSTATS_HARDCOVER_API_TOKEN?.trim();

export function metadataProviderStatuses(): MetadataProviderStatus[] {
  return [
    { id: "openlibrary", label: "Open Library", configured: true, role: "Open fallback and cover source" },
    { id: "googlebooks", label: "Google Books", configured: Boolean(googleKey()), role: "Edition metadata and descriptions" },
    { id: "hardcover", label: "Hardcover", configured: Boolean(hardcoverToken()), role: "Series and work intelligence" }
  ];
}

export function normalizeIsbn(value?: string): string | undefined {
  const isbn = value?.replace(/[^0-9Xx]/g, "").toUpperCase();
  return isbn && (isbn.length === 10 || isbn.length === 13) ? isbn : undefined;
}

export function isbn13From10(value: string): string | undefined {
  const isbn = normalizeIsbn(value);
  if (!isbn || isbn.length !== 10 || !/^\d{9}[\dX]$/.test(isbn)) return undefined;
  const body = `978${isbn.slice(0, 9)}`;
  let sum = 0;
  for (let index = 0; index < 12; index += 1) sum += Number(body[index]) * (index % 2 === 0 ? 1 : 3);
  return `${body}${(10 - (sum % 10)) % 10}`;
}

export function isbn10From13(value: string): string | undefined {
  const isbn = normalizeIsbn(value);
  if (!isbn || isbn.length !== 13 || !/^978\d{10}$/.test(isbn)) return undefined;
  const body = isbn.slice(3, 12);
  let sum = 0;
  for (let index = 0; index < 9; index += 1) sum += Number(body[index]) * (10 - index);
  const check = (11 - (sum % 11)) % 11;
  return `${body}${check === 10 ? "X" : check}`;
}

export function isbnEquivalent(left?: string, right?: string): boolean {
  const a = normalizeIsbn(left); const b = normalizeIsbn(right);
  if (!a || !b) return false;
  if (a === b) return true;
  return (a.length === 10 ? isbn13From10(a) : isbn10From13(a)) === b;
}

function preferredIsbn(values: Array<string | undefined>, requested?: string): string | undefined {
  const clean = [...new Set(values.map(normalizeIsbn).filter((value): value is string => Boolean(value)))];
  const wanted = normalizeIsbn(requested);
  if (wanted) {
    const exact = clean.find((value) => isbnEquivalent(value, wanted));
    if (exact) return exact.length === 13 ? exact : isbn13From10(exact) ?? exact;
  }
  return clean.find((value) => value.length === 13) ?? clean[0];
}

export async function searchAllMetadata(query: string, isbn?: string): Promise<MetadataCandidate[]> {
  const requestedIsbn = normalizeIsbn(isbn);
  const tasks: Array<Promise<MetadataCandidate[]>> = [
    requestedIsbn ? openLibraryByIsbn(requestedIsbn) : openLibrarySearch(query)
  ];
  if (googleKey()) tasks.push(requestedIsbn ? googleByIsbn(requestedIsbn) : googleSearch(query));
  if (hardcoverToken()) tasks.push(requestedIsbn ? hardcoverByIsbn(requestedIsbn) : hardcoverSearch(query));

  const settled = await Promise.allSettled(tasks);
  const candidates = settled.flatMap((result) => result.status === "fulfilled" ? result.value : []);
  if (requestedIsbn) {
    const exact = candidates.filter((candidate) => candidate.isbn && isbnEquivalent(candidate.isbn, requestedIsbn));
    if (exact.length === 0) return [];
    return [mergeCandidates(exact, requestedIsbn, true)];
  }
  return mergeSearchCandidates(candidates, query).slice(0, 24);
}

export async function getMetadataDetails(candidate: MetadataCandidate): Promise<MetadataCandidate> {
  const refs = candidate.sourceRefs?.length ? candidate.sourceRefs : legacyRefs(candidate);
  const details = await Promise.allSettled(refs.map((ref) => detailsForRef(ref, candidate.isbn)));
  const candidates = [candidate, ...details.flatMap((result) => result.status === "fulfilled" && result.value ? [result.value] : [])];
  const merged = mergeCandidates(candidates, candidate.exactEdition ? candidate.isbn : undefined, Boolean(candidate.exactEdition));
  if (merged.series && merged.seriesVolume && merged.seriesMetadata) return merged;
  const enrichment = await seriesEnrichmentFor(merged).catch(() => undefined);
  return enrichment ? applySeriesEnrichment(merged, enrichment) : merged;
}

function legacyRefs(candidate: MetadataCandidate): MetadataSourceRef[] {
  const provider = candidate.source === "aggregate" ? "openlibrary" : candidate.source;
  return [{ provider, workId: candidate.workId, editionId: candidate.editionId, exactIsbn: candidate.exactEdition ? candidate.isbn : undefined, sourceUrl: candidate.sourceUrl }];
}

async function detailsForRef(ref: MetadataSourceRef, isbn?: string): Promise<MetadataCandidate | undefined> {
  if (ref.provider === "openlibrary") return openLibraryDetails(ref, isbn);
  if (ref.provider === "googlebooks") return googleDetails(ref);
  return hardcoverDetails(ref);
}

function mergeSearchCandidates(candidates: MetadataCandidate[], query: string): MetadataCandidate[] {
  const groups = new Map<string, MetadataCandidate[]>();
  for (const candidate of candidates) {
    const key = candidate.isbn
      ? `isbn:${normalizeIsbn(candidate.isbn)}`
      : `text:${normalizeText(candidate.title)}|${normalizeText(candidate.author)}|${candidate.publicationYear ?? ""}`;
    groups.set(key, [...(groups.get(key) ?? []), candidate]);
  }
  return [...groups.values()]
    .map((group) => mergeCandidates(group, undefined, false))
    .sort((a, b) => searchRelevanceScore(b, query) - searchRelevanceScore(a, query) || (b.confidence ?? 0) - (a.confidence ?? 0));
}

const editionPriority: MetadataProvider[] = ["googlebooks", "hardcover", "openlibrary", "aggregate"];
const workPriority: MetadataProvider[] = ["hardcover", "googlebooks", "openlibrary", "aggregate"];

function dedupeSeriesMemberships(memberships: MetadataSeriesMembership[]): MetadataSeriesMembership[] {
  const byPair = new Map<string, MetadataSeriesMembership>();
  for (const membership of memberships) {
    if (!membership?.name?.trim()) continue;
    const normalizedName = normalizeText(membership.name);
    const normalizedVolume = normalizeText(membership.volume ?? "");
    const key = `${normalizedName}|${normalizedVolume}`;
    const existing = byPair.get(key);
    if (!existing) {
      byPair.set(key, { ...membership, name: membership.name.trim(), volume: membership.volume?.trim() || undefined });
      continue;
    }
    byPair.set(key, {
      ...existing,
      provider: existing.provider ?? membership.provider,
      seriesId: existing.seriesId ?? membership.seriesId,
      metadata: existing.metadata ?? membership.metadata
    });
  }
  return [...byPair.values()];
}

function candidateSeriesMemberships(candidate: MetadataCandidate): MetadataSeriesMembership[] {
  const memberships = [...(candidate.seriesMemberships ?? [])];
  const name = candidate.series ?? candidate.seriesMetadata?.name;
  if (name?.trim()) {
    memberships.unshift({
      provider: candidate.seriesMetadata?.provider ?? (candidate.source !== "aggregate" ? candidate.source : undefined),
      seriesId: candidate.seriesMetadata?.id,
      name,
      volume: candidate.seriesVolume,
      metadata: candidate.seriesMetadata
    });
  }
  return dedupeSeriesMemberships(memberships);
}

function preferredSeriesMembership(candidate: MetadataCandidate): MetadataSeriesMembership | undefined {
  const memberships = candidateSeriesMemberships(candidate);
  if (!memberships.length) return undefined;
  const preferredName = normalizeText(candidate.series ?? candidate.seriesMetadata?.name ?? "");
  return memberships.find((membership) => preferredName && normalizeText(membership.name) === preferredName) ?? memberships[0];
}

function sameSeriesName(left?: string, right?: string): boolean {
  return Boolean(left && right && normalizeText(left) === normalizeText(right));
}

export function mergeCandidates(candidates: MetadataCandidate[], requestedIsbn?: string, exactEdition = false): MetadataCandidate {
  const nonEmpty = candidates.filter((candidate) => candidate?.title);
  if (!nonEmpty.length) throw new Error("No metadata candidates were available to merge.");
  const sortedEdition = sortProviders(nonEmpty, editionPriority);
  const sortedWork = sortProviders(nonEmpty, workPriority);
  const fieldSources: Partial<Record<MetadataField, MetadataProvider>> = {};
  const choose = <K extends keyof MetadataCandidate>(field: K, list = sortedEdition): MetadataCandidate[K] | undefined => {
    const found = list.find((candidate) => hasValue(candidate[field]));
    if (found && metadataFieldNames.has(field as string)) fieldSources[field as MetadataField] = found.source === "aggregate" ? inferFieldProvider(found, field as MetadataField) : found.source;
    return found?.[field];
  };
  const refs = dedupeRefs(nonEmpty.flatMap((candidate) => candidate.sourceRefs?.length ? candidate.sourceRefs : legacyRefs(candidate)));
  const covers = uniqueStrings(sortedEdition.flatMap((candidate) => [...(candidate.coverUrls ?? []), ...(candidate.coverUrl ? [candidate.coverUrl] : [])]));
  const subjects = uniqueStrings(sortedWork.flatMap((candidate) => candidate.subjects ?? [])).slice(0, 30);
  const allSeriesMemberships = dedupeSeriesMemberships(sortedWork.flatMap(candidateSeriesMemberships));
  const seriesCandidate = sortedWork.find((candidate) => candidateSeriesMemberships(candidate).length > 0);
  const candidatePrimaryMembership = seriesCandidate ? preferredSeriesMembership(seriesCandidate) : undefined;
  const matchingMemberships = candidatePrimaryMembership
    ? allSeriesMemberships.filter((membership) => sameSeriesName(membership.name, candidatePrimaryMembership.name))
    : [];
  const primarySeriesMembership = candidatePrimaryMembership ? {
    ...candidatePrimaryMembership,
    volume: candidatePrimaryMembership.volume ?? matchingMemberships.find((membership) => membership.volume)?.volume,
    metadata: candidatePrimaryMembership.metadata ?? matchingMemberships.find((membership) => membership.metadata)?.metadata,
    provider: candidatePrimaryMembership.provider ?? matchingMemberships.find((membership) => membership.provider)?.provider,
    seriesId: candidatePrimaryMembership.seriesId ?? matchingMemberships.find((membership) => membership.seriesId)?.seriesId
  } : undefined;
  const primary = sortedEdition[0];
  const source: MetadataProvider = refs.length > 1 ? "aggregate" : refs[0]?.provider ?? primary.source;
  const isbn = preferredIsbn(nonEmpty.map((candidate) => candidate.isbn), requestedIsbn);
  const title = String(choose("title") ?? primary.title);
  const author = String(choose("author", sortedWork) ?? primary.author);
  const seriesName = primarySeriesMembership?.name;
  const seriesMetadata = primarySeriesMembership?.metadata;
  const seriesCatalogPosition = seriesMetadata?.books.find((book) => normalizeText(book.title) === normalizeText(title))?.position;
  const seriesVolume = primarySeriesMembership?.volume ?? seriesCatalogPosition;
  const seriesProvider = primarySeriesMembership?.provider ?? (seriesCandidate && seriesCandidate.source !== "aggregate" ? seriesCandidate.source : undefined);
  if (seriesName && seriesProvider) fieldSources.series = seriesProvider;
  if (seriesVolume && seriesProvider) fieldSources.seriesVolume = seriesProvider;
  const seriesMemberships = primarySeriesMembership
    ? dedupeSeriesMemberships([
        { ...primarySeriesMembership, volume: seriesVolume },
        ...allSeriesMemberships.filter((membership) => !sameSeriesName(membership.name, primarySeriesMembership.name) || normalizeText(membership.volume ?? "") !== normalizeText(seriesVolume ?? ""))
      ])
    : allSeriesMemberships;
  const additionalAuthors = uniqueStrings(nonEmpty.flatMap((candidate) => candidate.additionalAuthors ?? []).filter((name) => normalizeText(name) !== normalizeText(author)));
  const additionalAuthorSource = sortedWork.find((candidate) => candidate.additionalAuthors?.length);
  if (additionalAuthorSource) fieldSources.additionalAuthors = additionalAuthorSource.source === "aggregate" ? inferFieldProvider(additionalAuthorSource, "additionalAuthors") : additionalAuthorSource.source;
  const subjectSource = sortedWork.find((candidate) => candidate.subjects?.length);
  if (subjectSource) fieldSources.genre = subjectSource.source === "aggregate" ? inferFieldProvider(subjectSource, "genre") : subjectSource.source;
  if (isbn) fieldSources.isbn = sortedEdition.find((candidate) => candidate.isbn && isbnEquivalent(candidate.isbn, isbn))?.source ?? source;
  if (covers[0]) fieldSources.coverUrl = nonEmpty.find((candidate) => candidate.coverUrl === covers[0] || candidate.coverUrls?.includes(covers[0]))?.source ?? source;

  return {
    source,
    workId: primary.workId,
    editionId: primary.editionId,
    sourceRefs: refs,
    matchType: exactEdition ? "exact_isbn" : nonEmpty.some((candidate) => candidate.matchType === "edition") ? "edition" : "search",
    confidence: exactEdition ? 100 : Math.max(...nonEmpty.map((candidate) => candidate.confidence ?? 60)),
    exactEdition,
    title,
    author,
    additionalAuthors,
    isbn,
    publicationYear: choose("publicationYear"),
    publisher: choose("publisher"),
    language: choose("language"),
    pages: choose("pages"),
    format: choose("format"),
    series: seriesName,
    seriesVolume,
    seriesMetadata,
    seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined,
    subjects,
    description: choose("description", sortedWork),
    coverUrl: covers[0],
    coverUrls: covers,
    sourceUrl: primary.sourceUrl,
    fieldSources
  };
}

const metadataFieldNames = new Set<string>(["title", "author", "additionalAuthors", "isbn", "series", "seriesVolume", "publicationYear", "publisher", "language", "pages", "format", "genre", "description", "coverUrl"]);
function inferFieldProvider(candidate: MetadataCandidate, field: MetadataField): MetadataProvider { return candidate.fieldSources?.[field] ?? candidate.source; }
function sortProviders(candidates: MetadataCandidate[], priority: MetadataProvider[]): MetadataCandidate[] { return [...candidates].sort((a, b) => priority.indexOf(a.source) - priority.indexOf(b.source)); }
function hasValue(value: unknown): boolean { return value !== undefined && value !== null && value !== "" && (!Array.isArray(value) || value.length > 0); }
function uniqueStrings(values: Array<string | undefined | null>): string[] { return [...new Set(values.map((value) => value?.trim()).filter((value): value is string => Boolean(value)))]; }
function dedupeRefs(refs: MetadataSourceRef[]): MetadataSourceRef[] { const seen = new Set<string>(); return refs.filter((ref) => { const key = `${ref.provider}:${ref.workId}:${ref.editionId ?? ""}`; if (seen.has(key)) return false; seen.add(key); return true; }); }
function normalizeText(value?: string): string { return value?.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim() ?? ""; }
function searchRelevanceScore(candidate: MetadataCandidate, query: string): number {
  const wanted = normalizeText(query); const title = normalizeText(candidate.title); const author = normalizeText(candidate.author);
  if (!wanted) return candidate.confidence ?? 0;
  const wantedTokens = wanted.split(" ").filter(Boolean); const titleTokens = title.split(" ").filter(Boolean); const authorTokens = author.split(" ").filter(Boolean);
  const tokenCoverage = (needles: string[], haystack: string[]) => needles.length ? needles.filter((token) => haystack.includes(token)).length / needles.length : 0;
  let score = 0;
  if (title === wanted) score += 1000;
  else {
    if (title.startsWith(wanted)) score += 650;
    if (wanted.startsWith(title)) score += 575;
    score += tokenCoverage(wantedTokens, titleTokens) * 260;
    score += tokenCoverage(titleTokens, wantedTokens) * 340;
  }
  score += tokenCoverage(authorTokens, wantedTokens) * 130;
  score += tokenCoverage(wantedTokens, [...titleTokens, ...authorTokens]) * 120;
  const bundleTerms = ["box set", "boxed set", "trilogy", "omnibus", "collection", "complete series", "books 1 3", "books 1 2 3"];
  if (bundleTerms.some((term) => title.includes(term)) && !bundleTerms.some((term) => wanted.includes(term))) score -= 375;
  if (candidate.editionId) score += 20;
  if (candidate.isbn) score += 15;
  return score + (candidate.confidence ?? 0) * 0.25;
}
async function seriesEnrichmentFor(candidate: MetadataCandidate): Promise<MetadataCandidate | undefined> {
  const query = `${candidate.title} ${candidate.author}`.trim();
  let prospects: MetadataCandidate[] = [];
  if (hardcoverToken()) prospects = await hardcoverSearch(query).catch(() => []);
  if (!prospects.some((item) => item.series || item.seriesVolume || item.seriesMetadata)) prospects = await openLibrarySearch(query).catch(() => []);
  const best = prospects
    .filter((item) => item.series || item.seriesVolume || item.seriesMetadata)
    .map((item) => ({ item, score: workMatchScore(candidate, item) }))
    .filter(({ score }) => score >= 500)
    .sort((a, b) => b.score - a.score)[0]?.item;
  if (!best) return undefined;
  if (best.source === "hardcover") {
    const ref = best.sourceRefs?.find((item) => item.provider === "hardcover");
    if (ref) return await hardcoverDetails(ref).catch(() => best);
  }
  return best;
}
function workMatchScore(target: MetadataCandidate, candidate: MetadataCandidate): number {
  const leftTitle = normalizeWorkTitle(target.title); const rightTitle = normalizeWorkTitle(candidate.title);
  const leftAuthor = normalizeText(target.author); const rightAuthor = normalizeText(candidate.author);
  const leftTitleTokens = leftTitle.split(" ").filter(Boolean); const rightTitleTokens = rightTitle.split(" ").filter(Boolean);
  const coverage = (needles: string[], haystack: string[]) => needles.length ? needles.filter((token) => haystack.includes(token)).length / needles.length : 0;
  let score = leftTitle === rightTitle ? 500 : (coverage(leftTitleTokens, rightTitleTokens) + coverage(rightTitleTokens, leftTitleTokens)) * 190;
  const leftAuthorTokens = leftAuthor.split(" ").filter(Boolean); const rightAuthorTokens = rightAuthor.split(" ").filter(Boolean);
  const authorOverlap = Math.max(coverage(leftAuthorTokens, rightAuthorTokens), coverage(rightAuthorTokens, leftAuthorTokens));
  const usefulAuthor = (value: string) => Boolean(value && value !== "unknown author" && value !== "unknown");
  if (usefulAuthor(leftAuthor) && usefulAuthor(rightAuthor) && authorOverlap < 0.5) return 0;
  if (leftAuthor && rightAuthor) score += authorOverlap * 250;
  return score;
}
function normalizeWorkTitle(value: string): string {
  return normalizeText(value
    .replace(/\s*\([^)]*(?:#\s*\d|book\s+\d|vol(?:ume)?\.?\s*\d|series)[^)]*\)\s*$/i, "")
    .replace(/\s*:\s*(?:book|volume|vol\.?|part)\s+\d.*$/i, ""));
}
function applySeriesEnrichment(target: MetadataCandidate, enrichment: MetadataCandidate): MetadataCandidate {
  const incomingSeries = enrichment.series ?? enrichment.seriesMetadata?.name;
  const series = target.series ?? incomingSeries;
  const compatible = !target.series || !incomingSeries || seriesNamesCompatible(target.series, incomingSeries);
  const seriesVolume = target.seriesVolume ?? (compatible ? enrichment.seriesVolume : undefined);
  const seriesMetadata = target.seriesMetadata ?? (compatible ? enrichment.seriesMetadata : undefined);
  const seriesMemberships = dedupeSeriesMemberships([
    ...candidateSeriesMemberships(target),
    ...candidateSeriesMemberships(enrichment)
  ]);
  const fieldSources = { ...(target.fieldSources ?? {}) };
  if (!target.series && series) fieldSources.series = enrichment.fieldSources?.series ?? enrichment.source;
  if (!target.seriesVolume && seriesVolume) fieldSources.seriesVolume = enrichment.fieldSources?.seriesVolume ?? enrichment.source;
  return { ...target, series, seriesVolume, seriesMetadata, seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined, fieldSources };
}
function seriesNamesCompatible(left: string, right: string): boolean {
  const normalize = (value: string) => normalizeText(value).replace(/^the\s+/, "");
  const a = normalize(left); const b = normalize(right);
  return a === b || a.includes(b) || b.includes(a);
}
function yearFromDate(value?: string): number | undefined { const match = value?.match(/\b(1[4-9]\d{2}|20\d{2}|21\d{2})\b/); return match ? Number(match[1]) : undefined; }
function safeNumber(value: unknown): number | undefined { const number = Number(value); return Number.isFinite(number) && number > 0 ? number : undefined; }
function sourceRef(provider: MetadataSourceRef["provider"], workId: string, editionId?: string, exactIsbn?: string, sourceUrl?: string): MetadataSourceRef { return { provider, workId, editionId, exactIsbn, sourceUrl }; }

function googleSeriesPosition(entry?: { orderNumber?: number; bookDisplayNumber?: string }): string | undefined {
  // Google sometimes gives a single numeric orderNumber for an omnibus while
  // bookDisplayNumber contains the meaningful range (for example "Books 1-6").
  // Prefer that display value so multi-volume records remain multi-volume.
  const display = entry?.bookDisplayNumber?.trim().replace(/^(?:books?|vol(?:ume)?s?\.?|#)\s*/i, "").trim();
  return display || (entry?.orderNumber !== undefined ? String(entry.orderNumber) : undefined);
}

// ---------- Google Books ----------
interface GoogleVolumesResponse { items?: GoogleVolume[]; }
interface GoogleVolume {
  id: string;
  volumeInfo?: {
    title?: string; subtitle?: string; authors?: string[]; publisher?: string; publishedDate?: string; description?: string;
    industryIdentifiers?: Array<{ type?: string; identifier?: string }>; pageCount?: number; printType?: string; categories?: string[]; language?: string;
    imageLinks?: Record<string, string | undefined>;
    seriesInfo?: { volumeSeries?: Array<{ seriesId?: string; seriesBookType?: string; orderNumber?: number; bookDisplayNumber?: string }> };
    infoLink?: string;
  };
}
interface GoogleSeriesResponse { series?: Array<{ seriesId?: string; title?: string; isComplete?: boolean }>; }
interface GoogleSeriesMembershipResponse { member?: GoogleVolume[]; }

async function googleSearch(query: string): Promise<MetadataCandidate[]> {
  const url = new URL(`${GOOGLE_BOOKS_API}/volumes`);
  url.searchParams.set("q", query); url.searchParams.set("maxResults", "12"); url.searchParams.set("printType", "books"); url.searchParams.set("projection", "full");
  appendGoogleKey(url);
  const data = await jsonFetch<GoogleVolumesResponse>(url.toString());
  return (data.items ?? []).map((volume) => googleCandidate(volume, false)).filter((candidate): candidate is MetadataCandidate => Boolean(candidate));
}

async function googleByIsbn(isbn: string): Promise<MetadataCandidate[]> {
  const url = new URL(`${GOOGLE_BOOKS_API}/volumes`);
  url.searchParams.set("q", `isbn:${isbn}`); url.searchParams.set("maxResults", "10"); url.searchParams.set("printType", "books"); url.searchParams.set("projection", "full");
  appendGoogleKey(url);
  const data = await jsonFetch<GoogleVolumesResponse>(url.toString());
  return (data.items ?? []).filter((volume) => googleIdentifiers(volume).some((value) => isbnEquivalent(value, isbn))).map((volume) => googleCandidate(volume, true)).filter((candidate): candidate is MetadataCandidate => Boolean(candidate));
}

async function googleDetails(ref: MetadataSourceRef): Promise<MetadataCandidate | undefined> {
  const id = ref.editionId ?? ref.workId;
  const url = new URL(`${GOOGLE_BOOKS_API}/volumes/${encodeURIComponent(id)}`);
  url.searchParams.set("projection", "full"); url.searchParams.set("includeNonComicsSeries", "true"); appendGoogleKey(url);
  const volume = await jsonFetch<GoogleVolume>(url.toString());
  const candidate = googleCandidate(volume, Boolean(ref.exactIsbn));
  if (!candidate) return undefined;
  const seriesEntries = (volume.volumeInfo?.seriesInfo?.volumeSeries ?? []).filter((entry) => entry.seriesId);
  if (seriesEntries.length) {
    const resolved = await Promise.all(seriesEntries.map(async (entry) => ({
      entry,
      metadata: await googleSeriesMetadata(entry.seriesId!).catch(() => undefined)
    })));
    const memberships = resolved.flatMap(({ entry, metadata }) => metadata ? [{
      provider: "googlebooks" as const,
      seriesId: metadata.id,
      name: metadata.name,
      volume: googleSeriesPosition(entry),
      metadata
    }] : []);
    if (memberships.length) {
      candidate.seriesMemberships = dedupeSeriesMemberships(memberships);
      const primary = candidate.seriesMemberships[0];
      candidate.seriesMetadata = primary.metadata;
      candidate.series = primary.name;
      candidate.seriesVolume = primary.volume;
    }
  }
  return candidate;
}

function googleCandidate(volume: GoogleVolume, exactEdition: boolean): MetadataCandidate | undefined {
  const info = volume.volumeInfo; if (!info?.title) return undefined;
  const authors = info.authors ?? [];
  const isbn = preferredIsbn(googleIdentifiers(volume));
  const seriesEntry = info.seriesInfo?.volumeSeries?.[0];
  const coverUrls = uniqueStrings([info.imageLinks?.extraLarge, info.imageLinks?.large, info.imageLinks?.medium, info.imageLinks?.small, info.imageLinks?.thumbnail, info.imageLinks?.smallThumbnail].map(upgradeGoogleCover));
  const fieldSources = providerFieldSources("googlebooks", ["title", "author", "additionalAuthors", "isbn", "publicationYear", "publisher", "language", "pages", "description", "coverUrl"]);
  if (seriesEntry) { fieldSources.series = "googlebooks"; fieldSources.seriesVolume = "googlebooks"; }
  return {
    source: "googlebooks",
    workId: volume.id,
    editionId: volume.id,
    sourceRefs: [sourceRef("googlebooks", volume.id, volume.id, exactEdition ? isbn : undefined, info.infoLink)],
    matchType: exactEdition ? "exact_isbn" : "edition",
    confidence: exactEdition ? 100 : 82,
    exactEdition,
    title: info.title,
    author: authors[0] ?? "Unknown author",
    additionalAuthors: authors.slice(1),
    isbn,
    publicationYear: yearFromDate(info.publishedDate),
    publisher: info.publisher,
    language: info.language,
    pages: safeNumber(info.pageCount),
    seriesVolume: googleSeriesPosition(seriesEntry),
    subjects: info.categories ?? [],
    description: info.description,
    coverUrl: coverUrls[0],
    coverUrls,
    sourceUrl: info.infoLink,
    fieldSources
  };
}
function googleIdentifiers(volume: GoogleVolume): string[] { return (volume.volumeInfo?.industryIdentifiers ?? []).map((entry) => normalizeIsbn(entry.identifier)).filter((value): value is string => Boolean(value)); }
function appendGoogleKey(url: URL) { const key = googleKey(); if (key) url.searchParams.set("key", key); }
function upgradeGoogleCover(value?: string): string | undefined { return value?.replace(/^http:/, "https:").replace(/&zoom=\d/, "&zoom=2"); }

async function googleSeriesMetadata(seriesId: string): Promise<SeriesMetadata | undefined> {
  const seriesUrl = new URL(`${GOOGLE_BOOKS_API}/series/get`); seriesUrl.searchParams.set("series_id", seriesId); appendGoogleKey(seriesUrl);
  const membershipUrl = new URL(`${GOOGLE_BOOKS_API}/series/membership/get`); membershipUrl.searchParams.set("series_id", seriesId); membershipUrl.searchParams.set("page_size", "40"); appendGoogleKey(membershipUrl);
  const [seriesData, membership] = await Promise.all([jsonFetch<GoogleSeriesResponse>(seriesUrl.toString()), jsonFetch<GoogleSeriesMembershipResponse>(membershipUrl.toString())]);
  const series = seriesData.series?.[0];
  const books = (membership.member ?? []).map((volume) => {
    const info = volume.volumeInfo; const item = info?.seriesInfo?.volumeSeries?.find((entry) => entry.seriesId === seriesId) ?? info?.seriesInfo?.volumeSeries?.[0];
    return { providerId: volume.id, title: info?.title ?? "Untitled", position: googleSeriesPosition(item), isbn: preferredIsbn(googleIdentifiers(volume)), coverUrl: upgradeGoogleCover(info?.imageLinks?.thumbnail), author: info?.authors?.[0] };
  }).sort(seriesBookSort);
  if (!series?.title && books.length === 0) return undefined;
  return { provider: "googlebooks", id: seriesId, name: series?.title ?? "Series", totalBooks: books.length || undefined, isCompleted: series?.isComplete, books, updatedAt: new Date().toISOString() };
}

// ---------- Hardcover ----------
type JsonRecord = Record<string, unknown>;
interface HardcoverGraphqlResponse<T> { data?: T; errors?: Array<{ message?: string }>; }
interface HcImage { url?: string | null; }
interface HcAuthor { name?: string | null; }
interface HcContribution { author?: HcAuthor | null; contribution?: string | null; }
interface HcSeries { id: number; name: string; books_count?: number | null; primary_books_count?: number | null; is_completed?: boolean | null; }
interface HcBookSeries { position?: number | null; featured?: boolean | null; compilation?: boolean | null; series?: HcSeries | null; }
interface HcBook {
  id: number; title?: string | null; subtitle?: string | null; description?: string | null; pages?: number | null; release_year?: number | null; image?: HcImage | null;
  contributions?: HcContribution[]; book_series?: HcBookSeries[]; featured_book_series?: HcBookSeries | null;
  editions?: HcEdition[];
}
interface HcEdition {
  id: number; book_id: number; title?: string | null; subtitle?: string | null; isbn_10?: string | null; isbn_13?: string | null; pages?: number | null;
  physical_format?: string | null; edition_format?: string | null; release_date?: string | null; release_year?: number | null; image?: HcImage | null; images?: HcImage[];
  language?: { language?: string | null } | null; publisher?: { name?: string | null } | null; book: HcBook;
}

async function hardcoverByIsbn(isbn: string): Promise<MetadataCandidate[]> {
  const normalized = normalizeIsbn(isbn)!;
  const isbn13 = normalized.length === 13 ? normalized : isbn13From10(normalized) ?? normalized;
  const isbn10 = normalized.length === 10 ? normalized : isbn10From13(normalized) ?? normalized;
  const query = `query ExactEdition($isbn13: String!, $isbn10: String!) {
    editions(where: {_or: [{isbn_13: {_eq: $isbn13}}, {isbn_10: {_eq: $isbn10}}]}, limit: 10) {
      id book_id title subtitle isbn_10 isbn_13 pages physical_format edition_format release_date release_year
      image { url } images(limit: 8) { url } language { language } publisher { name }
      book { id title subtitle description pages release_year image { url }
        contributions { contribution author { name } }
        featured_book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
        book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
      }
    }
  }`;
  const data = await hardcoverGraphql<{ editions: HcEdition[] }>(query, { isbn13, isbn10 });
  return (data.editions ?? []).filter((edition) => [edition.isbn_13, edition.isbn_10].some((value) => isbnEquivalent(value ?? undefined, normalized))).map((edition) => hardcoverEditionCandidate(edition, true));
}

async function hardcoverSearch(queryText: string): Promise<MetadataCandidate[]> {
  const searchQuery = `query Search($query: String!) { search(query: $query, per_page: 12) { ids results } }`;
  const searched = await hardcoverGraphql<{ search?: { ids?: number[] | null; results?: unknown } }>(searchQuery, { query: queryText });
  const ids = (searched.search?.ids ?? []).filter((id): id is number => Number.isInteger(id)).slice(0, 12);
  if (!ids.length) return [];
  const query = `query Books($ids: [Int!]!) { books(where: {id: {_in: $ids}}, limit: 12) {
    id title subtitle description pages release_year image { url } contributions { contribution author { name } }
    editions(limit: 8, order_by: {users_count: desc}) { id book_id title isbn_10 isbn_13 pages physical_format edition_format release_date release_year image { url } language { language } publisher { name } }
    featured_book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
    book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
  } }`;
  const data = await hardcoverGraphql<{ books: HcBook[] }>(query, { ids });
  const searchOrder = new Map(ids.map((id, index) => [id, index]));
  return [...(data.books ?? [])]
    .sort((a, b) => (searchOrder.get(a.id) ?? Number.MAX_SAFE_INTEGER) - (searchOrder.get(b.id) ?? Number.MAX_SAFE_INTEGER))
    .map(hardcoverBookCandidate)
    .filter((candidate): candidate is MetadataCandidate => Boolean(candidate));
}

async function hardcoverDetails(ref: MetadataSourceRef): Promise<MetadataCandidate | undefined> {
  const editionId = Number(ref.editionId); const bookId = Number(ref.workId);
  let candidate: MetadataCandidate | undefined;
  if (Number.isInteger(editionId) && editionId > 0) {
    const query = `query Edition($id: Int!) { editions_by_pk(id: $id) {
      id book_id title subtitle isbn_10 isbn_13 pages physical_format edition_format release_date release_year
      image { url } images(limit: 12) { url } language { language } publisher { name }
      book { id title subtitle description pages release_year image { url } contributions { contribution author { name } }
        featured_book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
        book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
      }
    } }`;
    const data = await hardcoverGraphql<{ editions_by_pk?: HcEdition | null }>(query, { id: editionId });
    if (data.editions_by_pk) candidate = hardcoverEditionCandidate(data.editions_by_pk, Boolean(ref.exactIsbn));
  } else if (Number.isInteger(bookId) && bookId > 0) {
    const query = `query Book($id: Int!) { books_by_pk(id: $id) { id title subtitle description pages release_year image { url }
      contributions { contribution author { name } }
      editions(limit: 12, order_by: {users_count: desc}) { id book_id title isbn_10 isbn_13 pages physical_format edition_format release_date release_year image { url } language { language } publisher { name } }
      featured_book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
      book_series { position featured compilation series { id name books_count primary_books_count is_completed } }
    } }`;
    const data = await hardcoverGraphql<{ books_by_pk?: HcBook | null }>(query, { id: bookId });
    if (data.books_by_pk) candidate = hardcoverBookCandidate(data.books_by_pk);
  }
  if (!candidate) return undefined;
  const series = chooseHardcoverSeriesFromCandidate(candidate);
  if (series) candidate.seriesMetadata = await hardcoverSeriesMetadata(series.id).catch(() => candidate?.seriesMetadata);
  if (candidate.seriesMetadata) {
    candidate.series = candidate.seriesMetadata.name;
    candidate.seriesMemberships = dedupeSeriesMemberships((candidate.seriesMemberships ?? []).map((membership) =>
      membership.provider === "hardcover" && membership.seriesId === candidate.seriesMetadata?.id
        ? { ...membership, name: candidate.seriesMetadata!.name, metadata: candidate.seriesMetadata }
        : membership
    ));
  }
  return candidate;
}

function hardcoverEditionCandidate(edition: HcEdition, exactEdition: boolean): MetadataCandidate {
  const book = edition.book;
  const authors = hardcoverAuthors(book.contributions);
  const seriesLink = chooseHardcoverSeries(book);
  const seriesMemberships = hardcoverSeriesMemberships(book);
  const coverUrls = uniqueStrings([edition.image?.url, ...(edition.images ?? []).map((image) => image.url), book.image?.url]);
  const format = mapBookFormat(edition.physical_format ?? edition.edition_format);
  const isbn = preferredIsbn([edition.isbn_13 ?? undefined, edition.isbn_10 ?? undefined]);
  const fields = providerFieldSources("hardcover", ["title", "author", "additionalAuthors", "isbn", "publicationYear", "publisher", "language", "pages", "format", "description", "coverUrl", "series", "seriesVolume"]);
  return {
    source: "hardcover", workId: String(book.id), editionId: String(edition.id),
    sourceRefs: [sourceRef("hardcover", String(book.id), String(edition.id), exactEdition ? isbn : undefined, `https://hardcover.app/books/${book.id}`)],
    matchType: exactEdition ? "exact_isbn" : "edition", confidence: exactEdition ? 100 : 84, exactEdition,
    title: edition.title ?? book.title ?? "Untitled", author: authors[0] ?? "Unknown author", additionalAuthors: authors.slice(1), isbn,
    publicationYear: edition.release_year ?? yearFromDate(edition.release_date ?? undefined) ?? book.release_year ?? undefined,
    publisher: edition.publisher?.name ?? undefined, language: edition.language?.language ?? undefined, pages: edition.pages ?? book.pages ?? undefined, format,
    series: seriesLink?.series?.name, seriesVolume: seriesLink?.position !== undefined && seriesLink?.position !== null ? String(seriesLink.position) : undefined,
    seriesMetadata: seriesLink?.series ? hardcoverSeriesMetadataStub(seriesLink.series) : undefined,
    seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined,
    subjects: [], description: book.description ?? undefined, coverUrl: coverUrls[0], coverUrls, sourceUrl: `https://hardcover.app/books/${book.id}`, fieldSources: fields
  };
}
function hardcoverBookCandidate(book: HcBook): MetadataCandidate | undefined {
  if (!book.title) return undefined;
  const edition = book.editions?.[0];
  if (edition) { edition.book = book; return hardcoverEditionCandidate(edition, false); }
  const authors = hardcoverAuthors(book.contributions); const seriesLink = chooseHardcoverSeries(book); const cover = book.image?.url ?? undefined;
  const seriesMemberships = hardcoverSeriesMemberships(book);
  return { source: "hardcover", workId: String(book.id), sourceRefs: [sourceRef("hardcover", String(book.id), undefined, undefined, `https://hardcover.app/books/${book.id}`)], matchType: "work", confidence: 75, title: book.title, author: authors[0] ?? "Unknown author", additionalAuthors: authors.slice(1), publicationYear: book.release_year ?? undefined, pages: book.pages ?? undefined, series: seriesLink?.series?.name, seriesVolume: seriesLink?.position != null ? String(seriesLink.position) : undefined, seriesMetadata: seriesLink?.series ? hardcoverSeriesMetadataStub(seriesLink.series) : undefined, seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined, subjects: [], description: book.description ?? undefined, coverUrl: cover, coverUrls: cover ? [cover] : [], sourceUrl: `https://hardcover.app/books/${book.id}`, fieldSources: providerFieldSources("hardcover", ["title", "author", "additionalAuthors", "publicationYear", "pages", "series", "seriesVolume", "description", "coverUrl"]) };
}
function hardcoverAuthors(contributions?: HcContribution[]): string[] { const named = uniqueStrings((contributions ?? []).map((item) => item.author?.name)); const authorRole = uniqueStrings((contributions ?? []).filter((item) => !item.contribution || /author|writer/i.test(item.contribution)).map((item) => item.author?.name)); return authorRole.length ? authorRole : named; }
function hardcoverSeriesMetadataStub(series: HcSeries): SeriesMetadata {
  return { provider: "hardcover", id: String(series.id), name: series.name, totalBooks: series.books_count ?? undefined, primaryBooksCount: series.primary_books_count ?? undefined, isCompleted: series.is_completed ?? undefined, books: [] };
}
function rankedHardcoverSeries(book: HcBook): HcBookSeries[] {
  const entries = [book.featured_book_series, ...(book.book_series ?? [])]
    .filter((entry): entry is HcBookSeries => Boolean(entry?.series && !entry.compilation));
  const unique = [...new Map(entries.map((entry) => [entry.series!.id, entry] as const)).values()];
  return unique.sort((a, b) => {
    const aCount = a.series?.primary_books_count ?? a.series?.books_count ?? Number.MAX_SAFE_INTEGER;
    const bCount = b.series?.primary_books_count ?? b.series?.books_count ?? Number.MAX_SAFE_INTEGER;
    const aFinite = Number.isInteger(aCount) && aCount >= 2 && aCount <= 500;
    const bFinite = Number.isInteger(bCount) && bCount >= 2 && bCount <= 500;
    const aPositioned = Number.isFinite(Number(a.position)) && Number(a.position) > 0;
    const bPositioned = Number.isFinite(Number(b.position)) && Number(b.position) > 0;
    if (aPositioned !== bPositioned) return aPositioned ? -1 : 1;
    if (aFinite !== bFinite) return aFinite ? -1 : 1;
    if (aFinite && bFinite && aCount !== bCount) return aCount - bCount;
    if (Boolean(a.featured) !== Boolean(b.featured)) return a.featured ? -1 : 1;
    return (a.series?.name ?? "").localeCompare(b.series?.name ?? "");
  });
}
function hardcoverSeriesMemberships(book: HcBook): MetadataSeriesMembership[] {
  return rankedHardcoverSeries(book).flatMap((entry) => entry.series ? [{
    provider: "hardcover" as const,
    seriesId: String(entry.series.id),
    name: entry.series.name,
    volume: entry.position != null ? String(entry.position) : undefined,
    metadata: hardcoverSeriesMetadataStub(entry.series)
  }] : []);
}
function chooseHardcoverSeries(book: HcBook): HcBookSeries | undefined {
  const ranked = rankedHardcoverSeries(book);
  return ranked[0] ?? book.featured_book_series ?? book.book_series?.[0];
}
function chooseHardcoverSeriesFromCandidate(candidate: MetadataCandidate): { id: number } | undefined { const id = candidate.seriesMetadata?.provider === "hardcover" ? Number(candidate.seriesMetadata.id) : NaN; return Number.isInteger(id) ? { id } : undefined; }

async function hardcoverSeriesMetadata(seriesId: number): Promise<SeriesMetadata | undefined> {
  const query = `query Series($id: Int!) { series(where: {id: {_eq: $id}}, limit: 1) {
    id name books_count primary_books_count is_completed
    book_series(order_by: {position: asc}) { position featured compilation book { id title image { url } contributions { contribution author { name } } editions(limit: 6, order_by: {users_count: desc}) { isbn_13 isbn_10 image { url } } } }
  } }`;
  const data = await hardcoverGraphql<{ series: Array<HcSeries & { book_series: Array<HcBookSeries & { book?: HcBook | null }> }> }>(query, { id: seriesId });
  const series = data.series?.[0]; if (!series) return undefined;
  const primaryRows = (series.book_series ?? []).filter((row) => !row.compilation && row.book);
  const mappedBooks = primaryRows.map((row) => ({ providerId: String(row.book!.id), title: row.book!.title ?? "Untitled", position: row.position != null ? String(row.position) : undefined, isbn: preferredIsbn((row.book!.editions ?? []).flatMap((edition) => [edition.isbn_13 ?? undefined, edition.isbn_10 ?? undefined])), coverUrl: row.book!.image?.url ?? row.book!.editions?.find((edition) => edition.image?.url)?.image?.url ?? undefined, author: hardcoverAuthors(row.book!.contributions)[0], featured: Boolean(row.featured) }));
  const expectedPrimary = series.primary_books_count ?? undefined;
  const books = compactPrimarySeriesBooks(mappedBooks, expectedPrimary).map(({ featured: _featured, ...book }) => book).sort(seriesBookSort);
  // books_count can include translated/alternate work records. Completion tools care
  // about the mainline sequence, so expose the cleaned catalog size while retaining
  // primary_books_count as the provider hint that produced it.
  return { provider: "hardcover", id: String(series.id), name: series.name, totalBooks: books.length || expectedPrimary || series.books_count || undefined, primaryBooksCount: expectedPrimary ?? books.length, isCompleted: series.is_completed ?? undefined, books, updatedAt: new Date().toISOString() };
}

let lastHardcoverRequest = 0;
async function hardcoverGraphql<T>(query: string, variables: Record<string, unknown>): Promise<T> {
  const token = hardcoverToken(); if (!token) throw new Error("Hardcover is not configured.");
  const wait = Math.max(0, 1100 - (Date.now() - lastHardcoverRequest)); if (wait) await new Promise((resolve) => setTimeout(resolve, wait)); lastHardcoverRequest = Date.now();
  const response = await fetch(HARDCOVER_API, { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json", "User-Agent": metadataUserAgent() }, body: JSON.stringify({ query, variables }) });
  const payload = await response.json() as HardcoverGraphqlResponse<T>;
  if (!response.ok || payload.errors?.length || !payload.data) throw new Error(`Hardcover request failed (${response.status}): ${payload.errors?.map((error) => error.message).filter(Boolean).join("; ") || response.statusText}`);
  return payload.data;
}

// ---------- Open Library ----------
interface OpenLibrarySearchResponse { docs: OpenLibrarySearchDoc[]; }
interface OpenLibraryEditionSearchDoc { key?: string; title?: string; isbn?: string[]; cover_i?: number; publish_date?: string[]; publish_year?: number[]; publisher?: string[]; language?: string[]; number_of_pages?: number; physical_format?: string; series?: string[]; }
interface OpenLibrarySearchDoc { key?: string; title?: string; author_name?: string[]; first_publish_year?: number; isbn?: string[]; cover_i?: number; edition_key?: string[]; number_of_pages_median?: number; publisher?: string[]; language?: string[]; subject?: string[]; series?: string[]; editions?: { docs?: OpenLibraryEditionSearchDoc[] }; }
interface OpenLibraryEdition { key?: string; title?: string; subtitle?: string; number_of_pages?: number; publishers?: string[]; publish_date?: string; isbn_13?: string[]; isbn_10?: string[]; covers?: number[]; series?: string[]; physical_format?: string; languages?: Array<{ key?: string }>; works?: Array<{ key?: string }>; authors?: Array<{ key?: string }>; }
interface OpenLibraryWork { title?: string; description?: string | { value?: string }; subjects?: string[]; series?: string[]; authors?: Array<{ author?: { key?: string } }>; }
interface OpenLibraryAuthor { name?: string; }
interface OpenLibraryEditionsResponse { entries: Array<{ covers?: number[] }>; }

async function openLibrarySearch(query: string): Promise<MetadataCandidate[]> {
  const url = new URL(`${OPEN_LIBRARY_API}/search.json`); url.searchParams.set("q", query); url.searchParams.set("limit", "12"); url.searchParams.set("fields", "key,title,author_name,first_publish_year,isbn,cover_i,edition_key,number_of_pages_median,publisher,language,subject,series,editions,editions.key,editions.title,editions.isbn,editions.cover_i,editions.publish_date,editions.publish_year,editions.publisher,editions.language,editions.number_of_pages,editions.physical_format,editions.series");
  const data = await openLibraryFetch<OpenLibrarySearchResponse>(url.toString());
  return data.docs.map(candidateFromOpenLibrarySearch).filter((candidate): candidate is MetadataCandidate => Boolean(candidate));
}
async function openLibraryByIsbn(isbn: string): Promise<MetadataCandidate[]> {
  try {
    const edition = await openLibraryFetch<OpenLibraryEdition>(`${OPEN_LIBRARY_API}/isbn/${encodeURIComponent(isbn)}.json`);
    const ref = sourceRef("openlibrary", sanitizeOlid(edition.works?.[0]?.key, "W") ?? "", sanitizeOlid(edition.key, "M"), isbn, sanitizeOlid(edition.key, "M") ? `${OPEN_LIBRARY_API}/books/${sanitizeOlid(edition.key, "M")}` : undefined);
    const candidate = await openLibraryCandidateFromEdition(edition, ref, isbn);
    return candidate && candidate.isbn && isbnEquivalent(candidate.isbn, isbn) ? [candidate] : [];
  } catch {
    // ISBN endpoint can be missing even when the search index knows the edition.
    const url = new URL(`${OPEN_LIBRARY_API}/search.json`); url.searchParams.set("isbn", isbn); url.searchParams.set("limit", "5"); url.searchParams.set("fields", "key,title,author_name,first_publish_year,isbn,cover_i,edition_key,number_of_pages_median,publisher,language,subject,series");
    const data = await openLibraryFetch<OpenLibrarySearchResponse>(url.toString());
    // Search-index rows are work-centric and cannot safely prove edition-specific publisher,
    // pagination, binding, language, cover, or publication-date fields. Keep only fields that
    // are safe at the work/identifier level so an ISBN lookup never substitutes another edition.
    return data.docs.map(candidateFromOpenLibrarySearch)
      .filter((candidate): candidate is MetadataCandidate => Boolean(candidate?.isbn && isbnEquivalent(candidate.isbn, isbn)))
      .map((candidate) => ({
        ...candidate,
        editionId: undefined,
        exactEdition: true,
        matchType: "exact_isbn" as const,
        confidence: 88,
        publicationYear: undefined,
        publisher: undefined,
        language: undefined,
        pages: undefined,
        format: undefined,
        coverUrl: undefined,
        coverUrls: [],
        sourceRefs: candidate.sourceRefs?.map((ref) => ({ ...ref, editionId: undefined, exactIsbn: isbn })),
        fieldSources: providerFieldSources("openlibrary", ["title", "author", "additionalAuthors", "isbn", "series", "seriesVolume"])
      }));
  }
}
async function openLibraryDetails(ref: MetadataSourceRef, isbn?: string): Promise<MetadataCandidate | undefined> {
  const editionId = sanitizeOlid(ref.editionId, "M"); const workId = sanitizeOlid(ref.workId, "W");
  if (editionId) { const edition = await openLibraryFetch<OpenLibraryEdition>(`${OPEN_LIBRARY_API}/books/${editionId}.json`); return openLibraryCandidateFromEdition(edition, { ...ref, workId: workId ?? sanitizeOlid(edition.works?.[0]?.key, "W") ?? "", editionId }, ref.exactIsbn ?? isbn); }
  if (!workId) return undefined;
  const work = await openLibraryFetch<OpenLibraryWork>(`${OPEN_LIBRARY_API}/works/${workId}.json`); const authors = await openLibraryAuthorNames(work); const seriesMemberships = openLibrarySeriesMemberships(work.series); const seriesInfo = seriesMemberships[0];
  return { source: "openlibrary", workId, sourceRefs: [sourceRef("openlibrary", workId, undefined, undefined, `${OPEN_LIBRARY_API}/works/${workId}`)], matchType: "work", confidence: 65, title: work.title ?? "Untitled", author: authors[0] ?? "Unknown author", additionalAuthors: authors.slice(1), series: seriesInfo?.name, seriesVolume: seriesInfo?.volume, seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined, subjects: work.subjects ?? [], description: typeof work.description === "string" ? work.description : work.description?.value, sourceUrl: `${OPEN_LIBRARY_API}/works/${workId}`, fieldSources: providerFieldSources("openlibrary", ["title", "author", "additionalAuthors", "series", "seriesVolume", "description"]) };
}
async function openLibraryCandidateFromEdition(edition: OpenLibraryEdition, ref: MetadataSourceRef, requestedIsbn?: string): Promise<MetadataCandidate | undefined> {
  const workId = sanitizeOlid(ref.workId, "W") ?? sanitizeOlid(edition.works?.[0]?.key, "W"); const editionId = sanitizeOlid(edition.key, "M") ?? sanitizeOlid(ref.editionId, "M");
  const work = workId ? await openLibraryFetch<OpenLibraryWork>(`${OPEN_LIBRARY_API}/works/${workId}.json`).catch(() => undefined) : undefined;
  const authors = uniqueStrings(edition.authors?.length ? await Promise.all(edition.authors.map((author) => openLibraryAuthorName(author.key))) : work ? await openLibraryAuthorNames(work) : []);
  const isbn = preferredIsbn([...(edition.isbn_13 ?? []), ...(edition.isbn_10 ?? [])], requestedIsbn); const seriesMemberships = openLibrarySeriesMemberships([...(edition.series ?? []), ...(work?.series ?? [])]); const seriesInfo = seriesMemberships[0];
  const covers = uniqueStrings((edition.covers ?? []).filter((id) => id > 0).map(coverUrl)); const exact = Boolean(requestedIsbn && isbn && isbnEquivalent(isbn, requestedIsbn));
  return { source: "openlibrary", workId: workId ?? editionId ?? "unknown", editionId, sourceRefs: [sourceRef("openlibrary", workId ?? editionId ?? "unknown", editionId, exact ? requestedIsbn : undefined, editionId ? `${OPEN_LIBRARY_API}/books/${editionId}` : undefined)], matchType: exact ? "exact_isbn" : "edition", confidence: exact ? 96 : 75, exactEdition: exact, title: edition.title ?? work?.title ?? "Untitled", author: authors[0] ?? "Unknown author", additionalAuthors: authors.slice(1), isbn, publicationYear: yearFromDate(edition.publish_date), publisher: edition.publishers?.[0], language: languageFromOpenLibrary(edition.languages?.[0]?.key), pages: edition.number_of_pages, format: mapBookFormat(edition.physical_format), series: seriesInfo?.name, seriesVolume: seriesInfo?.volume, seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined, subjects: work?.subjects ?? [], description: typeof work?.description === "string" ? work.description : work?.description?.value, coverUrl: covers[0], coverUrls: covers, sourceUrl: editionId ? `${OPEN_LIBRARY_API}/books/${editionId}` : workId ? `${OPEN_LIBRARY_API}/works/${workId}` : undefined, fieldSources: providerFieldSources("openlibrary", ["title", "author", "additionalAuthors", "isbn", "publicationYear", "publisher", "language", "pages", "format", "series", "seriesVolume", "description", "coverUrl"]) };
}
function candidateFromOpenLibrarySearch(doc: OpenLibrarySearchDoc): MetadataCandidate | undefined {
  const workId = sanitizeOlid(doc.key, "W"); if (!workId || !doc.title) return undefined;
  const matchedEdition = doc.editions?.docs?.[0]; const authors = doc.author_name ?? [];
  const editionId = sanitizeOlid(matchedEdition?.key, "M") ?? sanitizeOlid(doc.edition_key?.[0], "M");
  const title = matchedEdition?.title?.trim() || doc.title;
  const seriesMemberships = openLibrarySeriesMemberships([...(matchedEdition?.series ?? []), ...(doc.series ?? [])]);
  const seriesInfo = seriesMemberships[0] ?? (() => { const parsed = parseSeriesFromTitle(title); return parsed ? { provider: "openlibrary" as const, name: parsed.name, volume: parsed.volume } : undefined; })();
  if (seriesInfo && !seriesMemberships.length) seriesMemberships.push(seriesInfo);
  const isbn = preferredIsbn(matchedEdition?.isbn?.length ? matchedEdition.isbn : doc.isbn ?? []);
  const coverId = matchedEdition?.cover_i ?? doc.cover_i; const cover = coverId ? coverUrl(coverId) : undefined;
  const publicationYear = matchedEdition?.publish_year?.[0] ?? yearFromDate(matchedEdition?.publish_date?.[0]) ?? doc.first_publish_year;
  const publisher = matchedEdition?.publisher?.[0] ?? doc.publisher?.[0]; const language = matchedEdition?.language?.[0] ?? doc.language?.[0];
  const pages = matchedEdition?.number_of_pages ?? doc.number_of_pages_median; const format = mapBookFormat(matchedEdition?.physical_format);
  const sourceUrl = editionId ? `${OPEN_LIBRARY_API}/books/${editionId}` : `${OPEN_LIBRARY_API}/works/${workId}`;
  return { source: "openlibrary", workId, editionId, sourceRefs: [sourceRef("openlibrary", workId, editionId, undefined, sourceUrl)], matchType: editionId ? "edition" : "search", confidence: editionId ? 74 : 68, title, author: authors[0] ?? "Unknown author", additionalAuthors: authors.slice(1), isbn, publicationYear, publisher, language, pages, format, series: seriesInfo?.name, seriesVolume: seriesInfo?.volume, seriesMemberships: seriesMemberships.length ? seriesMemberships : undefined, subjects: doc.subject?.slice(0, 20) ?? [], coverUrl: cover, coverUrls: cover ? [cover] : [], sourceUrl, fieldSources: providerFieldSources("openlibrary", ["title", "author", "additionalAuthors", "isbn", "publicationYear", "publisher", "language", "pages", "format", "series", "seriesVolume", "coverUrl"]) };
}
async function openLibraryAuthorNames(work: OpenLibraryWork): Promise<string[]> { return uniqueStrings(await Promise.all((work.authors ?? []).map((entry) => openLibraryAuthorName(entry.author?.key)))); }
async function openLibraryAuthorName(key?: string): Promise<string | undefined> { if (!key) return undefined; try { return (await openLibraryFetch<OpenLibraryAuthor>(`${OPEN_LIBRARY_API}${key}.json`)).name; } catch { return undefined; } }
function languageFromOpenLibrary(key?: string): string | undefined { return key?.split("/").filter(Boolean).at(-1); }
function coverUrl(id: number): string { return `https://covers.openlibrary.org/b/id/${id}-L.jpg`; }
async function openLibraryFetch<T>(url: string): Promise<T> { return jsonFetch<T>(url, { "User-Agent": metadataUserAgent(), Accept: "application/json" }); }
function sanitizeOlid(value: string | undefined, suffix: "W" | "M"): string | undefined { if (!value) return undefined; const id = value.replace(/^\/(works|books)\//, ""); return new RegExp(`^OL\\d+${suffix}$`).test(id) ? id : undefined; }

function openLibrarySeriesMemberships(values?: string[]): MetadataSeriesMembership[] {
  return dedupeSeriesMemberships((values ?? []).flatMap((value) => {
    const parsed = parseSeriesLabel(value);
    return parsed ? [{ provider: "openlibrary" as const, name: parsed.name, volume: parsed.volume }] : [];
  }));
}
function parseSeriesLabel(value?: string): { name: string; volume?: string } | undefined {
  if (!value?.trim()) return undefined; const clean = value.trim();
  const position = String.raw`\d+(?:\.\d+)?(?:\s*(?:&|and|,|[-–—])\s*\d+(?:\.\d+)?)*`;
  const patterns = [new RegExp(`^(.+?)[,;:]?\\s*#\\s*(${position})$`, "i"), new RegExp(`^(.+?)[,;:]?\\s*(?:book|bk\\.?|vol(?:ume)?\\.?)\\s*#?\\s*(${position})$`, "i"), new RegExp(`^(.+?)\\s*--\\s*(?:book|bk\\.?|vol(?:ume)?\\.?)\\s*#?\\s*(${position})$`, "i"), new RegExp(`^(.+?)\\s*\\((?:book|bk\\.?|vol(?:ume)?\\.?)?\\s*#?\\s*(${position})\\)$`, "i"), new RegExp(`^(.+?),\\s*(${position})$`, "i")];
  for (const pattern of patterns) { const match = clean.match(pattern); if (match) return { name: match[1].trim(), volume: match[2] }; }
  return { name: clean };
}
function parseSeriesFromTitle(title: string): { name: string; volume?: string } | undefined {
  const text = title.trim();
  if (!text.endsWith(")")) return undefined;
  let depth = 0; let opening = -1;
  for (let index = text.length - 1; index >= 0; index -= 1) {
    if (text[index] === ")") depth += 1;
    else if (text[index] === "(") { depth -= 1; if (depth === 0) { opening = index; break; } }
  }
  if (opening < 0) return undefined;
  const content = text.slice(opening + 1, -1).trim();
  const match = content.match(/^(.+),\s*#?\s*(\d+(?:\.\d+)?(?:\s*(?:&|and|,|[-–—])\s*\d+(?:\.\d+)?)*)$/i);
  return match ? { name: match[1].trim(), volume: match[2] } : undefined;
}

function providerFieldSources(provider: MetadataProvider, fields: MetadataField[]): Partial<Record<MetadataField, MetadataProvider>> { return Object.fromEntries(fields.map((field) => [field, provider])) as Partial<Record<MetadataField, MetadataProvider>>; }
function mapBookFormat(value?: string | null): BookFormat | undefined { if (!value) return undefined; const text = value.toLocaleLowerCase(); if (text.includes("mass market")) return "Mass Market Paperback"; if (text.includes("paperback") || text.includes("softcover")) return "Paperback"; if (text.includes("hardcover") || text.includes("hardback")) return "Hardcover"; if (text.includes("ebook") || text.includes("kindle") || text.includes("digital")) return "eBook"; if (text.includes("audio")) return "Audiobook"; if (text.includes("graphic")) return "Graphic Novel"; if (text.includes("omnibus")) return "Omnibus"; return "Other"; }
function compactPrimarySeriesBooks<T extends { providerId: string; title: string; position?: string; featured?: boolean }>(books: T[], expectedPrimary?: number): T[] {
  const unique = [...new Map(books.map((book) => [book.providerId, book] as const)).values()];
  const integerPositions = unique.map((book) => Number(book.position)).filter((position) => Number.isInteger(position) && position > 0 && position <= 500);
  const uniquePositions = [...new Set(integerPositions)].sort((a, b) => a - b);
  const maxPosition = uniquePositions.at(-1) ?? 0;
  const inferredPrimary = uniquePositions.length >= 2 && maxPosition > 0
    && uniquePositions.length / maxPosition >= 0.75 && integerPositions.length / uniquePositions.length >= 2
    ? maxPosition : undefined;
  // Hardcover can occasionally count translated/alternate work records as primary
  // books. A dense repeated 1..N position sequence is stronger evidence of the
  // collector-facing mainline when the reported count is wildly larger.
  const expected = inferredPrimary && (!expectedPrimary || expectedPrimary >= inferredPrimary * 3)
    ? inferredPrimary
    : expectedPrimary ?? inferredPrimary;
  if (!expected || expected <= 0 || expected > 500) return unique;
  const byPosition = new Map<number, T[]>();
  for (const book of unique) {
    const position = Number(book.position);
    if (!Number.isInteger(position) || position < 1 || position > expected) continue;
    const rows = byPosition.get(position) ?? []; rows.push(book); byPosition.set(position, rows);
  }
  const needed = expected <= 3 ? expected : Math.max(3, Math.ceil(expected * 0.6));
  if (byPosition.size < Math.min(expected, needed)) return unique;
  return [...byPosition.entries()].sort(([a], [b]) => a - b).map(([, rows]) => [...rows].sort((a, b) => Number(Boolean(b.featured)) - Number(Boolean(a.featured)) || a.title.length - b.title.length)[0]);
}
function seriesBookSort(a: { position?: string }, b: { position?: string }): number { const left = Number(a.position); const right = Number(b.position); if (Number.isFinite(left) && Number.isFinite(right)) return left - right; if (Number.isFinite(left)) return -1; if (Number.isFinite(right)) return 1; return String(a.position ?? "").localeCompare(String(b.position ?? "")); }
async function jsonFetch<T>(url: string, headers?: Record<string, string>): Promise<T> { const response = await fetch(url, { headers }); if (!response.ok) throw new Error(`Metadata request failed (${response.status}) for ${new URL(url).hostname}`); return response.json() as Promise<T>; }
