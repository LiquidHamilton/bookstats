import { useMemo, useState, type ReactNode } from "react";
import type { Book, ReadingGoal, ReadingGoalMetric, SeriesCompletionOverride, Shelf } from "@bookstats/domain";
import { activeReadingSession, shelfMatchesBook } from "@bookstats/domain";
import {
  authorLibraryDistribution,
  authorReadDistribution,
  averageRatingByAuthor,
  booksAddedByYear,
  booksReadByYear,
  formatDistribution,
  genreDistribution,
  libraryFlowByYear,
  monthlyReadingForYear,
  newAuthorsByYear,
  pageLengthDistribution,
  pagesReadByYear,
  publicationDecadeDistribution,
  ratingDistribution,
  ratingTrendByYear,
  readFormatDistribution,
  readGenreDistribution,
  readOwnershipDistribution,
  readingExtremes,
  readingGoalProgress,
  readingMonthDistribution,
  readingPaceSummary,
  seriesDistribution,
  seriesProgress,
  statusDistribution,
  summarizeLibrary
} from "@bookstats/statistics";
import { BookOpen, ChartNoAxesCombined, CheckCircle2, Gauge, LibraryBig, Pencil, Plus, Star, Target, Trash2, UsersRound, X } from "lucide-react";
import { StatChart } from "./StatChart";
import { SeriesCompletionEditor } from "./SeriesCompletionEditor";

type StatsCategory = "overview" | "goals" | "reading" | "authors" | "library" | "series" | "ratings";
const categoryLabels: Array<{ id: StatsCategory; label: string; icon: ReactNode }> = [
  { id: "overview", label: "Overview", icon: <ChartNoAxesCombined size={16} /> },
  { id: "goals", label: "Goals", icon: <Target size={16} /> },
  { id: "reading", label: "Reading", icon: <BookOpen size={16} /> },
  { id: "authors", label: "Authors", icon: <UsersRound size={16} /> },
  { id: "library", label: "Library", icon: <LibraryBig size={16} /> },
  { id: "series", label: "Series", icon: <Gauge size={16} /> },
  { id: "ratings", label: "Ratings", icon: <Star size={16} /> }
];

const goalMetricLabels: Record<ReadingGoalMetric, string> = {
  books: "Books read",
  pages: "Pages read",
  rereads: "Rereads",
  new_authors: "New-to-me authors",
  owned_books: "Owned books read"
};

export function StatisticsView({ books, shelves, goals, onSaveGoal, onDeleteGoal, onOpenSeries, onSaveSeriesCompletion }: {
  books: Book[];
  shelves: Shelf[];
  goals: ReadingGoal[];
  onSaveGoal: (goal: ReadingGoal) => Promise<void>;
  onDeleteGoal: (goal: ReadingGoal) => Promise<void>;
  onOpenSeries: (seriesName: string) => void;
  onSaveSeriesCompletion: (seriesName: string, override: SeriesCompletionOverride | undefined) => Promise<void>;
}) {
  const [category, setCategory] = useState<StatsCategory>("overview");
  const [goalEditor, setGoalEditor] = useState<ReadingGoal | null | undefined>(undefined);
  const [seriesEditorName, setSeriesEditorName] = useState<string>();
  const summary = useMemo(() => summarizeLibrary(books), [books]);
  const statusData = useMemo(() => statusDistribution(books), [books]);
  const ratingData = useMemo(() => ratingDistribution(books), [books]);
  const formatData = useMemo(() => formatDistribution(books), [books]);
  const genreData = useMemo(() => genreDistribution(books), [books]);
  const seriesData = useMemo(() => seriesDistribution(books), [books]);
  const authorLibraryData = useMemo(() => authorLibraryDistribution(books), [books]);
  const authorReadData = useMemo(() => authorReadDistribution(books), [books]);
  const readByYear = useMemo(() => booksReadByYear(books), [books]);
  const pagesByYear = useMemo(() => pagesReadByYear(books), [books]);
  const addedByYear = useMemo(() => booksAddedByYear(books), [books]);
  const decades = useMemo(() => publicationDecadeDistribution(books), [books]);
  const pageLengths = useMemo(() => pageLengthDistribution(books), [books]);
  const monthlyReading = useMemo(() => readingMonthDistribution(books), [books]);
  const ratingByAuthor = useMemo(() => averageRatingByAuthor(books), [books]);
  const pace = useMemo(() => readingPaceSummary(books), [books]);
  const extremes = useMemo(() => readingExtremes(books), [books]);
  const series = useMemo(() => seriesProgress(books), [books]);
  const flow = useMemo(() => libraryFlowByYear(books), [books]);
  const newAuthors = useMemo(() => newAuthorsByYear(books), [books]);
  const ratingTrend = useMemo(() => ratingTrendByYear(books), [books]);
  const readFormats = useMemo(() => readFormatDistribution(books), [books]);
  const readGenres = useMemo(() => readGenreDistribution(books), [books]);
  const readOwnership = useMemo(() => readOwnershipDistribution(books), [books]);
  const shelfData = useMemo(() => shelves.map((shelf) => ({
    name: shelf.name,
    value: books.filter((book) => shelfMatchesBook(shelf, book)).length
  })).filter((item) => item.value > 0).sort((a, b) => b.value - a.value || a.name.localeCompare(b.name)), [books, shelves]);

  const currentYearNumber = new Date().getFullYear();
  const currentYear = String(currentYearNumber);
  const monthlyCurrentYear = useMemo(() => monthlyReadingForYear(books, currentYearNumber), [books, currentYearNumber]);
  const readsThisYear = readByYear.find((datum) => datum.name === currentYear)?.value ?? 0;
  const pagesThisYear = pagesByYear.find((datum) => datum.name === currentYear)?.value ?? 0;
  const mostReadAuthor = authorReadData[0];
  const mostCollectedAuthor = authorLibraryData[0];
  const mostCommonGenre = genreData.find((datum) => datum.name !== "Unspecified");
  const busiestYear = [...readByYear].sort((a, b) => b.value - a.value)[0];
  const ratedBooks = books.filter((book) => typeof book.rating === "number");
  const reviewedBooks = books.filter((book) => Boolean(book.review?.trim()));
  const highlyRated = ratedBooks.filter((book) => (book.rating ?? 0) >= 4.5);
  const bestRatedAuthor = ratingByAuthor[0];
  const activeBooks = books.filter((book) => Boolean(activeReadingSession(book)));
  const goalProgress = goals.map((goal) => ({ goal, progress: readingGoalProgress(goal, books) }));
  const activeGoals = goalProgress.filter(({ goal }) => goal.endDate >= new Date().toISOString().slice(0, 10));
  const completedGoals = goalProgress.filter(({ progress }) => progress.complete).length;

  return <>
    <header className="page-header"><div><p className="eyebrow">Explore your reading</p><h1>Statistics</h1><p>Simple at a glance, deep when you want it: goals, reading pace, collection trends, authors, series and ratings.</p></div></header>
    <div className="stats-tabs" role="tablist" aria-label="Statistics categories">{categoryLabels.map((item) => <button key={item.id} className={category === item.id ? "active" : ""} onClick={() => setCategory(item.id)}>{item.icon}{item.label}</button>)}</div>

    {category === "overview" && <>
      <section className="metric-grid metric-grid-wide">
        <Metric label="Books" value={summary.totalBooks.toLocaleString()} note={`${summary.ownedBooks} owned`} />
        <Metric label="Read titles" value={summary.readBooks.toLocaleString()} note={summary.totalBooks ? `${Math.round((summary.readBooks / summary.totalBooks) * 100)}% of library` : "No books yet"} />
        <Metric label="Total readings" value={summary.totalReadings.toLocaleString()} note={`${summary.rereads} rereads`} />
        <Metric label="Pages read" value={summary.pagesRead.toLocaleString()} note={`${pagesThisYear.toLocaleString()} this year`} />
        <Metric label="Goals completed" value={completedGoals.toLocaleString()} note={`${goals.length} goals created`} />
        <Metric label="Avg. rating" value={summary.averageRating === null ? "—" : summary.averageRating.toFixed(2)} note={`${ratedBooks.length} rated books`} />
      </section>
      {activeGoals.length > 0 && <section className="stats-goal-strip"><div><Target size={17} /><strong>Goals in progress</strong></div><div>{activeGoals.slice(0, 3).map(({ goal, progress }) => <button key={goal.id} onClick={() => setCategory("goals")}><span>{goal.name}</span><strong>{formatGoalValue(goal.metric, progress.current)} / {formatGoalValue(goal.metric, progress.target)}</strong><div className="mini-progress"><i style={{ width: `${progress.percent}%` }} /></div></button>)}</div></section>}
      <section className="chart-grid">
        <ChartPanel title="Reading status" subtitle="Where your library stands"><DonutChart data={statusData} /></ChartPanel>
        <ChartPanel title="Reads by year" subtitle="Every completion, including rereads"><VerticalBar data={readByYear} /></ChartPanel>
        <ChartPanel title="Library in vs. books read" subtitle="Books added compared with completed readings"><DualLine data={flow.map((row) => ({ name: row.name, first: row.added, second: row.read }))} firstLabel="Added" secondLabel="Read" /></ChartPanel>
        <ChartPanel title="Most-read authors" subtitle="Ranked by recorded completion dates"><HorizontalBar data={authorReadData.slice(0, 10)} /></ChartPanel>
      </section>
    </>}

    {category === "goals" && <>
      <section className="goal-page-heading"><div><h2>Your goals</h2><p>Goals live entirely in Statistics. They update automatically from your reading history—no separate check-ins required.</p></div><button className="button primary" onClick={() => setGoalEditor(null)}><Plus size={16} />Add goal</button></section>
      {goals.length === 0 ? <section className="stats-empty"><Target size={38} /><h2>Set a goal that matters to you</h2><p>Track books, pages, rereads, new-to-you authors, or books read from your owned collection. BookStats does not create a default challenge.</p><button className="button primary" onClick={() => setGoalEditor(null)}><Plus size={16} />Create your first goal</button></section> : <section className="goal-grid">{goalProgress.map(({ goal, progress }) => <GoalCard key={goal.id} goal={goal} current={progress.current} target={progress.target} percent={progress.percent} complete={progress.complete} elapsedPercent={progress.elapsedPercent} daysRemaining={progress.daysRemaining} onPace={progress.onPace} projected={progress.projected} onEdit={() => setGoalEditor(goal)} onDelete={() => void onDeleteGoal(goal)} />)}</section>}
      {goals.length > 0 && <section className="chart-grid goal-chart-grid"><ChartPanel title="Goal completion" subtitle="Progress across all current and historical goals"><GoalComparison goals={goalProgress} /></ChartPanel><section className="chart-panel insight-panel"><p className="eyebrow">Tracking</p><h2>{completedGoals} / {goals.length}</h2><p>goals have reached their target. A goal keeps its historical date range, so it remains meaningful after the year or challenge ends.</p></section></section>}
    </>}

    {category === "reading" && <>
      <section className="metric-grid metric-grid-wide">
        <Metric label={`Reads in ${currentYear}`} value={readsThisYear.toLocaleString()} note={`${pagesThisYear.toLocaleString()} pages`} />
        <Metric label="All readings" value={summary.totalReadings.toLocaleString()} note={`${summary.readBooks} unique titles`} />
        <Metric label="Rereads" value={summary.rereads.toLocaleString()} note={summary.totalReadings ? `${Math.round((summary.rereads / summary.totalReadings) * 100)}% of readings` : "No reading history"} />
        <Metric label="Avg. finish time" value={pace.averageDaysToFinish === null ? "—" : `${pace.averageDaysToFinish.toFixed(1)} days`} note={`${pace.timedReads} reads with start + finish dates`} />
        <Metric label="Avg. pace" value={pace.averagePagesPerDay === null ? "—" : `${Math.round(pace.averagePagesPerDay)} pages/day`} note="For timed reads with page counts" />
        <Metric label="Currently reading" value={activeBooks.length.toLocaleString()} note={pace.fastestDays ? `Fastest completed: ${pace.fastestDays} days` : "Track a start date to measure pace"} />
      </section>
      {activeBooks.length > 0 && <section className="currently-reading-section"><div className="section-heading stats-subheading"><div><p className="eyebrow">In progress</p><h2>Currently reading</h2></div></div><div className="currently-reading-stats">{activeBooks.slice(0, 6).map((book) => { const session = activeReadingSession(book)!; const percent = book.pages && session.progressPages ? Math.min(100, session.progressPages / book.pages * 100) : 0; return <article className="reading-progress-card" key={book.id}><div><div><strong>{book.title}</strong><span>{book.author}</span></div>{session.startedAt && <span>Started {new Date(`${session.startedAt}T12:00:00`).toLocaleDateString(undefined, { month: "short", day: "numeric" })}</span>}</div><div className="reading-progress-label"><span>{session.progressPages ? `Page ${session.progressPages}${book.pages ? ` of ${book.pages}` : ""}` : "Started"}</span>{book.pages && session.progressPages ? <strong>{Math.round(percent)}%</strong> : null}</div>{book.pages && session.progressPages ? <div className="detail-progress-track"><i style={{ width: `${percent}%` }} /></div> : null}</article>; })}</div></section>}
      <section className="reading-extremes-grid">
        <article><span>Longest completed book</span><strong>{extremes.longest ? `${extremes.longest.pages.toLocaleString()} pages` : "—"}</strong><small>{extremes.longest ? `${extremes.longest.title} · ${extremes.longest.author}` : "Add page counts to completed books"}</small></article>
        <article><span>Shortest completed book</span><strong>{extremes.shortest ? `${extremes.shortest.pages.toLocaleString()} pages` : "—"}</strong><small>{extremes.shortest ? `${extremes.shortest.title} · ${extremes.shortest.author}` : "Add page counts to completed books"}</small></article>
        <article><span>Average pages per reading</span><strong>{pace.averagePagesPerRead === null ? "—" : Math.round(pace.averagePagesPerRead).toLocaleString()}</strong><small>Known page counts across recorded completions</small></article>
      </section>
      <div className="section-heading stats-subheading"><div><p className="eyebrow">This year</p><h2>{currentYear} month by month</h2><p>Monthly totals use recorded completion dates. Pages require a page count on the book.</p></div></div>
      <section className="chart-grid">
        <ChartPanel title="Books by month" subtitle={`Completed readings in ${currentYear}`}><VerticalBar data={monthlyCurrentYear.map((row) => ({ name: row.name, value: row.books }))} /></ChartPanel>
        <ChartPanel title="Pages by month" subtitle={`Known pages completed in ${currentYear}`}><LineChart data={monthlyCurrentYear.map((row) => ({ name: row.name, value: row.pages }))} /></ChartPanel>
      </section>
      <div className="section-heading stats-subheading"><div><p className="eyebrow">Long-term trends</p><h2>Your reading over time</h2></div></div>
      <section className="chart-grid">
        <ChartPanel title="Books read by year" subtitle="Completion history across years"><VerticalBar data={readByYear} /></ChartPanel>
        <ChartPanel title="Pages read by year" subtitle="Known pages, with rereads counted again"><LineChart data={pagesByYear} /></ChartPanel>
        <ChartPanel title="Reading by month" subtitle="Your most active months across all years"><VerticalBar data={monthlyReading} /></ChartPanel>
        <ChartPanel title="Genres read" subtitle="Completed readings by current genre metadata"><HorizontalBar data={readGenres.slice(0, 12)} /></ChartPanel>
        <ChartPanel title="Formats read" subtitle="Completed readings by current edition format"><DonutChart data={readFormats} /></ChartPanel>
        <ChartPanel title="Owned vs. unowned reads" subtitle="Based on each book's current ownership setting"><DonutChart data={readOwnership} /></ChartPanel>
        <ChartPanel title="Library flow" subtitle="Books added compared with completed readings"><DualLine data={flow.map((row) => ({ name: row.name, first: row.added, second: row.read }))} firstLabel="Added" secondLabel="Read" /></ChartPanel>
      </section>
    </>}

    {category === "authors" && <>
      <section className="metric-grid"><Metric label="Unique authors" value={summary.uniqueAuthors.toLocaleString()} note="Primary authors in library" /><Metric label="Most read" value={mostReadAuthor?.name ?? "—"} note={mostReadAuthor ? `${mostReadAuthor.value} readings` : "No reading history"} /><Metric label="Most collected" value={mostCollectedAuthor?.name ?? "—"} note={mostCollectedAuthor ? `${mostCollectedAuthor.value} books` : "No books yet"} /><Metric label="Highest avg. rating" value={bestRatedAuthor?.name ?? "—"} note={bestRatedAuthor ? `${bestRatedAuthor.value.toFixed(2)} ★ across ${bestRatedAuthor.count}` : "No ratings yet"} /></section>
      <section className="chart-grid"><ChartPanel title="Most-read authors" subtitle="Reading events, including rereads"><HorizontalBar data={authorReadData.slice(0, 15)} /></ChartPanel><ChartPanel title="New-to-you authors by year" subtitle="The year each primary author first appears in your reading history"><VerticalBar data={newAuthors} /></ChartPanel><ChartPanel title="Authors in your library" subtitle="Number of books by contributor"><HorizontalBar data={authorLibraryData.slice(0, 15)} /></ChartPanel><ChartPanel title="Average rating by author" subtitle="Authors ranked by your ratings"><HorizontalBar data={ratingByAuthor.slice(0, 15)} valueFormatter={(value) => value.toFixed(2)} /></ChartPanel><ChartPanel title="Author concentration" subtitle="How much of your library comes from repeat authors"><DonutChart data={authorLibraryData.slice(0, 10)} /></ChartPanel></section>
    </>}

    {category === "library" && <>
      <section className="metric-grid metric-grid-wide"><Metric label="Owned" value={summary.ownedBooks.toLocaleString()} note={`${summary.totalBooks - summary.ownedBooks} not owned`} /><Metric label="Unread" value={summary.unreadBooks.toLocaleString()} note={`${summary.wantToRead} marked want to read`} /><Metric label="Known pages" value={summary.pagesKnown.toLocaleString()} note="Across your collection" /><Metric label="Genres" value={summary.uniqueGenres.toLocaleString()} note={mostCommonGenre ? `${mostCommonGenre.name} is largest` : "Add genres for breakdowns"} /><Metric label="Series" value={summary.uniqueSeries.toLocaleString()} note={`${seriesData.reduce((sum, item) => sum + item.value, 0)} books in series`} /><Metric label="Shelves" value={shelves.length.toLocaleString()} note={shelfData[0] ? `${shelfData[0].name} is largest` : "Create shelves to organize books"} /></section>
      <section className="chart-grid"><ChartPanel title="Formats" subtitle="What is on your shelves"><HorizontalBar data={formatData} /></ChartPanel><ChartPanel title="Genres" subtitle="Genre composition of your library"><HorizontalBar data={genreData.slice(0, 12)} /></ChartPanel><ChartPanel title="Books added by year" subtitle="Growth of your BookStats library"><VerticalBar data={addedByYear} /></ChartPanel><ChartPanel title="Publication decades" subtitle="When your books were first published"><VerticalBar data={decades} /></ChartPanel><ChartPanel title="Book lengths" subtitle="Page-count distribution"><VerticalBar data={pageLengths} /></ChartPanel><ChartPanel title="Shelves" subtitle="Regular and smart shelf membership"><HorizontalBar data={shelfData.slice(0, 15)} /></ChartPanel></section>
    </>}

    {category === "series" && <>
      <section className="metric-grid">
        <Metric label="Series" value={series.length.toLocaleString()} note="Named series in your library" />
        <Metric label="Complete collections" value={series.filter((item) => item.catalogTotal && item.missingCatalogBooks.length === 0 && item.catalogBooks.filter((entry) => entry.owned).length >= item.catalogTotal).length.toLocaleString()} note="Mainline series currently complete" />
        <Metric label="Missing mainline books" value={series.reduce((sum, item) => sum + item.missingCatalogBooks.length, 0).toLocaleString()} note="After catalog cleanup and your overrides" />
        <Metric label="Catalog-connected" value={series.filter((item) => item.catalogBooks.length > 0).length.toLocaleString()} note="Series with a completion checklist" />
      </section>
      <section className="series-progress-list">{series.length === 0 ? <div className="stats-empty"><Gauge size={36} /><h2>No series yet</h2><p>Add series names and volume numbers to books and BookStats will track your collection and reading progress here.</p></div> : series.map((item) => {
        const collected = item.catalogBooks.filter((entry) => entry.owned).length;
        const target = item.catalogTotal ?? item.catalogBooks.length;
        const collectionComplete = Boolean(target && collected >= target && item.missingCatalogBooks.length === 0);
        return <article key={item.name} className={`series-progress-card ${collectionComplete ? "series-collection-complete" : ""}`}><div className="series-progress-heading"><div><div className="series-title-line"><h3><button className="series-filter-link" onClick={() => onOpenSeries(item.name)} title={`Show only ${item.name} in Library`}>{item.name}</button></h3>{item.completionOverride && <span className="series-provider-badge">Custom completion</span>}</div><span>{item.read} read · {item.owned} owned · {item.total} in library</span></div><div className="series-heading-actions"><button className="button secondary compact" onClick={() => setSeriesEditorName(item.name)}>Edit completion</button><div className="series-percent"><strong>{Math.round(item.completionPercent)}%</strong><span>read</span></div></div></div><div className="detail-progress-track"><i style={{ width: `${item.completionPercent}%` }} /></div>{item.catalogBooks.length > 0 || item.catalogTotal ? <><div className="series-catalog-summary"><div><strong>{collected} of {target || item.catalogBooks.length}</strong><span>mainline positions collected</span></div>{typeof item.collectionPercent === "number" && <div className="series-collection-meter"><div className="detail-progress-track secondary"><i style={{ width: `${Math.min(100, item.collectionPercent)}%` }} /></div><span>{Math.round(item.collectionPercent)}% collected</span></div>}</div>{item.missingCatalogBooks.length > 0 ? <div className="series-missing-books"><span className="series-missing-label">Missing from your owned collection</span><div className="series-missing-list">{item.missingCatalogBooks.slice(0, 8).map((entry) => <span key={`${entry.providerId}:${entry.position ?? entry.title}`}>{entry.position ? <b>#{entry.position}</b> : null}{entry.title}</span>)}{item.missingCatalogBooks.length > 8 && <span>+{item.missingCatalogBooks.length - 8} more</span>}</div></div> : collectionComplete ? <div className="series-complete-note"><CheckCircle2 size={15} />Mainline collection complete.</div> : null}</> : <div className="series-volume-row">{item.knownVolumes.length > 0 ? <span>Known volumes: {item.knownVolumes.join(", ")}</span> : <span>Volume numbers not recorded</span>}{item.missingVolumeGaps.length > 0 && <span className="series-gap">Number gaps: {item.missingVolumeGaps.join(", ")}</span>}<button className="series-inline-edit" onClick={() => setSeriesEditorName(item.name)}>Set expected series size manually</button></div>}</article>;
      })}</section>
      {series.length > 0 && <section className="chart-grid"><ChartPanel title="Largest series" subtitle="Books currently in your library"><HorizontalBar data={series.slice().sort((a, b) => b.total - a.total).slice(0, 15).map((item) => ({ name: item.name, value: item.total }))} /></ChartPanel><ChartPanel title="Reading completion" subtitle="Read percentage among series books already in your library"><HorizontalBar data={series.slice().sort((a, b) => b.completionPercent - a.completionPercent).slice(0, 15).map((item) => ({ name: item.name, value: Math.round(item.completionPercent) }))} valueFormatter={(value) => `${value}%`} /></ChartPanel></section>}
    </>}

    {category === "ratings" && <>
      <section className="metric-grid"><Metric label="Average rating" value={summary.averageRating === null ? "—" : `${summary.averageRating.toFixed(2)} ★`} note={`${ratedBooks.length} rated books`} /><Metric label="Highly rated" value={highlyRated.length.toLocaleString()} note="4½ stars or higher" /><Metric label="Reviews written" value={reviewedBooks.length.toLocaleString()} note={ratedBooks.length ? `${Math.round((reviewedBooks.length / Math.max(1, ratedBooks.length)) * 100)}% of rated books` : "No rated books"} /><Metric label="Favorite author rating" value={bestRatedAuthor ? `${bestRatedAuthor.value.toFixed(2)} ★` : "—"} note={bestRatedAuthor?.name ?? "No ratings yet"} /></section>
      <section className="chart-grid"><ChartPanel title="Rating distribution" subtitle="Ratings in half-star increments"><VerticalBar data={ratingData} /></ChartPanel><ChartPanel title="Average rating over time" subtitle="Average current rating of books completed in each year"><LineChart data={ratingTrend} /></ChartPanel><ChartPanel title="Average rating by author" subtitle="Authors ranked by your ratings"><HorizontalBar data={ratingByAuthor.slice(0, 15)} valueFormatter={(value) => value.toFixed(2)} /></ChartPanel><ChartPanel title="Genres" subtitle="Library genre mix for rating context"><DonutChart data={genreData.slice(0, 10)} /></ChartPanel><section className="chart-panel insight-panel"><p className="eyebrow">Reviews</p><h2>{reviewedBooks.length}</h2><p>books currently have a written review. Reviews stay separate from publisher descriptions so metadata cleanup never overwrites your own writing.</p></section></section>
    </>}

    {goalEditor !== undefined && <GoalEditor goal={goalEditor ?? undefined} onSave={async (goal) => { await onSaveGoal(goal); setGoalEditor(undefined); }} onClose={() => setGoalEditor(undefined)} />}
    {seriesEditorName && (() => { const selected = series.find((item) => item.name === seriesEditorName); return selected ? <SeriesCompletionEditor series={selected} onSave={(override) => onSaveSeriesCompletion(selected.name, override)} onClose={() => setSeriesEditorName(undefined)} /> : null; })()}
  </>;
}

function GoalEditor({ goal, onSave, onClose }: { goal?: ReadingGoal; onSave: (goal: ReadingGoal) => Promise<void>; onClose: () => void }) {
  const year = new Date().getFullYear();
  const [name, setName] = useState(goal?.name ?? `${year} reading goal`);
  const [metric, setMetric] = useState<ReadingGoalMetric>(goal?.metric ?? "books");
  const [target, setTarget] = useState(String(goal?.target ?? 30));
  const [startDate, setStartDate] = useState(goal?.startDate ?? `${year}-01-01`);
  const [endDate, setEndDate] = useState(goal?.endDate ?? `${year}-12-31`);
  const [working, setWorking] = useState(false);
  const valid = name.trim() && Number(target) > 0 && startDate && endDate && startDate <= endDate;
  async function save() { if (!valid) return; const now = new Date().toISOString(); setWorking(true); try { await onSave({ id: goal?.id ?? crypto.randomUUID(), name: name.trim(), metric, target: Math.max(1, Math.round(Number(target))), startDate, endDate, createdAt: goal?.createdAt ?? now, updatedAt: now }); } finally { setWorking(false); } }
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="goal-editor-modal"><div className="form-header"><div><p className="eyebrow">Statistics</p><h2>{goal ? "Edit reading goal" : "Add reading goal"}</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div><div className="goal-form-grid"><label>Goal name<input value={name} onChange={(event) => setName(event.target.value)} /></label><label>Track<select value={metric} onChange={(event) => setMetric(event.target.value as ReadingGoalMetric)}>{(Object.entries(goalMetricLabels) as Array<[ReadingGoalMetric, string]>).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label><label>Target<input type="number" min="1" step="1" value={target} onChange={(event) => setTarget(event.target.value)} /></label><label>Start date<input type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} /></label><label>End date<input type="date" value={endDate} onChange={(event) => setEndDate(event.target.value)} /></label></div><p className="goal-help">Progress is calculated from dated BookStats reading history. Books marked Read without a completion date still count in library totals, but cannot be placed inside a dated goal. Changing a goal never changes any book records.</p><div className="form-actions"><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={!valid || working} onClick={() => void save()}><Target size={16} />{working ? "Saving…" : "Save goal"}</button></div></section></div>;
}

function GoalCard({ goal, current, target, percent, complete, elapsedPercent, daysRemaining, onPace, projected, onEdit, onDelete }: {
  goal: ReadingGoal; current: number; target: number; percent: number; complete: boolean; elapsedPercent: number; daysRemaining: number; onPace: boolean; projected: number | null; onEdit: () => void; onDelete: () => void;
}) {
  const today = new Date().toISOString().slice(0, 10);
  const future = goal.startDate > today;
  const ended = goal.endDate < today;
  const projectedText = projected === null ? null : formatGoalValue(goal.metric, projected);
  return <article className={`goal-card ${complete ? "complete" : ended ? "ended" : ""}`}>
    <div className="goal-card-top"><div className="goal-icon">{complete ? <CheckCircle2 size={19} /> : <Target size={19} />}</div><div className="goal-card-actions"><button className="icon-button" onClick={onEdit} aria-label="Edit goal"><Pencil size={14} /></button><button className="icon-button danger-icon" onClick={() => { if (window.confirm(`Delete the goal “${goal.name}”? Your reading history will not be changed.`)) onDelete(); }} aria-label="Delete goal"><Trash2 size={14} /></button></div></div>
    <div><h3>{goal.name}</h3><span>{goalMetricLabels[goal.metric]} · {formatDateRange(goal.startDate, goal.endDate)}</span></div>
    <div className="goal-numbers"><strong>{formatGoalValue(goal.metric, current)}</strong><span>of {formatGoalValue(goal.metric, target)}</span></div>
    <div className="goal-progress-track"><i style={{ width: `${percent}%` }} /></div>
    <div className="goal-footer"><span>{Math.round(percent)}%</span><span>{complete ? "Goal reached" : ended ? "Goal period ended" : future ? `Starts ${new Date(`${goal.startDate}T12:00:00`).toLocaleDateString(undefined, { month: "short", day: "numeric" })}` : `${formatGoalValue(goal.metric, Math.max(0, target - current))} to go`}</span></div>
    {!complete && !ended && !future && <div className="goal-pace"><strong className={onPace ? "on-pace" : "behind-pace"}>{onPace ? "On pace" : "Behind pace"}</strong><span>{daysRemaining} {daysRemaining === 1 ? "day" : "days"} left · {Math.round(elapsedPercent)}% of goal period elapsed</span>{projectedText && <span>Projected finish: {projectedText}</span>}</div>}
  </article>;
}

function formatGoalValue(metric: ReadingGoalMetric, value: number): string { return metric === "pages" ? Math.round(value).toLocaleString() : Math.round(value).toLocaleString(); }
function formatDateRange(start: string, end: string): string { const options: Intl.DateTimeFormatOptions = { month: "short", day: "numeric", year: "numeric" }; return `${new Date(`${start}T12:00:00`).toLocaleDateString(undefined, options)} – ${new Date(`${end}T12:00:00`).toLocaleDateString(undefined, options)}`; }
function Metric({ label, value, note }: { label: string; value: string; note: string }) { return <article className="metric"><span>{label}</span><strong>{value}</strong><small>{note}</small></article>; }
function ChartPanel({ title, subtitle, children }: { title: string; subtitle: string; children: ReactNode }) { return <section className="chart-panel"><div><h2>{title}</h2><p>{subtitle}</p></div>{children}</section>; }
function VerticalBar({ data }: { data: Array<{ name: string; value: number }> }) { return <StatChart option={{ tooltip: { trigger: "axis" }, grid: { left: 42, right: 16, top: 18, bottom: 40 }, xAxis: { type: "category", data: data.map((d) => d.name), axisTick: { show: false }, axisLabel: { rotate: data.length > 8 ? 35 : 0 } }, yAxis: { type: "value", minInterval: 1 }, series: [{ type: "bar", data: data.map((d) => d.value), barMaxWidth: 42, itemStyle: { borderRadius: [6, 6, 0, 0] } }] }} />; }
function HorizontalBar({ data, valueFormatter }: { data: Array<{ name: string; value: number }>; valueFormatter?: (value: number) => string }) { const reversed = [...data].reverse(); return <StatChart option={{ tooltip: { trigger: "axis", valueFormatter: valueFormatter ? (value: unknown) => valueFormatter(Number(value)) : undefined }, grid: { left: 122, right: 24, top: 10, bottom: 24 }, xAxis: { type: "value", minInterval: valueFormatter ? undefined : 1 }, yAxis: { type: "category", data: reversed.map((d) => d.name), axisLabel: { width: 110, overflow: "truncate" } }, series: [{ type: "bar", data: reversed.map((d) => d.value), itemStyle: { borderRadius: [0, 6, 6, 0] } }] }} />; }
function LineChart({ data }: { data: Array<{ name: string; value: number }> }) { return <StatChart option={{ tooltip: { trigger: "axis" }, grid: { left: 52, right: 20, top: 22, bottom: 34 }, xAxis: { type: "category", boundaryGap: false, data: data.map((d) => d.name) }, yAxis: { type: "value" }, series: [{ type: "line", smooth: true, symbolSize: 7, areaStyle: { opacity: 0.08 }, data: data.map((d) => d.value) }] }} />; }
function DualLine({ data, firstLabel, secondLabel }: { data: Array<{ name: string; first: number; second: number }>; firstLabel: string; secondLabel: string }) { return <StatChart option={{ tooltip: { trigger: "axis" }, legend: { bottom: 0 }, grid: { left: 45, right: 20, top: 18, bottom: 50 }, xAxis: { type: "category", boundaryGap: false, data: data.map((d) => d.name) }, yAxis: { type: "value", minInterval: 1 }, series: [{ name: firstLabel, type: "line", smooth: true, data: data.map((d) => d.first) }, { name: secondLabel, type: "line", smooth: true, data: data.map((d) => d.second) }] }} />; }
function DonutChart({ data }: { data: Array<{ name: string; value: number }> }) { return <StatChart option={{ tooltip: { trigger: "item" }, series: [{ type: "pie", radius: ["52%", "76%"], avoidLabelOverlap: true, itemStyle: { borderRadius: 8, borderWidth: 3, borderColor: "#fffdf7" }, label: { color: "#4c4a43" }, data }] }} />; }
function GoalComparison({ goals }: { goals: Array<{ goal: ReadingGoal; progress: { current: number; target: number; percent: number } }> }) { const rows = goals.map(({ goal, progress }) => ({ name: goal.name, value: Math.round(progress.percent) })).sort((a, b) => b.value - a.value); return <HorizontalBar data={rows.slice(0, 12)} valueFormatter={(value) => `${value}%`} />; }
