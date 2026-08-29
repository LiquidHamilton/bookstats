import { normalizedReadDates, normalizedReadingSessions } from "@bookstats/domain";
function sortedDistribution(counts) {
    return [...counts.entries()]
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value || a.name.localeCompare(b.name));
}
function yearOf(value) {
    if (!value)
        return null;
    const year = Number(value.slice(0, 4));
    return Number.isFinite(year) && year > 0 ? year : null;
}
function daysBetween(start, end) {
    if (!start || !end)
        return undefined;
    const a = new Date(`${start}T12:00:00`).getTime();
    const b = new Date(`${end}T12:00:00`).getTime();
    if (!Number.isFinite(a) || !Number.isFinite(b) || b < a)
        return undefined;
    return Math.max(1, Math.round((b - a) / 86_400_000) + 1);
}
export function readEvents(books) {
    return books.flatMap((book) => {
        const completed = normalizedReadingSessions(book).filter((session) => Boolean(session.finishedAt));
        return completed.map((session, index) => {
            const date = session.finishedAt;
            const parsed = new Date(`${date}T12:00:00`);
            const year = yearOf(date) ?? parsed.getFullYear();
            return {
                bookId: book.id,
                title: book.title,
                author: book.author,
                date,
                startedAt: session.startedAt,
                year,
                month: Number.isNaN(parsed.getTime()) ? 0 : parsed.getMonth() + 1,
                pages: book.pages ?? 0,
                isReread: index > 0,
                durationDays: daysBetween(session.startedAt, date)
            };
        });
    }).sort((a, b) => a.date.localeCompare(b.date));
}
export function summarizeLibrary(books) {
    const rated = books.filter((book) => typeof book.rating === "number");
    const read = books.filter((book) => book.status === "read" || normalizedReadDates(book).length > 0);
    const pagesKnown = books.reduce((sum, book) => sum + (book.pages ?? 0), 0);
    const totalReadings = books.reduce((sum, book) => {
        const datedReads = normalizedReadDates(book).length;
        return sum + Math.max(datedReads, book.status === "read" ? 1 : 0);
    }, 0);
    const pagesRead = books.reduce((sum, book) => {
        const datedReads = normalizedReadDates(book).length;
        const readings = Math.max(datedReads, book.status === "read" ? 1 : 0);
        return sum + (book.pages ?? 0) * readings;
    }, 0);
    return {
        totalBooks: books.length,
        ownedBooks: books.filter((book) => book.owned).length,
        readBooks: read.length,
        unreadBooks: books.length - read.length,
        wantToRead: books.filter((book) => book.status === "want_to_read").length,
        pagesKnown,
        pagesRead,
        averageRating: rated.length === 0 ? null : rated.reduce((sum, book) => sum + (book.rating ?? 0), 0) / rated.length,
        uniqueAuthors: new Set(books.map((book) => book.author.trim()).filter(Boolean)).size,
        totalReadings,
        rereads: books.reduce((sum, book) => sum + Math.max(0, normalizedReadDates(book).length - 1), 0),
        uniqueGenres: new Set(books.map((book) => book.genre?.trim()).filter(Boolean)).size,
        uniqueSeries: new Set(books.map((book) => book.series?.trim()).filter(Boolean)).size
    };
}
export function statusDistribution(books) {
    const labels = {
        not_started: "Not Started",
        want_to_read: "Want to Read",
        currently_reading: "Currently Reading",
        read: "Read",
        did_not_finish: "Did Not Finish",
        on_hold: "On Hold"
    };
    const counts = new Map();
    for (const book of books)
        counts.set(book.status, (counts.get(book.status) ?? 0) + 1);
    return [...counts.entries()].map(([status, value]) => ({ name: labels[status], value }));
}
export function ratingDistribution(books) {
    const buckets = Array.from({ length: 10 }, (_, index) => ({ name: `${(index + 1) / 2}★`, value: 0 }));
    for (const book of books) {
        if (typeof book.rating !== "number" || book.rating <= 0)
            continue;
        const index = Math.min(9, Math.max(0, Math.round(book.rating * 2) - 1));
        buckets[index].value += 1;
    }
    return buckets;
}
export function formatDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        const format = book.format ?? "Unknown";
        counts.set(format, (counts.get(format) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function genreDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        const genre = book.genre?.trim() || "Unspecified";
        counts.set(genre, (counts.get(genre) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function seriesDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        if (!book.series?.trim())
            continue;
        counts.set(book.series.trim(), (counts.get(book.series.trim()) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function authorLibraryDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        const authors = [book.author, ...book.additionalAuthors].map((author) => author.trim()).filter(Boolean);
        for (const author of new Set(authors))
            counts.set(author, (counts.get(author) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function authorReadDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        const datedReads = normalizedReadDates(book).length;
        const readings = Math.max(datedReads, book.status === "read" ? 1 : 0);
        if (readings === 0)
            continue;
        counts.set(book.author, (counts.get(book.author) ?? 0) + readings);
    }
    return sortedDistribution(counts);
}
export function booksReadByYear(books) {
    const counts = new Map();
    for (const event of readEvents(books))
        counts.set(String(event.year), (counts.get(String(event.year)) ?? 0) + 1);
    return [...counts.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => a.name.localeCompare(b.name));
}
export function pagesReadByYear(books) {
    const counts = new Map();
    for (const event of readEvents(books))
        counts.set(String(event.year), (counts.get(String(event.year)) ?? 0) + event.pages);
    return [...counts.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => a.name.localeCompare(b.name));
}
export function booksAddedByYear(books) {
    const counts = new Map();
    for (const book of books) {
        const year = yearOf(book.dateAdded);
        if (!year)
            continue;
        counts.set(String(year), (counts.get(String(year)) ?? 0) + 1);
    }
    return [...counts.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => a.name.localeCompare(b.name));
}
export function publicationDecadeDistribution(books) {
    const counts = new Map();
    for (const book of books) {
        if (!book.publicationYear)
            continue;
        const decade = Math.floor(book.publicationYear / 10) * 10;
        const name = `${decade}s`;
        counts.set(name, (counts.get(name) ?? 0) + 1);
    }
    return [...counts.entries()]
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => Number(a.name.slice(0, 4)) - Number(b.name.slice(0, 4)));
}
export function pageLengthDistribution(books) {
    const buckets = [
        { name: "< 200", min: 0, max: 199, value: 0 },
        { name: "200–299", min: 200, max: 299, value: 0 },
        { name: "300–399", min: 300, max: 399, value: 0 },
        { name: "400–499", min: 400, max: 499, value: 0 },
        { name: "500–699", min: 500, max: 699, value: 0 },
        { name: "700+", min: 700, max: Number.POSITIVE_INFINITY, value: 0 }
    ];
    for (const book of books) {
        if (!book.pages)
            continue;
        const bucket = buckets.find((candidate) => book.pages >= candidate.min && book.pages <= candidate.max);
        if (bucket)
            bucket.value += 1;
    }
    return buckets.map(({ name, value }) => ({ name, value }));
}
export function averageRatingByAuthor(books) {
    const values = new Map();
    for (const book of books) {
        if (typeof book.rating !== "number")
            continue;
        const current = values.get(book.author) ?? [];
        current.push(book.rating);
        values.set(book.author, current);
    }
    return [...values.entries()]
        .map(([name, ratings]) => ({ name, value: ratings.reduce((sum, rating) => sum + rating, 0) / ratings.length, count: ratings.length }))
        .sort((a, b) => b.value - a.value || b.count - a.count || a.name.localeCompare(b.name));
}
export function readingMonthDistribution(books) {
    const labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const counts = Array(12).fill(0);
    for (const event of readEvents(books))
        if (event.month >= 1 && event.month <= 12)
            counts[event.month - 1] += 1;
    return labels.map((name, index) => ({ name, value: counts[index] }));
}
export function monthlyReadingForYear(books, year) {
    const labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const rows = labels.map((name) => ({ name, books: 0, pages: 0 }));
    for (const event of readEvents(books)) {
        if (event.year !== year || event.month < 1 || event.month > 12)
            continue;
        rows[event.month - 1].books += 1;
        rows[event.month - 1].pages += event.pages;
    }
    return rows;
}
export function newAuthorsByYear(books) {
    const firstReadByAuthor = new Map();
    for (const event of readEvents(books)) {
        const key = event.author.trim().toLocaleLowerCase();
        if (!key)
            continue;
        const existing = firstReadByAuthor.get(key);
        if (!existing || event.date < existing.date)
            firstReadByAuthor.set(key, { name: event.author.trim(), date: event.date });
    }
    const counts = new Map();
    for (const { date } of firstReadByAuthor.values()) {
        const year = yearOf(date);
        if (!year)
            continue;
        const key = String(year);
        counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return [...counts.entries()].map(([name, value]) => ({ name, value })).sort((a, b) => a.name.localeCompare(b.name));
}
export function ratingTrendByYear(books) {
    const bookById = new Map(books.map((book) => [book.id, book]));
    const values = new Map();
    for (const event of readEvents(books)) {
        const rating = bookById.get(event.bookId)?.rating;
        if (typeof rating !== "number")
            continue;
        const key = String(event.year);
        const list = values.get(key) ?? [];
        list.push(rating);
        values.set(key, list);
    }
    return [...values.entries()].map(([name, ratings]) => ({
        name, value: ratings.reduce((sum, rating) => sum + rating, 0) / ratings.length
    })).sort((a, b) => a.name.localeCompare(b.name));
}
export function readFormatDistribution(books) {
    const bookById = new Map(books.map((book) => [book.id, book]));
    const counts = new Map();
    for (const event of readEvents(books)) {
        const name = bookById.get(event.bookId)?.format ?? "Unknown";
        counts.set(name, (counts.get(name) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function readGenreDistribution(books) {
    const bookById = new Map(books.map((book) => [book.id, book]));
    const counts = new Map();
    for (const event of readEvents(books)) {
        const name = bookById.get(event.bookId)?.genre?.trim() || "Unspecified";
        counts.set(name, (counts.get(name) ?? 0) + 1);
    }
    return sortedDistribution(counts);
}
export function readOwnershipDistribution(books) {
    const bookById = new Map(books.map((book) => [book.id, book]));
    let owned = 0;
    let unowned = 0;
    for (const event of readEvents(books)) {
        if (bookById.get(event.bookId)?.owned)
            owned += 1;
        else
            unowned += 1;
    }
    return [{ name: "Currently owned", value: owned }, { name: "Not currently owned", value: unowned }].filter((item) => item.value > 0);
}
export function readingExtremes(books) {
    const completed = books.filter((book) => normalizedReadDates(book).length > 0 && typeof book.pages === "number" && book.pages > 0);
    if (!completed.length)
        return {};
    const longest = [...completed].sort((a, b) => (b.pages ?? 0) - (a.pages ?? 0) || a.title.localeCompare(b.title))[0];
    const shortest = [...completed].sort((a, b) => (a.pages ?? 0) - (b.pages ?? 0) || a.title.localeCompare(b.title))[0];
    return {
        longest: { title: longest.title, author: longest.author, pages: longest.pages },
        shortest: { title: shortest.title, author: shortest.author, pages: shortest.pages }
    };
}
function dayNumber(value) {
    const date = new Date(`${value}T12:00:00`);
    return Number.isNaN(date.getTime()) ? undefined : Math.floor(date.getTime() / 86_400_000);
}
export function readingGoalProgress(goal, books, today = new Date().toISOString().slice(0, 10)) {
    const events = readEvents(books);
    const inRange = events.filter((event) => event.date >= goal.startDate && event.date <= goal.endDate);
    let current = 0;
    switch (goal.metric) {
        case "books":
            current = inRange.length;
            break;
        case "pages":
            current = inRange.reduce((sum, event) => sum + event.pages, 0);
            break;
        case "rereads":
            current = inRange.filter((event) => event.isReread).length;
            break;
        case "owned_books": {
            const owned = new Set(books.filter((book) => book.owned).map((book) => book.id));
            current = inRange.filter((event) => owned.has(event.bookId)).length;
            break;
        }
        case "new_authors": {
            const firstReadByAuthor = new Map();
            for (const event of events) {
                const key = event.author.trim().toLocaleLowerCase();
                if (!key)
                    continue;
                const existing = firstReadByAuthor.get(key);
                if (!existing || event.date < existing)
                    firstReadByAuthor.set(key, event.date);
            }
            current = [...firstReadByAuthor.values()].filter((date) => date >= goal.startDate && date <= goal.endDate).length;
            break;
        }
    }
    const target = Math.max(1, goal.target);
    const percent = Math.min(100, (current / target) * 100);
    const start = dayNumber(goal.startDate);
    const end = dayNumber(goal.endDate);
    const now = dayNumber(today);
    let elapsedPercent = 0;
    let daysRemaining = 0;
    let projected = null;
    if (start !== undefined && end !== undefined && now !== undefined && end >= start) {
        const totalDays = end - start + 1;
        const elapsedDays = now < start ? 0 : Math.min(totalDays, now - start + 1);
        elapsedPercent = Math.min(100, (elapsedDays / totalDays) * 100);
        daysRemaining = now > end ? 0 : Math.max(0, end - Math.max(now, start) + 1);
        if (elapsedDays > 0 && now <= end)
            projected = current / elapsedDays * totalDays;
    }
    const complete = current >= target;
    return { current, target, percent, remaining: Math.max(0, target - current), complete, elapsedPercent, daysRemaining, onPace: complete || percent + 0.0001 >= elapsedPercent, projected };
}
export function readingPaceSummary(books) {
    const events = readEvents(books);
    const timed = events.filter((event) => typeof event.durationDays === "number" && event.durationDays > 0);
    const pageTimed = timed.filter((event) => event.pages > 0);
    const durations = timed.map((event) => event.durationDays);
    const totalPages = events.reduce((sum, event) => sum + event.pages, 0);
    return {
        timedReads: timed.length,
        averageDaysToFinish: durations.length ? durations.reduce((a, b) => a + b, 0) / durations.length : null,
        fastestDays: durations.length ? Math.min(...durations) : null,
        slowestDays: durations.length ? Math.max(...durations) : null,
        averagePagesPerDay: pageTimed.length ? pageTimed.reduce((sum, event) => sum + event.pages / event.durationDays, 0) / pageTimed.length : null,
        averagePagesPerRead: events.length && totalPages ? totalPages / events.length : null,
        activeReadings: books.reduce((sum, book) => sum + normalizedReadingSessions(book).filter((session) => !session.finishedAt && Boolean(session.startedAt || session.progressPages)).length, 0)
    };
}
function normalizeSeriesText(value) {
    return value?.normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(/[^a-z0-9]+/g, " ").trim() ?? "";
}
function normalizeSeriesIsbn(value) {
    const isbn = value?.replace(/[^0-9Xx]/g, "").toUpperCase();
    return isbn && (isbn.length === 10 || isbn.length === 13) ? isbn : undefined;
}
function seriesIsbn13From10(value) {
    const isbn = normalizeSeriesIsbn(value);
    if (!isbn || isbn.length !== 10 || !/^\d{9}[\dX]$/.test(isbn))
        return undefined;
    const body = `978${isbn.slice(0, 9)}`;
    let sum = 0;
    for (let index = 0; index < 12; index += 1)
        sum += Number(body[index]) * (index % 2 === 0 ? 1 : 3);
    return `${body}${(10 - (sum % 10)) % 10}`;
}
function seriesIsbnEquivalent(left, right) {
    const a = normalizeSeriesIsbn(left);
    const b = normalizeSeriesIsbn(right);
    if (!a || !b)
        return false;
    if (a === b)
        return true;
    return a.length === 10 ? seriesIsbn13From10(a) === b : seriesIsbn13From10(b) === a;
}
function isBookRead(book) {
    return book.status === "read" || normalizedReadDates(book).length > 0;
}
function chooseSeriesCatalog(items) {
    const catalogs = items.map((book) => book.seriesMetadata).filter((value) => Boolean(value));
    const providerRank = { hardcover: 3, googlebooks: 2, openlibrary: 1 };
    return catalogs.sort((a, b) => {
        const content = Number(b.books.length > 0) - Number(a.books.length > 0);
        if (content)
            return content;
        const provider = providerRank[b.provider] - providerRank[a.provider];
        if (provider)
            return provider;
        return b.books.length - a.books.length;
    })[0];
}
function chooseSeriesCompletionOverride(items) {
    return items
        .map((book) => book.seriesCompletionOverride)
        .filter((value) => Boolean(value))
        .sort((a, b) => (b.updatedAt ?? "").localeCompare(a.updatedAt ?? ""))[0];
}
function seriesPositions(value) {
    if (!value?.trim())
        return [];
    const clean = value.trim().replace(/[–—]/g, "-").replace(/^#\s*/, "");
    const single = clean.match(/^\d+(?:\.\d+)?$/);
    if (single)
        return [Number(clean)];
    const range = clean.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
        const start = Number(range[1]);
        const end = Number(range[2]);
        if (start > 0 && end >= start && end - start <= 50)
            return Array.from({ length: end - start + 1 }, (_, index) => start + index);
    }
    const parts = clean.split(/\s*(?:&|\+|,|\band\b)\s*/i).filter(Boolean);
    if (parts.length < 2 || parts.some((part) => !/^\d+(?:\.\d+)?$/.test(part)))
        return [];
    return [...new Set(parts.map(Number).filter((number) => Number.isFinite(number) && number > 0))];
}
function numericSeriesPosition(value) {
    const positions = seriesPositions(value);
    return positions.length === 1 ? positions[0] : undefined;
}
function seriesBookSort(a, b) {
    const left = Number(a.position);
    const right = Number(b.position);
    if (Number.isFinite(left) && Number.isFinite(right))
        return left - right || a.title.localeCompare(b.title);
    if (Number.isFinite(left))
        return -1;
    if (Number.isFinite(right))
        return 1;
    return String(a.position ?? "").localeCompare(String(b.position ?? "")) || a.title.localeCompare(b.title);
}
function seriesEntryPreference(entry, libraryBooks) {
    let score = 0;
    const title = normalizeSeriesText(entry.title);
    const author = normalizeSeriesText(entry.author);
    if (title && libraryBooks.some((book) => normalizeSeriesText(book.title) === title))
        score += 100;
    if (author && libraryBooks.some((book) => normalizeSeriesText(book.author) === author || (book.additionalAuthors ?? []).some((name) => normalizeSeriesText(name) === author)))
        score += 30;
    if (/^[\x00-\x7F]*$/.test(entry.title))
        score += 5;
    if (entry.isbn)
        score += 2;
    return score - Math.min(20, entry.title.length / 20);
}
function inferMainlineCountFromRepeatedPositions(rows) {
    const positions = rows.map((entry) => numericSeriesPosition(entry.position)).filter((value) => Boolean(value && Number.isInteger(value) && value > 0 && value <= 500));
    if (positions.length < 4)
        return undefined;
    const unique = [...new Set(positions)].sort((a, b) => a - b);
    const max = unique.at(-1) ?? 0;
    if (max < 2 || max > 500)
        return undefined;
    const dense = unique.filter((position) => position >= 1 && position <= max).length / max;
    const variantRatio = positions.length / unique.length;
    // Infer a mainline size only when many catalog rows reuse a compact, nearly
    // continuous sequence of integer positions. This avoids treating a genuine
    // 100-book series as a five-book series while still recognizing 100+ language
    // variants spread across positions 1–5.
    return dense >= 0.75 && variantRatio >= 2 ? max : undefined;
}
function cleanSeriesCatalog(catalog, libraryBooks, override) {
    if (!catalog)
        return override?.manualBooks?.map((entry) => ({ ...entry })) ?? [];
    const unique = [...new Map(catalog.books.map((entry) => [entry.providerId, entry])).values()];
    const included = new Set(override?.includedProviderIds ?? []);
    const excluded = new Set(override?.excludedProviderIds ?? []);
    let rows = unique.filter((entry) => !excluded.has(entry.providerId) && (included.size === 0 || included.has(entry.providerId)));
    // Provider series feeds frequently contain translated/alternate work records at
    // the same numbered position. If the provider tells us how many primary books
    // exist, collapse those variants to one representative per mainline position.
    // This is deliberately position-based rather than title/language-based so a
    // five-book mainline series cannot turn into 100+ "missing" translations.
    const inferredMainline = inferMainlineCountFromRepeatedPositions(rows);
    const providerPrimary = catalog.primaryBooksCount;
    // Some catalogs report every translated/alternate work as a "primary" book.
    // When the actual rows strongly repeat a compact numbered sequence, trust that
    // sequence over a wildly larger provider count. Keep close disagreements intact
    // so an incomplete six-book feed containing only positions 1–5 does not get
    // silently redefined as a five-book series.
    const automaticExpected = inferredMainline && (!providerPrimary || providerPrimary >= inferredMainline * 3)
        ? inferredMainline
        : providerPrimary ?? inferredMainline;
    const expected = override?.expectedCount ?? automaticExpected;
    if (included.size === 0 && expected && expected > 0 && expected <= 500) {
        const byPosition = new Map();
        for (const entry of rows) {
            const position = numericSeriesPosition(entry.position);
            if (!position || !Number.isInteger(position) || position < 1 || position > expected)
                continue;
            const group = byPosition.get(position) ?? [];
            group.push(entry);
            byPosition.set(position, group);
        }
        const requiredCoverage = expected <= 3 ? expected : Math.max(3, Math.ceil(expected * 0.6));
        if (byPosition.size >= Math.min(expected, requiredCoverage)) {
            rows = [...byPosition.entries()]
                .sort(([a], [b]) => a - b)
                .map(([, entries]) => [...entries].sort((a, b) => seriesEntryPreference(b, libraryBooks) - seriesEntryPreference(a, libraryBooks))[0]);
        }
    }
    // A provider catalog can be incomplete even when it reports the correct mainline
    // size. Clearly numbered books already in the user's library are strong evidence
    // for omitted positions. Supplement only missing integer positions, respect an
    // explicit provider whitelist, and never resurrect a deliberately excluded row.
    if (included.size === 0) {
        const existingPositions = new Set(rows.flatMap((entry) => seriesPositions(entry.position)).filter((position) => Number.isInteger(position)));
        const excludedPositions = new Set(unique.filter((entry) => excluded.has(entry.providerId)).flatMap((entry) => seriesPositions(entry.position)).filter((position) => Number.isInteger(position)));
        const expectedLimit = override?.expectedCount ?? automaticExpected;
        for (const book of libraryBooks) {
            for (const position of seriesPositions(book.seriesVolume)) {
                if (!Number.isInteger(position) || position < 1 || position > 500)
                    continue;
                if (expectedLimit && position > expectedLimit)
                    continue;
                const providerId = `library:${book.id}:position:${position}`;
                if (existingPositions.has(position) || excludedPositions.has(position) || excluded.has(providerId))
                    continue;
                rows.push({ providerId, title: book.title, position: String(position), isbn: book.isbn, coverUrl: book.coverUrl, author: book.author });
                existingPositions.add(position);
            }
        }
    }
    const manual = override?.manualBooks ?? [];
    if (manual.length) {
        const seen = new Set(rows.map((entry) => entry.providerId));
        for (const entry of manual)
            if (!seen.has(entry.providerId)) {
                rows.push({ ...entry });
                seen.add(entry.providerId);
            }
    }
    return [...rows].sort(seriesBookSort);
}
function catalogBookMatch(catalog, books) {
    if (catalog.isbn) {
        const exact = books.find((book) => seriesIsbnEquivalent(book.isbn, catalog.isbn));
        if (exact)
            return exact;
    }
    if (catalog.position) {
        const catalogPositions = seriesPositions(catalog.position);
        if (catalogPositions.length) {
            const byPosition = books.find((book) => {
                const ownedPositions = seriesPositions(book.seriesVolume);
                return catalogPositions.some((position) => ownedPositions.includes(position));
            });
            if (byPosition)
                return byPosition;
        }
        else {
            const byPosition = books.find((book) => normalizeSeriesText(book.seriesVolume) === normalizeSeriesText(catalog.position));
            if (byPosition)
                return byPosition;
        }
    }
    const title = normalizeSeriesText(catalog.title);
    return title ? books.find((book) => normalizeSeriesText(book.title) === title) : undefined;
}
export function seriesProgress(books) {
    const groups = new Map();
    for (const book of books) {
        const name = book.seriesMetadata?.name?.trim() || book.series?.trim();
        if (!name)
            continue;
        const key = normalizeSeriesText(name);
        const current = groups.get(key) ?? { name, books: [] };
        current.books.push(book);
        // Prefer the provider's canonical series name when it is available.
        if (book.seriesMetadata?.name?.trim())
            current.name = book.seriesMetadata.name.trim();
        groups.set(key, current);
    }
    return [...groups.values()].map(({ name, books: items }) => {
        const read = items.filter(isBookRead).length;
        const volumes = [...new Set(items.flatMap((book) => seriesPositions(book.seriesVolume)).filter((value) => Number.isInteger(value) && value > 0))].sort((a, b) => a - b);
        const gaps = [];
        if (volumes.length >= 2)
            for (let volume = volumes[0]; volume <= volumes[volumes.length - 1]; volume += 1)
                if (!volumes.includes(volume))
                    gaps.push(volume);
        const catalog = chooseSeriesCatalog(items);
        const completionOverride = chooseSeriesCompletionOverride(items);
        const cleanedCatalog = cleanSeriesCatalog(catalog, items, completionOverride);
        const catalogBooks = cleanedCatalog.map((entry) => {
            const match = catalogBookMatch(entry, items);
            return { ...entry, inLibrary: Boolean(match), owned: Boolean(match?.owned), read: Boolean(match && isBookRead(match)) };
        });
        const inferredMainline = catalog ? inferMainlineCountFromRepeatedPositions(catalog.books) : undefined;
        const providerPrimary = catalog?.primaryBooksCount;
        const automaticCatalogTotal = inferredMainline && (!providerPrimary || providerPrimary >= inferredMainline * 3)
            ? inferredMainline
            : providerPrimary ?? (catalogBooks.length || catalog?.totalBooks || undefined);
        const catalogTotal = completionOverride?.expectedCount ?? automaticCatalogTotal;
        const inCatalog = catalogBooks.filter((entry) => entry.owned).length;
        return {
            name,
            total: items.length,
            read,
            owned: items.filter((book) => book.owned).length,
            completionPercent: items.length ? (read / items.length) * 100 : 0,
            knownVolumes: volumes,
            missingVolumeGaps: gaps,
            catalogProvider: catalog?.provider,
            catalogTotal,
            catalogPrimaryBooksCount: catalog?.primaryBooksCount,
            catalogIsCompleted: catalog?.isCompleted,
            rawCatalogCount: catalog?.books.length,
            completionOverride,
            collectionPercent: catalogTotal ? Math.min(100, (inCatalog / catalogTotal) * 100) : undefined,
            catalogBooks,
            missingCatalogBooks: catalogBooks.filter((entry) => !entry.owned)
        };
    }).sort((a, b) => {
        const leftSize = a.catalogTotal ?? a.total;
        const rightSize = b.catalogTotal ?? b.total;
        return rightSize - leftSize || a.name.localeCompare(b.name);
    });
}
export function libraryFlowByYear(books) {
    const years = new Map();
    const ensure = (year) => { const row = years.get(year) ?? { name: year, added: 0, read: 0, pages: 0 }; years.set(year, row); return row; };
    for (const book of books) {
        const addedYear = yearOf(book.dateAdded);
        if (addedYear)
            ensure(String(addedYear)).added += 1;
    }
    for (const event of readEvents(books)) {
        const row = ensure(String(event.year));
        row.read += 1;
        row.pages += event.pages;
    }
    return [...years.values()].sort((a, b) => a.name.localeCompare(b.name));
}
