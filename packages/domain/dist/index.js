export const BOOK_CONDITIONS = ["New", "Like New", "Very Good", "Good", "Acceptable", "Poor"];
export const READING_STATUS_LABELS = {
    not_started: "Not Started",
    want_to_read: "Want to Read",
    currently_reading: "Currently Reading",
    read: "Read",
    did_not_finish: "Did Not Finish",
    on_hold: "On Hold"
};
export function normalizedReadingSessions(book) {
    const sessions = Array.isArray(book.readingSessions) ? book.readingSessions.filter(Boolean).map((session) => ({ ...session })) : [];
    if (sessions.length > 0)
        return sessions.sort((a, b) => (a.finishedAt ?? a.startedAt ?? "9999").localeCompare(b.finishedAt ?? b.startedAt ?? "9999"));
    const legacyDates = Array.isArray(book.readDates) ? book.readDates.filter(Boolean) : [];
    if (book.dateRead)
        legacyDates.push(book.dateRead);
    return [...new Set(legacyDates)].sort().map((date, index) => ({
        id: `legacy-${index}-${date}`,
        finishedAt: date,
        createdAt: `${date}T12:00:00.000Z`,
        updatedAt: `${date}T12:00:00.000Z`
    }));
}
export function normalizedReadDates(book) {
    const dates = [
        ...(Array.isArray(book.readDates) ? book.readDates.filter(Boolean) : []),
        ...(Array.isArray(book.readingSessions) ? book.readingSessions.map((session) => session.finishedAt).filter((value) => Boolean(value)) : []),
        ...(book.dateRead ? [book.dateRead] : [])
    ];
    return [...new Set(dates)].sort();
}
export function activeReadingSession(book) {
    return [...normalizedReadingSessions(book)].reverse().find((session) => !session.finishedAt);
}
export function normalizedLoans(book) {
    return (Array.isArray(book.loans) ? book.loans : [])
        .filter((loan) => Boolean(loan?.id && loan.borrower?.trim() && loan.loanedAt))
        .map((loan) => ({ ...loan, borrower: loan.borrower.trim() }))
        .sort((a, b) => (a.loanedAt || a.createdAt).localeCompare(b.loanedAt || b.createdAt));
}
export function activeLoan(book) {
    return [...normalizedLoans(book)].reverse().find((loan) => !loan.returnedAt);
}
export function loanIsOverdue(loan, today = new Date().toISOString().slice(0, 10)) {
    return Boolean(!loan.returnedAt && loan.dueAt && loan.dueAt < today);
}
export function normalizeIsbn(value) {
    return value.replace(/[^0-9Xx]/g, "").toUpperCase();
}
export function isSmartShelf(shelf) {
    return shelf.kind === "smart";
}
export function sortShelves(shelves) {
    const explicitOrders = shelves.map((shelf) => shelf.order).filter((value) => Number.isFinite(value));
    let fallbackOrder = explicitOrders.length ? Math.max(...explicitOrders) + 1 : 0;
    return [...shelves]
        .sort((a, b) => {
        const aOrder = typeof a.order === "number" && Number.isFinite(a.order) ? a.order : Number.MAX_SAFE_INTEGER;
        const bOrder = typeof b.order === "number" && Number.isFinite(b.order) ? b.order : Number.MAX_SAFE_INTEGER;
        return aOrder - bOrder || a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
    })
        .map((shelf) => typeof shelf.order === "number" && Number.isFinite(shelf.order) ? shelf : { ...shelf, order: fallbackOrder++ });
}
export function normalizedShelfRuleGroups(shelf) {
    if (Array.isArray(shelf.ruleGroups) && shelf.ruleGroups.length > 0) {
        return shelf.ruleGroups
            .filter((group) => Array.isArray(group.rules) && group.rules.length > 0)
            .map((group, index) => ({
            id: group.id || `group-${index + 1}`,
            match: group.match === "any" ? "any" : "all",
            rules: group.rules
        }));
    }
    const legacyRules = Array.isArray(shelf.rules) ? shelf.rules : [];
    if (legacyRules.length === 0)
        return [];
    return [{ id: "legacy", match: shelf.match === "any" ? "any" : "all", rules: legacyRules }];
}
export function shelfRuleGroupMatchesBook(group, book) {
    if (!group.rules.length)
        return false;
    const checks = group.rules.map((rule) => shelfRuleMatchesBook(rule, book));
    return group.match === "any" ? checks.some(Boolean) : checks.every(Boolean);
}
export function shelfMatchesBook(shelf, book) {
    if (!isSmartShelf(shelf))
        return (book.shelfIds ?? []).includes(shelf.id);
    const grouped = Array.isArray(shelf.ruleGroups) && shelf.ruleGroups.length > 0;
    const groups = normalizedShelfRuleGroups(shelf);
    if (groups.length === 0)
        return false;
    const checks = groups.map((group) => shelfRuleGroupMatchesBook(group, book));
    // On legacy flat shelves, match belongs to the rules inside the one generated group.
    // On v0.9.4+ grouped shelves, match combines the groups themselves.
    return grouped && shelf.match === "all" ? checks.every(Boolean) : grouped ? checks.some(Boolean) : checks[0];
}
export function shelfRuleMatchesBook(rule, book) {
    const text = (value) => String(value ?? "").trim().toLocaleLowerCase();
    const number = (value) => {
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : undefined;
    };
    const compareText = (actual) => {
        const left = text(actual);
        const right = text(rule.value);
        if (rule.operator === "contains")
            return Boolean(right) && left.includes(right);
        if (rule.operator === "not_contains")
            return Boolean(right) && !left.includes(right);
        if (rule.operator === "not_equals")
            return left !== right;
        return left === right;
    };
    const compareNumber = (actual) => {
        const left = number(actual);
        const right = number(rule.value);
        if (left === undefined || right === undefined)
            return false;
        if (rule.operator === "gte")
            return left >= right;
        if (rule.operator === "lte")
            return left <= right;
        if (rule.operator === "not_equals")
            return left !== right;
        return left === right;
    };
    const compareDate = (actual) => {
        const left = String(actual ?? "").slice(0, 10);
        const right = String(rule.value ?? "").slice(0, 10);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(left) || !/^\d{4}-\d{2}-\d{2}$/.test(right))
            return false;
        if (rule.operator === "gte")
            return left >= right;
        if (rule.operator === "lte")
            return left <= right;
        if (rule.operator === "not_equals")
            return left !== right;
        return left === right;
    };
    switch (rule.field) {
        case "owned": return rule.operator === "is_false" ? !book.owned : book.owned;
        case "status": return compareText(book.status);
        case "condition": return compareText(book.condition);
        case "rating": return compareNumber(book.rating);
        case "pages": return compareNumber(book.pages);
        case "publicationYear": return compareNumber(book.publicationYear);
        case "readCount": return compareNumber(normalizedReadDates(book).length);
        case "lastRead": return compareDate(normalizedReadDates(book).at(-1));
        case "dateAdded": return compareDate(book.dateAdded);
        case "tag": {
            const value = text(rule.value);
            const tags = (book.tags ?? []).map(text);
            if (rule.operator === "not_contains" || rule.operator === "not_equals")
                return !tags.some((tag) => rule.operator === "not_contains" ? tag.includes(value) : tag === value);
            return tags.some((tag) => rule.operator === "contains" ? tag.includes(value) : tag === value);
        }
        case "title": return compareText(book.title);
        case "author": {
            const value = text(rule.value);
            const authors = [book.author, ...(book.additionalAuthors ?? [])].map(text);
            if (rule.operator === "not_contains")
                return authors.every((author) => !author.includes(value));
            if (rule.operator === "not_equals")
                return authors.every((author) => author !== value);
            if (rule.operator === "contains")
                return Boolean(value) && authors.some((author) => author.includes(value));
            return authors.some((author) => author === value);
        }
        case "series": return compareText(book.series);
        case "format": return compareText(book.format);
        case "genre": return compareText(book.genre);
    }
}
