import type { CoverInspection } from "./covers";
import type { LibraryRepository } from "./libraryRepository";

const COVER_HEALTH_META_KEY = "libraryHealth.coverInspections:v1";

export interface PersistedCoverHealthEntry {
  revision: string;
  inspection: CoverInspection;
  checkedAt: string;
}

interface PersistedCoverHealthPayload {
  version: 1;
  entries: Record<string, PersistedCoverHealthEntry>;
}

export async function loadCoverHealth(repository: LibraryRepository): Promise<Map<string, PersistedCoverHealthEntry>> {
  try {
    const raw = await repository.getMeta(COVER_HEALTH_META_KEY);
    if (!raw) return new Map();
    const parsed = JSON.parse(raw) as Partial<PersistedCoverHealthPayload>;
    if (parsed.version !== 1 || !parsed.entries || typeof parsed.entries !== "object") return new Map();
    const result = new Map<string, PersistedCoverHealthEntry>();
    for (const [bookId, entry] of Object.entries(parsed.entries)) {
      if (!entry || typeof entry !== "object") continue;
      const candidate = entry as PersistedCoverHealthEntry;
      if (!candidate.revision || !candidate.inspection || typeof candidate.inspection.usable !== "boolean") continue;
      result.set(bookId, candidate);
    }
    return result;
  } catch {
    return new Map();
  }
}

export async function saveCoverHealth(repository: LibraryRepository, entries: ReadonlyMap<string, PersistedCoverHealthEntry>): Promise<void> {
  const serialized: Record<string, PersistedCoverHealthEntry> = {};
  for (const [bookId, entry] of entries) serialized[bookId] = entry;
  const payload: PersistedCoverHealthPayload = { version: 1, entries: serialized };
  await repository.setMeta(COVER_HEALTH_META_KEY, JSON.stringify(payload));
}
