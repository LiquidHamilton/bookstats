import { useMemo, useState } from "react";
import type { SeriesCatalogBook, SeriesCompletionOverride } from "@bookstats/domain";
import type { SeriesProgress } from "@bookstats/statistics";
import { CheckCircle2, Plus, RotateCcw, Settings2, Trash2, X } from "lucide-react";

interface Props {
  series: SeriesProgress;
  onSave: (override: SeriesCompletionOverride | undefined) => Promise<void>;
  onClose: () => void;
}

function hasCompletionRules(override?: SeriesCompletionOverride): boolean {
  return Boolean(
    override?.expectedCount
    || override?.excludedProviderIds?.length
    || override?.includedProviderIds?.length
    || override?.manualBooks?.length
  );
}

export function SeriesCompletionEditor({ series, onSave, onClose }: Props) {
  const existing = series.completionOverride;
  const automaticCount = series.catalogTotal?.toString() ?? series.catalogPrimaryBooksCount?.toString() ?? "";
  const [expectedCount, setExpectedCount] = useState(existing?.expectedCount?.toString() ?? automaticCount);
  const initialIds = useMemo(() => existing?.includedProviderIds?.length ? existing.includedProviderIds : series.catalogBooks.map((entry) => entry.providerId), [existing, series.catalogBooks]);
  const [includedIds, setIncludedIds] = useState<string[]>(initialIds);
  const [manualBooks, setManualBooks] = useState<SeriesCatalogBook[]>(existing?.manualBooks?.map((entry) => ({ ...entry })) ?? []);
  const [saving, setSaving] = useState(false);
  const automaticIds = series.catalogBooks.map((entry) => entry.providerId);
  const adjusted = hasCompletionRules(existing) || expectedCount !== automaticCount || includedIds.length !== automaticIds.length || includedIds.some((id) => !automaticIds.includes(id)) || manualBooks.length > 0;

  function toggle(id: string) { setIncludedIds((ids) => ids.includes(id) ? ids.filter((item) => item !== id) : [...ids, id]); }
  function addManual() { setManualBooks((items) => [...items, { providerId: `manual:${crypto.randomUUID()}`, title: "", position: "" }]); }
  function updateManual(id: string, patch: Partial<SeriesCatalogBook>) { setManualBooks((items) => items.map((item) => item.providerId === id ? { ...item, ...patch } : item)); }

  async function save() {
    setSaving(true);
    try {
      const count = Number(expectedCount);
      const validManual = manualBooks.filter((entry) => entry.title.trim()).map((entry) => ({ ...entry, title: entry.title.trim(), position: entry.position?.trim() || undefined }));
      const autoSet = new Set(automaticIds);
      const selectionChanged = includedIds.length !== automaticIds.length || includedIds.some((id) => !autoSet.has(id));
      await onSave({
        ignoredFromTracking: existing?.ignoredFromTracking || undefined,
        expectedCount: Number.isInteger(count) && count > 0 ? count : undefined,
        excludedProviderIds: existing?.excludedProviderIds?.length ? existing.excludedProviderIds : undefined,
        includedProviderIds: selectionChanged ? includedIds : undefined,
        manualBooks: validManual.length ? validManual : undefined,
        updatedAt: new Date().toISOString()
      });
      onClose();
    } finally { setSaving(false); }
  }

  async function reset() { setSaving(true); try { await onSave(existing?.ignoredFromTracking ? { ignoredFromTracking: true, updatedAt: new Date().toISOString() } : undefined); onClose(); } finally { setSaving(false); } }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}><section className="series-editor-modal"><div className="form-header"><div><p className="eyebrow">Series completion</p><h2>{series.name}</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div><div className="series-editor-summary"><Settings2 size={20} /><div><strong>Define the mainline collection</strong><p>BookStats reconciles duplicate catalog rows with clearly numbered books already in your library. Multi-volume entries such as 1 & 2 can satisfy more than one position. You can override the expected count or checklist when automatic detection still gets a series wrong.</p>{typeof series.rawCatalogCount === "number" && series.rawCatalogCount > series.catalogBooks.length && <span>Automatic reconciliation turned {series.rawCatalogCount.toLocaleString()} raw catalog entries into {series.catalogBooks.length.toLocaleString()} mainline positions.</span>}</div></div><div className="series-editor-grid"><label>Expected mainline positions<input type="number" min="1" step="1" value={expectedCount} onChange={(event) => setExpectedCount(event.target.value)} placeholder="e.g. 5" /><small>This controls the collection denominator. Leave blank to use the reconciled catalog.</small></label><div className="series-editor-checklist"><div><strong>Books that count toward completion</strong><span>Uncheck side stories or other positions you do not want included in the mainline sequence.</span></div>{series.catalogBooks.length === 0 ? <p className="detail-empty">No catalog checklist is available yet. You can still set an expected count or add manual entries below.</p> : series.catalogBooks.map((entry) => <label key={entry.providerId}><input type="checkbox" checked={includedIds.includes(entry.providerId)} onChange={() => toggle(entry.providerId)} /><span><b>{entry.position ? `#${entry.position}` : "—"}</b><strong>{entry.title}</strong>{entry.owned ? <em><CheckCircle2 size={12} />Owned</em> : entry.inLibrary ? <em>In library · not owned</em> : null}</span></label>)}</div><div className="series-manual-list"><div className="series-manual-heading"><div><strong>Manual entries</strong><span>Add a mainline book the catalog omitted.</span></div><button className="button secondary compact" onClick={addManual}><Plus size={14} />Add entry</button></div>{manualBooks.map((entry) => <div className="series-manual-row" key={entry.providerId}><input value={entry.position ?? ""} onChange={(event) => updateManual(entry.providerId, { position: event.target.value })} placeholder="#" aria-label="Series position" /><input value={entry.title} onChange={(event) => updateManual(entry.providerId, { title: event.target.value })} placeholder="Book title" aria-label="Book title" /><button className="icon-button danger-icon" onClick={() => setManualBooks((items) => items.filter((item) => item.providerId !== entry.providerId))} aria-label="Remove manual entry"><Trash2 size={15} /></button></div>)}</div></div><div className="form-actions series-editor-actions">{adjusted && <button className="button secondary" disabled={saving} onClick={() => void reset()}><RotateCcw size={15} />Use automatic detection</button>}<span /><button className="button secondary" onClick={onClose}>Cancel</button><button className="button primary" disabled={saving} onClick={() => void save()}>{saving ? "Saving…" : "Save completion rules"}</button></div></section></div>;
}
