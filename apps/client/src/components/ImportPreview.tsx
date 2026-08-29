import { useMemo, useState } from "react";
import { AlertTriangle, CheckCircle2, Download, X } from "lucide-react";
import type { ExternalImportPlan } from "../data/importHistory";

export function ImportPreview({ plan, onImport, onCancel }: { plan: ExternalImportPlan; onImport: () => Promise<void>; onCancel: () => void }) {
  const [working, setWorking] = useState(false);
  const [filter, setFilter] = useState<"all" | "new" | "update" | "ambiguous-new">("all");
  const rows = useMemo(() => filter === "all" ? plan.preview : plan.preview.filter((row) => row.action === filter), [filter, plan.preview]);
  async function commit() { setWorking(true); try { await onImport(); } finally { setWorking(false); } }
  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onCancel()}>
    <section className="import-preview-modal">
      <div className="form-header"><div><p className="eyebrow">Review before changing your library</p><h2>{plan.sourceName} import preview</h2></div><button className="icon-button" onClick={onCancel} aria-label="Close"><X size={20} /></button></div>
      <div className="import-preview-metrics"><Metric label="New" value={plan.createdBooks} /><Metric label="Matched" value={plan.updatedBooks} /><Metric label="Ambiguous" value={plan.ambiguousBooks} warning={plan.ambiguousBooks > 0} /><Metric label="New shelves" value={plan.createdShelves} /></div>
      {plan.ambiguousBooks > 0 && <div className="cleanup-intro"><AlertTriangle size={18} /><p>{plan.ambiguousBooks.toLocaleString()} record{plan.ambiguousBooks === 1 ? " has" : "s have"} more than one possible match. For safety, BookStats will import those as separate books so nothing is merged incorrectly. You can resolve them later in Library Cleanup.</p></div>}
      <div className="import-preview-tabs"><button className={filter === "all" ? "active" : ""} onClick={() => setFilter("all")}>All {plan.total}</button><button className={filter === "new" ? "active" : ""} onClick={() => setFilter("new")}>New {plan.createdBooks - plan.ambiguousBooks}</button><button className={filter === "update" ? "active" : ""} onClick={() => setFilter("update")}>Matched {plan.updatedBooks}</button>{plan.ambiguousBooks > 0 && <button className={filter === "ambiguous-new" ? "active" : ""} onClick={() => setFilter("ambiguous-new")}>Ambiguous {plan.ambiguousBooks}</button>}</div>
      <div className="import-preview-list">{rows.slice(0, 300).map((row, index) => <div className="import-preview-row" key={`${row.title}-${row.author}-${index}`}><div><strong>{row.title}</strong><span>{row.author}</span></div><span className={`import-action import-action-${row.action}`}>{row.action === "new" ? "New" : row.action === "update" ? "Update" : "New · review later"}</span>{row.matchedTitle && <small>matches “{row.matchedTitle}”</small>}</div>)}{rows.length > 300 && <p className="preview-truncated">Showing the first 300 of {rows.length.toLocaleString()} records in this view.</p>}</div>
      {plan.warnings.length > 0 && <div className="import-notes"><strong>Import notes</strong>{plan.warnings.map((warning) => <p key={warning}>• {warning}</p>)}</div>}
      <div className="form-actions"><button className="button secondary" onClick={onCancel}>Cancel</button><button className="button primary" disabled={working} onClick={() => void commit()}><Download size={16} />{working ? "Importing…" : `Import ${plan.total.toLocaleString()} books`}</button></div>
    </section>
  </div>;
}

function Metric({ label, value, warning = false }: { label: string; value: number; warning?: boolean }) {
  return <div className={warning ? "warning" : ""}><span>{label}</span><strong>{value.toLocaleString()}</strong>{!warning && value > 0 && <CheckCircle2 size={14} />}</div>;
}
