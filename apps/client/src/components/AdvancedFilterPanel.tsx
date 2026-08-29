import { useMemo, useState } from "react";
import type { Book, Shelf, ShelfMatchMode, ShelfRuleGroup } from "@bookstats/domain";
import { shelfMatchesBook } from "@bookstats/domain";
import { Filter, Save, X } from "lucide-react";
import { createRuleGroup, RuleBuilder, ruleGroupsAreValid } from "./RuleBuilder";

export interface AdvancedFilterState {
  /** How separate rule groups are combined: all = AND, any = OR. */
  match: ShelfMatchMode;
  ruleGroups: ShelfRuleGroup[];
}

function cloneGroups(groups: ShelfRuleGroup[]): ShelfRuleGroup[] {
  return groups.map((group) => ({ ...group, rules: group.rules.map((rule) => ({ ...rule })) }));
}

export function filterMatchesBook(filter: AdvancedFilterState | undefined, book: Book): boolean {
  if (!filter || !ruleGroupsAreValid(filter.ruleGroups)) return true;
  const shelf: Shelf = { id: "advanced-filter", name: "Advanced filter", kind: "smart", match: filter.match, ruleGroups: filter.ruleGroups, createdAt: "", updatedAt: "" };
  return shelfMatchesBook(shelf, book);
}

export function AdvancedFilterPanel({
  books,
  initial,
  onApply,
  onSaveAsShelf,
  onClose
}: {
  books: Book[];
  initial?: AdvancedFilterState;
  onApply: (filter: AdvancedFilterState | undefined) => void;
  onSaveAsShelf: (name: string, filter: AdvancedFilterState) => Promise<void>;
  onClose: () => void;
}) {
  const [match, setMatch] = useState<ShelfMatchMode>(initial?.match ?? "any");
  const [groups, setGroups] = useState<ShelfRuleGroup[]>(initial?.ruleGroups?.length ? cloneGroups(initial.ruleGroups) : [createRuleGroup("all")]);
  const [shelfName, setShelfName] = useState("");
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState<string>();
  const valid = ruleGroupsAreValid(groups);
  const preview = useMemo(() => valid ? books.filter((book) => filterMatchesBook({ match, ruleGroups: groups }, book)).length : books.length, [books, groups, match, valid]);

  async function saveShelf() {
    if (!valid || !shelfName.trim()) return;
    setSaving(true); setMessage(undefined);
    try {
      await onSaveAsShelf(shelfName.trim(), { match, ruleGroups: cloneGroups(groups) });
      setMessage(`Saved “${shelfName.trim()}” as a smart shelf.`);
      setShelfName("");
    } finally { setSaving(false); }
  }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
    <section className="advanced-filter-modal">
      <div className="form-header"><div><p className="eyebrow">Powerful when you need it</p><h2>Advanced filters</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div>
      <p className="shelf-manager-intro">Build a temporary library filter using the same grouped Boolean rules as smart shelves. Rules inside a group can use AND or OR, and separate groups can also be combined with AND or OR.</p>
      <RuleBuilder match={match} groups={groups} onMatchChange={setMatch} onGroupsChange={setGroups} />
      <div className="advanced-filter-summary"><Filter size={15} /><span>{valid ? <><strong>{preview.toLocaleString()}</strong> of {books.length.toLocaleString()} books match</> : "Complete the rules to preview matches"}</span></div>
      <div className="advanced-filter-actions"><button className="button primary" disabled={!valid} onClick={() => { onApply({ match, ruleGroups: cloneGroups(groups) }); onClose(); }}><Filter size={16} />Apply filter</button><button className="button secondary" onClick={() => { onApply(undefined); onClose(); }}>Clear filter</button></div>
      <div className="save-filter-row"><div><strong>Save this filter</strong><span>Creates a smart shelf that updates automatically.</span></div><input value={shelfName} onChange={(event) => setShelfName(event.target.value)} placeholder="Smart shelf name" /><button className="button secondary compact" disabled={!valid || !shelfName.trim() || saving} onClick={() => void saveShelf()}><Save size={15} />{saving ? "Saving…" : "Save as shelf"}</button></div>
      {message && <p className="inline-success">{message}</p>}
    </section>
  </div>;
}
