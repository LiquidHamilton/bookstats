import type { BookCondition, BookFormat, ReadingStatus, ShelfMatchMode, ShelfRule, ShelfRuleField, ShelfRuleGroup, ShelfRuleOperator } from "@bookstats/domain";
import { BOOK_CONDITIONS, READING_STATUS_LABELS } from "@bookstats/domain";
import { Layers3, Plus, Trash2 } from "lucide-react";

export type RuleDescriptor = { label: string; kind: "text" | "number" | "status" | "format" | "condition" | "boolean" | "date" };
export const ruleFields: Record<ShelfRuleField, RuleDescriptor> = {
  status: { label: "Status", kind: "status" },
  condition: { label: "Condition", kind: "condition" },
  owned: { label: "Ownership", kind: "boolean" },
  rating: { label: "Rating", kind: "number" },
  title: { label: "Title", kind: "text" },
  author: { label: "Author", kind: "text" },
  series: { label: "Series", kind: "text" },
  format: { label: "Format", kind: "format" },
  genre: { label: "Genre", kind: "text" },
  tag: { label: "Tag", kind: "text" },
  pages: { label: "Pages", kind: "number" },
  publicationYear: { label: "Publication year", kind: "number" },
  readCount: { label: "Read count", kind: "number" },
  lastRead: { label: "Last read date", kind: "date" },
  dateAdded: { label: "Date added", kind: "date" }
};

const formats: BookFormat[] = ["Hardcover", "Paperback", "Mass Market Paperback", "eBook", "Audiobook", "Graphic Novel", "Omnibus", "Other"];
const statuses = Object.keys(READING_STATUS_LABELS) as ReadingStatus[];
const conditions = BOOK_CONDITIONS as readonly BookCondition[];

export function createRule(field: ShelfRuleField = "status"): ShelfRule {
  return { id: crypto.randomUUID(), field, operator: field === "owned" ? "is_true" : "equals", value: field === "status" ? "read" : field === "format" ? "Paperback" : field === "condition" ? "Good" : "" };
}

export function createRuleGroup(match: ShelfMatchMode = "all"): ShelfRuleGroup {
  return { id: crypto.randomUUID(), match, rules: [createRule()] };
}

export function rulesAreValid(rules: ShelfRule[]): boolean {
  return rules.length > 0 && rules.every((rule) => ruleFields[rule.field].kind === "boolean" || Boolean(rule.value?.trim()));
}

export function ruleGroupsAreValid(groups: ShelfRuleGroup[]): boolean {
  return groups.length > 0 && groups.every((group) => rulesAreValid(group.rules));
}

export function updateRuleValue(rule: ShelfRule, patch: Partial<ShelfRule>): ShelfRule {
  const next = { ...rule, ...patch };
  if (patch.field) {
    next.operator = patch.field === "owned" ? "is_true" : "equals";
    next.value = patch.field === "status" ? "read" : patch.field === "format" ? "Paperback" : patch.field === "condition" ? "Good" : "";
  }
  return next;
}

export function RuleBuilder({
  match,
  groups,
  onMatchChange,
  onGroupsChange,
  compact = false
}: {
  /** How separate rule groups are combined: all = AND, any = OR. */
  match: ShelfMatchMode;
  groups: ShelfRuleGroup[];
  onMatchChange: (match: ShelfMatchMode) => void;
  onGroupsChange: (groups: ShelfRuleGroup[]) => void;
  compact?: boolean;
}) {
  const valid = ruleGroupsAreValid(groups);

  function updateGroup(id: string, patch: Partial<ShelfRuleGroup>) {
    onGroupsChange(groups.map((group) => group.id === id ? { ...group, ...patch } : group));
  }
  function updateRule(groupId: string, ruleId: string, patch: Partial<ShelfRule>) {
    onGroupsChange(groups.map((group) => group.id === groupId ? { ...group, rules: group.rules.map((rule) => rule.id === ruleId ? updateRuleValue(rule, patch) : rule) } : group));
  }
  function deleteRule(groupId: string, ruleId: string) {
    onGroupsChange(groups.map((group) => group.id === groupId ? { ...group, rules: group.rules.filter((rule) => rule.id !== ruleId) } : group));
  }
  function addRule(groupId: string) {
    onGroupsChange(groups.map((group) => group.id === groupId ? { ...group, rules: [...group.rules, createRule("status")] } : group));
  }
  function deleteGroup(groupId: string) {
    if (groups.length <= 1) return;
    onGroupsChange(groups.filter((group) => group.id !== groupId));
  }

  return <div className={`smart-shelf-builder ${compact ? "rule-builder-compact" : ""}`}>
    <div className="smart-shelf-heading">
      <div><strong>Rule groups</strong><span>Combine rules inside parentheses, then combine the groups with AND or OR.</span></div>
      <label>Between groups <select value={match} onChange={(event) => onMatchChange(event.target.value as ShelfMatchMode)}><option value="any">OR</option><option value="all">AND</option></select></label>
    </div>
    {!valid && <p className="inline-error">Complete each rule before applying this filter.</p>}

    <div className="smart-rule-groups">
      {groups.map((group, groupIndex) => <div key={group.id}>
        {groupIndex > 0 && <div className="smart-group-joiner"><span>{match === "all" ? "AND" : "OR"}</span></div>}
        <section className="smart-rule-group">
          <div className="smart-rule-group-heading">
            <div><span className="smart-rule-group-number">{groupIndex + 1}</span><strong>Group {groupIndex + 1}</strong></div>
            <div><label>Inside group <select value={group.match} onChange={(event) => updateGroup(group.id, { match: event.target.value as ShelfMatchMode })}><option value="all">AND</option><option value="any">OR</option></select></label><button className="icon-button danger-icon" disabled={groups.length <= 1} onClick={() => deleteGroup(group.id)} aria-label={`Remove group ${groupIndex + 1}`} title="Remove this rule group"><Trash2 size={14} /></button></div>
          </div>
          <div className="smart-rule-list">{group.rules.map((rule, ruleIndex) => <div key={rule.id}>
            {ruleIndex > 0 && <div className="smart-rule-inline-joiner">{group.match === "all" ? "AND" : "OR"}</div>}
            <SmartRuleRow rule={rule} onChange={(patch) => updateRule(group.id, rule.id, patch)} onDelete={() => deleteRule(group.id, rule.id)} />
          </div>)}</div>
          <div className="smart-rule-group-footer"><button className="button secondary compact" onClick={() => addRule(group.id)}><Plus size={14} />Add rule to group</button></div>
        </section>
      </div>)}
    </div>

    <div className="smart-shelf-footer"><button className="button secondary compact" onClick={() => onGroupsChange([...groups, createRuleGroup("all")])}><Layers3 size={14} />Add rule group</button></div>
  </div>;
}

function operatorsFor(field: ShelfRuleField): Array<{ value: ShelfRuleOperator; label: string }> {
  const kind = ruleFields[field].kind;
  if (kind === "boolean") return [{ value: "is_true", label: "is owned" }, { value: "is_false", label: "is not owned" }];
  if (kind === "number") return [{ value: "equals", label: "equals" }, { value: "not_equals", label: "does not equal" }, { value: "gte", label: "is at least" }, { value: "lte", label: "is at most" }];
  if (kind === "date") return [{ value: "equals", label: "is" }, { value: "not_equals", label: "is not" }, { value: "gte", label: "is on or after" }, { value: "lte", label: "is on or before" }];
  if (kind === "status" || kind === "format" || kind === "condition") return [{ value: "equals", label: "is" }, { value: "not_equals", label: "is not" }];
  return [{ value: "contains", label: "contains" }, { value: "not_contains", label: "does not contain" }, { value: "equals", label: "is exactly" }, { value: "not_equals", label: "is not" }];
}

function SmartRuleRow({ rule, onChange, onDelete }: { rule: ShelfRule; onChange: (patch: Partial<ShelfRule>) => void; onDelete: () => void }) {
  const descriptor = ruleFields[rule.field];
  const needsValue = descriptor.kind !== "boolean";
  return <div className="smart-rule-row">
    <select value={rule.field} onChange={(event) => onChange({ field: event.target.value as ShelfRuleField })}>{(Object.entries(ruleFields) as Array<[ShelfRuleField, RuleDescriptor]>).map(([field, item]) => <option value={field} key={field}>{item.label}</option>)}</select>
    <select value={rule.operator} onChange={(event) => onChange({ operator: event.target.value as ShelfRuleOperator })}>{operatorsFor(rule.field).map((operator) => <option key={operator.value} value={operator.value}>{operator.label}</option>)}</select>
    {needsValue ? descriptor.kind === "status" ? <select value={rule.value ?? "read"} onChange={(event) => onChange({ value: event.target.value })}>{statuses.map((status) => <option key={status} value={status}>{READING_STATUS_LABELS[status]}</option>)}</select> : descriptor.kind === "format" ? <select value={rule.value ?? "Paperback"} onChange={(event) => onChange({ value: event.target.value })}>{formats.map((format) => <option key={format} value={format}>{format}</option>)}</select> : descriptor.kind === "condition" ? <select value={rule.value ?? "Good"} onChange={(event) => onChange({ value: event.target.value })}>{conditions.map((condition) => <option key={condition} value={condition}>{condition}</option>)}</select> : <input type={descriptor.kind === "number" ? "number" : descriptor.kind === "date" ? "date" : "text"} step={descriptor.kind === "number" ? (rule.field === "rating" ? "0.5" : "1") : undefined} value={rule.value ?? ""} onChange={(event) => onChange({ value: event.target.value })} placeholder={descriptor.kind === "number" ? "Value" : descriptor.kind === "date" ? undefined : "Text"} /> : <span className="smart-rule-no-value">Automatically evaluated</span>}
    <button className="icon-button danger-icon" onClick={onDelete} aria-label="Remove rule"><Trash2 size={15} /></button>
  </div>;
}
