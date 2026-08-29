import { useMemo, useState } from "react";
import type { Book, Shelf, ShelfMatchMode, ShelfRuleGroup } from "@bookstats/domain";
import { isSmartShelf, shelfMatchesBook } from "@bookstats/domain";
import { ArrowDown, ArrowUp, FolderOpen, Pencil, Plus, Sparkles, Trash2, X } from "lucide-react";
import { createRuleGroup, RuleBuilder, ruleGroupsAreValid } from "./RuleBuilder";

interface Props {
  books: Book[];
  shelves: Shelf[];
  onCreate: (name: string, options?: Partial<Pick<Shelf, "kind" | "match" | "rules" | "ruleGroups">>) => Promise<Shelf>;
  onUpdate: (shelf: Shelf) => Promise<void>;
  onReorder: (shelves: Shelf[]) => Promise<void>;
  onDelete: (shelf: Shelf) => Promise<void>;
  onClose: () => void;
}

function cloneGroups(groups: ShelfRuleGroup[]): ShelfRuleGroup[] {
  return groups.map((group) => ({ ...group, rules: group.rules.map((rule) => ({ ...rule })) }));
}

function groupsFromShelf(shelf: Shelf): ShelfRuleGroup[] {
  if (shelf.ruleGroups?.length) return cloneGroups(shelf.ruleGroups);
  if (shelf.rules?.length) return [{ id: crypto.randomUUID(), match: shelf.match ?? "all", rules: shelf.rules.map((rule) => ({ ...rule })) }];
  return [createRuleGroup("all")];
}

export function ShelfManager({ books, shelves, onCreate, onUpdate, onReorder, onDelete, onClose }: Props) {
  const [name, setName] = useState("");
  const [kind, setKind] = useState<"manual" | "smart">("manual");
  const [match, setMatch] = useState<ShelfMatchMode>("any");
  const [groups, setGroups] = useState<ShelfRuleGroup[]>([createRuleGroup("all")]);
  const [working, setWorking] = useState(false);
  const [reordering, setReordering] = useState(false);
  const [editingShelf, setEditingShelf] = useState<Shelf>();
  const [error, setError] = useState<string>();

  const rulesValid = kind !== "smart" || ruleGroupsAreValid(groups);
  const previewCount = useMemo(() => {
    if (kind !== "smart" || !rulesValid) return 0;
    const preview: Shelf = { id: "preview", name: name || "Preview", kind: "smart", match, ruleGroups: groups, createdAt: "", updatedAt: "" };
    return books.filter((book) => shelfMatchesBook(preview, book)).length;
  }, [books, groups, kind, match, name, rulesValid]);

  async function saveShelf() {
    const trimmed = name.trim();
    if (!trimmed || !rulesValid) return;
    setWorking(true); setError(undefined);
    try {
      if (editingShelf) {
        await onUpdate({
          ...editingShelf,
          name: trimmed,
          kind,
          match: kind === "smart" ? match : undefined,
          rules: undefined,
          ruleGroups: kind === "smart" ? cloneGroups(groups) : undefined,
          updatedAt: new Date().toISOString()
        });
      } else {
        await onCreate(trimmed, kind === "smart" ? { kind: "smart", match, ruleGroups: cloneGroups(groups) } : { kind: "manual" });
      }
      resetBuilder();
    } catch (err) { setError(err instanceof Error ? err.message : "Could not save this shelf."); }
    finally { setWorking(false); }
  }

  function editShelf(shelf: Shelf) {
    setEditingShelf(shelf);
    setName(shelf.name);
    setKind(isSmartShelf(shelf) ? "smart" : "manual");
    // Legacy flat shelves keep their old all/any behavior inside the migrated first group.
    setMatch(shelf.ruleGroups?.length ? (shelf.match ?? "any") : "any");
    setGroups(groupsFromShelf(shelf));
  }

  function resetBuilder() {
    setEditingShelf(undefined);
    setName("");
    setKind("manual");
    setMatch("any");
    setGroups([createRuleGroup("all")]);
  }

  async function moveShelf(index: number, direction: -1 | 1) {
    const target = index + direction;
    if (target < 0 || target >= shelves.length || reordering) return;
    const next = shelves.slice();
    [next[index], next[target]] = [next[target], next[index]];
    setReordering(true); setError(undefined);
    try { await onReorder(next); }
    catch (err) { setError(err instanceof Error ? err.message : "Could not reorder shelves."); }
    finally { setReordering(false); }
  }

  return <div className="modal-backdrop" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
    <section className="shelf-manager shelf-manager-wide">
      <div className="form-header"><div><p className="eyebrow">Organize your library</p><h2>Add & manage shelves</h2></div><button className="icon-button" onClick={onClose} aria-label="Close"><X size={20} /></button></div>
      <p className="shelf-manager-intro">Regular shelves are collections you assign by hand. Smart shelves are saved rules that update automatically as your library changes. Smart shelves can combine parenthesized rule groups with AND/OR logic. Use the arrows below to choose the order shown throughout BookStats.</p>

      <div className="shelf-type-toggle" role="group" aria-label="Shelf type">
        <button disabled={Boolean(editingShelf)} className={kind === "manual" ? "active" : ""} onClick={() => setKind("manual")}><FolderOpen size={16} /><span><strong>Regular shelf</strong><small>Choose the books yourself</small></span></button>
        <button disabled={Boolean(editingShelf)} className={kind === "smart" ? "active" : ""} onClick={() => setKind("smart")}><Sparkles size={16} /><span><strong>Smart shelf</strong><small>Automatically matches Boolean rules</small></span></button>
      </div>

      <div className="new-shelf-row shelf-manager-add"><input autoFocus value={name} onChange={(event) => setName(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter" && kind === "manual") void saveShelf(); }} placeholder={kind === "smart" ? "Smart shelf name" : "Shelf name"} /><button className="button primary compact" disabled={!name.trim() || working || !rulesValid} onClick={() => void saveShelf()}>{editingShelf ? <Pencil size={15} /> : <Plus size={15} />}{editingShelf ? "Save shelf" : "Add shelf"}</button>{editingShelf && <button className="button secondary compact" onClick={resetBuilder}>Cancel edit</button>}</div>
      {error && <p className="inline-error">{error}</p>}

      {kind === "smart" && <div>
        <RuleBuilder match={match} groups={groups} onMatchChange={setMatch} onGroupsChange={setGroups} />
        <div className="smart-shelf-footer smart-shelf-preview"><span><Sparkles size={13} />Currently matches <strong>{previewCount.toLocaleString()}</strong> {previewCount === 1 ? "book" : "books"}</span></div>
      </div>}

      <div className="shelf-manager-list">
        {shelves.length === 0 ? <p className="read-history-empty">You haven't created any shelves yet.</p> : shelves.map((shelf, index) => {
          const count = books.filter((book) => shelfMatchesBook(shelf, book)).length;
          return <div className="shelf-manager-item" key={shelf.id}><div className="shelf-manager-name"><span className={`shelf-kind-icon ${isSmartShelf(shelf) ? "smart" : ""}`}>{isSmartShelf(shelf) ? <Sparkles size={14} /> : <FolderOpen size={14} />}</span><div><strong>{shelf.name}</strong><span>{isSmartShelf(shelf) ? `${count} ${count === 1 ? "book" : "books"} · smart shelf` : `${count} ${count === 1 ? "book" : "books"} · regular shelf`}</span></div></div><div className="shelf-manager-actions"><button className="icon-button" disabled={index === 0 || reordering} title={`Move ${shelf.name} up`} aria-label={`Move ${shelf.name} up`} onClick={() => void moveShelf(index, -1)}><ArrowUp size={15} /></button><button className="icon-button" disabled={index === shelves.length - 1 || reordering} title={`Move ${shelf.name} down`} aria-label={`Move ${shelf.name} down`} onClick={() => void moveShelf(index, 1)}><ArrowDown size={15} /></button><button className="icon-button" title={`Edit ${shelf.name}`} aria-label={`Edit ${shelf.name}`} onClick={() => editShelf(shelf)}><Pencil size={15} /></button><button className="icon-button danger-icon" title={`Delete ${shelf.name}`} aria-label={`Delete ${shelf.name}`} onClick={() => void onDelete(shelf)}><Trash2 size={16} /></button></div></div>;
        })}
      </div>
    </section>
  </div>;
}
