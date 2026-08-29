import { useEffect, useState, type FormEvent, type ReactNode } from "react";
import { Activity, BookOpen, Database, Gauge, KeyRound, LogOut, RefreshCw, Search, Shield, ShieldCheck, Trash2, UserCog, Users } from "lucide-react";
import {
  ApiError, type AdminAccount, type AdminRecord, type AdminUser, type AuditEntry, type DashboardData,
  clearCloudLibrary, currentAdmin, deleteRecord, deleteUser, forcePasswordReset, invalidateSessions, loadAudit,
  loadDashboard, loadRecords, loadUser, loadUsers, loginAdmin, logoutAdmin, saveRecord, setUserDisabled, updateUser
} from "./api";

type View = "dashboard" | "users" | "audit";

export function App() {
  const [admin, setAdmin] = useState<AdminAccount | null | undefined>(undefined);
  const [view, setView] = useState<View>("dashboard");
  const [selectedUserId, setSelectedUserId] = useState<string>();

  useEffect(() => { void currentAdmin().then(setAdmin); }, []);
  if (admin === undefined) return <div className="boot">Loading administration…</div>;
  if (!admin) return <Login onLogin={setAdmin} />;

  const signOut = async () => { await logoutAdmin(); setAdmin(null); };
  const navigate = (next: View) => { setView(next); if (next !== "users") setSelectedUserId(undefined); };

  return <div className="admin-shell">
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark"><ShieldCheck size={22}/></div><div><strong>BookStats</strong><span>Administration</span></div></div>
      <nav>
        <NavButton active={view === "dashboard"} icon={<Gauge size={17}/>} onClick={() => navigate("dashboard")}>Dashboard</NavButton>
        <NavButton active={view === "users"} icon={<Users size={17}/>} onClick={() => navigate("users")}>Users</NavButton>
        <NavButton active={view === "audit"} icon={<Activity size={17}/>} onClick={() => navigate("audit")}>Audit log</NavButton>
      </nav>
      <div className="sidebar-footer"><div className="admin-chip"><Shield size={15}/><div><strong>{admin.displayName}</strong><span>{admin.email}</span></div></div><button className="ghost wide" onClick={() => void signOut()}><LogOut size={15}/>Sign out</button></div>
    </aside>
    <main className="main">
      {view === "dashboard" && <Dashboard onOpenUser={(id) => { setSelectedUserId(id); setView("users"); }}/>} 
      {view === "users" && (selectedUserId ? <UserDetail userId={selectedUserId} onBack={() => setSelectedUserId(undefined)} /> : <UsersView onOpen={setSelectedUserId} />)}
      {view === "audit" && <AuditView />}
    </main>
  </div>;
}

function Login({ onLogin }: { onLogin: (admin: AdminAccount) => void }) {
  const [email, setEmail] = useState(""); const [password, setPassword] = useState(""); const [busy, setBusy] = useState(false); const [error, setError] = useState("");
  const submit = async (event: FormEvent) => { event.preventDefault(); setBusy(true); setError(""); try { onLogin(await loginAdmin(email, password)); } catch (err) { setError(message(err)); } finally { setBusy(false); } };
  return <div className="login-page"><form className="login-card" onSubmit={submit}><div className="login-mark"><ShieldCheck size={28}/></div><div><h1>BookStats Administration</h1><p>Authorized administrators only.</p></div>{error && <div className="alert error">{error}</div>}<label>Email<input type="email" autoComplete="username" value={email} onChange={(e) => setEmail(e.target.value)} required /></label><label>Password<input type="password" autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} required /></label><button className="primary" disabled={busy}>{busy ? "Signing in…" : "Sign in"}</button><span className="version">Admin client v{__BOOKSTATS_ADMIN_VERSION__}</span></form></div>;
}

function Dashboard({ onOpenUser }: { onOpenUser: (id: string) => void }) {
  const [data, setData] = useState<DashboardData>(); const [error, setError] = useState(""); const [loading, setLoading] = useState(true);
  const refresh = async () => { setLoading(true); setError(""); try { setData(await loadDashboard()); } catch (err) { setError(message(err)); } finally { setLoading(false); } };
  useEffect(() => { void refresh(); }, []);
  if (loading && !data) return <Loading/>;
  if (!data) return <ErrorBox text={error} onRetry={refresh}/>;
  const m = data.metrics;
  return <Page title="Dashboard" subtitle="BookStats account, library, and server overview" action={<button className="ghost" onClick={() => void refresh()}><RefreshCw size={15}/>Refresh</button>}>
    {error && <div className="alert error">{error}</div>}
    <div className="metrics">
      <Metric icon={<Users/>} label="Users" value={m.totalUsers} detail={`${m.active7d} active this week`} />
      <Metric icon={<BookOpen/>} label="Books" value={m.books.toLocaleString()} detail={`${m.shelves} shelves · ${m.goals} goals`} />
      <Metric icon={<Activity/>} label="Active today" value={m.active24h} detail={`${m.newUsers30d} new users / 30d`} />
      <Metric icon={<Shield/>} label="Disabled" value={m.disabledUsers} detail={`${m.admins} administrator${m.admins === 1 ? "" : "s"}`} />
    </div>
    <div className="two-col">
      <section className="panel"><PanelHead icon={<Database size={17}/>} title="Server health"/><dl className="facts"><Fact label="API version" value={data.server.version}/><Fact label="Schema version" value={data.server.schemaVersion}/><Fact label="Uptime" value={duration(data.server.uptimeSeconds)}/><Fact label="Database latency" value={`${data.server.databaseLatencyMs} ms`}/><Fact label="Database size" value={bytes(m.databaseBytes)}/><Fact label="Active sessions" value={m.activeSessions}/><Fact label="Metadata cache" value={m.metadataCacheEntries}/><Fact label="Email" value={data.server.emailConfigured ? "Configured" : "Not configured"}/></dl><div className="provider-row">{data.server.metadataProviders.map((p) => <span className={p.configured ? "pill good" : "pill"} key={p.id}>{p.label}</span>)}</div></section>
      <section className="panel"><PanelHead icon={<UserCog size={17}/>} title="Recent accounts"/><div className="rows">{data.recentUsers.map((u) => <button className="row-button" key={u.id} onClick={() => onOpenUser(u.id)}><div><strong>{u.displayName}</strong><span>{u.email}</span></div><time>{dateTime(u.createdAt)}</time></button>)}</div></section>
    </div>
  </Page>;
}

function UsersView({ onOpen }: { onOpen: (id: string) => void }) {
  const [q, setQ] = useState(""); const [result, setResult] = useState<{ users: AdminUser[]; total: number }>({ users: [], total: 0 }); const [loading, setLoading] = useState(true); const [error, setError] = useState("");
  const refresh = async (query = q) => { setLoading(true); setError(""); try { setResult(await loadUsers(query)); } catch (err) { setError(message(err)); } finally { setLoading(false); } };
  useEffect(() => { const timer = window.setTimeout(() => void refresh(q), 250); return () => window.clearTimeout(timer); }, [q]);
  return <Page title="Users" subtitle={`${result.total} account${result.total === 1 ? "" : "s"}`}>
    <div className="toolbar"><div className="search"><Search size={16}/><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search name or email" /></div><button className="ghost" onClick={() => void refresh()}><RefreshCw size={15}/>Refresh</button></div>
    {error && <div className="alert error">{error}</div>}
    <div className="table-panel"><table><thead><tr><th>User</th><th>Status</th><th>Books</th><th>Last active</th><th>Created</th></tr></thead><tbody>{result.users.map((u) => <tr key={u.id} onClick={() => onOpen(u.id)} className="click-row"><td><strong>{u.displayName}</strong><span>{u.email}</span></td><td><div className="status-cell"><span className={`pill ${u.disabled ? "danger" : "good"}`}>{u.disabled ? "Disabled" : "Active"}</span>{u.role === "admin" && <span className="pill admin">Admin</span>}</div></td><td>{u.bookCount.toLocaleString()}</td><td>{u.lastActiveAt ? dateTime(u.lastActiveAt) : "Never"}</td><td>{dateOnly(u.createdAt)}</td></tr>)}</tbody></table>{loading && <div className="table-loading">Loading…</div>}{!loading && result.users.length === 0 && <div className="empty">No matching accounts.</div>}</div>
  </Page>;
}

function UserDetail({ userId, onBack }: { userId: string; onBack: () => void }) {
  const [user, setUser] = useState<AdminUser>(); const [tab, setTab] = useState<"overview"|"records">("overview"); const [error, setError] = useState(""); const [notice, setNotice] = useState(""); const [confirm, setConfirm] = useState<{ title: string; body: ReactNode; phrase?: string; destructive?: boolean; action: (value: string) => Promise<void> }>();
  const refresh = async () => { setError(""); try { setUser((await loadUser(userId)).user); } catch (err) { setError(message(err)); } };
  useEffect(() => { void refresh(); }, [userId]);
  const run = async (fn: () => Promise<unknown>, success: string) => { setError(""); setNotice(""); try { await fn(); setNotice(success); await refresh(); } catch (err) { setError(message(err)); } };
  if (!user) return error ? <ErrorBox text={error} onRetry={refresh}/> : <Loading/>;
  return <Page title={user.displayName} subtitle={user.email} action={<button className="ghost" onClick={onBack}>← Back to users</button>}>
    {error && <div className="alert error">{error}</div>}{notice && <div className="alert success">{notice}</div>}
    <div className="user-hero"><div><span className={`pill ${user.disabled ? "danger" : "good"}`}>{user.disabled ? "Disabled" : "Active"}</span>{user.role === "admin" && <span className="pill admin">Administrator</span>}<span className={`pill ${user.emailVerified ? "good" : ""}`}>{user.emailVerified ? "Email verified" : "Unverified"}</span></div><div className="hero-actions"><button className="ghost" onClick={() => void run(() => setUserDisabled(user.id, !user.disabled), user.disabled ? "Account enabled." : "Account disabled and sessions revoked.")}>{user.disabled ? "Enable account" : "Disable account"}</button><button className="ghost" onClick={() => void run(() => invalidateSessions(user.id), "Sessions invalidated.")}><KeyRound size={15}/>Invalidate sessions</button><button className="ghost" onClick={() => void run(() => forcePasswordReset(user.id), "Password reset email sent and sessions invalidated.")}>Force password reset</button></div></div>
    <div className="tabs"><button className={tab === "overview" ? "active" : ""} onClick={() => setTab("overview")}>Overview</button><button className={tab === "records" ? "active" : ""} onClick={() => setTab("records")}>Library records</button></div>
    {tab === "overview" && <UserOverview user={user} onSaved={async () => { setNotice("Account updated."); await refresh(); }} onError={setError} onClear={() => setConfirm({ title: "Clear cloud library", phrase: `CLEAR ${user.email}`, destructive: true, body: <>This permanently removes every synchronized book, shelf, and goal for <strong>{user.email}</strong>. Local copies on a device are not erased.</>, action: async (value) => { const result = await clearCloudLibrary(user.id, value); setNotice(`${result.deleted} cloud records removed.`); await refresh(); } })} onDelete={() => setConfirm({ title: "Delete account", phrase: `DELETE ${user.email}`, destructive: true, body: <>This permanently deletes <strong>{user.email}</strong>, all server library data, sessions, and account tokens.</>, action: async (value) => { await deleteUser(user.id, value); onBack(); } })}/>} 
    {tab === "records" && <Records user={user} />}
    {confirm && <ConfirmModal {...confirm} onClose={() => setConfirm(undefined)} onError={setError}/>} 
  </Page>;
}

function UserOverview({ user, onSaved, onError, onClear, onDelete }: { user: AdminUser; onSaved: () => Promise<void>; onError: (s:string)=>void; onClear:()=>void; onDelete:()=>void }) {
  const [name, setName] = useState(user.displayName); const [email, setEmail] = useState(user.email); const [verified, setVerified] = useState(user.emailVerified); const [saving, setSaving] = useState(false);
  useEffect(() => { setName(user.displayName); setEmail(user.email); setVerified(user.emailVerified); }, [user]);
  const save = async (e: FormEvent) => { e.preventDefault(); setSaving(true); onError(""); try { await updateUser(user.id, { displayName: name, email, emailVerified: verified }); await onSaved(); } catch (err) { onError(message(err)); } finally { setSaving(false); } };
  return <div className="two-col user-columns"><section className="panel"><PanelHead icon={<UserCog size={17}/>} title="Account"/><form className="edit-form" onSubmit={save}><label>Display name<input value={name} onChange={(e) => setName(e.target.value)} /></label><label>Email<input type="email" value={email} onChange={(e) => { setEmail(e.target.value); if (e.target.value.trim().toLowerCase() !== user.email.toLowerCase()) setVerified(false); }} /></label><label className="check"><input type="checkbox" checked={verified} onChange={(e) => setVerified(e.target.checked)} />Email verified</label><button className="primary compact" disabled={saving}>{saving ? "Saving…" : "Save account"}</button></form><dl className="facts compact-facts"><Fact label="Created" value={dateTime(user.createdAt)}/><Fact label="Updated" value={user.updatedAt ? dateTime(user.updatedAt) : "—"}/><Fact label="Last active" value={user.lastActiveAt ? dateTime(user.lastActiveAt) : "Never"}/><Fact label="Active sessions" value={user.activeSessionCount ?? 0}/></dl></section><section className="panel"><PanelHead icon={<BookOpen size={17}/>} title="Cloud library"/><div className="metrics mini"><Metric label="Books" value={user.bookCount}/><Metric label="Owned" value={user.ownedBookCount ?? 0}/><Metric label="Readings" value={user.totalReadings ?? 0}/><Metric label="Shelves" value={user.shelfCount}/></div>{user.statusCounts && <div className="status-summary">{Object.entries(user.statusCounts).sort((a,b)=>b[1]-a[1]).map(([status,count])=><span className="pill" key={status}>{prettyStatus(status)} · {count}</span>)}</div>}<dl className="facts compact-facts"><Fact label="Goals" value={user.goalCount}/><Fact label="Deleted/tombstones" value={user.deletedRecordCount ?? 0}/></dl><div className="danger-zone"><h3>Destructive actions</h3><p>These operations affect server-side data and are written to the administrator audit log.</p><button className="danger-button" onClick={onClear}><Trash2 size={15}/>Clear cloud library</button><button className="danger-button strongest" onClick={onDelete}><Trash2 size={15}/>Permanently delete account</button></div></section></div>;
}

function Records({ user }: { user: AdminUser }) {
  const [type, setType] = useState(""); const [q, setQ] = useState(""); const [includeDeleted, setIncludeDeleted] = useState(false); const [records, setRecords] = useState<AdminRecord[]>([]); const [total, setTotal] = useState(0); const [error, setError] = useState(""); const [editing, setEditing] = useState<AdminRecord>();
  const refresh = async () => { setError(""); try { const result = await loadRecords(user.id, type, q, includeDeleted); setRecords(result.records); setTotal(result.total); } catch (err) { setError(message(err)); } };
  useEffect(() => { const t = window.setTimeout(() => void refresh(), 180); return () => window.clearTimeout(t); }, [type, q, includeDeleted, user.id]);
  return <section className="panel"><div className="record-toolbar"><div className="search"><Search size={16}/><input placeholder="Search title, author, or shelf" value={q} onChange={(e) => setQ(e.target.value)}/></div><select value={type} onChange={(e) => setType(e.target.value)}><option value="">All records</option><option value="book">Books</option><option value="shelf">Shelves</option><option value="goal">Goals</option></select><label className="check inline"><input type="checkbox" checked={includeDeleted} onChange={(e) => setIncludeDeleted(e.target.checked)}/>Include deleted</label><span className="muted">{total} records</span></div>{error && <div className="alert error">{error}</div>}<div className="record-list">{records.map((r) => <button className="record-card" key={r.id} onClick={() => setEditing(r)}><span className={`type-badge ${r.recordType}`}>{r.recordType}</span><div><strong>{recordLabel(r)}</strong><span>{r.id}</span></div><div className="record-meta"><span>rev {r.revision}</span>{r.deleted && <span className="danger-text">deleted</span>}<time>{dateTime(r.serverUpdatedAt)}</time></div></button>)}</div>{records.length === 0 && <div className="empty">No matching records.</div>}{editing && <RecordEditor userId={user.id} record={editing} onClose={() => setEditing(undefined)} onSaved={() => { setEditing(undefined); void refresh(); }} />}</section>;
}

function RecordEditor({ userId, record, onClose, onSaved }: { userId:string; record:AdminRecord; onClose:()=>void; onSaved:()=>void }) {
  const [json, setJson] = useState(JSON.stringify(record.data, null, 2)); const [error, setError] = useState(""); const [busy, setBusy] = useState(false);
  const save = async () => { setError(""); setBusy(true); try { const parsed = JSON.parse(json) as Record<string, unknown>; await saveRecord(userId, record, parsed); onSaved(); } catch (err) { setError(message(err)); } finally { setBusy(false); } };
  const remove = async () => { if (!window.confirm("Mark this synchronized record deleted? The operation is audited.")) return; setBusy(true); try { await deleteRecord(userId, record.id); onSaved(); } catch (err) { setError(message(err)); } finally { setBusy(false); } };
  return <div className="modal-backdrop" onMouseDown={(e) => e.currentTarget === e.target && onClose()}><div className="modal record-editor"><div className="modal-head"><div><span className={`type-badge ${record.recordType}`}>{record.recordType}</span><h2>Edit synchronized record</h2><p>{record.id}</p></div><button className="ghost" onClick={onClose}>Close</button></div>{error && <div className="alert error">{error}</div>}<div className="editor-warning">Direct record editing is a support tool. The JSON id must remain unchanged. Saving advances the sync revision so clients receive the change.</div><textarea spellCheck={false} value={json} onChange={(e) => setJson(e.target.value)}/><div className="modal-actions"><button className="danger-button" disabled={busy} onClick={() => void remove()}><Trash2 size={15}/>Delete record</button><div/><button className="ghost" onClick={onClose}>Cancel</button><button className="primary" disabled={busy} onClick={() => void save()}>{busy ? "Saving…" : "Save record"}</button></div></div></div>;
}

function AuditView() {
  const [q, setQ] = useState(""); const [entries, setEntries] = useState<AuditEntry[]>([]); const [error, setError] = useState("");
  const refresh = async () => { try { setEntries((await loadAudit(q)).entries); setError(""); } catch (err) { setError(message(err)); } };
  useEffect(() => { const t = window.setTimeout(() => void refresh(), 180); return () => window.clearTimeout(t); }, [q]);
  return <Page title="Audit log" subtitle="Administrative actions are append-only"><div className="toolbar"><div className="search"><Search size={16}/><input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search admin, action, or target" /></div><button className="ghost" onClick={() => void refresh()}><RefreshCw size={15}/>Refresh</button></div>{error && <div className="alert error">{error}</div>}<div className="audit-list">{entries.map((e) => <article className="audit-item" key={e.id}><div className="audit-dot"><Shield size={14}/></div><div><div className="audit-title"><strong>{humanAction(e.action)}</strong><span>{dateTime(e.createdAt)}</span></div><p><b>{e.adminEmail}</b>{e.targetEmail ? <> acted on <b>{e.targetEmail}</b></> : ""}{e.targetRecordId ? <> · record <code>{e.targetRecordId}</code></> : ""}</p>{Object.keys(e.details ?? {}).length > 0 && <pre>{JSON.stringify(e.details)}</pre>}<small>{e.ipAddress || "IP unavailable"}</small></div></article>)}</div>{entries.length === 0 && <div className="empty">No audit entries yet.</div>}</Page>;
}

function ConfirmModal({ title, body, phrase, destructive, action, onClose, onError }: { title:string; body:ReactNode; phrase?:string; destructive?:boolean; action:(v:string)=>Promise<void>; onClose:()=>void; onError:(s:string)=>void }) {
  const [value, setValue] = useState(""); const [busy, setBusy] = useState(false); const ready = !phrase || value === phrase;
  const run = async () => { setBusy(true); try { await action(value); onClose(); } catch (err) { onError(message(err)); } finally { setBusy(false); } };
  return <div className="modal-backdrop"><div className="modal confirm-modal"><h2>{title}</h2><p>{body}</p>{phrase && <label>Type <code>{phrase}</code> to continue<input value={value} onChange={(e) => setValue(e.target.value)} autoFocus /></label>}<div className="modal-actions"><div/><div/><button className="ghost" onClick={onClose}>Cancel</button><button className={destructive ? "danger-button strongest" : "primary"} disabled={!ready || busy} onClick={() => void run()}>{busy ? "Working…" : "Confirm"}</button></div></div></div>;
}

function Page({ title, subtitle, action, children }: { title:string; subtitle?:string; action?:ReactNode; children:ReactNode }) { return <><header className="page-head"><div><h1>{title}</h1>{subtitle && <p>{subtitle}</p>}</div>{action}</header><div className="page-body">{children}</div></>; }
function NavButton({ active, icon, onClick, children }: { active:boolean; icon:ReactNode; onClick:()=>void; children:ReactNode }) { return <button className={active ? "nav active" : "nav"} onClick={onClick}>{icon}{children}</button>; }
function Metric({ icon, label, value, detail }: { icon?:ReactNode; label:string; value:string|number; detail?:string }) { return <div className="metric">{icon && <div className="metric-icon">{icon}</div>}<div><span>{label}</span><strong>{value}</strong>{detail && <small>{detail}</small>}</div></div>; }
function PanelHead({ icon, title }: { icon:ReactNode; title:string }) { return <div className="panel-head">{icon}<h2>{title}</h2></div>; }
function Fact({ label, value }: { label:string; value:string|number }) { return <div><dt>{label}</dt><dd>{value}</dd></div>; }
function Loading() { return <div className="loading"><RefreshCw className="spin"/>Loading…</div>; }
function ErrorBox({ text, onRetry }: { text:string; onRetry:()=>void }) { return <div className="center-error"><div className="alert error">{text || "Could not load administrator data."}</div><button className="ghost" onClick={onRetry}>Try again</button></div>; }
function message(error: unknown): string { return error instanceof ApiError || error instanceof Error ? error.message : "An unexpected error occurred."; }
function dateTime(value:string) { return new Intl.DateTimeFormat(undefined,{dateStyle:"medium",timeStyle:"short"}).format(new Date(value)); }
function dateOnly(value:string) { return new Intl.DateTimeFormat(undefined,{dateStyle:"medium"}).format(new Date(value)); }
function duration(seconds:number) { const d=Math.floor(seconds/86400), h=Math.floor((seconds%86400)/3600), m=Math.floor((seconds%3600)/60); return [d&&`${d}d`,h&&`${h}h`,`${m}m`].filter(Boolean).join(" "); }
function bytes(value:number) { if (!Number.isFinite(value)) return "—"; const units=["B","KB","MB","GB","TB"]; let n=value,i=0; while(n>=1024&&i<units.length-1){n/=1024;i++;} return `${n.toFixed(i ? 1 : 0)} ${units[i]}`; }
function recordLabel(record:AdminRecord) { const d=record.data; if (!d) return "Deleted record"; if (record.recordType === "book") return `${String(d.title ?? "Untitled")} · ${String(d.author ?? "Unknown author")}`; if (record.recordType === "shelf") return String(d.name ?? "Unnamed shelf"); if (record.recordType === "goal") return `${String(d.year ?? "Goal")} reading goal`; return record.id; }
function humanAction(action:string) { return action.toLowerCase().split("_").map((s)=>s.charAt(0).toUpperCase()+s.slice(1)).join(" "); }
function prettyStatus(status:string) { return status.replace(/_/g," ").replace(/\b\w/g,(c)=>c.toUpperCase()); }
