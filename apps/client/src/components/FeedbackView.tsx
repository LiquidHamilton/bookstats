import { useMemo, useState } from "react";
import type { UserAccount } from "@bookstats/domain";
import { Bug, CheckCircle2, Lightbulb, Send, ShieldCheck } from "lucide-react";
import { submitFeedback, type FeedbackDiagnostics } from "../data/api";

interface Props {
  account: UserAccount | null;
  storageKind?: "indexeddb" | "sqlite";
  bookCount: number;
  shelfCount: number;
}

export function FeedbackView({ account, storageKind, bookCount, shelfCount }: Props) {
  const [kind, setKind] = useState<"bug" | "feature">("bug");
  const [message, setMessage] = useState("");
  const [contactEmail, setContactEmail] = useState(account?.email ?? "");
  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string>();
  const [sent, setSent] = useState(false);

  const diagnostics = useMemo<FeedbackDiagnostics>(() => ({
    version: __BOOKSTATS_VERSION__,
    platform: detectPlatform(),
    storage: storageKind === "sqlite" ? "SQLite" : storageKind === "indexeddb" ? "IndexedDB" : "Opening…",
    signedIn: Boolean(account),
    emailVerified: Boolean(account?.emailVerified),
    bookCount,
    shelfCount
  }), [account, bookCount, shelfCount, storageKind]);

  async function send() {
    if (!message.trim()) return;
    setSending(true); setError(undefined); setSent(false);
    try {
      await submitFeedback(kind, message.trim(), contactEmail.trim() || undefined, diagnostics);
      setSent(true); setMessage("");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not send feedback.");
    } finally { setSending(false); }
  }

  return <>
    <header className="page-header"><div><p className="eyebrow">Help & feedback</p><h1>Tell me what you think</h1><p>Found something broken or have an idea that would make BookStats better? Send it directly from the app.</p></div></header>
    <section className="feedback-layout">
      <article className="feedback-card">
        <div className="feedback-kind-toggle">
          <button className={kind === "bug" ? "active" : ""} onClick={() => { setKind("bug"); setSent(false); }}><Bug size={18} /><span><strong>Report a Bug</strong><small>Something isn't working correctly</small></span></button>
          <button className={kind === "feature" ? "active" : ""} onClick={() => { setKind("feature"); setSent(false); }}><Lightbulb size={18} /><span><strong>Suggest a Feature</strong><small>An idea for BookStats</small></span></button>
        </div>
        <label className="feedback-message">{kind === "bug" ? "What happened?" : "What would you like to see?"}<textarea rows={10} maxLength={6000} value={message} onChange={(event) => setMessage(event.target.value)} placeholder={kind === "bug" ? "Tell me what you were doing, what you expected to happen, and what happened instead…" : "Describe the feature or improvement and how you'd use it…"} /><small>{message.length.toLocaleString()} / 6,000</small></label>
        <label className="feedback-contact">Contact email <input type="email" value={contactEmail} onChange={(event) => setContactEmail(event.target.value)} placeholder="Optional" /><small>Optional. Include this if you'd like a reply.</small></label>
        {error && <p className="inline-error">{error}</p>}
        {sent && <p className="inline-success"><CheckCircle2 size={15} />Thanks — your {kind === "bug" ? "bug report" : "feature suggestion"} was sent.</p>}
        <button className="button primary" disabled={sending || !message.trim()} onClick={() => void send()}><Send size={16} />{sending ? "Sending…" : kind === "bug" ? "Send bug report" : "Send suggestion"}</button>
      </article>

      <aside className="diagnostic-card"><div className="diagnostic-heading"><ShieldCheck size={19} /><div><strong>Included diagnostics</strong><span>Only these non-sensitive details are attached automatically.</span></div></div><dl>{Object.entries(diagnostics).map(([key, value]) => <div key={key}><dt>{diagnosticLabel(key)}</dt><dd>{typeof value === "boolean" ? (value ? "Yes" : "No") : String(value)}</dd></div>)}</dl><p>No book titles, reviews, notes, passwords, tokens, or library contents are included.</p></aside>
    </section>
  </>;
}

function detectPlatform(): string {
  const tauri = typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
  const ua = typeof navigator !== "undefined" ? navigator.userAgent : "";
  if (!tauri) return "Web";
  if (/Windows/i.test(ua)) return "Windows desktop";
  if (/Macintosh|Mac OS X/i.test(ua)) return "macOS desktop";
  return "Desktop";
}

function diagnosticLabel(key: string): string {
  return ({ version: "BookStats version", platform: "Platform", storage: "Local storage", signedIn: "Signed in", emailVerified: "Email verified", bookCount: "Book count", shelfCount: "Shelf count" } as Record<string, string>)[key] ?? key;
}
