import { useEffect, useRef, useState, type FormEvent } from "react";
import type { UserAccount } from "@bookstats/domain";
import { Cloud, Database, KeyRound, LogIn, LogOut, RefreshCw, ShieldCheck, Trash2, UserPlus, X } from "lucide-react";
import {
  changeAccountPassword,
  currentAccount,
  getAuthToken,
  loginAccount,
  registerAccount,
  requestPasswordReset,
  resendVerificationEmail,
  resetPassword,
  verifyEmail
} from "../data/api";

interface Props {
  account: UserAccount | null;
  initialMode?: "login" | "register";
  storageKind?: "indexeddb" | "sqlite";
  syncing: boolean;
  lastSync?: string;
  syncError?: string;
  onAccountChange: (account: UserAccount | null) => void;
  onSync: () => Promise<void>;
  onLogout: () => Promise<void>;
  onDeleteCloudData: () => Promise<void>;
  onDeleteAccount: (password: string) => Promise<void>;
}

type AuthMode = "login" | "register" | "forgot" | "reset";

export function AccountView({ account, initialMode = "login", storageKind, syncing, lastSync, syncError, onAccountChange, onSync, onLogout, onDeleteCloudData, onDeleteAccount }: Props) {
  const [mode, setMode] = useState<AuthMode>(initialMode);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [resetToken, setResetToken] = useState<string>();
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const [message, setMessage] = useState<string>();
  const handledLink = useRef(false);

  useEffect(() => {
    if (handledLink.current) return;
    handledLink.current = true;
    const params = new URLSearchParams(window.location.search);
    const verificationToken = params.get("verify");
    const passwordToken = params.get("reset");
    if (passwordToken) { setResetToken(passwordToken); setMode("reset"); return; }
    if (!verificationToken) return;
    setBusy(true);
    void verifyEmail(verificationToken)
      .then(async () => {
        setMessage("Email verified. Cloud synchronization is now enabled for this account.");
        if (getAuthToken()) { const refreshed = await currentAccount(); if (refreshed) onAccountChange(refreshed); }
        clearAccountLinkQuery();
      })
      .catch((err) => setError(err instanceof Error ? err.message : "Could not verify this email address."))
      .finally(() => setBusy(false));
  }, [onAccountChange]);

  function changeMode(next: AuthMode) { setMode(next); setError(undefined); setMessage(undefined); setPassword(""); setConfirmPassword(""); }
  async function submit(event: FormEvent) {
    event.preventDefault(); setBusy(true); setError(undefined); setMessage(undefined);
    try {
      if (mode === "register") {
        if (password !== confirmPassword) throw new Error("Passwords do not match.");
        const result = await registerAccount(email, password, confirmPassword, displayName); onAccountChange(result.user);
        setMessage(result.emailVerificationSent ? "Account created. Check your email for the verification link before using cloud sync." : "Account created, but the verification email could not be sent. Check the server email configuration, then resend the verification message.");
      } else if (mode === "login") {
        const result = await loginAccount(email, password); onAccountChange(result.user);
      } else if (mode === "forgot") {
        setMessage((await requestPasswordReset(email)).message);
      } else {
        if (!resetToken) throw new Error("The password reset link is missing its token.");
        if (password !== confirmPassword) throw new Error("Passwords do not match.");
        await resetPassword(resetToken, password, confirmPassword); onAccountChange(null);
        setMessage("Password changed. All existing BookStats sessions were signed out; sign in with your new password."); setResetToken(undefined); clearAccountLinkQuery(); setMode("login"); setPassword(""); setConfirmPassword("");
      }
    } catch (err) { setError(err instanceof Error ? err.message : "Account request failed."); }
    finally { setBusy(false); }
  }
  async function resendVerification() { setBusy(true); setError(undefined); setMessage(undefined); try { const result = await resendVerificationEmail(); setMessage(result.alreadyVerified ? "This email address is already verified." : result.throttled ? "A verification email was sent recently. Check your inbox and spam folder." : "Verification email sent."); } catch (err) { setError(err instanceof Error ? err.message : "Could not send a verification email."); } finally { setBusy(false); } }
  async function logout() { setBusy(true); setError(undefined); try { await onLogout(); } catch (err) { setError(err instanceof Error ? err.message : "Could not sign out."); } finally { setBusy(false); } }

  return <>
    <header className="page-header"><div><p className="eyebrow">Cloud library</p><h1>Account</h1><p>Keep your library synchronized, manage account security, and stay in control of your data.</p></div></header>
    <section className="account-grid account-grid-compact">
      {account && mode !== "reset" ? <>
        <article className="account-card account-card-compact account-cloud-card">
          <div className="tool-icon"><Cloud size={20} /></div>
          <div className="account-details"><h2>{account.displayName}</h2><p>{account.email}</p><small>{account.emailVerified ? (lastSync ? `Last synchronized ${new Date(lastSync).toLocaleString()}` : "Cloud sync is ready on this device.") : "Verify this email address before cloud synchronization can begin."}</small>{!account.emailVerified && <div className="verification-note"><Cloud size={16} /><span>Verification is required before cloud sync begins.</span><button className="text-button" disabled={busy} onClick={() => void resendVerification()}>Resend</button></div>}</div>
          <span className={`status-pill ${account.emailVerified ? "good" : "warning"}`}>{account.emailVerified ? "Verified" : "Verify email"}</span>
          {syncError && <p className="inline-error sync-error">{syncError}</p>}
          {message && <p className="inline-success sync-error"><LogIn size={15} />{message}</p>}
          {error && <p className="inline-error sync-error">{error}</p>}
          <div className="account-actions"><button className="button primary compact" disabled={syncing || !account.emailVerified} onClick={() => void onSync()}><RefreshCw className={syncing ? "spin" : ""} size={16} />{syncing ? "Syncing…" : "Sync now"}</button><button className="button secondary compact" disabled={busy || syncing} onClick={() => void logout()}><LogOut size={16} />Sign out</button></div>
        </article>
        <StorageCard storageKind={storageKind} account={account} />
        <SecuritySettings />
        <DataSettings storageKind={storageKind} onDeleteCloudData={onDeleteCloudData} onDeleteAccount={onDeleteAccount} />
      </> : <>
        <article className="account-card account-card-wide auth-card">
          <div className="auth-tabs"><button className={mode === "login" ? "active" : ""} onClick={() => changeMode("login")}><LogIn size={16} />Sign in</button><button className={mode === "register" ? "active" : ""} onClick={() => changeMode("register")}><UserPlus size={16} />Create account</button>{mode === "reset" && <button className="active"><RefreshCw size={16} />Reset password</button>}</div>
          <form onSubmit={submit} className="auth-form">{mode === "register" && <label>Display name<input required value={displayName} onChange={(event) => setDisplayName(event.target.value)} /></label>}{(mode === "login" || mode === "register" || mode === "forgot") && <label>Email<input required type="email" autoComplete="email" value={email} onChange={(event) => setEmail(event.target.value)} /></label>}{(mode === "login" || mode === "register" || mode === "reset") && <label>Password<input required minLength={mode === "login" ? 1 : 10} autoComplete={mode === "login" ? "current-password" : "new-password"} type="password" value={password} onChange={(event) => setPassword(event.target.value)} />{mode !== "login" && <small>At least 10 characters.</small>}</label>}{(mode === "register" || mode === "reset") && <label>Confirm password<input required minLength={10} autoComplete="new-password" type="password" value={confirmPassword} onChange={(event) => setConfirmPassword(event.target.value)} /></label>}{mode === "login" && <div className="auth-assist"><button type="button" className="text-button" onClick={() => changeMode("forgot")}>Forgot password?</button></div>}{mode === "forgot" && <p className="auth-help">Enter your account email and BookStats will send a one-hour password reset link if the account exists.</p>}{error && <p className="inline-error">{error}</p>}{message && <p className="inline-success"><LogIn size={15} />{message}</p>}<button className="button primary" disabled={busy}>{busy ? "Working…" : mode === "login" ? "Sign in" : mode === "register" ? "Create account" : mode === "forgot" ? "Send reset link" : "Change password"}</button>{mode === "forgot" && <button type="button" className="button secondary" onClick={() => changeMode("login")}>Back to sign in</button>}</form>
        </article>
        <StorageCard storageKind={storageKind} account={null} />
      </>}
    </section>
  </>;
}

function StorageCard({ storageKind, account }: { storageKind?: "indexeddb" | "sqlite"; account: UserAccount | null }) {
  const desktop = storageKind === "sqlite";
  const synced = Boolean(account?.emailVerified);
  const awaitingVerification = Boolean(account && !account.emailVerified);

  const storageLabel = desktop
    ? synced ? "Desktop SQLite + account sync" : awaitingVerification ? "Desktop SQLite — local until verification" : "Desktop SQLite — local only"
    : synced ? "Browser IndexedDB + account sync" : awaitingVerification ? "Browser IndexedDB — local until verification" : "Browser IndexedDB — local only";

  const explanation = synced
    ? `Your library is stored ${desktop ? "on this computer" : "in this browser on this device"} and synchronized with your BookStats account for use across devices. Exported BookStats backups are still recommended as an independent recovery copy.`
    : awaitingVerification
      ? `Your library is currently stored ${desktop ? "on this computer" : "in this browser on this device"}. Verify your email to enable cloud synchronization across devices.`
      : desktop
        ? "Your library persists on this computer without an account. Create an account to synchronize a cloud copy across devices; exported BookStats backups provide an additional recovery copy."
        : "Your library persists in this browser on this device. Clearing site data, using private browsing, or switching browsers or devices can make it unavailable. Create an account to synchronize a cloud copy across devices.";

  return <article className="account-card account-card-compact">
    <div className="tool-icon"><Database size={20} /></div>
    <div className="account-details"><h2>Storage</h2><p>{storageLabel}</p><small>{explanation}</small></div>
    <span className={`status-pill ${synced ? "good" : "warning"}`}>{synced ? "Cloud sync" : awaitingVerification ? "Verify to sync" : "Local only"}</span>
  </article>;
}

function SecuritySettings() {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string>();
  const [error, setError] = useState<string>();
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirm, setConfirm] = useState("");

  function close() { if (busy) return; setOpen(false); setError(undefined); setCurrentPassword(""); setNewPassword(""); setConfirm(""); }
  async function changePassword(event: FormEvent) {
    event.preventDefault(); setBusy(true); setError(undefined); setMessage(undefined);
    try {
      if (newPassword !== confirm) throw new Error("Passwords do not match.");
      await changeAccountPassword(currentPassword, newPassword, confirm);
      setMessage("Password changed. Other signed-in devices were disconnected.");
      setCurrentPassword(""); setNewPassword(""); setConfirm(""); setOpen(false);
    } catch (err) { setError(err instanceof Error ? err.message : "Could not change the password."); }
    finally { setBusy(false); }
  }

  return <>
    <article className="account-card account-card-compact account-simple-card">
      <div className="tool-icon"><ShieldCheck size={20} /></div>
      <div><h2>Account security</h2><p>Update the password used to sign in to BookStats.</p>{message && <p className="inline-success"><ShieldCheck size={14} />{message}</p>}</div>
      <button className="button secondary compact account-card-action" onClick={() => { setError(undefined); setOpen(true); }}><KeyRound size={16} />Change password</button>
    </article>
    {open && <div className="modal-backdrop account-modal-backdrop" role="presentation">
      <section className="account-password-modal" role="dialog" aria-modal="true" aria-labelledby="change-password-title">
        <div className="form-header"><div><p className="eyebrow">Account security</p><h2 id="change-password-title">Change password</h2></div><button type="button" className="icon-button" aria-label="Close" onClick={close}><X size={18} /></button></div>
        <p className="modal-intro">Enter your current password, then choose a new password with at least 10 characters.</p>
        <form onSubmit={changePassword} className="password-modal-form">
          <label>Current password<input type="password" autoComplete="current-password" required value={currentPassword} onChange={(event) => setCurrentPassword(event.target.value)} /></label>
          <label>New password<input type="password" autoComplete="new-password" minLength={10} required value={newPassword} onChange={(event) => setNewPassword(event.target.value)} /></label>
          <label>Confirm new password<input type="password" autoComplete="new-password" minLength={10} required value={confirm} onChange={(event) => setConfirm(event.target.value)} /></label>
          {error && <p className="inline-error">{error}</p>}
          <div className="form-actions"><button type="button" className="button secondary" disabled={busy} onClick={close}>Cancel</button><button className="button primary" disabled={busy}>{busy ? "Saving…" : "Save password"}</button></div>
        </form>
      </section>
    </div>}
  </>;
}

function DataSettings({ storageKind, onDeleteCloudData, onDeleteAccount }: { storageKind?: "indexeddb" | "sqlite"; onDeleteCloudData: () => Promise<void>; onDeleteAccount: (password: string) => Promise<void> }) {
  const [busy, setBusy] = useState(false); const [deletePassword, setDeletePassword] = useState(""); const [confirmDelete, setConfirmDelete] = useState(false); const [error, setError] = useState<string>();
  async function clearCloud() { if (!window.confirm(`Delete the cloud copy of this BookStats library and sign out? ${storageKind === "sqlite" ? "Your desktop SQLite library will remain on this computer." : "The browser cache will be cleared when you sign out."}`)) return; setBusy(true); try { await onDeleteCloudData(); } catch (err) { setError(err instanceof Error ? err.message : "Could not delete cloud data."); } finally { setBusy(false); } }
  async function removeAccount() { if (!deletePassword) return; if (!window.confirm("Permanently delete this BookStats account and its cloud library? This cannot be undone on the server. Create a backup first if you want a recovery copy.")) return; setBusy(true); try { await onDeleteAccount(deletePassword); } catch (err) { setError(err instanceof Error ? err.message : "Could not delete the account."); } finally { setBusy(false); } }
  return <article className="account-card account-card-compact account-simple-card"><div className="tool-icon"><Database size={20} /></div><div><h2>Data & privacy</h2><p>Disconnect this device from the cloud or permanently delete the account.</p></div><div className="account-data-actions"><button className="button secondary compact" disabled={busy} onClick={() => void clearCloud()}>Delete cloud copy & disconnect</button><button className="button danger-button compact" onClick={() => setConfirmDelete(!confirmDelete)}><Trash2 size={15} />Delete account</button></div>{confirmDelete && <div className="danger-zone"><strong>Permanently delete account</strong><p>This removes the account and cloud library. Enter your password to confirm.</p><div><input type="password" autoComplete="current-password" value={deletePassword} onChange={(event) => setDeletePassword(event.target.value)} placeholder="Current password" /><button className="button danger-button compact" disabled={busy || !deletePassword} onClick={() => void removeAccount()}>Delete permanently</button></div></div>}{error && <p className="inline-error sync-error">{error}</p>}</article>;
}

function clearAccountLinkQuery() { const url = new URL(window.location.href); url.searchParams.delete("verify"); url.searchParams.delete("reset"); window.history.replaceState({}, "", `${url.pathname}${url.search}${url.hash}`); }
