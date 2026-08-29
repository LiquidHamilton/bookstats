import { useEffect, useMemo, useState } from "react";
import { Download, MoreHorizontal, Share, X } from "lucide-react";
import type { UserAccount } from "@bookstats/domain";
import { isStandaloneWebApp } from "../pwa";

const DISMISS_KEY = "bookstats.installPromptDismissedAt.v1";
const DISMISS_FOR_MS = 30 * 24 * 60 * 60 * 1000;

type BeforeInstallPromptEvent = Event & {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed"; platform: string }>;
};

interface InstallAppPromptProps {
  account: UserAccount | null;
  storageKind?: string;
}

function recentlyDismissed(): boolean {
  const raw = localStorage.getItem(DISMISS_KEY);
  if (!raw) return false;
  const dismissedAt = Number(raw);
  return Number.isFinite(dismissedAt) && Date.now() - dismissedAt < DISMISS_FOR_MS;
}

function isIos(): boolean {
  const ua = navigator.userAgent;
  return /iPhone|iPad|iPod/i.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
}

function isAndroid(): boolean { return /Android/i.test(navigator.userAgent); }

export function InstallAppPrompt({ account, storageKind }: InstallAppPromptProps) {
  const [deferredPrompt, setDeferredPrompt] = useState<BeforeInstallPromptEvent>();
  const [visible, setVisible] = useState(false);
  const platform = useMemo(() => isIos() ? "ios" : isAndroid() ? "android" : "other", []);

  useEffect(() => {
    if (isStandaloneWebApp() || platform === "other" || recentlyDismissed()) return;

    const showTimer = window.setTimeout(() => setVisible(true), 1400);
    const onBeforeInstallPrompt = (event: Event) => {
      event.preventDefault();
      setDeferredPrompt(event as BeforeInstallPromptEvent);
      setVisible(true);
    };
    const onInstalled = () => {
      localStorage.setItem(DISMISS_KEY, String(Date.now()));
      setVisible(false);
    };

    window.addEventListener("beforeinstallprompt", onBeforeInstallPrompt);
    window.addEventListener("appinstalled", onInstalled);
    return () => {
      window.clearTimeout(showTimer);
      window.removeEventListener("beforeinstallprompt", onBeforeInstallPrompt);
      window.removeEventListener("appinstalled", onInstalled);
    };
  }, [platform]);

  if (!visible) return null;

  const localOnlyIos = platform === "ios" && !account && storageKind === "indexeddb";

  const dismiss = () => {
    localStorage.setItem(DISMISS_KEY, String(Date.now()));
    setVisible(false);
  };

  const installAndroid = async () => {
    if (!deferredPrompt) return;
    await deferredPrompt.prompt();
    const choice = await deferredPrompt.userChoice;
    if (choice.outcome === "accepted") {
      localStorage.setItem(DISMISS_KEY, String(Date.now()));
      setVisible(false);
    }
    setDeferredPrompt(undefined);
  };

  return <div className="modal-backdrop install-app-backdrop" onMouseDown={(event) => event.target === event.currentTarget && dismiss()}>
    <section className="install-app-modal" role="dialog" aria-modal="true" aria-labelledby="bookstats-install-title">
      <button className="icon-button install-app-close" type="button" onClick={dismiss} aria-label="Dismiss install suggestion"><X size={18} /></button>
      <img className="install-app-logo" src={`${import.meta.env.BASE_URL}icons/icon-192.png`} alt="" />
      <div className="install-app-copy">
        <p className="eyebrow">BookStats on your phone</p>
        <h2 id="bookstats-install-title">Add BookStats to your Home Screen</h2>
        <p>Open BookStats in its own app window with more screen space and quicker access from your Home Screen.</p>
      </div>

      {platform === "ios" && <>
        <div className="install-app-steps" aria-label="iPhone installation steps">
          <span><MoreHorizontal size={18} /><strong>More</strong></span><b>→</b><span><Share size={18} /><strong>Share</strong></span><b>→</b><span><Download size={18} /><strong>Add to Home Screen</strong></span>
        </div>
        <p className="install-app-hint">Leave <strong>Open as Web App</strong> enabled, then tap <strong>Add</strong>.</p>
        {account && <p className="install-app-warning">After installation, sign in to BookStats in the new Home Screen app so your synchronized library can be restored there.</p>}
        {localOnlyIos && <p className="install-app-warning"><strong>Local-only library:</strong> iPhone keeps the Home Screen app's local storage separate from the browser you installed it from. Export or back up your library before installing, then import it into the Home Screen app, or create an account and sync first.</p>}
      </>}

      {platform === "android" && <>
        {deferredPrompt
          ? <button className="button primary install-app-action" type="button" onClick={() => void installAndroid()}><Download size={17} />Install BookStats</button>
          : <div className="install-app-note"><MoreHorizontal size={18} /><span><strong>Use your browser menu.</strong><small>Choose <em>Install app</em> or <em>Add to Home screen</em>. Chrome may show its native install prompt automatically.</small></span></div>}
      </>}

      <button className="text-button install-app-later" type="button" onClick={dismiss}>Maybe later</button>
    </section>
  </div>;
}
