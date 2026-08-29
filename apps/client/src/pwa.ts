export function isStandaloneWebApp(): boolean {
  if (typeof window === "undefined") return false;
  const navigatorWithStandalone = navigator as Navigator & { standalone?: boolean };
  return window.matchMedia("(display-mode: standalone)").matches || navigatorWithStandalone.standalone === true;
}

export function registerBookStatsServiceWorker(): void {
  if (typeof window === "undefined") return;
  document.documentElement.classList.toggle("standalone-webapp", isStandaloneWebApp());
  if (!("serviceWorker" in navigator) || import.meta.env.DEV || "__TAURI_INTERNALS__" in window) return;
  window.addEventListener("load", () => {
    void navigator.serviceWorker.register(`${import.meta.env.BASE_URL}sw.js`, { scope: import.meta.env.BASE_URL }).catch((error) => {
      console.warn("BookStats service worker registration failed", error);
    });
  }, { once: true });
}
