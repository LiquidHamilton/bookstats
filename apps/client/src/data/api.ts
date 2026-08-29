import type { AuthResponse, MetadataCandidate, SyncMutation, SyncResponse, UserAccount } from "@bookstats/domain";

const TOKEN_KEY = "bookstats.authToken";
const DEFAULT_API = "http://127.0.0.1:8787/api/v1";
const UPDATE_REQUIRED_EVENT = "bookstats:update-required";

class ApiError extends Error {
  constructor(message: string, readonly status: number) { super(message); this.name = "ApiError"; }
}

interface ServerInfo {
  name: string;
  version: number;
  appVersion?: string;
  minimumClientVersion?: string;
  status?: string;
}

function isTauriRuntime(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export function apiBaseUrl(): string {
  const configured = (import.meta.env.VITE_BOOKSTATS_API_URL as string | undefined)?.replace(/\/$/, "");
  // Browser releases should talk back to the same host that served the app. This avoids
  // www/non-www CORS mismatches and prevents a production web build from ever targeting localhost.
  if (typeof window !== "undefined" && import.meta.env.PROD && !isTauriRuntime()) {
    return `${window.location.origin}/bookstats/api/v1`;
  }
  return configured ?? DEFAULT_API;
}

function notifyUpdateRequired(serverVersion?: string): void {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(UPDATE_REQUIRED_EVENT, { detail: { serverVersion } }));
}

function semverParts(version: string): [number, number, number] {
  const match = version.trim().match(/^(\d+)\.(\d+)\.(\d+)/);
  return match ? [Number(match[1]), Number(match[2]), Number(match[3])] : [0, 0, 0];
}

function compareVersions(left: string, right: string): number {
  const a = semverParts(left); const b = semverParts(right);
  for (let i = 0; i < 3; i += 1) { if (a[i] !== b[i]) return a[i] - b[i]; }
  return 0;
}

export function updateRequiredEventName(): string { return UPDATE_REQUIRED_EVENT; }

export async function checkServerCompatibility(): Promise<ServerInfo | null> {
  try {
    const response = await fetch(apiBaseUrl(), {
      cache: "no-store",
      headers: { "x-bookstats-client-version": __BOOKSTATS_VERSION__ }
    });
    if (!response.ok) return null;
    const info = await response.json() as ServerInfo;
    const minimum = info.minimumClientVersion;
    const serverVersion = info.appVersion;
    if ((minimum && compareVersions(__BOOKSTATS_VERSION__, minimum) < 0) || (serverVersion && compareVersions(__BOOKSTATS_VERSION__, serverVersion) < 0)) {
      notifyUpdateRequired(serverVersion);
    }
    return info;
  } catch {
    return null;
  }
}

export function getAuthToken(): string | null { return localStorage.getItem(TOKEN_KEY); }
export function setAuthToken(token: string | null): void {
  if (token) localStorage.setItem(TOKEN_KEY, token); else localStorage.removeItem(TOKEN_KEY);
}

async function apiFetch<T>(path: string, init: RequestInit = {}, authenticated = false): Promise<T> {
  const headers = new Headers(init.headers);
  if (init.body && !headers.has("content-type")) headers.set("content-type", "application/json");
  headers.set("x-bookstats-client-version", __BOOKSTATS_VERSION__);
  if (authenticated) {
    const token = getAuthToken();
    if (!token) throw new Error("You are not signed in.");
    headers.set("authorization", `Bearer ${token}`);
  }
  let response: Response;
  try {
    response = await fetch(`${apiBaseUrl()}${path}`, { ...init, headers });
  } catch (error) {
    console.error("BookStats API request failed", error);
    throw new ApiError("Could not reach the BookStats server. Check your connection and try again.", 0);
  }
  const payload = await response.json().catch(() => ({})) as { error?: string; serverVersion?: string } & T;
  if (response.status === 426) notifyUpdateRequired(payload.serverVersion);
  if (response.status === 413 && path === "/sync") {
    throw new ApiError(
      "Synchronization data was too large for the server. Your changes are still saved locally and will remain queued for the next sync.",
      413
    );
  }
  if (!response.ok) throw new ApiError(payload.error || `BookStats server returned ${response.status}.`, response.status);
  return payload;
}

export async function registerAccount(email: string, password: string, confirmPassword: string, displayName: string): Promise<AuthResponse> {
  const result = await apiFetch<AuthResponse>("/auth/register", { method: "POST", body: JSON.stringify({ email, password, confirmPassword, displayName }) });
  setAuthToken(result.token);
  return result;
}
export async function loginAccount(email: string, password: string): Promise<AuthResponse> {
  const result = await apiFetch<AuthResponse>("/auth/login", { method: "POST", body: JSON.stringify({ email, password }) });
  setAuthToken(result.token);
  return result;
}
export async function currentAccount(): Promise<UserAccount | null> {
  if (!getAuthToken()) return null;
  try { return (await apiFetch<{ user: UserAccount }>("/auth/me", {}, true)).user; }
  catch (error) { if (error instanceof ApiError && error.status === 401) setAuthToken(null); return null; }
}
export async function logoutAccount(): Promise<void> {
  try { if (getAuthToken()) await apiFetch("/auth/logout", { method: "POST" }, true); } finally { setAuthToken(null); }
}

export async function changeAccountPassword(currentPassword: string, password: string, confirmPassword: string): Promise<{ ok: boolean }> {
  return apiFetch<{ ok: boolean }>("/auth/change-password", { method: "POST", body: JSON.stringify({ currentPassword, password, confirmPassword }) }, true);
}
export async function changeAccountEmail(currentPassword: string, email: string): Promise<{ ok: boolean; user: UserAccount; verificationSent?: boolean }> {
  return apiFetch<{ ok: boolean; user: UserAccount; verificationSent?: boolean }>("/auth/change-email", { method: "POST", body: JSON.stringify({ currentPassword, email }) }, true);
}
export async function deleteCloudLibrary(): Promise<{ ok: boolean; deleted: number }> {
  return apiFetch<{ ok: boolean; deleted: number }>("/account/cloud-library", { method: "DELETE" }, true);
}
export async function deleteAccount(password: string): Promise<{ ok: boolean }> {
  const result = await apiFetch<{ ok: boolean }>("/account/delete", { method: "POST", body: JSON.stringify({ password }) }, true);
  if (result.ok) setAuthToken(null);
  return result;
}
export async function resendVerificationEmail(): Promise<{ ok: boolean; alreadyVerified?: boolean; throttled?: boolean }> {
  return apiFetch("/auth/resend-verification", { method: "POST" }, true);
}
export async function verifyEmail(token: string): Promise<{ ok: boolean; user: UserAccount }> {
  return apiFetch("/auth/verify-email", { method: "POST", body: JSON.stringify({ token }) });
}
export async function requestPasswordReset(email: string): Promise<{ ok: boolean; message: string }> {
  return apiFetch("/auth/forgot-password", { method: "POST", body: JSON.stringify({ email }) });
}
export async function resetPassword(token: string, password: string, confirmPassword: string): Promise<{ ok: boolean }> {
  const result = await apiFetch<{ ok: boolean }>("/auth/reset-password", { method: "POST", body: JSON.stringify({ token, password, confirmPassword }) });
  setAuthToken(null);
  return result;
}
export async function syncAccountLibrary(cursor: string | undefined, changes: SyncMutation[]): Promise<SyncResponse> {
  return apiFetch<SyncResponse>("/sync", { method: "POST", body: JSON.stringify({ cursor, changes }) }, true);
}

export async function fetchCatalogCoverForInspection(coverUrl: string): Promise<Blob | undefined> {
  const params = new URLSearchParams({ url: coverUrl });
  try {
    const response = await fetch(`${apiBaseUrl()}/covers/inspect?${params.toString()}`, {
      cache: "force-cache",
      headers: { "x-bookstats-client-version": __BOOKSTATS_VERSION__ }
    });
    if (!response.ok) return undefined;
    const blob = await response.blob();
    return blob.type.startsWith("image/") ? blob : undefined;
  } catch {
    return undefined;
  }
}

export async function archiveSelectedCover(coverUrl: string): Promise<{ assetId: string; assetToken: string; sourceUrl?: string }> {
  return apiFetch<{ assetId: string; assetToken: string; sourceUrl?: string }>("/covers/archive", { method: "POST", body: JSON.stringify({ coverUrl }) }, true);
}
export function cloudCoverUrl(assetId: string, assetToken: string, revision?: string): string {
  const base = `${apiBaseUrl()}/covers/${encodeURIComponent(assetId)}/${encodeURIComponent(assetToken)}`;
  return revision ? `${base}?v=${encodeURIComponent(revision)}` : base;
}
export async function searchMetadata(query: string, isIsbn: boolean): Promise<MetadataCandidate[]> {
  const params = new URLSearchParams(isIsbn ? { isbn: query } : { q: query });
  return (await apiFetch<{ results: MetadataCandidate[] }>(`/metadata/search?${params}`)).results;
}
export async function metadataDetails(candidate: MetadataCandidate): Promise<MetadataCandidate> {
  return (await apiFetch<{ details: MetadataCandidate }>("/metadata/details", { method: "POST", body: JSON.stringify({ candidate }) })).details;
}

export interface FeedbackDiagnostics {
  version: string;
  platform: string;
  storage?: string;
  signedIn: boolean;
  emailVerified: boolean;
  bookCount: number;
  shelfCount: number;
}

export async function submitFeedback(kind: "bug" | "feature", message: string, contactEmail: string | undefined, diagnostics: FeedbackDiagnostics): Promise<{ ok: boolean }> {
  return apiFetch<{ ok: boolean }>("/feedback", { method: "POST", body: JSON.stringify({ kind, message, contactEmail, diagnostics }) });
}
