const TOKEN_KEY = "bookstats.adminAuthToken";
const DEFAULT_API = "http://127.0.0.1:8787/api/v1";

export class ApiError extends Error {
  constructor(message: string, readonly status: number) { super(message); this.name = "ApiError"; }
}

export interface AdminAccount { id: string; email: string; displayName: string; createdAt: string; }
export interface AdminUser {
  id: string; email: string; displayName: string; role: string; disabled: boolean; disabledAt?: string; emailVerified: boolean;
  createdAt: string; updatedAt?: string; bookCount: number; shelfCount: number; goalCount: number; deletedRecordCount?: number;
  activeSessionCount?: number; lastActiveAt?: string; ownedBookCount?: number; totalReadings?: number; statusCounts?: Record<string, number>;
}
export interface DashboardData {
  metrics: { totalUsers: number; disabledUsers: number; admins: number; active24h: number; active7d: number; newUsers30d: number; books: number; shelves: number; goals: number; deletedRecords: number; activeSessions: number; metadataCacheEntries: number; databaseBytes: number; };
  server: { version: string; schemaVersion: number; uptimeSeconds: number; databaseLatencyMs: number; emailConfigured: boolean; metadataProviders: Array<{ id: string; label: string; configured: boolean; }> };
  recentUsers: Array<{ id: string; email: string; displayName: string; createdAt: string; }>;
}
export interface AdminRecord { id: string; recordType: "book" | "shelf" | "goal"; data: Record<string, unknown> | null; clientUpdatedAt: string; serverUpdatedAt: string; revision: number; deleted: boolean; deletedAt?: string; }
export interface AuditEntry { id: string; adminEmail: string; action: string; targetUserId?: string; targetEmail?: string; targetRecordId?: string; details: Record<string, unknown>; ipAddress?: string; createdAt: string; }

function isTauriRuntime(): boolean { return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window; }
function apiBaseUrl(): string {
  const configured = (import.meta.env.VITE_BOOKSTATS_API_URL as string | undefined)?.replace(/\/$/, "");
  if (typeof window !== "undefined" && import.meta.env.PROD && !isTauriRuntime()) return `${window.location.origin}/bookstats/api/v1`;
  return configured ?? DEFAULT_API;
}
export function getAdminToken(): string | null { return sessionStorage.getItem(TOKEN_KEY); }
function setAdminToken(token: string | null): void { token ? sessionStorage.setItem(TOKEN_KEY, token) : sessionStorage.removeItem(TOKEN_KEY); }

async function request<T>(path: string, init: RequestInit = {}, authenticated = true): Promise<T> {
  const headers = new Headers(init.headers);
  if (init.body && !headers.has("content-type")) headers.set("content-type", "application/json");
  headers.set("x-bookstats-client-version", __BOOKSTATS_ADMIN_VERSION__);
  if (authenticated) {
    const token = getAdminToken();
    if (!token) throw new ApiError("Administrator sign-in required.", 401);
    headers.set("authorization", `Bearer ${token}`);
  }
  const response = await fetch(`${apiBaseUrl()}${path}`, { ...init, headers });
  const payload = await response.json().catch(() => ({})) as T & { error?: string };
  if (!response.ok) {
    if (response.status === 401 && authenticated) setAdminToken(null);
    throw new ApiError(payload.error || `BookStats returned ${response.status}.`, response.status);
  }
  return payload;
}

export async function loginAdmin(email: string, password: string): Promise<AdminAccount> {
  const result = await request<{ token: string; admin: AdminAccount }>("/admin/auth/login", { method: "POST", body: JSON.stringify({ email, password }) }, false);
  setAdminToken(result.token); return result.admin;
}
export async function currentAdmin(): Promise<AdminAccount | null> {
  if (!getAdminToken()) return null;
  try { return (await request<{ admin: AdminAccount }>("/admin/auth/me")).admin; }
  catch { setAdminToken(null); return null; }
}
export async function logoutAdmin(): Promise<void> { try { if (getAdminToken()) await request("/admin/auth/logout", { method: "POST" }); } finally { setAdminToken(null); } }
export const loadDashboard = () => request<DashboardData>("/admin/dashboard");
export const loadUsers = (q = "", offset = 0) => request<{ users: AdminUser[]; total: number }>(`/admin/users?${new URLSearchParams({ q, limit: "50", offset: String(offset) })}`);
export const loadUser = (id: string) => request<{ user: AdminUser }>(`/admin/users/${id}`);
export const updateUser = (id: string, input: { displayName: string; email: string; emailVerified: boolean }) => request<{ ok: boolean }>(`/admin/users/${id}`, { method: "PATCH", body: JSON.stringify(input) });
export const setUserDisabled = (id: string, disabled: boolean) => request<{ ok: boolean }>(`/admin/users/${id}/${disabled ? "disable" : "enable"}`, { method: "POST" });
export const invalidateSessions = (id: string) => request<{ ok: boolean; deleted: number }>(`/admin/users/${id}/invalidate-sessions`, { method: "POST" });
export const forcePasswordReset = (id: string) => request<{ ok: boolean }>(`/admin/users/${id}/force-password-reset`, { method: "POST" });
export const clearCloudLibrary = (id: string, confirmation: string) => request<{ ok: boolean; deleted: number }>(`/admin/users/${id}/cloud-library`, { method: "DELETE", body: JSON.stringify({ confirmation }) });
export const deleteUser = (id: string, confirmation: string) => request<{ ok: boolean }>(`/admin/users/${id}`, { method: "DELETE", body: JSON.stringify({ confirmation }) });
export const loadRecords = (id: string, type: string, q: string, includeDeleted: boolean) => request<{ records: AdminRecord[]; total: number }>(`/admin/users/${id}/records?${new URLSearchParams({ type, q, includeDeleted: String(includeDeleted), limit: "100" })}`);
export const saveRecord = (userId: string, record: AdminRecord, data: Record<string, unknown>) => request<{ ok: boolean; revision: number }>(`/admin/users/${userId}/records/${record.id}`, { method: "PUT", body: JSON.stringify({ recordType: record.recordType, data }) });
export const deleteRecord = (userId: string, recordId: string) => request<{ ok: boolean }>(`/admin/users/${userId}/records/${recordId}`, { method: "DELETE", body: JSON.stringify({ confirmation: "DELETE RECORD" }) });
export const loadAudit = (q = "") => request<{ entries: AuditEntry[] }>(`/admin/audit?${new URLSearchParams({ q, limit: "150" })}`);
