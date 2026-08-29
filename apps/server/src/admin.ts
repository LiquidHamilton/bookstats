import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { createHash, randomBytes, scrypt as scryptCallback, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";
import type { Pool, PoolClient } from "pg";
import { emailConfigured, sendPasswordResetEmail } from "./email.js";
import { metadataProviderStatuses } from "./metadata.js";
import { listCoverAssetPathsForUser, removeStoredCoverFiles } from "./covers.js";

const scrypt = promisify(scryptCallback);
const ADMIN_SESSION_DAYS = 30;
const RESET_HOURS = 1;

interface AdminActor {
  id: string;
  email: string;
  displayName: string;
  createdAt: string;
}

interface AdminUserRow {
  id: string;
  email: string;
  display_name: string;
  password_hash: string;
  role: "user" | "admin";
  disabled_at: Date | null;
  created_at: Date;
}

function normalizeEmail(value?: string): string | undefined {
  const email = value?.trim().toLowerCase();
  return email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : undefined;
}

function sha256(value: string): string { return createHash("sha256").update(value).digest("hex"); }

async function verifyPassword(password: string, encoded: string): Promise<boolean> {
  const [algorithm, saltHex, hashHex] = encoded.split(":");
  if (algorithm !== "scrypt" || !saltHex || !hashHex) return false;
  const expected = Buffer.from(hashHex, "hex");
  const actual = await scrypt(password, Buffer.from(saltHex, "hex"), expected.length) as Buffer;
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}

async function createSession(client: PoolClient, userId: string): Promise<string> {
  const token = randomBytes(32).toString("base64url");
  await client.query(
    `INSERT INTO sessions (user_id, token_hash, expires_at) VALUES ($1, $2, now() + ($3 || ' days')::interval)`,
    [userId, sha256(token), ADMIN_SESSION_DAYS]
  );
  return token;
}

async function issueResetToken(client: PoolClient, userId: string): Promise<string> {
  const token = randomBytes(32).toString("base64url");
  await client.query("UPDATE password_reset_tokens SET used_at = now() WHERE user_id = $1 AND used_at IS NULL", [userId]);
  await client.query(
    `INSERT INTO password_reset_tokens (user_id, token_hash, expires_at) VALUES ($1, $2, now() + ($3 || ' hours')::interval)`,
    [userId, sha256(token), RESET_HOURS]
  );
  return token;
}

async function writeAudit(
  client: PoolClient,
  request: FastifyRequest,
  action: string,
  targetUserId?: string,
  targetRecordId?: string,
  details: Record<string, unknown> = {}
): Promise<void> {
  const actor = request.bookstatsAdmin;
  if (!actor) throw new Error("Administrator context is missing.");
  await client.query(
    `INSERT INTO admin_audit_log (admin_user_id, admin_email, action, target_user_id, target_record_id, details, ip_address)
     VALUES ($1, $2, $3, $4, $5, $6::jsonb, $7)`,
    [actor.id, actor.email, action, targetUserId ?? null, targetRecordId ?? null, JSON.stringify(details), request.ip]
  );
}

async function withAuditTransaction<T>(
  pool: Pool,
  request: FastifyRequest,
  action: string,
  targetUserId: string | undefined,
  targetRecordId: string | undefined,
  details: Record<string, unknown>,
  operation: (client: PoolClient) => Promise<T>
): Promise<T> {
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const value = await operation(client);
    await writeAudit(client, request, action, targetUserId, targetRecordId, details);
    await client.query("COMMIT");
    return value;
  } catch (error) {
    await client.query("ROLLBACK");
    throw error;
  } finally {
    client.release();
  }
}

function actorFromRow(row: Pick<AdminUserRow, "id" | "email" | "display_name" | "created_at">): AdminActor {
  return { id: row.id, email: row.email, displayName: row.display_name, createdAt: row.created_at.toISOString() };
}

export function registerAdminRoutes(app: FastifyInstance, pool: Pool | undefined, appVersion: string): void {
  const adminLoginWindows = new Map<string, number[]>();
  const requirePool = (reply: FastifyReply): Pool | undefined => {
    if (pool) return pool;
    void reply.code(503).send({ error: "The BookStats database is not configured." });
    return undefined;
  };

  const authenticateAdmin = async (request: FastifyRequest, reply: FastifyReply) => {
    const db = requirePool(reply); if (!db) return;
    const header = request.headers.authorization;
    const token = header?.startsWith("Bearer ") ? header.slice(7).trim() : "";
    if (!token) return reply.code(401).send({ error: "Administrator authentication required." });
    const tokenHash = sha256(token);
    const result = await db.query<AdminUserRow>(
      `SELECT u.id, u.email, u.display_name, u.password_hash, u.role, u.disabled_at, u.created_at
       FROM sessions s JOIN users u ON u.id = s.user_id
       WHERE s.token_hash = $1 AND s.expires_at > now() AND u.role = 'admin' AND u.disabled_at IS NULL`,
      [tokenHash]
    );
    const row = result.rows[0];
    if (!row) return reply.code(401).send({ error: "Administrator session expired or invalid." });
    request.bookstatsAdmin = actorFromRow(row);
    request.bookstatsAdminTokenHash = tokenHash;
    void db.query("UPDATE sessions SET last_used_at = now() WHERE token_hash = $1", [tokenHash]);
  };

  app.post<{ Body: { email?: string; password?: string } }>("/api/v1/admin/auth/login", async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const email = normalizeEmail(request.body?.email);
    const password = request.body?.password ?? "";
    if (!email || !password) return reply.code(400).send({ error: "Email and password are required." });
    const now = Date.now();
    const cutoff = now - 15 * 60 * 1000;
    const attempts = (adminLoginWindows.get(request.ip) ?? []).filter((timestamp) => timestamp > cutoff);
    if (attempts.length >= 8) return reply.code(429).send({ error: "Too many administrator sign-in attempts. Try again later." });
    attempts.push(now); adminLoginWindows.set(request.ip, attempts);
    const client = await db.connect();
    try {
      const result = await client.query<AdminUserRow>(
        `SELECT id, email, display_name, password_hash, role, disabled_at, created_at
         FROM users WHERE email = $1`, [email]
      );
      const row = result.rows[0];
      if (!row || row.role !== "admin" || row.disabled_at || !(await verifyPassword(password, row.password_hash))) {
        return reply.code(403).send({ error: "This account is not authorized for BookStats administration." });
      }
      const token = await createSession(client, row.id);
      adminLoginWindows.delete(request.ip);
      const actor = actorFromRow(row);
      try {
        request.bookstatsAdmin = actor;
        await writeAudit(client, request, "ADMIN_LOGIN", undefined, undefined, {});
      } finally {
        request.bookstatsAdmin = undefined;
      }
      return { token, admin: actor };
    } finally { client.release(); }
  });

  app.get("/api/v1/admin/auth/me", { preHandler: authenticateAdmin }, async (request) => ({ admin: request.bookstatsAdmin }));

  app.post("/api/v1/admin/auth/logout", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db || !request.bookstatsAdmin) return;
    const client = await db.connect();
    try {
      await client.query("BEGIN");
      await writeAudit(client, request, "ADMIN_LOGOUT", undefined, undefined, {});
      if (request.bookstatsAdminTokenHash) await client.query("DELETE FROM sessions WHERE token_hash = $1", [request.bookstatsAdminTokenHash]);
      await client.query("COMMIT");
      return { ok: true };
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally { client.release(); }
  });

  app.get("/api/v1/admin/dashboard", { preHandler: authenticateAdmin }, async (_request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const started = process.hrtime.bigint();
    const result = await db.query<{
      total_users: string; disabled_users: string; admins: string; active_24h: string; active_7d: string;
      books: string; shelves: string; goals: string; deleted_records: string; sessions: string; metadata_cache: string;
      new_users_30d: string; database_bytes: string;
    }>(`SELECT
      (SELECT count(*) FROM users) AS total_users,
      (SELECT count(*) FROM users WHERE disabled_at IS NOT NULL) AS disabled_users,
      (SELECT count(*) FROM users WHERE role = 'admin') AS admins,
      (SELECT count(DISTINCT user_id) FROM sessions WHERE last_used_at > now() - interval '24 hours' AND expires_at > now()) AS active_24h,
      (SELECT count(DISTINCT user_id) FROM sessions WHERE last_used_at > now() - interval '7 days' AND expires_at > now()) AS active_7d,
      (SELECT count(*) FROM library_records WHERE deleted_at IS NULL AND record_type = 'book') AS books,
      (SELECT count(*) FROM library_records WHERE deleted_at IS NULL AND record_type = 'shelf') AS shelves,
      (SELECT count(*) FROM library_records WHERE deleted_at IS NULL AND record_type = 'goal') AS goals,
      (SELECT count(*) FROM library_records WHERE deleted_at IS NOT NULL) AS deleted_records,
      (SELECT count(*) FROM sessions WHERE expires_at > now()) AS sessions,
      (SELECT count(*) FROM metadata_cache WHERE expires_at > now()) AS metadata_cache,
      (SELECT count(*) FROM users WHERE created_at > now() - interval '30 days') AS new_users_30d,
      pg_database_size(current_database())::text AS database_bytes`);
    const dbLatencyMs = Number(process.hrtime.bigint() - started) / 1_000_000;
    const row = result.rows[0];
    const recent = await db.query<{ id: string; email: string; display_name: string; created_at: Date }>(
      `SELECT id, email, display_name, created_at FROM users ORDER BY created_at DESC LIMIT 8`
    );
    return {
      metrics: {
        totalUsers: Number(row.total_users), disabledUsers: Number(row.disabled_users), admins: Number(row.admins),
        active24h: Number(row.active_24h), active7d: Number(row.active_7d), newUsers30d: Number(row.new_users_30d),
        books: Number(row.books), shelves: Number(row.shelves), goals: Number(row.goals), deletedRecords: Number(row.deleted_records),
        activeSessions: Number(row.sessions), metadataCacheEntries: Number(row.metadata_cache), databaseBytes: Number(row.database_bytes)
      },
      server: {
        version: appVersion,
        schemaVersion: 6,
        uptimeSeconds: Math.round(process.uptime()),
        databaseLatencyMs: Math.round(dbLatencyMs * 10) / 10,
        emailConfigured: emailConfigured(),
        metadataProviders: metadataProviderStatuses()
      },
      recentUsers: recent.rows.map((user) => ({ id: user.id, email: user.email, displayName: user.display_name, createdAt: user.created_at.toISOString() }))
    };
  });

  app.get<{ Querystring: { q?: string; limit?: string; offset?: string } }>("/api/v1/admin/users", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const q = request.query.q?.trim() ?? "";
    const limit = Math.min(100, Math.max(1, Number(request.query.limit ?? 50) || 50));
    const offset = Math.max(0, Number(request.query.offset ?? 0) || 0);
    const search = `%${q}%`;
    const count = await db.query<{ count: string }>(
      `SELECT count(*) FROM users WHERE $1 = '%%' OR email ILIKE $1 OR display_name ILIKE $1`, [search]
    );
    const result = await db.query<{
      id: string; email: string; display_name: string; role: string; disabled_at: Date | null; email_verified_at: Date | null; created_at: Date;
      book_count: string; shelf_count: string; goal_count: string; last_active_at: Date | null;
    }>(`SELECT u.id, u.email, u.display_name, u.role, u.disabled_at, u.email_verified_at, u.created_at,
        COALESCE(records.book_count, 0)::text AS book_count,
        COALESCE(records.shelf_count, 0)::text AS shelf_count,
        COALESCE(records.goal_count, 0)::text AS goal_count,
        activity.last_active_at
      FROM users u
      LEFT JOIN LATERAL (
        SELECT
          count(*) FILTER (WHERE record_type = 'book' AND deleted_at IS NULL) AS book_count,
          count(*) FILTER (WHERE record_type = 'shelf' AND deleted_at IS NULL) AS shelf_count,
          count(*) FILTER (WHERE record_type = 'goal' AND deleted_at IS NULL) AS goal_count
        FROM library_records WHERE user_id = u.id
      ) records ON true
      LEFT JOIN LATERAL (
        SELECT max(last_used_at) AS last_active_at FROM sessions WHERE user_id = u.id
      ) activity ON true
      WHERE $1 = '%%' OR u.email ILIKE $1 OR u.display_name ILIKE $1
      ORDER BY u.created_at DESC LIMIT $2 OFFSET $3`, [search, limit, offset]);
    return {
      total: Number(count.rows[0].count),
      users: result.rows.map((row) => ({
        id: row.id, email: row.email, displayName: row.display_name, role: row.role,
        disabled: Boolean(row.disabled_at), disabledAt: row.disabled_at?.toISOString(), emailVerified: Boolean(row.email_verified_at),
        createdAt: row.created_at.toISOString(), bookCount: Number(row.book_count), shelfCount: Number(row.shelf_count),
        goalCount: Number(row.goal_count), lastActiveAt: row.last_active_at?.toISOString()
      }))
    };
  });

  app.get<{ Params: { userId: string } }>("/api/v1/admin/users/:userId", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const result = await db.query<{
      id: string; email: string; display_name: string; role: string; disabled_at: Date | null; email_verified_at: Date | null; created_at: Date; updated_at: Date;
    }>(`SELECT id, email, display_name, role, disabled_at, email_verified_at, created_at, updated_at FROM users WHERE id = $1`, [request.params.userId]);
    const row = result.rows[0];
    if (!row) return reply.code(404).send({ error: "User not found." });
    const counts = await db.query<{ books: string; shelves: string; goals: string; deleted: string; sessions: string; last_active_at: Date | null; owned_books: string; total_readings: string; status_counts: Record<string, number> }>(
      `SELECT
        (SELECT count(*) FROM library_records WHERE user_id = $1 AND record_type='book' AND deleted_at IS NULL) AS books,
        (SELECT count(*) FROM library_records WHERE user_id = $1 AND record_type='shelf' AND deleted_at IS NULL) AS shelves,
        (SELECT count(*) FROM library_records WHERE user_id = $1 AND record_type='goal' AND deleted_at IS NULL) AS goals,
        (SELECT count(*) FROM library_records WHERE user_id = $1 AND deleted_at IS NOT NULL) AS deleted,
        (SELECT count(*) FROM sessions WHERE user_id = $1 AND expires_at > now()) AS sessions,
        (SELECT max(last_used_at) FROM sessions WHERE user_id = $1) AS last_active_at,
        (SELECT count(*) FROM library_records WHERE user_id=$1 AND record_type='book' AND deleted_at IS NULL AND book_data->>'owned'='true') AS owned_books,
        (SELECT COALESCE(sum(CASE WHEN jsonb_typeof(book_data->'readingSessions')='array' THEN jsonb_array_length(book_data->'readingSessions') WHEN jsonb_typeof(book_data->'readDates')='array' THEN jsonb_array_length(book_data->'readDates') ELSE 0 END),0) FROM library_records WHERE user_id=$1 AND record_type='book' AND deleted_at IS NULL) AS total_readings,
        (SELECT COALESCE(jsonb_object_agg(status, count_value), '{}'::jsonb) FROM (SELECT COALESCE(book_data->>'status','unknown') AS status, count(*)::int AS count_value FROM library_records WHERE user_id=$1 AND record_type='book' AND deleted_at IS NULL GROUP BY COALESCE(book_data->>'status','unknown')) status_rows) AS status_counts`, [row.id]
    );
    const c = counts.rows[0];
    return { user: {
      id: row.id, email: row.email, displayName: row.display_name, role: row.role, disabled: Boolean(row.disabled_at),
      disabledAt: row.disabled_at?.toISOString(), emailVerified: Boolean(row.email_verified_at), createdAt: row.created_at.toISOString(),
      updatedAt: row.updated_at.toISOString(), bookCount: Number(c.books), shelfCount: Number(c.shelves), goalCount: Number(c.goals),
      deletedRecordCount: Number(c.deleted), activeSessionCount: Number(c.sessions), lastActiveAt: c.last_active_at?.toISOString(),
      ownedBookCount: Number(c.owned_books), totalReadings: Number(c.total_readings), statusCounts: c.status_counts ?? {}
    }};
  });

  app.patch<{ Params: { userId: string }; Body: { displayName?: string; email?: string; emailVerified?: boolean } }>("/api/v1/admin/users/:userId", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const displayName = request.body?.displayName?.trim();
    const email = request.body?.email === undefined ? undefined : normalizeEmail(request.body.email);
    if (request.body?.displayName !== undefined && !displayName) return reply.code(400).send({ error: "Display name cannot be blank." });
    if (request.body?.email !== undefined && !email) return reply.code(400).send({ error: "Enter a valid email address." });
    try {
      const updated = await withAuditTransaction(db, request, "USER_PROFILE_UPDATE", request.params.userId, undefined,
        { displayNameChanged: request.body.displayName !== undefined, emailChanged: request.body.email !== undefined, verificationChanged: request.body.emailVerified !== undefined },
        async (client) => {
          const existing = await client.query<{ email: string; display_name: string; email_verified_at: Date | null }>(
            "SELECT email, display_name, email_verified_at FROM users WHERE id = $1 FOR UPDATE", [request.params.userId]
          );
          if (!existing.rows[0]) throw Object.assign(new Error("User not found."), { statusCode: 404 });
          const nextEmail = email ?? existing.rows[0].email;
          const nextName = displayName ?? existing.rows[0].display_name;
          let verifiedAt: Date | null = existing.rows[0].email_verified_at;
          if (email && email !== existing.rows[0].email) verifiedAt = null;
          if (request.body.emailVerified === true) verifiedAt = new Date();
          if (request.body.emailVerified === false) verifiedAt = null;
          const result = await client.query(
            `UPDATE users SET email=$1, display_name=$2, email_verified_at=$3, updated_at=now() WHERE id=$4
             RETURNING id, email, display_name, role, disabled_at, email_verified_at, created_at, updated_at`,
            [nextEmail, nextName, verifiedAt, request.params.userId]
          );
          return result.rows[0];
        }
      );
      return { ok: true, user: updated };
    } catch (error) {
      if (error && typeof error === "object" && "code" in error && (error as { code?: string }).code === "23505") return reply.code(409).send({ error: "That email address is already in use." });
      if (error && typeof error === "object" && "statusCode" in error) {
        const typed = error as { statusCode?: unknown; message?: unknown };
        return reply.code(Number(typed.statusCode)).send({ error: typeof typed.message === "string" ? typed.message : "Could not update the account." });
      }
      throw error;
    }
  });

  app.post<{ Params: { userId: string } }>("/api/v1/admin/users/:userId/disable", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db || !request.bookstatsAdmin) return;
    if (request.params.userId === request.bookstatsAdmin.id) return reply.code(400).send({ error: "You cannot disable your own administrator account." });
    const affected = await withAuditTransaction(db, request, "USER_DISABLE", request.params.userId, undefined, {}, async (client) => {
      const result = await client.query("UPDATE users SET disabled_at = COALESCE(disabled_at, now()), updated_at=now() WHERE id=$1", [request.params.userId]);
      if (!result.rowCount) throw Object.assign(new Error("User not found."), { statusCode: 404 });
      await client.query("DELETE FROM sessions WHERE user_id=$1", [request.params.userId]);
      return result.rowCount;
    }).catch((error) => { throw error; });
    return { ok: true, affected };
  });

  app.post<{ Params: { userId: string } }>("/api/v1/admin/users/:userId/enable", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const affected = await withAuditTransaction(db, request, "USER_ENABLE", request.params.userId, undefined, {}, async (client) => {
      const result = await client.query("UPDATE users SET disabled_at = NULL, updated_at=now() WHERE id=$1", [request.params.userId]);
      if (!result.rowCount) throw Object.assign(new Error("User not found."), { statusCode: 404 });
      return result.rowCount;
    });
    return { ok: true, affected };
  });

  app.post<{ Params: { userId: string } }>("/api/v1/admin/users/:userId/invalidate-sessions", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db || !request.bookstatsAdmin) return;
    if (request.params.userId === request.bookstatsAdmin.id) return reply.code(400).send({ error: "Use Sign out to end your own administrator session." });
    const deleted = await withAuditTransaction(db, request, "SESSIONS_INVALIDATE", request.params.userId, undefined, {}, async (client) => {
      const exists = await client.query("SELECT 1 FROM users WHERE id=$1", [request.params.userId]);
      if (!exists.rows[0]) throw Object.assign(new Error("User not found."), { statusCode: 404 });
      const result = await client.query("DELETE FROM sessions WHERE user_id=$1", [request.params.userId]);
      return result.rowCount ?? 0;
    });
    return { ok: true, deleted };
  });

  app.post<{ Params: { userId: string } }>("/api/v1/admin/users/:userId/force-password-reset", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db || !request.bookstatsAdmin) return;
    if (request.params.userId === request.bookstatsAdmin.id) return reply.code(400).send({ error: "Use the normal account password-change flow for your own administrator account." });
    if (!emailConfigured()) return reply.code(503).send({ error: "Transactional email is not configured, so BookStats cannot deliver a reset link." });
    let target: { email: string; display_name: string } | undefined;
    let token = "";
    await withAuditTransaction(db, request, "FORCE_PASSWORD_RESET", request.params.userId, undefined, {}, async (client) => {
      const result = await client.query<{ email: string; display_name: string }>("SELECT email, display_name FROM users WHERE id=$1 FOR UPDATE", [request.params.userId]);
      target = result.rows[0];
      if (!target) throw Object.assign(new Error("User not found."), { statusCode: 404 });
      token = await issueResetToken(client, request.params.userId);
      await client.query("DELETE FROM sessions WHERE user_id=$1", [request.params.userId]);
    });
    try {
      await sendPasswordResetEmail(target!.email, target!.display_name, token);
      return { ok: true };
    } catch (error) {
      request.log.error(error, "Administrator reset email could not be delivered");
      return reply.code(502).send({ error: "The account was signed out and a reset token was created, but the reset email could not be delivered." });
    }
  });

  app.delete<{ Params: { userId: string }; Body: { confirmation?: string } }>("/api/v1/admin/users/:userId/cloud-library", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const user = await db.query<{ email: string }>("SELECT email FROM users WHERE id=$1", [request.params.userId]);
    if (!user.rows[0]) return reply.code(404).send({ error: "User not found." });
    const expected = `CLEAR ${user.rows[0].email}`;
    if (request.body?.confirmation !== expected) return reply.code(400).send({ error: `Confirmation must exactly match: ${expected}` });
    const deleted = await withAuditTransaction(db, request, "CLOUD_LIBRARY_CLEAR", request.params.userId, undefined, {}, async (client) => {
      const result = await client.query("DELETE FROM library_records WHERE user_id=$1", [request.params.userId]);
      return result.rowCount ?? 0;
    });
    return { ok: true, deleted };
  });

  app.delete<{ Params: { userId: string }; Body: { confirmation?: string } }>("/api/v1/admin/users/:userId", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db || !request.bookstatsAdmin) return;
    if (request.params.userId === request.bookstatsAdmin.id) return reply.code(400).send({ error: "You cannot delete your own administrator account." });
    const user = await db.query<{ email: string }>("SELECT email FROM users WHERE id=$1", [request.params.userId]);
    if (!user.rows[0]) return reply.code(404).send({ error: "User not found." });
    const expected = `DELETE ${user.rows[0].email}`;
    if (request.body?.confirmation !== expected) return reply.code(400).send({ error: `Confirmation must exactly match: ${expected}` });
    // Write the audit entry before deleting the FK target; ON DELETE SET NULL preserves the entry.
    const coverPaths = await listCoverAssetPathsForUser(db, request.params.userId);
    const client = await db.connect();
    try {
      await client.query("BEGIN");
      await writeAudit(client, request, "USER_DELETE", request.params.userId, undefined, { email: user.rows[0].email });
      await client.query("DELETE FROM users WHERE id=$1", [request.params.userId]);
      await client.query("COMMIT");
      await removeStoredCoverFiles(coverPaths);
      return { ok: true };
    } catch (error) {
      await client.query("ROLLBACK");
      throw error;
    } finally { client.release(); }
  });

  app.get<{ Params: { userId: string }; Querystring: { type?: string; q?: string; limit?: string; offset?: string; includeDeleted?: string } }>("/api/v1/admin/users/:userId/records", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const type = ["book", "shelf", "goal"].includes(request.query.type ?? "") ? request.query.type! : "";
    const q = request.query.q?.trim() ?? "";
    const limit = Math.min(200, Math.max(1, Number(request.query.limit ?? 100) || 100));
    const offset = Math.max(0, Number(request.query.offset ?? 0) || 0);
    const includeDeleted = request.query.includeDeleted === "true";
    const result = await db.query<{
      id: string; record_type: string; book_data: unknown; client_updated_at: Date; updated_at: Date; revision: string; deleted_at: Date | null;
    }>(`SELECT id, record_type, book_data, client_updated_at, updated_at, revision, deleted_at
       FROM library_records
       WHERE user_id=$1
         AND ($2 = '' OR record_type=$2)
         AND ($3::boolean OR deleted_at IS NULL)
         AND ($4 = '' OR COALESCE(book_data->>'title','') ILIKE '%' || $4 || '%' OR COALESCE(book_data->>'author','') ILIKE '%' || $4 || '%' OR COALESCE(book_data->>'name','') ILIKE '%' || $4 || '%')
       ORDER BY updated_at DESC LIMIT $5 OFFSET $6`, [request.params.userId, type, includeDeleted, q, limit, offset]);
    const count = await db.query<{ count: string }>(`SELECT count(*) FROM library_records
       WHERE user_id=$1 AND ($2 = '' OR record_type=$2) AND ($3::boolean OR deleted_at IS NULL)
         AND ($4 = '' OR COALESCE(book_data->>'title','') ILIKE '%' || $4 || '%' OR COALESCE(book_data->>'author','') ILIKE '%' || $4 || '%' OR COALESCE(book_data->>'name','') ILIKE '%' || $4 || '%')`,
      [request.params.userId, type, includeDeleted, q]);
    return { total: Number(count.rows[0].count), records: result.rows.map((row) => ({
      id: row.id, recordType: row.record_type, data: row.book_data, clientUpdatedAt: row.client_updated_at.toISOString(),
      serverUpdatedAt: row.updated_at.toISOString(), revision: Number(row.revision), deleted: Boolean(row.deleted_at), deletedAt: row.deleted_at?.toISOString()
    })) };
  });

  app.put<{ Params: { userId: string; recordId: string }; Body: { recordType?: string; data?: unknown } }>("/api/v1/admin/users/:userId/records/:recordId", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const recordType = request.body?.recordType;
    const data = request.body?.data;
    if (!recordType || !["book", "shelf", "goal"].includes(recordType)) return reply.code(400).send({ error: "recordType must be book, shelf, or goal." });
    if (!data || typeof data !== "object" || Array.isArray(data)) return reply.code(400).send({ error: "Record data must be a JSON object." });
    const dataId = (data as { id?: unknown }).id;
    if (dataId !== request.params.recordId) return reply.code(400).send({ error: "The JSON record id must match the record being edited." });
    const revision = await withAuditTransaction(db, request, "LIBRARY_RECORD_UPDATE", request.params.userId, request.params.recordId, { recordType }, async (client) => {
      const result = await client.query<{ revision: string }>(
        `UPDATE library_records SET record_type=$1, book_data=$2::jsonb, client_updated_at=now(), deleted_at=NULL, revision=revision+1, updated_at=now()
         WHERE user_id=$3 AND id=$4 RETURNING revision`, [recordType, JSON.stringify(data), request.params.userId, request.params.recordId]
      );
      if (!result.rows[0]) throw Object.assign(new Error("Record not found."), { statusCode: 404 });
      return Number(result.rows[0].revision);
    });
    return { ok: true, revision };
  });

  app.delete<{ Params: { userId: string; recordId: string }; Body: { confirmation?: string } }>("/api/v1/admin/users/:userId/records/:recordId", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    if (request.body?.confirmation !== "DELETE RECORD") return reply.code(400).send({ error: "Confirmation must exactly match: DELETE RECORD" });
    const revision = await withAuditTransaction(db, request, "LIBRARY_RECORD_DELETE", request.params.userId, request.params.recordId, {}, async (client) => {
      const result = await client.query<{ revision: string }>(
        `UPDATE library_records SET book_data=NULL, client_updated_at=now(), deleted_at=now(), revision=revision+1, updated_at=now()
         WHERE user_id=$1 AND id=$2 RETURNING revision`, [request.params.userId, request.params.recordId]
      );
      if (!result.rows[0]) throw Object.assign(new Error("Record not found."), { statusCode: 404 });
      return Number(result.rows[0].revision);
    });
    return { ok: true, revision };
  });

  app.get<{ Querystring: { limit?: string; offset?: string; q?: string } }>("/api/v1/admin/audit", { preHandler: authenticateAdmin }, async (request, reply) => {
    const db = requirePool(reply); if (!db) return;
    const limit = Math.min(200, Math.max(1, Number(request.query.limit ?? 100) || 100));
    const offset = Math.max(0, Number(request.query.offset ?? 0) || 0);
    const q = request.query.q?.trim() ?? "";
    const result = await db.query<{
      id: string; admin_email: string; action: string; target_user_id: string | null; target_record_id: string | null; details: Record<string, unknown>; ip_address: string | null; created_at: Date;
      target_email: string | null;
    }>(`SELECT a.id::text, a.admin_email, a.action, a.target_user_id, a.target_record_id, a.details, a.ip_address, a.created_at, u.email AS target_email
       FROM admin_audit_log a LEFT JOIN users u ON u.id=a.target_user_id
       WHERE $1 = '' OR a.admin_email ILIKE '%' || $1 || '%' OR a.action ILIKE '%' || $1 || '%' OR COALESCE(u.email,'') ILIKE '%' || $1 || '%'
       ORDER BY a.created_at DESC LIMIT $2 OFFSET $3`, [q, limit, offset]);
    return { entries: result.rows.map((row) => ({
      id: row.id, adminEmail: row.admin_email, action: row.action, targetUserId: row.target_user_id, targetEmail: row.target_email,
      targetRecordId: row.target_record_id, details: row.details, ipAddress: row.ip_address, createdAt: row.created_at.toISOString()
    })) };
  });
}

declare module "fastify" {
  interface FastifyRequest {
    bookstatsAdmin?: AdminActor;
    bookstatsAdminTokenHash?: string;
  }
}
