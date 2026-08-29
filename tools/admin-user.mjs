#!/usr/bin/env node
import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import pg from "pg";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const { Client } = pg;

function usage(code = 0) {
  console.log(`BookStats administrator role utility

Usage:
  node tools/admin-user.mjs grant user@example.com
  node tools/admin-user.mjs revoke user@example.com
  node tools/admin-user.mjs status user@example.com

Administrator promotion is intentionally server-side only. The web API does not expose a role-promotion endpoint.`);
  process.exit(code);
}

const [command, rawEmail] = process.argv.slice(2);
if (command === "--help" || command === "-h" || command === "help") usage(0);
if (!command || !rawEmail || !["grant", "revoke", "status"].includes(command)) usage(2);
const email = rawEmail.trim().toLowerCase();
if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
  console.error("Enter a valid BookStats account email.");
  process.exit(2);
}

function readDatabaseUrl() {
  if (process.env.DATABASE_URL?.trim()) return process.env.DATABASE_URL.trim();
  const envPath = resolve(root, ".env");
  if (!existsSync(envPath)) return "";
  for (const rawLine of readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const index = line.indexOf("=");
    if (index < 1 || line.slice(0, index).trim() !== "DATABASE_URL") continue;
    let value = line.slice(index + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) value = value.slice(1, -1);
    return value;
  }
  return "";
}

const databaseUrl = readDatabaseUrl();
if (!databaseUrl) {
  console.error("DATABASE_URL is not set. Run this from the BookStats server directory or export DATABASE_URL.");
  process.exit(2);
}

const client = new Client({ connectionString: databaseUrl });

try {
  await client.connect();

  const existing = await client.query(
    "SELECT id, email, display_name, role, disabled_at FROM users WHERE lower(email) = lower($1)",
    [email]
  );

  if (existing.rowCount !== 1) {
    console.error(`No BookStats account was found for ${email}. No role was changed.`);
    process.exitCode = 1;
  } else if (command === "status") {
    console.table(existing.rows);
  } else {
    const role = command === "grant" ? "admin" : "user";
    const updated = await client.query(
      `UPDATE users
          SET role = $1, updated_at = now()
        WHERE lower(email) = lower($2)
      RETURNING id, email, display_name, role, disabled_at`,
      [role, email]
    );
    console.table(updated.rows);
    console.log(`${updated.rows[0].email} is now ${role}.`);
  }
} catch (error) {
  console.error(`Administrator role update failed: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
} finally {
  await client.end().catch(() => {});
}
