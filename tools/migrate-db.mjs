#!/usr/bin/env node
import { readFileSync, readdirSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const envPath = resolve(root, ".env");
let databaseUrl = process.env.DATABASE_URL;

if (!databaseUrl && existsSync(envPath)) {
  for (const rawLine of readFileSync(envPath, "utf8").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const index = line.indexOf("=");
    if (index < 1) continue;
    const key = line.slice(0, index).trim();
    let value = line.slice(index + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) value = value.slice(1, -1);
    if (key === "DATABASE_URL") databaseUrl = value;
  }
}

if (!databaseUrl) {
  console.error("DATABASE_URL is not set. Add it to .env or export it in your shell.");
  process.exit(2);
}

const migrationsDir = resolve(root, "database", "migrations");
const migrations = readdirSync(migrationsDir).filter((name) => name.endsWith(".sql")).sort();
for (const migration of migrations) {
  console.log(`==> ${migration}`);
  const result = spawnSync("psql", [databaseUrl, "-v", "ON_ERROR_STOP=1", "-f", resolve(migrationsDir, migration)], { stdio: "inherit" });
  if (result.error?.code === "ENOENT") {
    console.error("psql was not found. Install PostgreSQL client tools first.");
    process.exit(2);
  }
  if (result.status !== 0) process.exit(result.status ?? 1);
}
console.log("BookStats database migrations complete.");
