#!/usr/bin/env node
import { existsSync, readFileSync } from "node:fs";
import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import pg from "pg";
import { absoluteCoverAssetPath, archiveCoverForUser } from "../apps/server/dist/covers.js";

const { Pool } = pg;
const scriptDir = resolve(fileURLToPath(new URL(".", import.meta.url)));
const projectRoot = resolve(scriptDir, "..");

function usage(exitCode = 0) {
  console.log(`BookStats selected-cover migration / repair (v1.0.1)

Usage:
  node tools/migrate-cover-assets.mjs --email you@example.com
  node tools/migrate-cover-assets.mjs --email you@example.com --apply

Options:
  --email <address>   BookStats account to migrate/repair (required)
  --apply             Store missing covers and update book records. Default is dry-run.
  --concurrency <n>   Parallel cover downloads, 1-8 (default: 3)
  --help              Show this help text.

Behavior:
  • Archives pre-v1.0.1 selected covers that do not have a coverAssetId yet.
  • Checks existing coverAssetId records and reports archived files that are missing.
  • With --apply, repairs a missing archived file from the book's retained source URL when possible.
  • Existing healthy cover assets are never replaced.
  • Existing custom data-URL covers are moved out of JSON into the server asset store.
  • Failed downloads/repairs are left recoverable and reported so the script can be rerun safely.

Safety:
  Run once without --apply first. Close BookStats on other devices and keep a normal BookStats/database backup before applying.
`);
  process.exit(exitCode);
}

const args = process.argv.slice(2);
let email = "";
let apply = false;
let concurrency = 3;
for (let i = 0; i < args.length; i += 1) {
  const arg = args[i];
  if (arg === "--help" || arg === "-h") usage(0);
  if (arg === "--apply") { apply = true; continue; }
  if (arg === "--email") { email = args[++i]?.trim().toLowerCase() ?? ""; continue; }
  if (arg === "--concurrency") { concurrency = Number(args[++i]); continue; }
  console.error(`Unknown argument: ${arg}`); usage(2);
}
if (!email) { console.error("--email is required."); usage(2); }
if (!Number.isInteger(concurrency) || concurrency < 1 || concurrency > 8) { console.error("--concurrency must be an integer from 1 to 8."); process.exit(2); }

function readDatabaseUrl() {
  if (process.env.DATABASE_URL?.trim()) return process.env.DATABASE_URL.trim();
  for (const envPath of [resolve(projectRoot, ".env"), resolve(process.cwd(), ".env")]) {
    if (!existsSync(envPath)) continue;
    for (const rawLine of readFileSync(envPath, "utf8").split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;
      const index = line.indexOf("=");
      if (index < 1 || line.slice(0, index).trim() !== "DATABASE_URL") continue;
      let value = line.slice(index + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) value = value.slice(1, -1);
      if (value) return value;
    }
  }
  return "";
}

function selectedRepairSource(book) {
  const values = [book?.coverSourceUrl, book?.coverUrl];
  return values.find((value) => typeof value === "string" && (/^https?:\/\//i.test(value) || value.startsWith("data:"))) ?? "";
}

const databaseUrl = readDatabaseUrl();
if (!databaseUrl) { console.error("DATABASE_URL is not set. Export it or place it in the BookStats .env file."); process.exit(2); }

const pool = new Pool({ connectionString: databaseUrl });
try {
  const userResult = await pool.query("SELECT id, email FROM users WHERE lower(email)=lower($1) LIMIT 1", [email]);
  const user = userResult.rows[0];
  if (!user) { console.error(`No BookStats account found for ${email}.`); process.exitCode = 2; }
  else {
    const records = await pool.query(
      `SELECT id, book_data
         FROM library_records
        WHERE user_id=$1 AND record_type='book' AND deleted_at IS NULL
          AND (COALESCE(book_data->>'coverUrl','')<>'' OR COALESCE(book_data->>'coverAssetId','')<>'')
        ORDER BY id`, [user.id]
    );

    const assetIds = [...new Set(records.rows.map((row) => String(row.book_data?.coverAssetId ?? "")).filter(Boolean))];
    const assetRows = assetIds.length
      ? await pool.query("SELECT id::text, storage_path FROM cover_assets WHERE user_id=$1 AND id::text = ANY($2::text[])", [user.id, assetIds])
      : { rows: [] };
    const assets = new Map(assetRows.rows.map((row) => [String(row.id), row]));

    const unarchived = records.rows.filter((row) => !row.book_data?.coverAssetId && row.book_data?.coverUrl);
    const custom = unarchived.filter((row) => String(row.book_data?.coverUrl ?? "").startsWith("data:")).length;
    const remote = unarchived.length - custom;
    const broken = [];
    for (const row of records.rows) {
      const assetId = String(row.book_data?.coverAssetId ?? "");
      if (!assetId) continue;
      const asset = assets.get(assetId);
      if (!asset) {
        broken.push({ row, reason: "asset metadata is missing" });
        continue;
      }
      try { await stat(absoluteCoverAssetPath(asset.storage_path)); }
      catch { broken.push({ row, reason: "archived image file is missing" }); }
    }

    console.log(`BookStats cover migration/repair for ${user.email}`);
    console.log(`  covers needing first archive: ${unarchived.length}`);
    console.log(`    remote URLs:                ${remote}`);
    console.log(`    custom embedded covers:     ${custom}`);
    console.log(`  archived covers needing repair: ${broken.length}`);

    const tasks = [
      ...unarchived.map((row) => ({ row, kind: "archive", source: String(row.book_data?.coverUrl ?? "") })),
      ...broken.map(({ row, reason }) => ({ row, kind: "repair", source: selectedRepairSource(row.book_data), reason }))
    ];

    if (!apply) {
      if (broken.length) {
        const repairable = broken.filter(({ row }) => Boolean(selectedRepairSource(row.book_data))).length;
        console.log(`  broken assets repairable from retained source: ${repairable}/${broken.length}`);
      }
      console.log("\nDRY RUN only. Re-run with --apply to archive/repair the covers listed above.");
    } else if (!tasks.length) {
      console.log("\nNothing to migrate or repair.");
    } else {
      let nextIndex = 0;
      let migrated = 0;
      let repaired = 0;
      let failed = 0;
      const failures = [];

      async function worker() {
        while (true) {
          const index = nextIndex++;
          if (index >= tasks.length) return;
          const task = tasks[index];
          const row = task.row;
          const book = row.book_data ?? {};
          if (!task.source) {
            failed += 1;
            failures.push({ id: row.id, title: book.title ?? "Untitled", error: `${task.reason ?? "Archived cover is unavailable"}; no retained source URL is available for server-side repair.` });
            continue;
          }
          try {
            const asset = await archiveCoverForUser(pool, user.id, task.source);
            const sourceIsRemote = /^https?:\/\//i.test(task.source);
            const updated = {
              ...book,
              coverAssetId: asset.id,
              coverAssetToken: asset.accessToken,
              coverSourceUrl: sourceIsRemote ? task.source : (book.coverSourceUrl ?? asset.sourceUrl),
              coverArchivePending: undefined,
              coverUrl: task.source.startsWith("data:") ? undefined : (book.coverUrl ?? task.source)
            };
            for (const key of Object.keys(updated)) if (updated[key] === undefined) delete updated[key];
            await pool.query(
              `UPDATE library_records
                  SET book_data=$1::jsonb, client_updated_at=now(), revision=revision+1, updated_at=now()
                WHERE user_id=$2 AND id=$3`,
              [JSON.stringify(updated), user.id, row.id]
            );
            if (task.kind === "repair") repaired += 1; else migrated += 1;
            const complete = migrated + repaired + failed;
            if (complete % 50 === 0 || complete === tasks.length) console.log(`  processed ${complete}/${tasks.length} (${migrated} archived, ${repaired} repaired, ${failed} failed)`);
          } catch (error) {
            failed += 1;
            failures.push({ id: row.id, title: book.title ?? "Untitled", error: error instanceof Error ? error.message : String(error) });
            console.error(`  FAILED ${book.title ?? row.id}: ${failures.at(-1).error}`);
          }
        }
      }

      await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, () => worker()));
      console.log(`\nCover maintenance complete: ${migrated} archived, ${repaired} repaired, ${failed} failed.`);
      if (failures.length) {
        console.log("Failed records were left recoverable. It is safe to rerun this command later.");
        console.log("If a custom cover has no retained source, open it on a device that still has the local cache and save the book again.");
        for (const failure of failures.slice(0, 25)) console.log(`  ${failure.id}  ${failure.title}: ${failure.error}`);
        if (failures.length > 25) console.log(`  ...and ${failures.length - 25} more failures.`);
      }
    }
  }
} finally {
  await pool.end();
}
