#!/usr/bin/env node

import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const scriptDir = resolve(fileURLToPath(new URL(".", import.meta.url)));
const projectRoot = resolve(scriptDir, "..");

function usage(exitCode = 0) {
  console.log(`BookStats one-time condition tag migration

Usage:
  node tools/bookstats-migrate-condition-tags.mjs --email you@example.com
  node tools/bookstats-migrate-condition-tags.mjs --email you@example.com --apply

Options:
  --email <address>   BookStats account to update (required)
  --apply             Perform the migration. Without this flag, the script is dry-run only.
  --help              Show this help text.

Behavior:
  • Recognizes exact tags (case-insensitive): New, Like New, Very Good, Good, Acceptable, Poor.
  • Copies the recognized tag value into book.condition.
  • Removes only recognized condition tags; every other tag is preserved in its existing order.
  • If a book has conflicting condition tags, it is skipped and reported.
  • If book.condition is already set to a different value, it is skipped and reported.
  • Advances BookStats sync timestamps/revision so the server-side change is pulled by clients.

Safety:
  Run once without --apply first. Close BookStats on other devices while running --apply,
  and make a normal BookStats/database backup before applying the migration.
`);
  process.exit(exitCode);
}

const args = process.argv.slice(2);
let email = "";
let apply = false;
for (let i = 0; i < args.length; i += 1) {
  const arg = args[i];
  if (arg === "--help" || arg === "-h") usage(0);
  if (arg === "--apply") {
    apply = true;
    continue;
  }
  if (arg === "--email") {
    email = args[i + 1]?.trim() ?? "";
    i += 1;
    continue;
  }
  console.error(`Unknown argument: ${arg}`);
  usage(2);
}

if (!email) {
  console.error("--email is required so the script cannot accidentally modify another account.");
  usage(2);
}

function readDatabaseUrl() {
  if (process.env.DATABASE_URL?.trim()) return process.env.DATABASE_URL.trim();

  // Normal use is tools/bookstats-migrate-condition-tags.mjs inside the BookStats repo.
  // Also try the current working directory to make one-off use from another location convenient.
  const candidates = [resolve(projectRoot, ".env"), resolve(process.cwd(), ".env")];
  for (const envPath of candidates) {
    if (!existsSync(envPath)) continue;
    for (const rawLine of readFileSync(envPath, "utf8").split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;
      const index = line.indexOf("=");
      if (index < 1) continue;
      if (line.slice(0, index).trim() !== "DATABASE_URL") continue;
      let value = line.slice(index + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }
      if (value) return value;
    }
  }
  return "";
}

const databaseUrl = readDatabaseUrl();
if (!databaseUrl) {
  console.error("DATABASE_URL is not set. Export it or place it in the BookStats .env file.");
  process.exit(2);
}

const psqlCheck = spawnSync("psql", ["--version"], { encoding: "utf8" });
if (psqlCheck.error?.code === "ENOENT") {
  console.error("psql was not found. Install PostgreSQL client tools or run this on the BookStats server.");
  process.exit(2);
}
if (psqlCheck.status !== 0) {
  console.error(psqlCheck.stderr || "Could not run psql.");
  process.exit(psqlCheck.status ?? 2);
}

function runPsql(sql, { capture = false } = {}) {
  const result = spawnSync(
    "psql",
    [databaseUrl, "-X", "-v", "ON_ERROR_STOP=1", "-v", `target_email=${email}`],
    {
      input: sql,
      encoding: "utf8",
      stdio: capture ? ["pipe", "pipe", "pipe"] : ["pipe", "inherit", "inherit"]
    }
  );
  if (result.error) throw result.error;
  if (result.status !== 0) {
    if (capture) {
      if (result.stdout) process.stdout.write(result.stdout);
      if (result.stderr) process.stderr.write(result.stderr);
    }
    process.exit(result.status ?? 1);
  }
  return result.stdout ?? "";
}

const accountRows = runPsql(
  `COPY (SELECT id::text || '|' || email FROM users WHERE lower(email) = lower(:'target_email')) TO STDOUT;\n`,
  { capture: true }
).trim().split(/\r?\n/).filter(Boolean);

if (accountRows.length === 0) {
  console.error(`No BookStats user was found for ${email}. No changes were made.`);
  process.exit(1);
}
if (accountRows.length > 1) {
  console.error(`More than one BookStats user matched ${email}. No changes were made.`);
  process.exit(1);
}

const [userId, canonicalEmail] = accountRows[0].split("|");
console.log(`BookStats condition-tag migration for ${canonicalEmail} (${userId})`);
console.log(apply ? "MODE: APPLY\n" : "MODE: DRY RUN (no data will be changed)\n");

const canonicalCase = `CASE lower(btrim(tag))
  WHEN 'new' THEN 'New'
  WHEN 'like new' THEN 'Like New'
  WHEN 'very good' THEN 'Very Good'
  WHEN 'good' THEN 'Good'
  WHEN 'acceptable' THEN 'Acceptable'
  WHEN 'poor' THEN 'Poor'
  ELSE NULL
END`;

const migrationSql = `
\\pset pager off
\\set ON_ERROR_STOP on
BEGIN;

CREATE TEMP TABLE bookstats_condition_tag_candidates ON COMMIT DROP AS
SELECT
  lr.user_id,
  lr.id,
  COALESCE(NULLIF(lr.book_data->>'title', ''), '(untitled)') AS title,
  NULLIF(btrim(lr.book_data->>'condition'), '') AS existing_condition,
  recognized.conditions,
  recognized.conditions[1] AS chosen_condition,
  remaining.tags AS remaining_tags,
  (
    cardinality(recognized.conditions) = 1
    AND (
      NULLIF(btrim(lr.book_data->>'condition'), '') IS NULL
      OR lower(btrim(lr.book_data->>'condition')) = lower(recognized.conditions[1])
    )
  ) AS safe_to_migrate,
  CASE
    WHEN cardinality(recognized.conditions) > 1
      THEN 'multiple condition tags: ' || array_to_string(recognized.conditions, ', ')
    WHEN NULLIF(btrim(lr.book_data->>'condition'), '') IS NOT NULL
         AND lower(btrim(lr.book_data->>'condition')) <> lower(recognized.conditions[1])
      THEN 'existing condition is ' || quote_literal(lr.book_data->>'condition')
    ELSE NULL
  END AS conflict_reason
FROM library_records lr
JOIN users u ON u.id = lr.user_id
CROSS JOIN LATERAL (
  SELECT array_agg(DISTINCT canonical ORDER BY canonical) AS conditions
  FROM (
    SELECT ${canonicalCase} AS canonical
    FROM jsonb_array_elements_text(
      CASE
        WHEN jsonb_typeof(lr.book_data->'tags') = 'array' THEN lr.book_data->'tags'
        ELSE '[]'::jsonb
      END
    ) AS condition_tags(tag)
  ) mapped
  WHERE canonical IS NOT NULL
) recognized
CROSS JOIN LATERAL (
  SELECT COALESCE(jsonb_agg(to_jsonb(tag) ORDER BY ord), '[]'::jsonb) AS tags
  FROM jsonb_array_elements_text(
    CASE
      WHEN jsonb_typeof(lr.book_data->'tags') = 'array' THEN lr.book_data->'tags'
      ELSE '[]'::jsonb
    END
  ) WITH ORDINALITY AS all_tags(tag, ord)
  WHERE (${canonicalCase}) IS NULL
) remaining
WHERE lower(u.email) = lower(:'target_email')
  AND lr.record_type = 'book'
  AND lr.deleted_at IS NULL
  AND lr.book_data IS NOT NULL
  AND cardinality(recognized.conditions) > 0;

\\echo 'Summary'
SELECT
  count(*) AS books_with_condition_tags,
  count(*) FILTER (WHERE safe_to_migrate) AS safe_to_migrate,
  count(*) FILTER (WHERE NOT safe_to_migrate) AS conflicts_skipped
FROM bookstats_condition_tag_candidates;

\\echo ''
\\echo 'Safe migrations by condition'
SELECT chosen_condition AS condition, count(*) AS books
FROM bookstats_condition_tag_candidates
WHERE safe_to_migrate
GROUP BY chosen_condition
ORDER BY CASE chosen_condition
  WHEN 'New' THEN 1
  WHEN 'Like New' THEN 2
  WHEN 'Very Good' THEN 3
  WHEN 'Good' THEN 4
  WHEN 'Acceptable' THEN 5
  WHEN 'Poor' THEN 6
  ELSE 99 END;

\\echo ''
\\echo 'Conflicts (these are never changed)'
SELECT id, title, existing_condition, array_to_string(conditions, ', ') AS condition_tags, conflict_reason
FROM bookstats_condition_tag_candidates
WHERE NOT safe_to_migrate
ORDER BY title, id;

${apply ? `
WITH stamp AS (
  SELECT
    now() AS ts,
    to_char(now() AT TIME ZONE 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"') AS iso
)
UPDATE library_records lr
SET
  book_data = jsonb_set(
    jsonb_set(
      jsonb_set(lr.book_data, '{condition}', to_jsonb(c.chosen_condition), true),
      '{tags}', c.remaining_tags, true
    ),
    '{updatedAt}', to_jsonb(stamp.iso), true
  ),
  client_updated_at = stamp.ts,
  revision = lr.revision + 1,
  updated_at = stamp.ts
FROM bookstats_condition_tag_candidates c
CROSS JOIN stamp
WHERE lr.user_id = c.user_id
  AND lr.id = c.id
  AND c.safe_to_migrate;

\\echo ''
\\echo 'Applied migration.'
SELECT count(*) AS books_updated
FROM bookstats_condition_tag_candidates
WHERE safe_to_migrate;
COMMIT;
` : `
\\echo ''
\\echo 'Dry run complete. Nothing was changed.'
ROLLBACK;
`}
`;

runPsql(migrationSql);

if (apply) {
  console.log("\nDone. The next BookStats cloud sync should pull the changed condition/tags into the client.");
} else {
  console.log(`\nReview the counts/conflicts above. To apply exactly these safe migrations, run:\n  node tools/bookstats-migrate-condition-tags.mjs --email ${email} --apply`);
}
