BookStats v1.1.0 patch
======================

Baseline required: BookStats v1.0.4 source.

Apply by copying the contents of this archive over the project root, preserving paths and replacing existing files. This patch does not replace .gitignore or any local environment file.

Major changes:
- Incremental per-record cloud synchronization backed by a persistent local outbox.
- One-book edits upload that book only; repeated edits coalesce to the latest state.
- Large backlogs are split into batches of at most 100 records and approximately 900 KiB per request.
- Deletions remain durable tombstones and successful/stale retries are explicitly acknowledged by the server.
- The sync cursor is saved after each successful batch, so later failures retain only unacknowledged work for retry.
- v1.0.x upgrades seed the outbox only from records newer than the device's last successful sync.
- HTTP 413 sync failures now explain that local work remains saved/queued instead of reporting a generic connection error.
- Account UI remains intentionally simple: no pending-change counters or backend sync diagnostics are exposed.
- Catalog cover choices filter obvious placeholder URLs, broken/non-image results, nearly blank white images, and common "Image Not Available" cards when browser inspection is possible.
- PWA shell cache generation is bumped for v1.1.0.

Storage/deployment:
- No PostgreSQL migration is required.
- IndexedDB upgrades locally from schema 5 to schema 6 by adding syncOutbox.
- Desktop SQLite creates sync_outbox automatically on startup.
- Keep the current 25 MB NGINX/Fastify ceiling as an emergency limit; normal v1.1 sync batches target <= ~900 KiB.
- Deploy the v1.1 server/web build before publishing the v1.1 desktop updater. v1.0.x clients remain compatible with the v1.1 server.

After applying:
  ./set-version.sh 1.1.0
  npm install
  npm run typecheck
  npm test
  npm run build

Running set-version.sh after overlaying the patch is intentional: if your cleaned repository now contains package-lock.json, it updates the lockfile's workspace version metadata too.
