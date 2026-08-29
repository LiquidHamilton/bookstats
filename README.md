# BookStats

BookStats is a local-first personal library manager and reading-statistics application built primarily for the web, with optional Windows and macOS desktop packaging through Tauri. It is a ground-up TypeScript/React rewrite of the earlier Python BookStats prototype.

**Current version:** `1.1.1` — incremental cloud synchronization, safer retry batching, and cleaner catalog cover selection

## Mobile Home Screen app

BookStats can now be installed from the web build on iPhone/iPad and Android. The web manifest uses the same brown open-book mark as the BookStats header and desktop application, launches with `display: standalone`, respects mobile safe areas, and registers a small service-worker app shell. The worker intentionally excludes `/bookstats/api/` and does not cache account/library API responses.

On iPhone/iPad, BookStats shows a dismissible first-use guide for Safari's **More → Share → Add to Home Screen** flow and reminds the user to leave **Open as Web App** enabled. The prompt never appears once BookStats is already running in standalone mode. Android uses the browser's native install prompt when available and falls back to the browser-menu **Install app / Add to Home screen** flow.

There is one important iOS local-data distinction: the browser and an installed Home Screen web app have separate local storage. A signed-in user can simply sign in again inside the installed app and synchronize. A local-only Safari user should export/back up the library before installing and import it into the Home Screen app, or create an account and sync first. BookStats calls this out in the iPhone install prompt rather than allowing the installed copy to look unexpectedly empty.

## Desktop auto-update

Starting with v1.0.3, the Tauri desktop applications use signed updates. If the BookStats API reports a newer required application version, the web build still offers **Refresh BookStats**, while macOS/Windows offers **Update BookStats**. The desktop client checks the signed Tauri update manifest, downloads the platform package, installs it, and relaunches BookStats.

Tauri updater signatures require a long-lived signing key. Generate it once on the trusted release machine:

```bash
./tools/setup-updater-key.sh
```

The helper stores the private key under `$HOME/.config/bookstats/` by default and writes only the corresponding public key into `apps/client/src-tauri/tauri.conf.json`. Back up the private key separately and never commit or upload it. Before each desktop release build, export the variables printed by the helper:

Because the keypair lives outside the source tree, replacing BookStats with a future full source archive is safe: if `tauri.conf.json` contains the updater placeholder, `export.sh` automatically re-embeds `$HOME/.config/bookstats/updater.key.pub`. If you store the public key elsewhere, set `BOOKSTATS_UPDATER_PUBLIC_KEY_PATH` to that `.pub` file.

```bash
export TAURI_SIGNING_PRIVATE_KEY="$HOME/.config/bookstats/updater.key"
export TAURI_SIGNING_PRIVATE_KEY_PASSWORD="YOUR_KEY_PASSWORD"
./export.sh all
```

Desktop release builds produce the normal downloadable installers plus signed updater files under `dist/releases/updater/` and a convenient `BookStats-Updater-vX.Y.Z.zip` deployment bundle. The updater manifest URL is platform-specific under:

```text
https://kylecarroll.com/downloads/bookstats/latest-<target>-<arch>.json
```

The server deployment helper publishes the updater bundle to `/var/www/kylecarroll/downloads/bookstats/` without deleting older signed binaries. It also supports an updater-only deployment when the desktop packages are built separately from the web/server release.

**Bootstrap note:** v1.0.2 desktop builds do not contain updater code, so v1.0.2 → v1.0.3 is the one manual desktop upgrade. Once v1.0.3 is installed, later compatible releases can use the in-app updater.

## What works now

### Library and reading

- Cover-grid and sortable data-table library views
- Add, edit and delete books
- Search title, author, series, ISBN, publisher, genre, condition, tags and shelf names
- Reading-status and shelf filters plus an advanced all/any rule builder that can be saved as a smart shelf
- One primary reading status plus any number of user-defined regular shelves
- Rule-based smart shelves that update automatically, with all/any rule matching and no default shelves
- Shelf creation/editing, persistent custom shelf ordering, assignment from the book editor and shelf-level statistics
- A read-only book detail view with cover, metadata, user-entered condition, review, shelves, tags and complete reading history
- Per-device Hide / Show table-column preferences
- Series name + volume
- Structured reading sessions with start/finish dates, current-page progress, rereads and per-session notes
- Half-star ratings from 0.5–5.0
- Description, personal review and private notes
- Portable BookStats JSON export/import plus separate recovery snapshots under **Tools**
- Goodreads CSV import with status, rating, dates, ownership, reviews and shelf mapping
- LibraryThing JSON import with collections→shelves, ratings, read dates, tags, genres and edition metadata
- Random unread-book picker
- Bulk selection/editing for status, condition, ownership, regular shelves, tags, export and deletion
- Import preview, conservative matching and undoable recent-import history
- Strong duplicate reconciliation in Library Health with side-by-side edition comparison, canonical ISBN matching, merge safety backups, and persistent “keep separate” decisions
- Mainline series-completion tracking with automatic alternate-language/edition cleanup plus manually editable expected counts and completion checklists
- Lending dashboard with borrower, due-date, overdue and returned-loan history
- Camera ISBN/barcode scanning with automatic exact-ISBN catalog lookup and a cross-browser ZXing fallback
- Incremental large-library rendering on web so thousands of books do not create thousands of DOM rows/cards at once
- Targeted Edit Book repair actions for catalog covers and missing metadata without overwriting fields that are already populated
- Built-in Report a Bug / Suggest a Feature screen with safe build diagnostics
- Categorized statistics for overview, goals, reading history, authors, collection, series and ratings
- User-created date-range goals for books, pages, rereads, new-to-you authors and owned books read, with pace/projection tracking

### Lending, scanning and collection completion

The **Lending** view records who currently has an owned book, when it was loaned, an optional due date and notes. Returning a book closes the active loan but keeps the history on that Book record. Due-soon and overdue counts are derived locally and synchronize with the rest of the book when an account is enabled.

**Scan ISBN** opens the camera and reads common book barcodes. A successful ISBN immediately opens Add Book with that identifier prefilled and triggers the existing edition-first metadata lookup. BookStats uses the browser's native barcode detector when present and lazily loads ZXing when it is not, so scanning can work on a wider range of browsers without increasing the normal startup bundle. Camera access is requested only after the user opens the scanner.

**Statistics → Series** treats collection completion as an owned-mainline question. Provider feeds sometimes contain many translated or alternate work records; BookStats uses the known primary-book count and numeric series positions to collapse those variants to one likely representative per mainline position. When automatic data is still wrong, **Edit completion** lets the user set the expected count, include/exclude catalog entries, add a missing manual entry, or reset to automatic detection. The UI intentionally does not expose provider branding in this completion workflow—the important result is the sequence the user considers canonical.

### Administration

BookStats includes a separate administrator web application intended for server operators. It is built and deployed under `/bookstats/admin/` but is not linked, imported, or referenced anywhere in the normal BookStats web/desktop client UI. The administrator bundle has its own login storage and can only use `/api/v1/admin/*` endpoints after the server independently verifies that the account has the `admin` role.

The administrator console provides a dashboard with user/library/server statistics, searchable user accounts, account detail and cloud-record inspection, profile corrections, account disable/enable, session invalidation, forced password-reset delivery, cloud-library reset, permanent account deletion, targeted synchronized-record JSON repair, metadata/email/database health, and an append-only administrator audit log. Destructive operations require explicit confirmation and every administrative write is audited. Raw SQL, shell commands, systemd/NGINX controls, and administrator-role promotion are intentionally not exposed in the browser.

Administrator roles are granted or revoked only from the server command line after migrations have been applied:

```bash
cd /opt/bookstats
sudo -u bookstats npm run admin:user -- grant admin@example.com
sudo -u bookstats npm run admin:user -- status admin@example.com
sudo -u bookstats npm run admin:user -- revoke admin@example.com
```


### Book lookup

The Add/Edit Book window has one unified catalog search for ISBN, title, or title + author. The BookStats API combines configured metadata providers rather than making the user choose a source:

- **Google Books** — edition metadata, publisher descriptions, identifiers, page counts and series metadata
- **Hardcover** — work/series relationships, ordered series catalogs and edition enrichment
- **Open Library** — open fallback metadata and additional cover/edition information

**ISBN lookup is edition-first.** When the user enters an ISBN, BookStats asks every configured provider for that identifier and only accepts results whose ISBN is the same ISBN-10/ISBN-13 edition. It does not silently substitute a different printing or format. Edition-specific fields such as publisher, publication year, page count, binding/format, ISBN and cover therefore come from the matched edition when a provider exposes them. Work-level fields such as description and series relationships can still be enriched from the associated work without replacing those edition facts.

Title/author search is broader and can return multiple editions. Results from the providers are normalized, deduplicated and ranked by query relevance before provider confidence, so exact-title matches stay ahead of box sets/omnibuses and relevant Open Library editions are not crowded out by higher-confidence providers. BookStats then merges useful fields with field-level provenance. Exact ISBN matches receive the strongest confidence; fuzzy title matches never become exact-edition records automatically.

Catalog-managed fields remain editable and use simple blank-only refill behavior. Metadata lookup and targeted repair fill fields that are empty and leave any field that already contains data untouched. Clearing a catalog field makes it eligible to be filled on the next lookup; there is no separate saved-field or manual-override state. Provider references and per-field provenance are still retained for diagnostics and future refresh work.

Alternate covers are aggregated across the available providers and presented in one cover picker. Exact-ISBN artwork is shown first and title/author edition alternatives follow. Once a signed-in user chooses a cover, BookStats archives the selected image on its own server so the library does not permanently depend on a third-party cover URL.


### Covers

BookStats supports catalog/URL covers, alternate catalog cover selection, and user-selected local image files. Local uploads are resized before storage. Every device still keeps a local cover cache when possible, but verified accounts now also have a durable server-side cover archive. The display priority is the device-local cache, then the definitive BookStats archive, then the retained original source as a graceful fallback if either cached layer is temporarily unavailable. The original external URL remains provenance and a repair/re-archival source when it is still available.

Cloud cover images live as files under BookStats server storage while PostgreSQL stores ownership, content hash, MIME/size metadata, provenance, and an opaque access token. Custom uploaded images are archived the same way and are removed from synchronized JSON after archival succeeds. Users without an account remain fully local-first: their custom image data and remote cover cache stay on that browser/desktop until an account is available.

### Shelves, statuses and imports

A **status** describes the book's single reading lifecycle state (`Want to Read`, `Currently Reading`, `Read`, `Did Not Finish`, etc.). **Condition** is a separate user-entered assessment of the physical copy (`New`, `Like New`, `Very Good`, `Good`, `Acceptable`, or `Poor`) and is never filled or overwritten by catalog providers. **Regular shelves** are independent user-created groupings, and one book can belong to many of them at the same time. **Smart shelves** are saved rule sets evaluated dynamically against the library; they are never manually assigned to books. Smart rules support grouped Boolean logic, so a shelf can express combinations such as `(Status = Read AND Ownership = Not owned) OR (Status = Did Not Finish AND Ownership = Not owned)`. Rules inside a group and the groups themselves can independently use AND/OR. Shelves can be rearranged in Add & manage shelves, and that order is used throughout the UI. BookStats creates no default shelves, so users decide whether concepts such as Favorites are regular shelves or whether combinations such as Owned But Unread are smart shelves.

The external importers are designed to be repeatable. Goodreads matches its stable book ID first, then ISBN, then normalized title + author. LibraryThing is intentionally collector-safe: its `books_id` identifies one catalog entry/copy, so only the same `books_id` is treated as a repeat import; separate LibraryThing entries remain separate even when they share an ISBN or title + author. When LibraryThing supplies several translated series labels for the same entry, BookStats prefers a likely English series label when one is available and then uses collection-wide consistency to choose among plausible candidates. Existing BookStats fields generally win when a repeated import would otherwise overwrite manual edits; new reading history, tags, source IDs and shelf assignments are merged. The one exception is an unmodified series value that originally came from the same LibraryThing entry: a repeat import may refresh that importer-owned series choice so improved English-series selection can repair earlier imports. A series field explicitly edited in BookStats remains protected. Imports open a preview before they change the library. Ambiguous non-LibraryThing matches stay separate rather than being merged automatically, and recent imports can be undone when the affected records have not been edited afterward.

### Reading sessions, progress and goals

Reading history is stored as sessions rather than only completion dates. A session can record when a book was started, when it was finished, current-page progress and an optional note. Legacy `readDates` are still understood and exported for backward compatibility.

The **Statistics → Goals** area lets the user create any number of date-range goals. Supported metrics are books read, pages read, rereads, new-to-you authors and owned books read. BookStats creates no default challenge. Progress is calculated from the library automatically and, while a goal is active, shows time elapsed, days remaining, whether the current pace is on target and a projected finish value.

### Advanced filters and bulk editing

The Library's advanced filter builder uses the same grouped Boolean rule language as smart shelves, including status, condition, ownership, ratings, title/author/series text, format, genre, tags, pages, publication year, read count, last-read date and date-added rules. A temporary filter can be saved directly as a smart shelf. Grid and table views also support multi-selection for bulk status/condition/ownership changes, regular-shelf and tag changes, subset export and deletion.

### Export vs. backup

These intentionally solve different problems:

- **BookStats Export** is portable, merge-friendly data. Importing an export adds new records and merges matching records into the current library; unrelated books already in the destination library stay there. Use it for interchange, moving data, selected-book exports or combining libraries.
- **BookStats Backup** is a recovery snapshot. Restoring a backup replaces the local books, shelves, goals and saved library-view preferences with the state captured in that snapshot. Use it when you need to roll BookStats back to a known point in time.

Backups preserve reading sessions and cover references but omit passwords, account authentication tokens, server credentials and recreatable device-local cover caches. Server operators should back up the PostgreSQL database together with the durable cover asset directory; a future asset-inclusive portable archive can build on the same cover IDs. BookStats can keep up to five local safety snapshots and automatically creates them before high-impact actions such as large imports, import undo, duplicate merges, bulk deletion and restores. A manual backup can also be downloaded as a file.

### Library intelligence

The cleanup tool now includes a **Library Health** score based only on objective metadata completeness: covers, descriptions, ISBNs, page counts, publication years and series positions where applicable. Personal choices such as ratings, reviews and notes never lower the score. The same tool retains user-reviewed duplicate merging.

Statistics also include series progress, monthly/yearly reading and page trends, reading pace, formats and genres read, new-to-you authors by year, rating trends and current owned/unowned reading breakdowns. Series collection completion focuses on the likely mainline sequence: BookStats collapses alternate-language/edition catalog rows that repeat the same numbered positions and can reject wildly inflated catalog totals when the numbered data strongly supports a smaller sequence. The Series page shows owned and missing mainline titles without exposing provider branding. **Edit completion** lets the user set the expected count, include/exclude catalog entries, add an omitted title manually, or reset to automatic detection. If no usable catalog checklist is available, BookStats falls back to conservative numeric-gap behavior rather than inventing titles.

### Local data architecture

BookStats is now genuinely platform-aware:

- **Web:** IndexedDB via Dexie
- **macOS / Windows Tauri:** SQLite via `@tauri-apps/plugin-sql`

When a v0.3 desktop build starts for the first time, it checks the WebView IndexedDB used by the older desktop builds. If SQLite is empty, existing v0.2 books are copied into the new desktop SQLite database automatically.

A new installation starts with an empty library; demo records are no longer automatically inserted.

### Accounts and cloud sync

When PostgreSQL is configured, the **Account** section supports:

- account creation with password confirmation
- sign in / sign out
- email verification before cloud synchronization
- resend-verification flow
- forgotten-password email + one-hour reset links
- password resets that invalidate existing sessions
- password hashing with Node's scrypt implementation
- random server-side session tokens
- PostgreSQL-backed cloud library records
- deletion tombstones
- per-device sync cursors
- manual synchronization plus visible last-successful-sync and error state
- automatic sync after local edits while signed in
- incremental per-record uploads with a persistent local outbox and bounded retry batches
- change password from a dedicated security dialog
- delete the cloud library while keeping desktop-local data, or permanently delete the account
- synchronized reading goals alongside books and shelves
- browser ↔ desktop synchronization through the same API
- browser cache removal on logout so the previous account's library is not left visible
- desktop logout keeps the local SQLite library and only disconnects cloud sync

The application still works without an account or server connection.

### Help / feedback

The **Help / Feedback** section can send a bug report or feature suggestion through the same Resend-backed BookStats API. Reports automatically include only a small diagnostic summary: BookStats version, Web/macOS/Windows runtime, local-storage type, signed-in/verification state and library/shelf counts. Book titles, reviews, notes, passwords and tokens are not attached.

Feedback is delivered to `BOOKSTATS_FEEDBACK_TO` when set; otherwise the server falls back to `BOOKSTATS_EMAIL_REPLY_TO`, then the configured sender address. The public endpoint has basic per-IP throttling.

### Web and mobile layout

The web build is the primary BookStats experience. Desktop browsers keep the full sidebar and dense library controls, while phone-sized screens use a labeled bottom navigation bar, stacked/compact cards, touch-sized controls, horizontally scrollable table data, and bottom-sheet style dialogs. Shelf management remains available on phones through the Library toolbar even though the desktop shelf sidebar is hidden.

### Desktop packaging

The repository now includes the Tauri icon assets that were missing from v0.2, plus the Tauri SQL plugin and required capability permissions.

## Repository layout

```text
BookStats/
├── apps/
│   ├── client/       Normal React/Vite web UI + Tauri desktop shell
│   ├── admin/        Separate administrator-only React/Vite web UI
│   └── server/       Fastify account/sync/metadata/admin API
├── packages/
│   ├── domain/       Shared BookStats types
│   └── statistics/   Shared statistics calculations
├── database/
│   └── migrations/   PostgreSQL migrations
├── docs/
│   ├── DESIGN.md
│   └── SERVER_SETUP.md
├── tools/
│   ├── migrate-db.mjs
│   └── admin-user.mjs
├── export.sh
├── set-version.sh
├── unpack_bookstats.sh
└── CHANGELOG.md
```

## Required dependencies

### All development

Install:

1. Node.js 20 or newer
2. npm

```bash
node --version
npm --version
```

### macOS desktop builds

Install Apple's command-line tools:

```bash
xcode-select --install
```

Install Rust with rustup:

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

Restart your shell or run:

```bash
source "$HOME/.cargo/env"
```

Verify:

```bash
rustc --version
cargo --version
```

### Windows desktop builds

Install:

1. Microsoft C++ Build Tools / Visual Studio Build Tools with **Desktop development with C++**
2. WebView2 if Windows does not already provide it
3. Rust with the MSVC toolchain
4. Node.js 20+

Tauri's current platform prerequisites are documented at <https://v2.tauri.app/start/prerequisites/>.

### PostgreSQL

PostgreSQL is optional for local-only BookStats use. It is required for:

- user accounts
- cloud storage
- web/desktop synchronization
- server-side metadata caching

Transactional email is required in production for email verification and password recovery. BookStats uses the Resend HTTPS API, so the server does not store or authenticate with a mailbox password.

## First-time setup

From the repository root:

```bash
npm install
npm run check
```

`npm install` also installs the Tauri SQL JavaScript bindings. Cargo downloads the Rust-side SQL plugin the first time the desktop project builds.

## Run the web app only

```bash
npm run dev:client
```

Open:

```text
http://localhost:5173
```

Library CRUD/statistics work without the API. Catalog lookup and accounts require the API.

## Run web + API

```bash
npm run dev
```

This is the recommended web-development command because book lookup and accounts talk to the API.

API health check:

```text
http://127.0.0.1:8787/health
```

The metadata endpoints work without PostgreSQL. Account and sync endpoints return a clear `503` until `DATABASE_URL` is configured and the migrations are applied.

## Run the macOS/Windows desktop app

```bash
npm run tauri:dev
```

The root `tauri:dev` command starts both:

- the local BookStats API
- the Tauri application (which starts Vite itself)

This means catalog lookup works during desktop development without opening a second terminal.

The native desktop library is stored in SQLite rather than the WebView's IndexedDB.

## PostgreSQL / cloud setup

Copy the environment template:

```bash
cp .env.example .env
```

Configure at least:

```dotenv
BOOKSTATS_HOST=127.0.0.1
BOOKSTATS_PORT=8787
DATABASE_URL=postgresql://bookstats:YOUR_PASSWORD@127.0.0.1:5432/bookstats
BOOKSTATS_METADATA_USER_AGENT=BookStats/1.1.1 (your-real-contact@example.com)
# Optional but recommended for the full v0.9 metadata stack:
BOOKSTATS_GOOGLE_BOOKS_API_KEY=your_google_books_api_key
BOOKSTATS_HARDCOVER_API_TOKEN=your_hardcover_api_token
BOOKSTATS_PUBLIC_URL=https://kylecarroll.com/bookstats/
BOOKSTATS_RESEND_API_KEY=re_your_api_key
BOOKSTATS_EMAIL_FROM=BookStats <bookstats@kylecarroll.com>
BOOKSTATS_EMAIL_REPLY_TO=bookstats@kylecarroll.com
# Optional. Feedback defaults to EMAIL_REPLY_TO/EMAIL_FROM if omitted.
BOOKSTATS_FEEDBACK_TO=bookstats@kylecarroll.com
```

Then apply the database migrations:

```bash
npm run db:migrate
```

The migration runner requires the PostgreSQL `psql` client.

Complete Debian/NGINX/systemd deployment notes are in [`docs/SERVER_SETUP.md`](docs/SERVER_SETUP.md).

## Client API URL

During local development the client defaults to:

```text
http://127.0.0.1:8787/api/v1
```

For a deployed web or desktop build, set the API address when building:

```bash
VITE_BOOKSTATS_API_URL=https://your-domain.example/bookstats/api/v1 npm run build -w @bookstats/client
```

For a Tauri production build:

```bash
VITE_BOOKSTATS_API_URL=https://your-domain.example/bookstats/api/v1 npm run tauri:build
```

## Build the web app

```bash
npm run build -w @bookstats/client
```

Output:

```text
apps/client/dist/
```

## Build the desktop application

```bash
npm run tauri:build
```

Tauri build artifacts are produced under:

```text
apps/client/src-tauri/target/release/bundle/
```

For the simplest release process, build macOS on macOS and Windows on Windows.

## Validate everything

```bash
npm run check
```

This runs type checking, statistics tests and production builds.

Individual commands:

```bash
npm run typecheck
npm test
npm run build
```

## Change the project version

Use the new helper from the repository root:

```bash
./set-version.sh 0.3.1
```

or:

```bash
npm run version:set -- 0.3.1
```

The script accepts exactly `x.x.x` and updates:

- root `package.json`
- all workspace `package.json` files
- `package-lock.json` workspace versions when the lockfile exists
- Tauri `tauri.conf.json`
- Rust `Cargo.toml`
- the API-reported application version
- the README current-version line
- the sample Open Library user-agent version

It intentionally does **not** rewrite historical entries in `CHANGELOG.md`.

## Tauri icons

The complete icon set is committed under:

```text
apps/client/src-tauri/icons/
```

If the source icon is replaced later, regenerate the Tauri set with:

```bash
cd apps/client
npx tauri icon public/bookstats-icon.png
```

## Current sync design

BookStats v1.1 uses an incremental per-record outbox on each device. Saving a book, shelf, or reading goal updates one compact local outbox entry; repeated edits to the same record replace that entry with its latest timestamp instead of adding duplicate work. Deletions continue to use durable tombstones. A normal one-book edit therefore uploads that book record rather than rebuilding an upload from the complete library.

Large backlogs are split into bounded requests of at most 100 records and approximately 900 KiB. The server returns explicit acknowledgements for newly accepted mutations as well as safe retries that it has already applied or superseded, so the client can clear only the exact outbox/tombstone entries that are durably accounted for. The server cursor is saved after every successful batch; any later failed batch simply remains queued locally. Pull-only sync requests still run when the device has no local edits so changes made on other devices arrive normally.

The v1.1 upgrade seeds the outbox once from records newer than the device's last successful v1.0.x sync timestamp. That preserves unsynced offline edits without forcing an already-synchronized multi-thousand-book library through the network again. Sync implementation details remain internal; the Account UI intentionally stays simple.

Each synchronized item is still stored as a complete BookStats record in PostgreSQL JSONB. Synchronization is incremental at the record level rather than field-diff level: editing one book sends one book, not one field and not the whole library. Newer local saves/deletes are protected from older in-flight server responses. Notes/review field-level conflict UI remains a later hardening task.

## Security note

Passwords are hashed server-side and raw session tokens are never stored in PostgreSQL (only SHA-256 token hashes are stored).

For this development milestone, the client stores its bearer session token in local storage. Before a public production release, the desktop build should move that credential into the operating-system keychain/credential store and the web deployment should receive an additional security review.

## Metadata providers

Open Library is always available as the no-key fallback. Google Books and Hardcover are optional server-side enrichments configured with `BOOKSTATS_GOOGLE_BOOKS_API_KEY` and `BOOKSTATS_HARDCOVER_API_TOKEN`. The keys/tokens are read only by Fastify and are never emitted into the Vite/browser bundle.

BookStats uses the providers by field rather than treating one catalog as globally authoritative. In general, exact edition facts prefer an exact Google Books/Hardcover/Open Library edition match, while work/series relationships prefer Hardcover, then Google Books, then Open Library. Covers are pooled for user selection. If a provider is unavailable, the other configured providers continue to work.

For a deployed server, set `BOOKSTATS_METADATA_USER_AGENT` to include a real contact address for Open Library. The server caches human-triggered lookup responses in PostgreSQL when available. Hardcover requests are spaced to stay comfortably below its public API rate limit.


## Git setup

The supplied archive contains the initialized `.git` repository. If an extraction/copy tool strips hidden files:

```bash
git init
git branch -M main
git add .
git commit -m "BookStats v0.3.0"
```

Then add your remote normally.

## Planned next work

BookStats v1.0.0 is the feature-complete stable baseline. Ongoing 1.0.x work should be limited to regression fixes, performance tuning, accessibility/security hardening, and issues found through real-world beta/production use. Larger model changes and new collector workflows belong in later 1.x releases.

## Documentation

- [`docs/DESIGN.md`](docs/DESIGN.md) — product and architecture direction
- [`docs/SERVER_SETUP.md`](docs/SERVER_SETUP.md) — cloud/API deployment
- [`CHANGELOG.md`](CHANGELOG.md) — release history

## Troubleshooting

### `cargo metadata ... No such file or directory`

Rust/Cargo is not installed or is not on `PATH`:

```bash
source "$HOME/.cargo/env"
cargo --version
```

### Tauri reports a missing icon

v0.3 includes the icon directory. Confirm this exists:

```text
apps/client/src-tauri/icons/icon.png
```

If upgrading an older working tree with a partial patch, make sure the entire `src-tauri/icons/` directory was copied.

### Catalog lookup says the server is unavailable

For browser development, use:

```bash
npm run dev
```

rather than `npm run dev:client` when you want lookup/account features.

For Tauri development, use the root:

```bash
npm run tauri:dev
```

which starts the API automatically in this release.


## Release bundles and server deployment

Create release bundles with:

```bash
./export.sh all
```

The server artifact is now a ZIP, so the normal server upload is simply the matching pair:

```text
BookStats-Web-vX.Y.Z.zip
BookStats-Server-vX.Y.Z.zip
```

After SCPing those files into `/home/kcarroll/Downloads/`, the server-side helper can deploy both, preserve `/opt/bookstats/.env`, install production dependencies, apply database migrations, restart BookStats and NGINX, run a health check, and remove the consumed archives:

```bash
./unpack_bookstats.sh
```

The deployment helper also accepts the older server `.tar.gz` artifact during the transition.
