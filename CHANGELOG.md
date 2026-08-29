# Changelog

## 1.0.4

- Added a small Library Health easter egg for a perfectly clean catalog: when every applicable metadata check passes, BookStats now displays **The Librarian Approves.** and a brief completion message.
- Cleaned up desktop updater initialization so Tauri propagates plugin setup errors instead of emitting an unused `Result` compiler warning.
- This release is intentionally small so the complete v1.0.3 → v1.0.4 desktop auto-update and download/deployment workflow can be verified end to end.
- No database migration is required for v1.0.4.

## 1.0.3

- Replaced BookStats browser, Home Screen/PWA, and desktop application icons with the brown BookStats open-book mark used in the application header. Added dedicated favicon, Apple touch icon, standard PWA sizes, maskable icons, and regenerated Tauri icon assets.
- Added installable mobile-web-app support with a web app manifest, standalone display mode, safe-area handling for notches/Home indicators, and a small service-worker app shell so the BookStats interface can start more gracefully when connectivity is interrupted. API responses and account data are never service-worker cached.
- Added a dismissible iPhone/iPad install guide that explains Safari's More → Share → Add to Home Screen flow and detects when BookStats is already running as an installed web app. Signed-in users are reminded to sign in inside the new app instance; local-only iOS users receive an explicit warning that Safari and the installed Home Screen app do not share the same IndexedDB library.
- Added the Android equivalent. Chromium browsers use the native install prompt when available, with a browser-menu Install app / Add to Home screen fallback on other Android browsers.
- Added signed Tauri desktop auto-updates. When the server requires a newer version, macOS/Windows builds now offer **Update BookStats**, show download/install progress, verify the signed update, install it, and relaunch instead of presenting the web-only Refresh action.
- Added updater signing and release tooling. Desktop exports now create Tauri updater artifacts/signatures, per-platform static update manifests for `https://kylecarroll.com/downloads/bookstats/`, plus a `BookStats-Updater-vX.Y.Z.zip` deployment bundle. The deployment helper can publish that updater bundle together with a normal web/server release or by itself.
- Added a one-time updater-key setup helper. The private signing key is generated outside the repository and must be retained securely; only its public key is written into Tauri configuration. Existing v1.0.2 desktop installations require one manual upgrade to v1.0.3 because the updater did not exist in v1.0.2. Automatic desktop updates apply from v1.0.3 forward.
- No database migration is required for v1.0.3.

## 1.0.2

- Added Previous/Next navigation to Book Details, following the current filtered/sorted Library order when possible. Left and Right arrow keys provide the same navigation, while the buttons disable cleanly at the ends of the list.
- Fixed alignment in Edit Book → Your library & reading so Status and Condition controls share the same top alignment even when Condition includes helper text.
- Refined the reading-history editor into clearer per-session cards with structured date/page/note fields, compact status/progress treatment, and better responsive layouts.
- Added a dismissible first-visit account introduction for signed-out users. It explains cloud-sync and browser-safety benefits, offers direct Create account / Sign in actions, and explicitly allows continued local-only use. Dismissing it is remembered in that browser.
- Simplified metadata refresh behavior: catalog lookup now fills blank fields only and never overwrites populated values. Removed the user-facing saved-field/manual-override concept; legacy override markers from older records are ignored and stripped as records are normalized.
- Added sanitized rich-HTML rendering for Book Details descriptions so common catalog markup such as bold, italics, paragraphs, lists, headings, blockquotes and links displays correctly instead of showing raw tags.
- Added a dismiss button to the Random Book card so a generated suggestion no longer stays on screen until another view change.
- Preserved provider-supplied multi-volume series positions during Google Books lookup. Display positions such as `Books 1-6` now win over a lossy numeric order value, allowing omnibus/multi-volume records to remain `1-6` instead of collapsing to `1`.
- Added editing for additional authors in Edit Book. The primary author remains separate, while co-authors/editors can be added, changed or removed individually.
- Fixed an intermittent save/sync race that could make a freshly edited book appear to revert shortly after saving. BookStats now queues a follow-up sync when a save occurs during synchronization, refuses to apply an older server book response over a newer live local edit, and prevents a background refresh snapshot captured mid-save from replacing newer in-memory data.

## 1.0.1

- Kept catalog series names and positions correlated as a single metadata pair. Multi-series provider memberships now retain each series with its own volume; changing a provider-managed series replaces/clears the old provider-managed position, while a manually chosen series can resolve its matching catalog position on the next metadata lookup. Manual series-volume edits remain protected.
- Fixed the phone bottom navigation so all six primary destinations fit on one row and **Account** is always accessible; the Statistics label is shortened to **Stats** on small screens to preserve tap-target width.
- Added stale-client protection. API requests now identify the BookStats client version; incompatible pre-1.0.1 clients receive a clear refresh-required response, and current clients periodically compare against the server and show a blocking **Refresh BookStats** dialog when a newer server release is detected.
- Hardened web API routing so production browser builds use the same origin that served BookStats instead of depending on an absolute hostname. This avoids www/non-www CORS mismatches while desktop/Tauri builds continue to use the configured public API URL.
- CORS configuration now recognizes the matching www/non-www alias of configured HTTP(S) origins and accepts the BookStats client-version header.
- Replaced opaque browser network failures with a user-facing BookStats server connection message while preserving the underlying error in the developer console.

- Fixed archived cover rendering so the server uses a deterministic installation-relative cover path instead of depending on its process working directory.
- Cover images now fall back through local cache, BookStats archive, and retained original source rather than leaving a broken image when one layer is unavailable.
- Archived cover URLs include a book revision cache key so repaired/reselected images refresh immediately in Library and detail views.
- The cover migration utility now detects and can repair missing archived files when a retained source URL is available.
- The Library now defaults to Author ascending (using BookStats family-name author sorting) instead of Title.

- Added durable server-side cover assets for signed-in libraries. Once a user selects a catalog, URL, or custom cover, BookStats can archive the actual image instead of depending permanently on the provider URL.
- Added opaque per-asset access tokens with device-local cache → BookStats archive rendering; original external URLs are retained as provenance and a repair/re-archival source.
- Custom uploaded images are moved out of synchronized book JSON after the cloud archive succeeds, avoiding large base64 payloads in PostgreSQL sync records.
- Advanced portable BookStats JSON exports to format version 12 for the new cover-asset references.
- Added migration `0007_cover_assets.sql` and a safe dry-run-first `covers:migrate` utility for archiving covers already selected before v1.0.1.
- Updated deployment to preserve `/opt/bookstats/data/` across server releases.
- Retained local-first behavior: users without an account continue to keep selected covers in local browser/desktop storage, and failed cloud archiving never blocks saving a book.

## 1.0.0

- Added **Lending** as a first-class collection workflow. Books can be loaned to a named borrower with loan/due dates and notes, marked returned without losing history, and reviewed from a dedicated Lending dashboard with overdue and due-soon summaries. Active loans also appear on Book Details.
- Added **camera ISBN/barcode scanning** to the normal client. Scan ISBN opens the device camera, recognizes ISBN-10/ISBN-13 book barcodes, and launches Add Book with the exact ISBN prefilled and automatically searched. Native BarcodeDetector is used when available, with a lazily loaded ZXing browser fallback for browsers without that API; manual ISBN entry remains available. macOS desktop bundles now include an explicit camera-usage description.
- Rebuilt Library Health duplicate handling into a reconciliation workflow. Candidate groups are ranked by source IDs, canonical ISBN edition identity, and normalized title/author; ISBN-10 and equivalent ISBN-13 now count as the same edition signal. Different ISBNs are called out as likely intentional editions.
- Added side-by-side duplicate comparison before merging and a persistent **Keep as separate editions/copies** decision so intentionally duplicated collector records stop reappearing in Library Health. Merges continue to create a safety backup and now also preserve lending history, series-completion rules, and duplicate decisions.
- Added **Series collection-completion tools** under Statistics → Series. Completion now focuses on owned mainline positions rather than every catalog row or every book merely recorded in the library.
- Strengthened completion reconciliation so an incomplete provider checklist can be repaired by clearly numbered books already present in the user's library. A provider reporting 12 Everworld books but omitting positions 7 and 11 now reconciles against owned `#7` and `#11` records instead of incorrectly reporting 10/12.
- Added multi-volume series-position parsing. Values such as `1 & 2`, `1 and 2`, `1, 2`, and integer ranges such as `3-4` can satisfy multiple mainline positions while remaining a single physical library record; Goodreads/title metadata parsing now preserves these compound position expressions as well.
- Tightened noisy series catalogs by using numbered-position repetition and sane catalog counts to collapse translated editions, alternate-language work records, omnibuses, and side entries that share the same mainline volume positions. When a provider itself reports an obviously inflated primary-book count, a strong compact 1..N sequence can override it; a noisy 131-row/131-count catalog repeating positions 1–5 is therefore reduced to a five-book completion sequence.
- Added a manual **Edit completion** tool for imperfect series data: users can set the expected mainline count, choose which catalog entries count, add omitted books manually, or reset to automatic detection. Completion overrides synchronize as part of the existing Book JSON records.
- Removed provider-brand labels such as “Hardcover catalog” from Series completion UI. BookStats may combine catalog sources internally, while the collection view focuses on the resulting sequence and lets the user correct it when necessary.
- Polished completed Series cards so reading and collection progress retain the same two-bar layout at 100%, with a subtle completed-state treatment instead of collapsing into a horizontal row. Removed internal reconciliation/explanation text from the end-user Series view.
- Made the Series name on Book Details navigable: selecting it now closes the detail view and opens the Library with the same exact-series filter used by Statistics → Series.
- Expanded **Load catalog covers** so cover discovery always combines exact-ISBN artwork with likely alternate editions found by title + author. Exact-ISBN covers stay first, duplicates are removed, and choosing alternate artwork never changes the stored ISBN or any other metadata.
- Upgraded portable BookStats JSON exports to format version 11 so lending history, duplicate-ignore decisions, and series-completion overrides round-trip through export/import. Repeated imports preserve these local collection-management fields.
- Fixed the new 1.0 Lending, Series Completion, duplicate-reconciliation, and barcode-scanner dialogs so their modal surfaces are fully opaque while the surrounding page remains dimmed.
- Added no new PostgreSQL schema requirement for these collection features; they live inside the existing synchronized Book JSON records.

## 0.10.0

- Reordered the Account page so sign-in/account controls appear before storage information, and expanded the Storage card to explain local-only persistence, browser-storage risks, email-verification state, cloud synchronization, and the value of independent exports/backups.
- Added a completely separate administrator web client under `/bookstats/admin/`; the normal BookStats web/desktop client contains no admin navigation, components, API helpers, or UI references.
- Added PostgreSQL migration `0006_admin_console.sql` with `user`/`admin` roles, account disable state, and an append-only administrator audit log.
- Added server-only `npm run admin:user -- grant|revoke|status <email>` tooling. There is deliberately no browser/API endpoint that can promote an account to administrator.
- Hardened the administrator role utility to use the existing PostgreSQL `pg` driver with parameterized queries instead of shelling out to `psql`, avoiding client-side variable-substitution/quoting failures.
- Added administrator-only authentication and `/api/v1/admin/*` authorization. Disabled accounts cannot sign into the normal client, and disabled/non-admin accounts cannot obtain an administrator session.
- Added an administrator dashboard with total/active/disabled users, library-record totals, active sessions, database size/latency, API uptime/version, metadata provider state, transactional-email state, and recent signups.
- Added searchable user management with account details, library/shelf/goal counts, verification state, recent activity, profile corrections, enable/disable, session invalidation, and forced password-reset delivery.
- Added guarded cloud-library reset and permanent account deletion. Destructive account actions require exact typed confirmation and are audited.
- Added administrator inspection of synchronized book/shelf/goal records, including targeted JSON repair and audited record deletion for support/recovery work.
- Added administrator audit-log browsing with actor, target, action, record, timestamp, IP address, and non-sensitive action details.
- Web release bundles now contain both the normal app and a separately built `/admin/` application; desktop bundles still contain only the normal BookStats client.
- Server release bundles now include the administrator role utility.
- Retained all cumulative v0.9.4 fixes: English-first LibraryThing series choice, grouped Smart Shelf/advanced-filter logic, Statistics series links, centered goal icon, family-name author sorting, and refillable manually-cleared catalog fields.

## 0.9.4

- Library author sorting now uses the primary author's family name rather than the first character of the display name. A book by `A. C. Crispin` therefore sorts under `C`; common multi-word surnames such as `Le Guin` are kept together. Books with multiple contributors continue to sort by their primary/first-listed author for a stable single position.
- Clearing any catalog-managed field in Edit Book now removes that field's manual-override protection. A later metadata lookup or targeted repair can refill the blank field instead of treating an intentionally emptied value as permanently protected. Existing records that already contain a blank manual override are normalized when opened for editing.
- Clicking a series name in Statistics → Series now opens the Library with an exact filter for that series.
- Added grouped Boolean logic to Smart Shelves and Advanced Filters. Rules inside each group can use **AND** or **OR**, and separate parenthesized groups can also be combined with **AND** or **OR**.
- Preserved all pre-v0.9.4 smart shelves without changing their existing flat `match all` / `match any` behavior; editing an older smart shelf transparently places its rules into the first Boolean group.
- Upgraded portable BookStats JSON exports to format version 10 so grouped smart-shelf definitions round-trip through export/import, backups and cloud sync.
- Improved LibraryThing series selection for multilingual exports. Clear foreign-language candidates are deprioritized when an English-looking alternative is present, while collection-wide frequency still chooses among plausible English/base-series candidates. Re-importing the same LibraryThing entry can refresh an importer-owned series value so libraries created before this fix can be corrected without discarding other BookStats edits; manually protected series fields are still preserved.
- Verified the English-first importer against the 2,499-book LibraryThing regression library, including Foundation, Wheel of Time, Narnia, Lord of the Rings, Discworld, Culture, Magic Tree House and other multilingual series lists.
- Fixed the Statistics reading-goal target/check icon alignment by preventing the goal-card flex rule from overriding the icon container's centered grid layout.
- No database migration is required; grouped shelf rules remain inside the existing synchronized shelf JSON.

## 0.9.3

- Added a user-entered **Condition** field with New, Like New, Very Good, Good, Acceptable and Poor values. Catalog metadata lookup/repair never overwrites it.
- Added Condition to Add/Edit Book and Book Details while intentionally leaving the library table columns unchanged.
- Added Condition rules to Smart Shelves and Advanced Filters, including `is` and `is not` matching against the six supported values.
- Added bulk Condition editing so large imported libraries can be classified efficiently, including an option to clear the field.
- Added Condition to normal library text search and preserved it through BookStats export/import, backups, cloud sync, duplicate merges and repeated imports.
- Added persistent manual shelf ordering in **Add & manage shelves** with Move Up / Move Down controls. The chosen order is used in the sidebar, shelf selectors and book editing.
- New shelves are appended to the current order; older libraries without saved order data continue to open safely and receive stable fallback ordering until rearranged.
- Upgraded portable BookStats JSON exports to format version 9 so shelf order and condition travel with exported libraries.
- No IndexedDB or PostgreSQL schema migration is required; both fields live in the existing JSON book/shelf records.

## 0.9.2

- Improved large-library web performance by avoiding full IndexedDB reloads after ordinary local writes, precomputing shelf membership/counts, deferring text search work, and rendering the Library in 120-book batches instead of creating thousands of cards or table rows at once.
- Limited Library Health metadata cleanup rows to 100 at a time and reduced repeated health/duplicate calculations, keeping the cleanup screen responsive with libraries containing thousands of books.
- Added **Load catalog covers** to Edit Book. It searches the current ISBN first, falls back to title + author only for cover discovery, and loads the existing multi-provider cover pool without replacing any other book metadata.
- Added **Fill missing metadata** to Edit Book for targeted repair of blank catalog fields such as description, ISBN, series/position, pages, publication year, publisher, language and genre while preserving existing values and manually protected fields.
- Kept targeted ISBN metadata repair edition-safe: when an ISBN is present, BookStats will not use another edition to fill missing edition fields.
- Updated Library Health to route cleanup work into the targeted repair workflow and to show missing-field percentages explicitly.
- No database migration is required.

## 0.9.1

- Made the Library header count follow the currently visible filtered/search result set, matching shelf count behavior.
- Fixed LibraryThing imports for collectors: each stable LibraryThing `books_id` is now treated as an individual catalog entry, so separate copies/editions with the same title or ISBN are preserved instead of merged. Re-importing the same LibraryThing entry remains idempotent.
- Improved metadata title/author search ranking so exact-title matches outrank box sets, omnibuses and loose provider ordering; expanded the result window so relevant Open Library editions are not crowded out by Google Books/Hardcover results.
- Improved Open Library title search to use its query-matched edition information instead of blindly taking the first edition key for a work when edition data is available.
- Improved series/volume merging by allowing the best compatible provider to supply a missing volume even when another provider supplied the series name, and added targeted Hardcover/Open Library series enrichment when selected-edition details are incomplete.
- Preserved the v0.9 exact-ISBN rule: ISBN lookups still accept only the requested ISBN (or its equivalent ISBN-10/ISBN-13 form) and never substitute a different edition.
- No database migration is required.

## 0.9.0

- Added a server-side metadata-provider layer combining Open Library, Google Books and Hardcover behind one BookStats search UI.
- Made ISBN lookup edition-first: BookStats accepts only the requested ISBN or its equivalent ISBN-10/ISBN-13 form and will not silently substitute another printing or format.
- Added Google Books enrichment for edition facts, descriptions, identifiers, page counts, covers and explicit series order metadata.
- Added Hardcover enrichment for editions, authors and ordered series catalogs; provider requests are rate-spaced and the API token remains server-side.
- Added field-by-field metadata merging with edition fields and work/series fields using different provider priorities rather than trusting whichever service responds first.
- Added per-field metadata provenance plus retained provider work/edition references and exact-ISBN match information on Book records.
- Added manual metadata protection: catalog fields edited by the user are preserved when a later lookup is applied.
- Expanded alternate-cover selection to aggregate available artwork across configured metadata providers.
- Upgraded Series statistics to compare the local library with an available Hardcover/Google series catalog, including provider-backed missing-title tracking and collection percentage; conservative numeric-gap detection remains the fallback.
- Updated Library Health copy and lookup paths to use the configured provider stack.
- Added `BOOKSTATS_GOOGLE_BOOKS_API_KEY` and `BOOKSTATS_HARDCOVER_API_TOKEN` server configuration plus provider status reporting in `/health`.
- No PostgreSQL schema migration is required; metadata provenance and series catalogs are stored inside existing synchronized book JSON records.

## 0.8.0

- Reorganized Tools into six consistent action cards: Export BookStats, Import BookStats, Import Goodreads, Import LibraryThing, Library Health & Cleanup, and Backup & Restore.
- Removed the redundant Safe Repeated Imports explainer card while keeping conservative matching behavior in the import workflow itself.
- Moved local backup history below the main Tools grid and made Backup & Restore match the size and visual weight of the other tool cards.
- Redesigned Account into a compact two-column settings layout on desktop with smaller cards, tighter spacing, and corrected icon centering.
- Replaced the expanding Change Password section with a dedicated modal dialog and removed the Change Email control from the Account UI.
- Removed the Account-page shortcut to exports/backups; those tools now live only in the Tools section.
- Added a focused mobile-web UX pass: labeled bottom navigation, hidden desktop shelf sidebar on phones, a mobile Manage Shelves action, responsive toolbar controls, touch-friendly buttons, compact cards, scroll-friendly tables, and phone-sized bottom-sheet dialogs.
- Improved responsive behavior across Library, Statistics, Tools, Account, Help/Feedback, book details, cleanup, import preview, goals, filters, shelf management, and bulk editing while preserving desktop web as the primary layout.
- No database migration is required for v0.8.0; this release is intentionally focused on web UX/UI polish before metadata-provider work in v0.9.0.

## 0.7.0

- Merged the planned library-management, reading-tracking and library-intelligence milestones into one focused release while preserving the simple default interface.
- Added structured reading sessions with start/finish dates, current-page progress and session notes; legacy completion dates remain backward compatible.
- Added user-created Statistics goals for books read, pages read, rereads, new-to-you authors and owned books read across custom date ranges.
- Added goal pace tracking with elapsed-time percentage, days remaining, on/behind-pace status and projected finish values.
- Expanded Statistics with monthly current-year books/pages, long-term reading and page trends, genres/formats read, current ownership breakdown, new-to-you authors by year, rating trends, reading extremes and richer pace metrics.
- Added a dedicated Series statistics view with collection/read progress, known volume positions and conservative numeric gap detection between recorded volumes only.
- Added structured Advanced Filters using the same all/any rule engine as smart shelves, including last-read/date-added rules and one-click Save as Smart Shelf.
- Added grid/table multi-select and bulk status, ownership, shelf and tag editing, selected-record export, and bulk deletion.
- Added import preview for BookStats, Goodreads and LibraryThing imports with new/matched/ambiguous counts and conservative handling of uncertain matches.
- Added recent import history and safe Undo Import; records edited after an import are skipped rather than overwritten by undo.
- Separated portable BookStats Export from recovery Backup. Export/import is merge-friendly; backup restore replaces the local library with the saved snapshot.
- Added manual downloadable backups plus up to five device-local safety snapshots, including automatic daily snapshots and pre-change snapshots before imports, import undo, duplicate merges, bulk deletes and restores.
- Added Library Health with an objective completeness score and cleanup categories for cover, description, ISBN, pages, publication year and missing series position.
- Expanded account management with change-password, change-email/reverification, explicit sync status/last-successful-sync, cloud-library deletion and permanent account deletion.
- Reading goals synchronize through the existing free-text cloud record system; reading sessions remain inside book JSON, so no new PostgreSQL migration is required.
- Browser logout/data removal now clears local metadata too, including device-local safety backups and import history, so another browser user cannot see private remnants.
- Kept Open Library as the sole metadata provider in this release. Provider expansion/ranking and the planned 1.0 release-hardening work are intentionally deferred.

## 0.6.0

- Added user-created smart shelves alongside regular shelves; BookStats still creates no default shelves.
- Smart shelves support all/any matching and rules for reading status, ownership, rating, title, author, series, format, genre, tags, pages, publication year and read count.
- Smart shelves update automatically throughout sidebar counts, library filtering, book cards, book details and shelf statistics; they are not manually assignable from the book editor.
- Added shelf editing so smart-shelf rules and shelf names can be adjusted after creation.
- Added a polished read-only Book Detail screen with cover art, status, series, rating, edition facts, regular/smart shelves, description, review, notes, tags and complete reread history. Clicking a book now opens details first; editing is an explicit action.
- Added Help / Feedback with Report a Bug and Suggest a Feature flows. Reports include only safe diagnostics (version, platform, storage type, account-state booleans and library counts) and are delivered through Resend.
- Added a throttled `/api/v1/feedback` endpoint plus optional `BOOKSTATS_FEEDBACK_TO` configuration.
- Added a Library Cleanup tool with strong duplicate detection, user-reviewed record merging and metadata completeness checks for covers, descriptions, ISBNs, pages and publication years.
- Duplicate merges preserve the chosen record while combining missing metadata, reading history, shelves, tags, additional authors and external source IDs.
- Added alternate-cover selection from Open Library edition cover data while keeping local-file upload and device-local remote-cover caching.
- Removed the Owned Only checkbox. Table view now has a Hide / Show Columns control whose choices are remembered on the current device.
- Fixed the v0.5 sync type guard around locally cached covers.
- Upgraded portable BookStats JSON exports to format version 6 so smart-shelf definitions are preserved.
- No new PostgreSQL migration is required for v0.6; smart-shelf definitions use the existing synchronized shelf JSON record.

## 0.5.0

- Added first-class user-defined shelves separate from the single reading status.
- Books can belong to any number of shelves; shelves can be created, filtered, managed and synchronized across devices.
- Added shelf chips to cover cards, a Shelves table column, sidebar shelf navigation and shelf statistics.
- Added Goodreads CSV import using the real Goodreads export structure, including exclusive-status mapping, ratings, read dates, ownership, reviews, notes, ISBNs and custom shelves.
- Added LibraryThing JSON import using collections as shelf/status/ownership inputs and preserving ratings, read dates, tags, genres, formats, ISBNs and publisher data.
- External imports merge repeat runs by source ID, ISBN and title+author instead of blindly duplicating records.
- Improved series lookup by reading Open Library series fields and parsing common series/volume labels; bumped metadata-cache keys so stale lookups do not hide the fix.
- Replaced the old hard-coded `The Expanse` / `1` series placeholders with neutral field hints.
- Added user-selected local cover images with resizing plus device-local caching of remote catalog covers.
- Added PostgreSQL migration `0005_shelves.sql` so shelf records synchronize as first-class entities.
- Increased the API JSON body limit to support custom-cover records.
- Upgraded portable BookStats JSON exports to format version 5 and included shelf definitions/assignments.

## 0.4.3

- Replaced Microsoft Graph OAuth2 transactional email with the Resend HTTPS API.
- Removed the Microsoft tenant/client/secret configuration requirement.
- Added `BOOKSTATS_RESEND_API_KEY`, `BOOKSTATS_EMAIL_FROM`, and optional `BOOKSTATS_EMAIL_REPLY_TO` configuration.
- Health checks now report `emailProvider: "resend"` when transactional email is configured.
- Added Resend domain/DNS setup instructions for the production server.
- Kept the visible app version indicator and noninteractive DMG export behavior from v0.4.2.

## 0.4.2

- Replaced username/password SMTP delivery with Microsoft Graph OAuth2 client-credentials email delivery.
- Added Microsoft Graph `Mail.Send` application-permission configuration and server setup documentation.
- Added `emailProvider` to the health response for easier production diagnostics.
- Added a small app version label to the bottom of the sidebar; it is sourced from the package version and follows `set-version.sh` automatically.
- Removed the Nodemailer runtime dependency.

## 0.4.0

- Added confirmation-password validation during account creation.
- Raised the new/reset password minimum to 10 characters while retaining login compatibility with existing accounts.
- Added SMTP-backed verification emails and email-verification tokens.
- Cloud synchronization now requires a verified email address.
- Added resend-verification support.
- Added forgotten-password and one-hour password-reset links.
- Password resets revoke all existing sessions and send a password-changed notification when email is configured.
- Added migration `0004_account_security.sql`; pre-v0.4 accounts are grandfathered as verified.
- Web logout performs a final sync when possible and clears IndexedDB library/tombstone data so another browser user cannot see the previous account's synced library.
- Desktop logout intentionally keeps the local SQLite library while disconnecting the cloud account.
- Added `unpack_bookstats.sh` for matched Web/Server deployment from `/home/kcarroll/Downloads`.
- Server release bundles are now ZIP files; the deployment script still accepts legacy `.tar.gz` bundles.
- Added Windows export preflight for `makensis`.
- Added Nodemailer SMTP configuration and account-email documentation.

## 0.3.0

- Added Open Library title/author/ISBN catalog search through the BookStats API.
- Added catalog-detail enrichment for descriptions, covers, page counts, publishers and identifiers.
- Added publisher, language and metadata provenance fields to Book records.
- Added real Tauri desktop SQLite persistence using the official SQL plugin.
- Added one-time migration from the older desktop WebView IndexedDB into SQLite.
- Added Account navigation, account creation, login/logout and sync status UI.
- Added PostgreSQL-backed sessions and cloud library records.
- Added browser/desktop synchronization with deletion tombstones and per-device cursors.
- Added automatic sync after edits while signed in and manual Sync Now.
- Added PostgreSQL metadata caching.
- Added migration `0003_accounts_sync.sql`.
- Added `tools/migrate-db.mjs` and `npm run db:migrate`.
- Added complete Tauri icon assets and SQL plugin capabilities.
- Added `docs/SERVER_SETUP.md`.
- Added `set-version.sh` / `npm run version:set -- x.x.x`.
- Removed automatic demo-book seeding for new libraries.

## 0.2.0

- Moved JSON import/export into a dedicated Tools section.
- Added sortable library table columns with ascending/descending indicators.
- Added series name and volume fields and a Series table column.
- Added half-star ratings from 0.5 to 5.0.
- Added separate book description, personal review and private notes fields.
- Added multiple read dates per book, including rereads.
- Added an IndexedDB v2 migration that preserves legacy `dateRead` values.
- Expanded statistics into Overview, Reading, Authors, Library and Ratings categories.
- Added read/page history by year, monthly reading activity, reread metrics, author rankings, genres, series, collection growth, publication decades, page-length distributions and more.
- Upgraded BookStats JSON exports to format version 2.
- Added PostgreSQL migration scaffolding for future reading sessions and series data.
- Fixed statistics-package clean-build/typecheck resolution so it no longer depends on stale generated output.

## 0.1.1

- Fixed clean-workspace package resolution for the client.

## 0.1.0

- Initial TypeScript/React/Tauri BookStats foundation.
