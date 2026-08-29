# BookStats — Product & Technical Design Document

**Version:** 0.1  
**Status:** Initial architecture and product design  
**Targets:** Web, Windows, macOS  
**Primary client language:** TypeScript  
**Desktop framework:** Tauri 2  
**Cloud database:** PostgreSQL  
**Desktop/local data:** SQLite planned; IndexedDB used by the first web-capable iteration

## 1. Product vision

BookStats is a personal library manager, reading tracker, and reading-data visualization application. It combines the useful cataloging ideas of services such as Goodreads and LibraryThing without making the product a social network. Its defining feature is deep, visual exploration of a user's own reading and collection data.

At a basic level users can add and remove books, search by title/author/ISBN, track ownership and reading state, rate and review books, organize shelves and tags, and retrieve useful metadata and covers. At a deeper level BookStats should answer questions about how a user's reading, tastes, acquisition habits, unread collection, authors, genres, formats, and ratings change over time.

## 2. Product principles

1. **The user owns their data.** Complete export and portable backups are requirements.
2. **Power without clutter.** Basic workflows stay simple; advanced controls appear when needed.
3. **Statistics are first-class.** Historical events are modeled instead of flattening everything into booleans.
4. **Metadata remains editable.** Provider data assists the user but never overrides their own edition-specific facts.
5. **Desktop can work locally.** A cloud account is optional on desktop and useful for sync/backup.
6. **No accidental social network.** No followers, public feeds, likes, comments, or messaging in the core product.

## 3. Lessons retained from the original Python application

The abandoned project already contained valuable concepts: SQLite storage, Goodreads CSV and LibraryThing JSON imports, import IDs for deduplication, metadata overrides, Open Library cover lookups and cache, manual editing, read/unread/to-read/owned filters, tags, formats, customizable columns, backups, a filtered random-book picker, and substantial visual statistics.

The rewrite retains those product ideas but replaces the old single-record model with a richer domain model capable of editions, rereads, sync, and long-term statistics.

The original statistics define the baseline: total/owned/read/unread books, known and read pages, average ratings, unique authors, reading years and streaks, top authors/genres/tags/formats/collections, books and pages by year, ratings by year, acquisition history, rating distribution, format distribution, tag distribution, and page-length distribution.

## 4. Selected stack

| Area | Decision |
| --- | --- |
| Main language | TypeScript |
| UI | React |
| Web tooling | Vite |
| Desktop | Tauri 2 |
| Native desktop layer | Rust when needed |
| Client query/cache | TanStack Query planned |
| Local browser DB | IndexedDB/Dexie |
| Desktop DB | SQLite planned |
| Cloud database | PostgreSQL |
| API | TypeScript/Fastify |
| Charts | Apache ECharts |
| Metadata | Multi-provider adapter |
| Initial providers | Open Library + secondary provider |
| IDs | Client-generated UUIDs |

The browser and desktop applications share the same React UI. Tauri wraps that frontend for Windows and macOS and allows native Rust integration without requiring a separate desktop UI codebase.

## 5. Repository architecture

```text
BookStats/
├── apps/
│   ├── client/            # React/Vite + Tauri shell
│   └── server/            # Fastify API
├── packages/
│   ├── domain/            # Shared types and rules
│   └── statistics/        # UI-independent calculations
├── database/
│   └── migrations/        # PostgreSQL schema
├── docs/
├── tests/
└── tools/
```

Additional packages will be introduced for metadata, importers, synchronization and the reusable UI system when those milestones begin.

## 6. Core domain model

A **work**, an **edition**, a user's relationship with a work, a physical/digital library copy, and a reading session are different objects.

- **Work:** conceptual title such as *The Hobbit*.
- **Edition:** a specific published edition, ISBN, page count, publisher, format and cover.
- **Contributor:** authors, editors, translators, illustrators, narrators and other roles.
- **User book:** reading state, personal rating, review, notes and favorite status for a work.
- **Library item:** a particular edition owned by the user, with acquisition and physical-copy information.
- **Reading session:** one reading or rereading with start/finish dates and optional progress.
- **Shelf/collection:** arbitrary user organization independent of status and ownership.
- **Tags:** user-defined descriptors.
- **External IDs:** normalized mappings for ISBN/Open Library/Google Books/LibraryThing/etc.

This model permits multiple owned editions and multiple readings without destroying history.

## 7. Library experience

The main library supports a visual cover grid and a dense table view. Users can search title, author, ISBN, series and tags; combine filters; sort; save useful views; adjust cover density; and customize table columns.

The book detail view should feel like a book page rather than a database form. Editing is broken into focused sections. Bulk editing is important for large collections.

## 8. Book lookup and metadata

Search supports ISBN, title, and title+author. Search results lead to explicit edition selection when editions are known. Metadata providers implement a common normalized interface so the application is never coupled to a single external catalog.

Metadata provenance is retained. User overrides always win over provider data. Covers follow an edition cover → work cover → alternate provider → BookStats placeholder fallback and users may choose or upload their own cover.

## 9. Imports

Initial import targets:

- Goodreads CSV
- LibraryThing JSON/tab-delimited
- legacy BookStats SQLite
- BookStats JSON backup
- generic CSV mapping

Each importer follows: detect → parse → normalize → match/deduplicate → preview → import → report. Import batches record changes so an undo-import feature can be implemented safely.

Matching priority: prior source ID, ISBN-13, ISBN-10, catalog ID, title+author+edition, then fuzzy candidates that require confirmation.

## 10. Statistics architecture

Statistics are calculated independently of chart rendering. The engine accepts library/reading data plus metrics, dimensions, filters, grouping and time ranges, then visualization components render those results.

Metrics include books/pages read or owned, ratings, reading duration, rereads, DNF counts, unique/new authors, acquisition totals, TBR growth and reading velocity. Dimensions include time, author, genre, subject, tag, series, format, language, publisher, publication decade, shelf and ownership status.

Visualization types include metric cards, line/area/bar/stacked charts, histograms, donut charts, scatter plots, heatmaps, calendar heatmaps, treemaps, timelines and tables.

Long-term signature features include custom statistics dashboards and a visually polished annual “Year in Books” report.

## 11. Accounts and synchronization

Desktop supports local mode and synced mode. Web primarily uses a signed-in cloud library. Accounts initially use email/password, verification and password recovery; passkeys or identity providers can follow later.

Cloud data lives in PostgreSQL. Desktop will move to a local SQLite database as sync is implemented. Syncable records use client-generated stable UUIDs, revision counters, timestamps and tombstones for deletions. Clients exchange changes using cursors rather than replacing entire libraries.

Conflicting freeform text such as reviews and notes must be preserved rather than silently overwritten.

## 12. Security and privacy

- Passwords use a modern memory-hard hash such as Argon2id.
- HTTPS is mandatory for production API traffic.
- Desktop tokens belong in OS credential/keychain storage.
- Every server-side resource operation verifies ownership.
- Login endpoints are rate-limited.
- Libraries, ratings, history, reviews and notes are private by default.
- Telemetry is absent or deliberately minimal/documented.

## 13. Backup and export

Sync and backup are separate concerns. The server performs its own database and asset backups; users can also download a full personal backup.

The native backup format is a versioned ZIP containing a manifest, normalized library JSON and optional covers/attachments. CSV and JSON exports ensure portability outside BookStats.

## 14. Interface direction

BookStats should look like a personal library and data-exploration tool rather than a generic SaaS/admin dashboard. Cover art supplies much of the color while application chrome stays restrained. Use strong typography, generous spacing, subtle borders/shadows, responsive layouts, accessible controls and restrained animation.

Initial themes: Light, Dark and System. Later options may include Warm Paper, Midnight, high contrast and custom accents/density.

Primary navigation: Library, Reading, Statistics, Pick a Book; secondary actions include Add Books, Import/Export, Settings and Account.

## 15. Performance and accessibility

The library must remain comfortable around 10,000 books and should scale beyond 50,000 records through virtualization, indexes, lazy cover loading, thumbnails and cached statistics. Controls must be keyboard accessible; charts should expose textual/tabular equivalents when practical; reduced motion and sufficient contrast are requirements.

## 16. Development roadmap

### Milestone 0 — Foundation
Monorepo, React/Vite client, Tauri shell, Fastify API, PostgreSQL schema, local persistence, shared domain package, statistics package, test framework and initial design system.

### Milestone 1 — Core Library
Work/edition model, library grid/table, add/edit/delete, search, filters, sorting, tags, shelves, ownership and ratings.

### Milestone 2 — Accounts & Sync
Registration/login, local/cloud identity, desktop SQLite, sync cursor, revisions, tombstones and conflict handling.

### Milestone 3 — Metadata
Provider interface, Open Library, secondary provider, ISBN/title/author search, edition selection, covers/descriptions and overrides.

### Milestone 4 — Imports
Preview/dedup framework plus Goodreads, LibraryThing, legacy BookStats and generic CSV importers.

### Milestone 5 — Reading History
Reading sessions, start/finish dates, rereads, DNF, reviews, notes and reading timeline.

### Milestone 6 — Statistics
Recreate the original statistics, then add interactive filtering, historical analysis and richer charts.

### Milestone 7 — Customization & Polish
Custom dashboards, themes, saved filters, configurable views, shortcuts, accessibility and performance work.

### Milestone 8 — Fun Features
Pick My Next Book, Year in Books, library cover mosaic, TBR archaeology, rediscovery, reading challenges and library time machine.

## 17. Definition of BookStats 1.0

A user can install BookStats on Windows/macOS or open the web app; use local mode or an account; import an existing library; find books by ISBN/title/author and select editions; edit retrieved metadata; organize shelves/tags/ownership; track readings and rereads; rate/review books; browse and filter large collections; explore meaningful visual statistics; synchronize desktop/web libraries; and export their complete data at any time.
