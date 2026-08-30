# BookStats Server Setup (v1.2.1)

BookStats uses one public site and one private local API process:

- `https://kylecarroll.com/bookstats/` — static Vite web application
- `https://kylecarroll.com/bookstats/admin/` — separate administrator frontend (not linked from the normal client)
- `https://kylecarroll.com/bookstats/api/` — NGINX proxy to the Fastify API
- `127.0.0.1:<BOOKSTATS_PORT>` — private Fastify listener
- PostgreSQL — accounts, cloud libraries, synchronization, metadata cache, cover-asset references, account-security tokens, administrator roles and audit history
- `/opt/bookstats/data/covers/` — durable copies of user-selected cover images (default location)
- `/var/www/kylecarroll/downloads/bookstats/` — signed desktop updater manifests and packages

The client remains usable without an account, but cloud synchronization now requires a verified email address.

## Environment

The production `/opt/bookstats/.env` should contain at least:

```dotenv
BOOKSTATS_HOST=127.0.0.1
BOOKSTATS_PORT=8790
DATABASE_URL=postgresql://bookstats:CHANGE_ME@127.0.0.1:5432/bookstats
BOOKSTATS_METADATA_USER_AGENT=BookStats/1.2.5 (your-real-contact@example.com)
BOOKSTATS_CORS_ORIGIN=https://kylecarroll.com,tauri://localhost,http://tauri.localhost
# HTTP(S) origins automatically allow their matching www/non-www counterpart.
BOOKSTATS_PUBLIC_URL=https://kylecarroll.com/bookstats/
# Optional override; the deployment helper preserves the default directory across upgrades.
BOOKSTATS_COVER_DIR=/opt/bookstats/data/covers
# Optional metadata enrichments; keep these on the server only.
BOOKSTATS_GOOGLE_BOOKS_API_KEY=your_google_books_api_key
BOOKSTATS_HARDCOVER_API_TOKEN=your_hardcover_api_token
```

### Metadata providers — Google Books + Hardcover

Open Library works without credentials and remains the fallback provider. For the full multi-provider lookup stack, configure Google Books and Hardcover on the API server:

1. In Google Cloud Console, enable the **Books API** for a project.
2. Create an API key and restrict it to the Books API where practical.
3. Put the key in `BOOKSTATS_GOOGLE_BOOKS_API_KEY`.
4. In Hardcover, create an API token from your account API settings.
5. Put that token in `BOOKSTATS_HARDCOVER_API_TOKEN`. Treat it like a password.
6. Restart `bookstats.service` and inspect `/health`.

Neither credential belongs in `VITE_*` variables or the static web directory. The browser talks only to the BookStats API. A missing optional provider does not take lookup offline; `/health` reports which providers are configured.

ISBN lookup is intentionally strict: BookStats accepts only the requested ISBN (including its equivalent ISBN-10/ISBN-13 form) for edition-level results. Work-level series/description enrichment may be layered on afterward, but cannot replace the matched edition's ISBN, publisher, pagination, format or publication metadata.

### Transactional email — Resend

BookStats sends verification, password-reset and password-change messages through the Resend HTTPS API. The BookStats server does not need the password for `bookstats@kylecarroll.com`, and MFA on the GoDaddy/Microsoft 365 mailbox can remain enabled.

Set up Resend once:

1. Create a Resend account and add `kylecarroll.com` as a sending domain.
2. Add the DNS records Resend shows to the DNS zone that is authoritative for `kylecarroll.com` (typically the SPF and DKIM records shown by Resend).
3. Wait for the domain to show as verified in Resend.
4. Create a Resend API key with permission to send email.
5. Add the values below to `/opt/bookstats/.env`.

```dotenv
BOOKSTATS_RESEND_API_KEY=re_your_api_key
BOOKSTATS_EMAIL_FROM=BookStats <bookstats@kylecarroll.com>
BOOKSTATS_EMAIL_REPLY_TO=bookstats@kylecarroll.com
# Optional; defaults to BOOKSTATS_EMAIL_REPLY_TO when omitted.
BOOKSTATS_FEEDBACK_TO=bookstats@kylecarroll.com
```

`BOOKSTATS_EMAIL_REPLY_TO` is optional. `BOOKSTATS_EMAIL_FROM` must use a sender on a domain verified in Resend. The sender address does not need to be the mailbox credentials BookStats uses; Resend is responsible for authenticated outbound delivery.

After changing the environment, restart BookStats and check `/health`:

```bash
sudo systemctl restart bookstats
curl -s http://127.0.0.1:8790/health
```

A configured server reports `"emailConfigured":true`, `"emailProvider":"resend"`, `"feedbackConfigured":true`, plus a `metadataProviders` array showing Open Library and whether Google Books/Hardcover are configured.

If Resend is not configured, normal local use and metadata lookup continue to work. Accounts can still be created, but they cannot become verified and therefore cannot cloud-sync until email delivery is configured.


## Database migrations

Apply every migration after installing a new server bundle:

```bash
cd /opt/bookstats
sudo -u bookstats npm run db:migrate
```

Migration `0004_account_security.sql` adds verification and password-reset token tables. Accounts created before v0.4 are grandfathered as verified so the upgrade does not disable existing cloud libraries. Migration `0005_shelves.sql` adds the record-type discriminator used to synchronize shelves alongside books. Migration `0006_admin_console.sql` adds administrator roles, account suspension state, and the append-only administrator audit log. Migration `0007_cover_assets.sql` adds the durable cover-asset index; image bytes live in the server cover directory rather than PostgreSQL JSON. Smart shelves, reading sessions, goals, metadata provenance and series-catalog data remain stored inside the existing synchronized JSON records.


## Durable cover archive (v1.0.1)

For verified accounts, newly selected catalog/URL/custom covers are archived automatically. The default filesystem location is:

```text
/opt/bookstats/data/covers/
```

The deployment helper preserves `/opt/bookstats/data/` across server upgrades. If `BOOKSTATS_COVER_DIR` points somewhere else, ensure the `bookstats` service user can create/read/write files there. Back up this directory together with PostgreSQL; the database alone contains references and metadata, not the image bytes.

If `/usr/local/sbin/unpack_bookstats` was installed before v1.0.1, replace it **immediately after deploying v1.0.1 and before running the cover migration**. Older copies do not know that `/opt/bookstats/data/` has become persistent application data:

```bash
sudo install -m 755 /opt/bookstats/tools/unpack_bookstats.sh /usr/local/sbin/unpack_bookstats
```

This is important because older deploy helpers used `rsync --delete` without preserving the new `data/` directory.

Existing pre-1.0.1 selections are intentionally not bulk-downloaded during an ordinary sync. After deploying v1.0.1 and migration `0007_cover_assets.sql`, inspect the one-time migration first:

```bash
cd /opt/bookstats
sudo -u bookstats npm run covers:migrate -- --email you@example.com
```

If the dry run looks correct, close BookStats on your other devices for the migration, then archive the covers:

```bash
sudo -u bookstats npm run covers:migrate -- --email you@example.com --apply
```

The utility is restartable and idempotent. Its dry run also checks already-archived records for missing asset files. With `--apply`, missing files are repaired from the retained source URL when possible; healthy assets are never replaced. Failed external URLs are left recoverable and reported. If an external URL has died but a client device still has that selected cover in its local cache, editing/saving that book in v1.0.1 can archive the cached image as a recovery fallback. Custom data-URL covers are moved out of the synchronized JSON after they are archived successfully.


## Administrator console

The v0.10 web archive contains a second Vite application at:

```text
https://kylecarroll.com/bookstats/admin/
```

The regular BookStats application does not link to or expose the administrator console. Authorization is enforced by the API on every `/api/v1/admin/*` request, so knowing the URL is not sufficient for access.

After deploying v0.10 and applying migration `0006_admin_console.sql`, grant the first administrator role from the server only:

```bash
cd /opt/bookstats
sudo -u bookstats npm run admin:user -- grant kyle.carroll@icloud.com
```

Inspect or revoke it with:

```bash
sudo -u bookstats npm run admin:user -- status kyle.carroll@icloud.com
sudo -u bookstats npm run admin:user -- revoke kyle.carroll@icloud.com
```

There is intentionally no HTTP endpoint for role promotion. The admin console can inspect aggregate database/server health, manage accounts and sessions, inspect/repair synchronized records, clear a user's cloud library, permanently delete a user, and review its own audit trail. It cannot run arbitrary SQL, shell commands, restart services, or manage NGINX/systemd.

## NGINX

Inside the existing HTTPS `server {}` block for `kylecarroll.com`:

```nginx
location = /bookstats {
    return 301 /bookstats/;
}

location = /bookstats/admin {
    return 301 /bookstats/admin/;
}

location ^~ /bookstats/admin/ {
    root /var/www/kylecarroll;
    try_files $uri $uri/ /bookstats/admin/index.html;
}

location ^~ /bookstats/api/ {
    proxy_pass http://127.0.0.1:8790/api/;
    proxy_http_version 1.1;

    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;

    client_max_body_size 25m;
    proxy_read_timeout 60s;
}

location ^~ /bookstats/ {
    root /var/www/kylecarroll;
    try_files $uri $uri/ /bookstats/index.html;
}
```

Adjust `8790` if `BOOKSTATS_PORT` uses another local port.

The `25m` request limit is an emergency ceiling, not the expected sync size. BookStats keeps a persistent local outbox and normally sends only changed records in batches capped at roughly 900 KiB and 100 records. v1.2 also applies bounded client timeouts plus capped exponential retry/backoff for transient failures and records structured batch telemetry server-side. Retaining a larger reverse-proxy ceiling protects unusual single records and upgrade/recovery cases without making normal edits large transfers.

## systemd

Example `/etc/systemd/system/bookstats.service`:

```ini
[Unit]
Description=BookStats API Server
After=network-online.target postgresql.service
Wants=network-online.target

[Service]
Type=simple
User=bookstats
Group=bookstats
WorkingDirectory=/opt/bookstats
Environment=NODE_ENV=production
ExecStart=/usr/bin/npm start
Restart=on-failure
RestartSec=3
NoNewPrivileges=true
PrivateTmp=true

[Install]
WantedBy=multi-user.target
```

Apply changes with:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now bookstats
```

## Release/deployment workflow

Local `export.sh` produces matching server/web ZIPs:

```text
BookStats-Web-v1.2.5.zip
BookStats-Server-v1.2.5.zip
```

A signed desktop build also produces a deployment wrapper such as:

```text
BookStats-Updater-v1.0.3.zip
```

That updater ZIP is only a convenient server-transfer bundle. Desktop clients do **not** download the wrapper ZIP. They fetch the platform-specific `latest-<target>-<arch>.json` manifest and the signed `.app.tar.gz` (macOS) or NSIS `.exe` (Windows) referenced by that manifest.

Before the first v1.0.3 desktop release, generate the updater signing key on the trusted build machine:

```bash
./tools/setup-updater-key.sh
```

Keep the generated private key off the server and out of Git. Only the public key belongs in `tauri.conf.json`. Future desktop release builds require `TAURI_SIGNING_PRIVATE_KEY` (and `TAURI_SIGNING_PRIVATE_KEY_PASSWORD` when the key has a password). The same signing key must be retained for later releases or installed clients will reject replacement updates.

The source archive intentionally does not contain that long-lived public/private keypair. If a later full source replacement restores the placeholder in `tauri.conf.json`, `export.sh` automatically re-embeds `$HOME/.config/bookstats/updater.key.pub` before a desktop build. Set `BOOKSTATS_UPDATER_PUBLIC_KEY_PATH` if the `.pub` file is stored elsewhere.

Copy the web/server bundles and, when available, the updater bundle into:

```text
/home/kcarroll/Downloads/
```

Install the repository's `unpack_bookstats.sh` on the server once, for example:

```bash
sudo install -m 755 unpack_bookstats.sh /usr/local/sbin/unpack_bookstats
```

Future deployments are then:

```bash
sudo unpack_bookstats
```

The helper:

1. selects the highest-version Web and Server archives when present and refuses mismatched versions;
2. accepts `BookStats-Updater-v*.zip` alongside them, or as an updater-only deployment;
3. validates all archives before touching production and checks updater manifest version/signature fields;
4. publishes signed updater files to `/var/www/kylecarroll/downloads/bookstats/` first, without deleting older updater packages, so the new update is available before the API reports the new version;
5. replaces `/var/www/kylecarroll/bookstats/`;
6. replaces release-managed files under `/opt/bookstats/` while preserving `.env` and the durable `data/` asset directory;
7. installs production npm dependencies and applies PostgreSQL migrations;
8. restarts the `bookstats` service, validates NGINX, and checks the local API health endpoint;
9. removes the consumed transfer ZIPs from Downloads after a successful deployment.

If your website's `/downloads/` directory is elsewhere, override it when running the helper:

```bash
BOOKSTATS_PUBLIC_DOWNLOADS_DIR=/your/web/root/downloads/bookstats sudo -E unpack_bookstats
```

The public updater files are ordinary HTTPS static assets. `latest-*.json` should be served as JSON and the `.sig`, `.app.tar.gz`, and `.exe` files must remain byte-for-byte unchanged after export. There is no updater private key on the server.

**One-time bootstrap:** the v1.0.2 desktop application predates the updater plugin, so it cannot install v1.0.3 itself. Users must manually install v1.0.3 once. Beginning with v1.0.3, future desktop releases can be applied from BookStats' blocking update dialog.

The helper also accepts the old `BookStats-Server-v*.tar.gz` format while upgrading older installations.

## Account behavior

### Web

When a user signs out, BookStats attempts one final sync and then clears the IndexedDB book/tombstone cache. This prevents a second person using the same browser from seeing the previous user's cloud library. The account-specific sync cursor is reset so signing in again pulls the full cloud library.

### Desktop

Signing out revokes the cloud session but deliberately keeps the local SQLite library. This preserves BookStats' local-first/offline behavior. Cloud synchronization resumes only after a verified account signs in again.

## Account email links

Verification emails open:

```text
https://kylecarroll.com/bookstats/?verify=<token>
```

Password-reset emails open:

```text
https://kylecarroll.com/bookstats/?reset=<token>
```

The React application detects those parameters and opens the Account screen automatically. Verification links expire after 24 hours; password-reset links expire after one hour. Password resets invalidate all existing sessions for that account.
