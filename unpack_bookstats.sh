#!/usr/bin/env bash
set -Eeuo pipefail

DOWNLOADS_DIR="${BOOKSTATS_DOWNLOADS_DIR:-/home/kcarroll/Downloads}"
WEB_DIR="${BOOKSTATS_WEB_DIR:-/var/www/kylecarroll/bookstats}"
PUBLIC_DOWNLOADS_DIR="${BOOKSTATS_PUBLIC_DOWNLOADS_DIR:-/var/www/kylecarroll/downloads/bookstats}"
SERVER_DIR="${BOOKSTATS_SERVER_DIR:-/opt/bookstats}"
SERVICE_NAME="${BOOKSTATS_SERVICE_NAME:-bookstats}"
RUN_USER="${BOOKSTATS_RUN_USER:-bookstats}"
TMP_ROOT=""

say() { printf '\n==> %s\n' "$*"; }
die() { printf '\nERROR: %s\n' "$*" >&2; exit 1; }
cleanup() { [[ -n "$TMP_ROOT" && -d "$TMP_ROOT" ]] && rm -rf "$TMP_ROOT"; }
trap cleanup EXIT

if [[ ${EUID:-$(id -u)} -ne 0 ]]; then
  exec sudo -E "$0" "$@"
fi

for command in unzip rsync npm systemctl nginx sort find curl tar node; do
  command -v "$command" >/dev/null 2>&1 || die "Required command '$command' was not found."
done

[[ -d "$DOWNLOADS_DIR" ]] || mkdir -p "$DOWNLOADS_DIR"
[[ -d "$PUBLIC_DOWNLOADS_DIR" ]] || mkdir -p "$PUBLIC_DOWNLOADS_DIR"

latest_match_in() {
  local directory="$1"
  local pattern="$2"
  find "$directory" -maxdepth 1 -type f -name "$pattern" -printf '%f\n' 2>/dev/null | sort -V | tail -n 1
}

# Prefer the traditional staging directory. If it has no BookStats release,
# fall back to the public downloads directory used by upload_bookstats_releases.sh.
SOURCE_DIR="$DOWNLOADS_DIR"

web_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Web-v*.zip')"
server_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Server-v*.zip')"
if [[ -z "$server_name" ]]; then
  server_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Server-v*.tar.gz')"
fi
updater_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Updater-v*.zip')"

if [[ -z "$web_name" && -z "$server_name" && -z "$updater_name" ]]; then
  SOURCE_DIR="$PUBLIC_DOWNLOADS_DIR"
  web_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Web-v*.zip')"
  server_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Server-v*.zip')"
  if [[ -z "$server_name" ]]; then
    server_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Server-v*.tar.gz')"
  fi
  updater_name="$(latest_match_in "$SOURCE_DIR" 'BookStats-Updater-v*.zip')"
fi

has_app_release=false
if [[ -n "$web_name" || -n "$server_name" ]]; then
  [[ -n "$web_name" ]] || die "Found a server bundle but no matching BookStats-Web-v*.zip in $SOURCE_DIR"
  [[ -n "$server_name" ]] || die "Found a web bundle but no matching BookStats-Server-v*.zip (or legacy .tar.gz) in $SOURCE_DIR"
  has_app_release=true
fi

# The direct upload workflow also publishes loose updater artifacts directly
# into PUBLIC_DOWNLOADS_DIR. If those exist, an updater ZIP is optional here.
has_published_updater=false
if find "$PUBLIC_DOWNLOADS_DIR" -maxdepth 1 -type f -name 'latest-*.json' -print -quit 2>/dev/null | grep -q .; then
  has_published_updater=true
fi

[[ "$has_app_release" == true || -n "$updater_name" || "$has_published_updater" == true ]] || die \
  "No BookStats release found in $DOWNLOADS_DIR or $PUBLIC_DOWNLOADS_DIR"

say "Using release source: $SOURCE_DIR"

release_version=""
if [[ "$has_app_release" == true ]]; then
  web_version="${web_name#BookStats-Web-v}"; web_version="${web_version%.zip}"
  server_version="${server_name#BookStats-Server-v}"; server_version="${server_version%.zip}"; server_version="${server_version%.tar.gz}"
  [[ "$web_version" == "$server_version" ]] || die "Bundle versions do not match: web=$web_version server=$server_version"
  release_version="$web_version"
fi

updater_version=""
if [[ -n "$updater_name" ]]; then
  updater_version="${updater_name#BookStats-Updater-v}"; updater_version="${updater_version%.zip}"
  if [[ -n "$release_version" ]]; then
    [[ "$release_version" == "$updater_version" ]] || die \
      "Updater version does not match the web/server release: app=$release_version updater=$updater_version"
  else
    release_version="$updater_version"
  fi
fi

TMP_ROOT="$(mktemp -d /tmp/bookstats-deploy.XXXXXX)"
WEB_STAGE="$TMP_ROOT/web"
SERVER_STAGE="$TMP_ROOT/server"
UPDATER_STAGE="$TMP_ROOT/updater"
mkdir -p "$WEB_STAGE" "$SERVER_STAGE" "$UPDATER_STAGE"

say "Deploying BookStats v$release_version"

if [[ "$has_app_release" == true ]]; then
  WEB_ARCHIVE="$SOURCE_DIR/$web_name"
  SERVER_ARCHIVE="$SOURCE_DIR/$server_name"

  say "Staging web bundle: $web_name"
  unzip -q "$WEB_ARCHIVE" -d "$WEB_STAGE"
  [[ -f "$WEB_STAGE/index.html" ]] || die "Web archive does not contain index.html at its root."
  [[ -f "$WEB_STAGE/admin/index.html" ]] || die "Web archive does not contain the administrator bundle at admin/index.html."

  say "Staging server bundle: $server_name"
  if [[ "$SERVER_ARCHIVE" == *.zip ]]; then
    unzip -q "$SERVER_ARCHIVE" -d "$SERVER_STAGE"
  else
    tar -xzf "$SERVER_ARCHIVE" -C "$SERVER_STAGE"
  fi
  [[ -f "$SERVER_STAGE/package.json" ]] || die "Server archive does not contain package.json at its root."
  [[ -f "$SERVER_STAGE/apps/server/package.json" ]] || die "Server archive is missing apps/server/package.json."
fi

if [[ -n "$updater_name" ]]; then
  UPDATER_ARCHIVE="$SOURCE_DIR/$updater_name"
  say "Staging signed desktop updater bundle: $updater_name"
  unzip -q "$UPDATER_ARCHIVE" -d "$UPDATER_STAGE"

  mapfile -t updater_manifests < <(find "$UPDATER_STAGE" -maxdepth 1 -type f -name 'latest-*.json' -print | sort)
  [[ ${#updater_manifests[@]} -gt 0 ]] || die "Updater archive contains no latest-*.json manifests."
  for manifest in "${updater_manifests[@]}"; do
    node - "$manifest" "$updater_version" <<'NODE'
const fs = require('fs');
const [manifestPath, expectedVersion] = process.argv.slice(2);
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
if (manifest.version !== expectedVersion) {
  throw new Error(`${manifestPath}: manifest version ${manifest.version} does not match ${expectedVersion}`);
}
const platforms = Object.values(manifest.platforms || {});
if (!platforms.length || platforms.some((entry) => !entry?.url || !entry?.signature)) {
  throw new Error(`${manifestPath}: updater manifest is missing a signed platform URL`);
}
NODE
  done
fi

# Publish the signed updater before the API begins reporting the new app version.
# This avoids even a short deployment window where existing desktop clients are
# told to update before their platform manifest is available.
if [[ -n "$updater_name" ]]; then
  say "Publishing signed desktop updater files"
  mkdir -p "$PUBLIC_DOWNLOADS_DIR"
  # Do not --delete here: older signed binaries may remain useful while clients transition.
  rsync -a "$UPDATER_STAGE/" "$PUBLIC_DOWNLOADS_DIR/"
  find "$PUBLIC_DOWNLOADS_DIR" -type d -exec chmod 755 {} +
  find "$PUBLIC_DOWNLOADS_DIR" -type f -exec chmod 644 {} +
elif [[ "$has_published_updater" == true ]]; then
  say "Signed desktop updater files are already published"
fi

if [[ "$has_app_release" == true ]]; then
  say "Installing web application"
  mkdir -p "$WEB_DIR"
  rsync -a --delete "$WEB_STAGE/" "$WEB_DIR/"
  find "$WEB_DIR" -type d -exec chmod 755 {} +
  find "$WEB_DIR" -type f -exec chmod 644 {} +

  say "Installing BookStats server"
  mkdir -p "$SERVER_DIR"
  systemctl stop "$SERVICE_NAME" 2>/dev/null || true
  # Preserve the production .env and durable cover asset directory while replacing release-managed files.
  rsync -a --delete --exclude='.env' --exclude='data/' "$SERVER_STAGE/" "$SERVER_DIR/"
  mkdir -p "$SERVER_DIR/data/covers"
  chown -R "$RUN_USER:$RUN_USER" "$SERVER_DIR"
  if [[ -f "$SERVER_DIR/.env" ]]; then
    chown "$RUN_USER:$RUN_USER" "$SERVER_DIR/.env"
    chmod 600 "$SERVER_DIR/.env"
  else
    die "$SERVER_DIR/.env is missing. Restore the production environment file before starting BookStats."
  fi

  say "Installing production Node dependencies"
  sudo -u "$RUN_USER" sh -c "cd '$SERVER_DIR' && npm install --omit=dev"

  say "Applying database migrations"
  sudo -u "$RUN_USER" sh -c "cd '$SERVER_DIR' && npm run db:migrate"

  say "Restarting BookStats"
  systemctl restart "$SERVICE_NAME"
  systemctl --no-pager --full status "$SERVICE_NAME" >/dev/null || die "BookStats service failed to start."
fi

if [[ "$has_app_release" == true ]]; then
  say "Validating and restarting NGINX"
  nginx -t
  systemctl restart nginx

  port="$(awk -F= '/^[[:space:]]*BOOKSTATS_PORT=/{gsub(/[[:space:]"]/, "", $2); print $2; exit}' "$SERVER_DIR/.env" 2>/dev/null || true)"
  port="${port:-8787}"
  say "Checking BookStats API on 127.0.0.1:$port"
  for _ in 1 2 3 4 5; do
    if curl -fsS "http://127.0.0.1:${port}/health" >/dev/null; then
      break
    fi
    sleep 1
  done
  curl -fsS "http://127.0.0.1:${port}/health" || die "BookStats health check failed after deployment."
  printf '\n'
fi

if [[ "$SOURCE_DIR" == "$DOWNLOADS_DIR" ]]; then
  say "Removing deployed staging archives from Downloads"
  if [[ "$has_app_release" == true ]]; then
    rm -f "$WEB_ARCHIVE" "$SERVER_ARCHIVE"
  fi
  if [[ -n "$updater_name" ]]; then
    rm -f "$UPDATER_ARCHIVE"
  fi
else
  say "Keeping release archives in the public downloads directory"
fi

say "BookStats v$release_version deployment complete"
if [[ "$has_app_release" == true ]]; then
  printf 'Web:     https://kylecarroll.com/bookstats/\n'
  printf 'API:     https://kylecarroll.com/bookstats/api/v1\n'
  printf 'Server:  %s\n' "$SERVER_DIR"
fi
if [[ -n "$updater_name" ]]; then
  printf 'Updates: https://kylecarroll.com/downloads/bookstats/\n'
  printf 'Files:   %s\n' "$PUBLIC_DOWNLOADS_DIR"
fi
