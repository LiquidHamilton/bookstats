#!/usr/bin/env bash
set -Eeuo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

TARGET="${1:-all}"
PUBLIC_API_URL="${BOOKSTATS_PUBLIC_API_URL:-https://kylecarroll.com/bookstats/api/v1}"
WEB_BASE="${BOOKSTATS_WEB_BASE:-/bookstats/}"
RELEASE_DIR="$ROOT_DIR/dist/releases"
STAGING_DIR="$ROOT_DIR/dist/.release-staging"
UPDATER_DIR="$RELEASE_DIR/updater"
UPDATER_BASE_URL="${BOOKSTATS_UPDATER_BASE_URL:-https://kylecarroll.com/downloads/bookstats}"
UPDATER_PUBLIC_KEY_PATH="${BOOKSTATS_UPDATER_PUBLIC_KEY_PATH:-$HOME/.config/bookstats/updater.key.pub}"
VERSION="$(node -p "require('./package.json').version")"
export TAURI_SIGNING_PRIVATE_KEY="/Users/kylecarroll/.config/bookstats/updater.key"
export TAURI_SIGNING_PRIVATE_KEY_PASSWORD=""
rm -rf dist/releases/*

say() { printf '\n==> %s\n' "$*"; }
die() { printf '\nERROR: %s\n' "$*" >&2; exit 1; }
need() { command -v "$1" >/dev/null 2>&1 || die "Required command '$1' was not found."; }

usage() {
  cat <<USAGE
BookStats release exporter

Usage:
  ./export.sh [all|web|server|mac|windows|desktop|clean]

Environment overrides:
  BOOKSTATS_PUBLIC_API_URL   API embedded in web/desktop builds
                             default: https://kylecarroll.com/bookstats/api/v1
  BOOKSTATS_WEB_BASE         Vite public base for web deployment
                             default: /bookstats/
  BOOKSTATS_UPDATER_BASE_URL Public URL containing signed desktop updater artifacts
                             default: https://kylecarroll.com/downloads/bookstats
  BOOKSTATS_UPDATER_PUBLIC_KEY_PATH
                             Existing updater public key to embed when tauri.conf.json
                             still contains the source-archive placeholder
                             default: $HOME/.config/bookstats/updater.key.pub

Desktop release signing:
  Run ./tools/setup-updater-key.sh once, then export TAURI_SIGNING_PRIVATE_KEY
  and TAURI_SIGNING_PRIVATE_KEY_PASSWORD before mac/windows/desktop/all builds.

Outputs are written to:
  dist/releases/

Examples:
  ./export.sh all
  ./export.sh web
  ./export.sh server
  ./export.sh desktop
  BOOKSTATS_PUBLIC_API_URL=https://example.com/bookstats/api/v1 ./export.sh all
USAGE
}

case "$TARGET" in
  -h|--help|help) usage; exit 0 ;;
  all|web|server|mac|windows|desktop|clean) ;;
  *) usage; die "Unknown target '$TARGET'." ;;
esac

need node
need npm
need zip
need tar

if [[ "$TARGET" == "clean" ]]; then
  say "Cleaning release output"
  rm -rf "$RELEASE_DIR" "$STAGING_DIR"
  exit 0
fi

[[ -d node_modules ]] || die "node_modules is missing. Run 'npm install' first."
mkdir -p "$RELEASE_DIR"
rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR"

if [[ "$TARGET" == "all" || "$TARGET" == "desktop" || "$TARGET" == "mac" || "$TARGET" == "windows" ]]; then
  rm -rf "$UPDATER_DIR"
fi
mkdir -p "$UPDATER_DIR"


require_updater_signing() {
  local configured_pubkey
  configured_pubkey="$(node -e 'const c=require("./apps/client/src-tauri/tauri.conf.json"); process.stdout.write(c.plugins?.updater?.pubkey || "")')"

  if [[ -z "$configured_pubkey" || "$configured_pubkey" == "BOOKSTATS_UPDATER_PUBLIC_KEY_NOT_CONFIGURED" ]]; then
    if [[ -s "$UPDATER_PUBLIC_KEY_PATH" ]]; then
      say "Restoring the existing updater public key into Tauri configuration"
      node tools/configure-updater-key.mjs "$UPDATER_PUBLIC_KEY_PATH"
      configured_pubkey="$(node -e 'const c=require("./apps/client/src-tauri/tauri.conf.json"); process.stdout.write(c.plugins?.updater?.pubkey || "")')"
    else
      die "Desktop auto-update is not configured. Run ./tools/setup-updater-key.sh once, or set BOOKSTATS_UPDATER_PUBLIC_KEY_PATH to your existing .pub key."
    fi
  fi

  [[ -n "${TAURI_SIGNING_PRIVATE_KEY:-}" ]] || die \
    "TAURI_SIGNING_PRIVATE_KEY is not set. Point it at your private updater key before desktop release builds."
}

write_updater_manifest() {
  local platform="$1" artifact="$2" signature="$3" public_name="$4"
  [[ -s "$artifact" ]] || die "Updater artifact missing: $artifact"
  [[ -s "$signature" ]] || die "Updater signature missing: $signature"
  cp "$artifact" "$UPDATER_DIR/$public_name"
  cp "$signature" "$UPDATER_DIR/$public_name.sig"
  node tools/write-updater-manifest.mjs \
    "$UPDATER_DIR/latest-${platform}.json" \
    "$VERSION" \
    "$platform" \
    "${UPDATER_BASE_URL%/}/$public_name" \
    "$signature"
}

package_updater_bundle() {
  local manifests
  manifests="$(find "$UPDATER_DIR" -maxdepth 1 -type f -name 'latest-*.json' -print -quit 2>/dev/null || true)"
  [[ -n "$manifests" ]] || die "No updater manifests were generated."

  local out="$RELEASE_DIR/BookStats-Updater-v${VERSION}.zip"
  rm -f "$out"
  (cd "$UPDATER_DIR" && zip -qry "$out" .)
  say "Created ${out#$ROOT_DIR/} (server deployment bundle)"
}

validate() {
  say "Validating BookStats v$VERSION"
  npm run typecheck
  npm test
}

build_web() {
  say "Building web bundle for $WEB_BASE"
  rm -rf apps/client/dist apps/admin/dist
  VITE_BOOKSTATS_API_URL="$PUBLIC_API_URL" \
    npm exec -w @bookstats/client -- vite build --base="$WEB_BASE"

  say "Building separate administrator bundle for ${WEB_BASE}admin/"
  VITE_BOOKSTATS_API_URL="$PUBLIC_API_URL" \
    npm exec -w @bookstats/admin -- vite build --base="${WEB_BASE}admin/"
  mkdir -p apps/client/dist/admin
  cp -R apps/admin/dist/. apps/client/dist/admin/

  local out="$RELEASE_DIR/BookStats-Web-v${VERSION}.zip"
  rm -f "$out"
  (cd apps/client/dist && zip -qry "$out" .)
  say "Created ${out#$ROOT_DIR/}"
}

build_server() {
  say "Building server bundle"
  npm run build -w @bookstats/domain
  npm run build -w @bookstats/server

  local stage="$STAGING_DIR/server"
  rm -rf "$stage"
  mkdir -p \
    "$stage/apps/server" \
    "$stage/packages/domain" \
    "$stage/database" \
    "$stage/tools" \
    "$stage/docs"

  cp -R apps/server/dist "$stage/apps/server/dist"
  cp apps/server/package.json "$stage/apps/server/package.json"
  cp -R packages/domain/dist "$stage/packages/domain/dist"
  cp packages/domain/package.json "$stage/packages/domain/package.json"
  cp -R database/migrations "$stage/database/migrations"
  cp tools/migrate-db.mjs "$stage/tools/migrate-db.mjs"
  cp tools/admin-user.mjs "$stage/tools/admin-user.mjs"
  cp tools/migrate-cover-assets.mjs "$stage/tools/migrate-cover-assets.mjs"
  cp unpack_bookstats.sh "$stage/tools/unpack_bookstats.sh"
  cp .env.example "$stage/.env.example"
  cp docs/SERVER_SETUP.md "$stage/docs/SERVER_SETUP.md"

  cat > "$stage/package.json" <<JSON
{
  "name": "bookstats-server-release",
  "version": "$VERSION",
  "private": true,
  "type": "module",
  "workspaces": [
    "apps/server",
    "packages/domain"
  ],
  "scripts": {
    "start": "npm run start -w @bookstats/server",
    "db:migrate": "node tools/migrate-db.mjs",
    "admin:user": "node tools/admin-user.mjs",
    "covers:migrate": "node tools/migrate-cover-assets.mjs"
  },
  "engines": {
    "node": ">=20"
  }
}
JSON

  cat > "$stage/INSTALL.txt" <<'TXT'
BookStats Server Bundle

1. Extract under /opt/bookstats.
2. Run: npm install --omit=dev
3. Copy .env.example to .env and edit DATABASE_URL, CORS origin, metadata User-Agent, public URL, and Resend email settings.
4. Run: npm run db:migrate
5. Start with: npm start
6. To grant the first administrator: npm run admin:user -- grant you@example.com
7. v1.0.1+: install tools/unpack_bookstats.sh as your deploy helper so data/covers is preserved.
8. Existing accounts can dry-run cover archival with: npm run covers:migrate -- --email you@example.com

See docs/SERVER_SETUP.md for the complete deployment procedure.
TXT

  local out="$RELEASE_DIR/BookStats-Server-v${VERSION}.zip"
  rm -f "$out"
  (cd "$stage" && zip -qry "$out" .)
  say "Created ${out#$ROOT_DIR/}"
}

build_mac() {
  [[ "$(uname -s)" == "Darwin" ]] || die "The macOS desktop bundle must be built on macOS."
  need cargo
  need rustc

  say "Building macOS desktop bundle"
  (
    cd apps/client
    CI=true \
    VITE_BOOKSTATS_API_URL="$PUBLIC_API_URL" \
      npx tauri build --bundles app,dmg
  )

  local bundle_root="$ROOT_DIR/apps/client/src-tauri/target/release/bundle"
  local app_path dmg_path updater_path updater_sig
  app_path="$(find "$bundle_root/macos" -maxdepth 1 -name '*.app' -print -quit 2>/dev/null || true)"
  dmg_path="$(find "$bundle_root/dmg" -maxdepth 1 -name '*.dmg' -print -quit 2>/dev/null || true)"
  updater_path="$(find "$bundle_root/macos" -maxdepth 1 -name '*.app.tar.gz' -print -quit 2>/dev/null || true)"
  updater_sig="${updater_path}.sig"

  [[ -n "$app_path" ]] || die "Tauri completed but no .app bundle was found."
  [[ -n "$updater_path" && -s "$updater_sig" ]] || die "Tauri completed but no signed macOS updater artifact was found."

  local zip_out="$RELEASE_DIR/BookStats-macOS-v${VERSION}.zip"
  rm -f "$zip_out"
  if command -v ditto >/dev/null 2>&1; then
    ditto -c -k --sequesterRsrc --keepParent "$app_path" "$zip_out"
  else
    (cd "$(dirname "$app_path")" && zip -qry "$zip_out" "$(basename "$app_path")")
  fi
  say "Created ${zip_out#$ROOT_DIR/}"

  if [[ -n "$dmg_path" ]]; then
    cp "$dmg_path" "$RELEASE_DIR/BookStats-macOS-v${VERSION}.dmg"
    say "Created dist/releases/BookStats-macOS-v${VERSION}.dmg"
  fi

  local mac_arch platform
  mac_arch="$(uname -m)"
  case "$mac_arch" in
    arm64|aarch64) platform="darwin-aarch64" ;;
    x86_64|amd64) platform="darwin-x86_64" ;;
    *) die "Unsupported macOS updater architecture: $mac_arch" ;;
  esac
  write_updater_manifest "$platform" "$updater_path" "$updater_sig" "BookStats-macOS-v${VERSION}.app.tar.gz"
}

prepare_windows_cross_compile() {
  if [[ "$(uname -s)" == "Darwin" ]]; then
    need cargo
    need rustup

    if command -v brew >/dev/null 2>&1; then
      local llvm_prefix
      llvm_prefix="$(brew --prefix llvm 2>/dev/null || true)"
      if [[ -n "$llvm_prefix" && -d "$llvm_prefix/bin" ]]; then
        export PATH="$llvm_prefix/bin:$PATH"
      fi
    fi

    command -v llvm-rc >/dev/null 2>&1 || die \
      "Windows cross-build needs LLVM. On macOS run: brew install llvm"
    command -v cargo-xwin >/dev/null 2>&1 || die \
      "Windows cross-build needs cargo-xwin. Run: cargo install --locked cargo-xwin"
    command -v makensis >/dev/null 2>&1 || die \
      "Windows NSIS packaging needs makensis. On macOS run: brew install makensis"

    if ! rustup target list --installed | grep -qx 'x86_64-pc-windows-msvc'; then
      say "Installing Rust Windows target"
      rustup target add x86_64-pc-windows-msvc
    fi
  fi
}

build_windows() {
  need cargo

  local os="$(uname -s)"
  local target_args=()
  if [[ "$os" == "Darwin" || "$os" == "Linux" ]]; then
    prepare_windows_cross_compile
    target_args=(--runner cargo-xwin --target x86_64-pc-windows-msvc --bundles nsis)
  else
    target_args=(--bundles nsis)
  fi

  say "Building Windows x64 NSIS installer"
  (
    cd apps/client
    VITE_BOOKSTATS_API_URL="$PUBLIC_API_URL" \
      npx tauri build "${target_args[@]}"
  )

  local bundle_root installer
  if [[ "$os" == "Darwin" || "$os" == "Linux" ]]; then
    bundle_root="$ROOT_DIR/apps/client/src-tauri/target/x86_64-pc-windows-msvc/release/bundle/nsis"
  else
    bundle_root="$ROOT_DIR/apps/client/src-tauri/target/release/bundle/nsis"
  fi
  installer="$(find "$bundle_root" -maxdepth 1 -name '*.exe' -print -quit 2>/dev/null || true)"
  [[ -n "$installer" ]] || die "Tauri completed but no Windows NSIS installer was found."
  local updater_sig="${installer}.sig"
  [[ -s "$updater_sig" ]] || die "Tauri completed but no signed Windows updater artifact was found."

  local exe_out="$RELEASE_DIR/BookStats-Windows-v${VERSION}-Setup.exe"
  local zip_out="$RELEASE_DIR/BookStats-Windows-v${VERSION}.zip"
  cp "$installer" "$exe_out"
  rm -f "$zip_out"
  zip -qj "$zip_out" "$exe_out"
  say "Created ${exe_out#$ROOT_DIR/}"
  say "Created ${zip_out#$ROOT_DIR/}"
  write_updater_manifest "windows-x86_64" "$installer" "$updater_sig" "BookStats-Windows-v${VERSION}-Setup.exe"
}

if [[ "$TARGET" == "all" || "$TARGET" == "desktop" || "$TARGET" == "mac" || "$TARGET" == "windows" ]]; then
  require_updater_signing
fi

validate

case "$TARGET" in
  all)
    build_web
    build_server
    build_mac
    build_windows
    ;;
  web) build_web ;;
  server) build_server ;;
  mac) build_mac ;;
  windows) build_windows ;;
  desktop)
    build_mac
    build_windows
    ;;
esac

if [[ "$TARGET" == "all" || "$TARGET" == "desktop" || "$TARGET" == "mac" || "$TARGET" == "windows" ]]; then
  package_updater_bundle
fi

rm -rf "$STAGING_DIR"
say "Release export complete"
printf 'API URL embedded in clients: %s\n' "$PUBLIC_API_URL"
printf 'Web base path: %s\n' "$WEB_BASE"
if [[ "$TARGET" == "all" || "$TARGET" == "desktop" || "$TARGET" == "mac" || "$TARGET" == "windows" ]]; then
  printf 'Updater public base: %s\n' "$UPDATER_BASE_URL"
  printf 'Signed updater manifests/artifacts: %s\n' "$UPDATER_DIR"
fi
printf 'Artifacts: %s\n' "$RELEASE_DIR"
