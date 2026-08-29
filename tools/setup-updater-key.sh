#!/usr/bin/env bash
set -Eeuo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

KEY_PATH="${BOOKSTATS_UPDATER_KEY_PATH:-$HOME/.config/bookstats/updater.key}"
PUB_PATH="${KEY_PATH}.pub"

if [[ -e "$KEY_PATH" || -e "$PUB_PATH" ]]; then
  echo "Updater key already exists at $KEY_PATH (or $PUB_PATH)." >&2
  echo "Refusing to replace it. Losing/changing this key would strand existing desktop installs." >&2
  exit 1
fi

mkdir -p "$(dirname "$KEY_PATH")"
chmod 700 "$(dirname "$KEY_PATH")" 2>/dev/null || true

echo "Generating the one-time BookStats updater signing key."
echo "Keep the private key safe and backed up; never commit or upload it."
(
  cd apps/client
  npx tauri signer generate -w "$KEY_PATH"
)

[[ -s "$KEY_PATH" ]] || { echo "Private key was not created." >&2; exit 1; }
[[ -s "$PUB_PATH" ]] || { echo "Public key was not created at $PUB_PATH." >&2; exit 1; }
chmod 600 "$KEY_PATH" "$PUB_PATH" 2>/dev/null || true
node tools/configure-updater-key.mjs "$PUB_PATH"

cat <<EOF

Updater key configured.

Before building a desktop release, export:
  export TAURI_SIGNING_PRIVATE_KEY="$KEY_PATH"
  export TAURI_SIGNING_PRIVATE_KEY_PASSWORD="<the password you chose, or empty if none>"

The public key is now embedded in tauri.conf.json and is safe to commit.
Back up the private key separately. Do not put it in this repository.
EOF
