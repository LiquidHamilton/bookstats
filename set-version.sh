#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "Usage: $0 x.x.x" >&2
  exit 2
fi

VERSION="$1"
if [[ ! "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
  echo "Error: version must use x.x.x format (example: 0.3.0)." >&2
  exit 2
fi

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

BOOKSTATS_VERSION="$VERSION" node <<'NODE'
const fs = require('fs');
const path = require('path');
const version = process.env.BOOKSTATS_VERSION;

function read(file) { return fs.readFileSync(path.resolve(file), 'utf8'); }
function write(file, text) { fs.writeFileSync(path.resolve(file), text); }

const jsonFiles = [
  'package.json',
  'apps/client/package.json',
  'apps/admin/package.json',
  'apps/server/package.json',
  'packages/domain/package.json',
  'packages/statistics/package.json',
  'apps/client/src-tauri/tauri.conf.json'
];

for (const file of jsonFiles) {
  const data = JSON.parse(read(file));
  data.version = version;
  write(file, JSON.stringify(data, null, 2) + '\n');
}

const lockPath = path.resolve('package-lock.json');
if (fs.existsSync(lockPath)) {
  const lock = JSON.parse(fs.readFileSync(lockPath, 'utf8'));
  if (lock.version) lock.version = version;
  if (lock.packages) {
    for (const key of ['', 'apps/client', 'apps/admin', 'apps/server', 'packages/domain', 'packages/statistics']) {
      if (lock.packages[key]) lock.packages[key].version = version;
    }
  }
  fs.writeFileSync(lockPath, JSON.stringify(lock, null, 2) + '\n');
}

let cargo = read('apps/client/src-tauri/Cargo.toml');
if (!/^version\s*=\s*"[^"]+"/m.test(cargo)) throw new Error('Could not find Cargo.toml package version');
cargo = cargo.replace(/^version\s*=\s*"[^"]+"/m, `version = "${version}"`);
write('apps/client/src-tauri/Cargo.toml', cargo);

let server = read('apps/server/src/index.ts');
if (!/const APP_VERSION = "[^"]+";/.test(server)) throw new Error('Could not find server APP_VERSION');
server = server.replace(/const APP_VERSION = "[^"]+";/, `const APP_VERSION = "${version}";`);
write('apps/server/src/index.ts', server);

let metadata = read('apps/server/src/metadata.ts');
metadata = metadata.replace(/BookStats\/\d+\.\d+\.\d+ \(local-development\)/g, `BookStats/${version} (local-development)`);
write('apps/server/src/metadata.ts', metadata);

let readme = read('README.md');
readme = readme.replace(/\*\*Current version:\*\* `[^`]+`/, `**Current version:** \`${version}\``);
readme = readme.replace(/BookStats\/\d+\.\d+\.\d+/g, `BookStats/${version}`);
write('README.md', readme);

if (fs.existsSync(path.resolve('docs/SERVER_SETUP.md'))) {
  let serverSetup = read('docs/SERVER_SETUP.md');
  serverSetup = serverSetup.replace(/BookStats\/\d+\.\d+\.\d+/g, `BookStats/${version}`);
  serverSetup = serverSetup.replace(/BookStats-Web-v\d+\.\d+\.\d+\.zip/g, `BookStats-Web-v${version}.zip`);
  serverSetup = serverSetup.replace(/BookStats-Server-v\d+\.\d+\.\d+\.zip/g, `BookStats-Server-v${version}.zip`);
  write('docs/SERVER_SETUP.md', serverSetup);
}

if (fs.existsSync(path.resolve('.env.example'))) {
  let env = read('.env.example');
  env = env.replace(/^BOOKSTATS_METADATA_USER_AGENT=BookStats\/[^ ]+/m, `BOOKSTATS_METADATA_USER_AGENT=BookStats/${version}`);
  write('.env.example', env);
}
NODE

echo "BookStats version set to $VERSION"
