#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";

const [outputPath, version, platform, url, signaturePath] = process.argv.slice(2);
if (![outputPath, version, platform, url, signaturePath].every(Boolean)) {
  console.error("Usage: node tools/write-updater-manifest.mjs OUTPUT VERSION PLATFORM URL SIGNATURE_FILE");
  process.exit(2);
}

const signature = fs.readFileSync(signaturePath, "utf8").trim();
if (!signature) throw new Error(`Updater signature is empty: ${signaturePath}`);
const manifest = {
  version,
  notes: `BookStats v${version}`,
  pub_date: new Date().toISOString(),
  platforms: {
    [platform]: { signature, url }
  }
};
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, JSON.stringify(manifest, null, 2) + "\n");
console.log(`Created ${outputPath}`);
