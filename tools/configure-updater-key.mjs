#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";

const input = process.argv[2];
if (!input) {
  console.error("Usage: node tools/configure-updater-key.mjs /path/to/bookstats-updater.key.pub");
  process.exit(2);
}

const publicKey = fs.readFileSync(path.resolve(input), "utf8").trim();
if (!publicKey) throw new Error("Updater public key file is empty.");

const configPath = path.resolve("apps/client/src-tauri/tauri.conf.json");
const config = JSON.parse(fs.readFileSync(configPath, "utf8"));
config.plugins ??= {};
config.plugins.updater ??= {};
config.plugins.updater.pubkey = publicKey;
fs.writeFileSync(configPath, JSON.stringify(config, null, 2) + "\n");
console.log("Configured BookStats updater public key in apps/client/src-tauri/tauri.conf.json");
