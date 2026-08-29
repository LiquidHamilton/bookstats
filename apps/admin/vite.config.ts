import { readFileSync } from "node:fs";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const packageJson = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as { version: string };

export default defineConfig({
  plugins: [react()],
  define: { __BOOKSTATS_ADMIN_VERSION__: JSON.stringify(packageJson.version) },
  server: { port: 5174, strictPort: true },
  clearScreen: false
});
