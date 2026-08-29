import { readFileSync } from "node:fs";
import { fileURLToPath, URL } from "node:url";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const packageJson = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf8")) as { version: string };

export default defineConfig({
  plugins: [react()],
  define: {
    __BOOKSTATS_VERSION__: JSON.stringify(packageJson.version)
  },
  resolve: {
    alias: {
      "@bookstats/domain": fileURLToPath(new URL("../../packages/domain/src/index.ts", import.meta.url)),
      "@bookstats/statistics": fileURLToPath(new URL("../../packages/statistics/src/index.ts", import.meta.url))
    }
  },
  server: {
    port: 5173,
    strictPort: true,
    watch: { ignored: ["**/src-tauri/**"] }
  },
  clearScreen: false
});
