export type DesktopUpdatePhase = "idle" | "checking" | "downloading" | "installing" | "relaunching" | "error";

export interface DesktopUpdateProgress {
  phase: DesktopUpdatePhase;
  downloaded?: number;
  total?: number;
}

export function isTauriRuntime(): boolean {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

export async function applyDesktopUpdate(onProgress: (progress: DesktopUpdateProgress) => void): Promise<string> {
  if (!isTauriRuntime()) throw new Error("Desktop updater is only available in the BookStats desktop app.");

  onProgress({ phase: "checking" });
  const [{ check }, { relaunch }] = await Promise.all([
    import("@tauri-apps/plugin-updater"),
    import("@tauri-apps/plugin-process")
  ]);

  const update = await check({ timeout: 20_000 });
  if (!update) throw new Error("The server requires a newer BookStats version, but no signed desktop update is available for this platform yet.");

  let downloaded = 0;
  let total: number | undefined;
  await update.downloadAndInstall((event) => {
    if (event.event === "Started") {
      total = event.data.contentLength ?? undefined;
      onProgress({ phase: "downloading", downloaded, total });
    } else if (event.event === "Progress") {
      downloaded += event.data.chunkLength;
      onProgress({ phase: "downloading", downloaded, total });
    } else if (event.event === "Finished") {
      onProgress({ phase: "installing", downloaded, total });
    }
  });

  onProgress({ phase: "relaunching", downloaded, total });
  await relaunch();
  return update.version;
}
