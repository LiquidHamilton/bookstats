import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import { registerBookStatsServiceWorker } from "./pwa";

registerBookStatsServiceWorker();

createRoot(document.getElementById("root")!).render(<StrictMode><App /></StrictMode>);
