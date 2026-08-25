import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

const appBase = process.env.VITE_APP_BASE_PATH || "/";
// Dev-only: proxies API calls to the standalone backend (see ../backend) so
// the browser never needs CORS. In production the backend itself serves this
// app's built dist/ as static files, so frontend and backend share an origin
// and no proxy is involved -- see backend/src/server.ts.
const backendPort = process.env.TB_BACKEND_PORT || 8097;

export default defineConfig({
  base: appBase,
  plugins: [react(), tailwindcss()],
  server: {
    host: "0.0.0.0",
    port: 8095,
    proxy: {
      "/api": `http://localhost:${backendPort}`,
    },
  },
  preview: {
    host: "0.0.0.0",
  },
});
