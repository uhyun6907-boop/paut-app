import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";
import { execSync } from "node:child_process";

/** 설정 화면에 보이는 빌드 번호: 짧은 커밋 + UTC 시각(예: 1eff3c6+20261113T0412Z). */
function buildId(): string {
  let sha = "nogit";
  try {
    sha = execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim() || sha;
  } catch {
    /* git 없음 */
  }
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z").slice(0, 13) + "Z";
  return process.env.VITE_BUILD_ID || `${sha}+${stamp}`;
}
const BUILD_ID = buildId();

export default defineConfig(({ mode }) => {
  const base = mode === "pages" ? "/paut-app/" : "/";
  return {
  base,
  define: { "import.meta.env.VITE_BUILD_ID": JSON.stringify(BUILD_ID) },
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      // 등록은 main.tsx 가 workbox-window 로 직접 한다(updateViaCache: "none").
      injectRegister: false,
      includeAssets: ["favicon.svg", "pwa-192.png", "pwa-512.png"],
      manifest: {
        name: "PAUT 오늘 작업",
        short_name: "PAUT",
        description: "위상배열초음파 현장 라인·포인트 입력. OneDrive 엑셀과 동기화.",
        lang: "ko",
        dir: "ltr",
        display: "standalone",
        orientation: "portrait",
        id: base,
        start_url: base,
        scope: base,
        background_color: "#0f1419",
        theme_color: "#0f1419",
        icons: [
          { src: `${base}pwa-192.png`, sizes: "192x192", type: "image/png", purpose: "any" },
          { src: `${base}pwa-512.png`, sizes: "512x512", type: "image/png", purpose: "any maskable" },
        ],
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,svg,png,woff2,xlsx}"],
        maximumFileSizeToCacheInBytes: 8 * 1024 * 1024,
        navigateFallback: `${base}index.html`,
        skipWaiting: true,
        clientsClaim: true,
      },
    }),
  ],
  server: { host: "0.0.0.0", port: 4317, strictPort: true },
  preview: { host: "0.0.0.0", port: 4317, strictPort: true },
  test: {
    environment: "node",
    include: ["src/**/*.test.ts"],
    testTimeout: 120000,
    hookTimeout: 120000,
  },
};
});
