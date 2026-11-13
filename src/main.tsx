import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { Workbox } from "workbox-window";
import { App } from "./ui/App";
import { initAuth, resolveClientId } from "./sync/msal";
import "./styles.css";

/**
 * MSAL 의 숨은 iframe 이 이 앱 주소로 돌아오면, 부모 창이 주소의 응답만 읽고 iframe 을 지운다.
 * 그 안에서 앱 전체(서비스 워커·동기화)를 띄우면 동기화가 두 번 돌 수 있으므로 iframe 안에서는 아무것도 하지 않는다.
 */
function inAuthIframe(): boolean {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

if (!inAuthIframe()) void boot();

/**
 * 서비스 워커: updateViaCache "none" 으로 등록해 GitHub Pages 의 10분 HTTP 캐시가 옛 sw.js 를 붙잡지 않게 한다.
 * (virtual:pwa-register 는 이 옵션을 넘기지 않아 workbox-window 를 직접 쓴다.) skipWaiting·clientsClaim 은
 * vite.config 의 workbox 설정, 새 판이 화면을 넘겨받으면 한 번 다시 불러온다. 화면 복귀 때마다 새 판을 확인한다.
 */
function registerServiceWorker(): void {
  if (!("serviceWorker" in navigator) || import.meta.env.DEV) return;
  const base = import.meta.env.BASE_URL || "/";
  const wb = new Workbox(`${base}sw.js`, { scope: base, updateViaCache: "none" });
  let reloaded = false;
  wb.addEventListener("waiting", () => wb.messageSkipWaiting());
  wb.addEventListener("controlling", (event) => {
    if (!event.isUpdate || reloaded) return;
    reloaded = true;
    window.location.reload();
  });
  void wb.register({ immediate: true }).then((registration) => {
    if (!registration) return;
    const check = () => {
      registration.update().catch(() => undefined);
    };
    check();
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") check();
    });
  });
}

async function boot(): Promise<void> {
  registerServiceWorker();
  // 화면을 그리기 전에: Microsoft 응답 처리 + 캐시 계정이 있으면 토큰으로 실제 확인(만료면 바로 '로그인 필요').
  // 확인이 오래 걸리면(최대 수 초) 화면부터 띄운다. 실패해도 앱은 뜬다.
  await initAuth(resolveClientId("")).catch((error) => console.warn("initAuth 실패", error));
  createRoot(document.getElementById("root")!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}
