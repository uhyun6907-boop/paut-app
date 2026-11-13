/**
 * 로그인 중계 서버(paut-api, Cloudflare Worker) 모드.
 *
 * 빌드 변수 VITE_PAUT_API 가 있으면 MSAL(SPA, refresh token 24시간) 대신 이 서버를 쓴다.
 * - 로그인: 서버가 Microsoft 로그인을 받아 refresh token 을 서버에 보관하고, 앱에는 기기 열쇠(device token)만 준다.
 * - 토큰: Graph access token 을 필요할 때 서버에서 받는다. 서버 쪽 로그인은 쓰는 동안 계속 유지된다.
 * - 서버가 'login_required' 를 줄 때만 로그인 만료. 네트워크·서버 일시 오류는 만료가 아니다.
 */

import { AuthExpiredError, LoginRequiredError, markAuthOk, markExpired, markSignedOut } from "./authState";

const DEVICE_KEY = "paut-device-token";
const EMAIL_KEY = "paut-device-email";
const STATE_KEY = "paut-login-state";

export function backendUrl(): string {
  return (import.meta.env.VITE_PAUT_API ?? "").trim().replace(/\/+$/, "");
}

export function backendEnabled(): boolean {
  return !!backendUrl();
}

function read(key: string): string | null {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

function write(key: string, value: string | null): void {
  try {
    if (value === null) localStorage.removeItem(key);
    else localStorage.setItem(key, value);
  } catch {
    /* 저장 불가 */
  }
}

export function backendEmail(): string | null {
  return read(DEVICE_KEY) ? read(EMAIL_KEY) : null;
}

export interface BackendRedirect {
  email: string | null;
  fromRedirect: boolean;
  error?: string;
}

let handled: BackendRedirect | null = null;

/**
 * 앱 시작 때 한 번: 서버에서 돌아온 주소(#paut_device=…)를 읽어 기기 열쇠를 저장하고 주소창에서 지운다.
 */
export function completeBackendRedirect(): BackendRedirect {
  if (handled) return handled;
  const hash = typeof window !== "undefined" ? window.location.hash.replace(/^#/, "") : "";
  const params = new URLSearchParams(hash);
  const device = params.get("paut_device");
  const error = params.get("paut_error");
  const state = params.get("paut_state");
  if (device || error) {
    try {
      history.replaceState(null, "", window.location.pathname + window.location.search);
    } catch {
      /* 무시 */
    }
  }
  const expected = read(STATE_KEY);
  write(STATE_KEY, null);
  if (device) {
    if (expected && state !== expected) {
      handled = { email: backendEmail(), fromRedirect: false, error: "로그인 응답이 이 기기의 요청과 맞지 않습니다. 다시 로그인하세요." };
      return handled;
    }
    write(DEVICE_KEY, device);
    write(EMAIL_KEY, params.get("paut_email") ?? "");
    cached = null;
    markAuthOk();
    handled = { email: backendEmail(), fromRedirect: true };
    return handled;
  }
  handled = { email: backendEmail(), fromRedirect: false, ...(error ? { error: `Microsoft 로그인 실패: ${error}` } : {}) };
  if (!read(DEVICE_KEY)) markSignedOut();
  return handled;
}

export function backendLogin(hint?: string | null): void {
  const state = crypto.randomUUID();
  write(STATE_KEY, state);
  const q = new URLSearchParams({ state });
  if (hint) q.set("hint", hint);
  window.location.assign(`${backendUrl()}/auth/login?${q}`);
}

export async function backendLogout(): Promise<void> {
  const device = read(DEVICE_KEY);
  write(DEVICE_KEY, null);
  write(EMAIL_KEY, null);
  cached = null;
  if (device) {
    await fetch(`${backendUrl()}/auth/logout`, { method: "POST", headers: { authorization: `Bearer ${device}` } }).catch(() => undefined);
  }
  markSignedOut();
}

let cached: { token: string; expiresAt: number } | null = null;
let inflight: Promise<string> | null = null;

/** 서버에서 Graph access token 을 받는다. 만료 판정은 서버가 login_required/device_unknown 을 줄 때만. */
export function backendToken(force = false): Promise<string> {
  if (!force && cached && cached.expiresAt - Date.now() > 2 * 60_000) return Promise.resolve(cached.token);
  if (inflight && !force) return inflight;
  const p = (async () => {
    const device = read(DEVICE_KEY);
    if (!device) {
      markSignedOut();
      throw new LoginRequiredError();
    }
    const res = await fetch(`${backendUrl()}/api/token${force ? "?force=1" : ""}`, {
      method: "POST",
      headers: { authorization: `Bearer ${device}` },
    });
    const body = (await res.json().catch(() => ({}))) as { access_token?: string; expires_at?: number; email?: string; error?: string; message?: string };
    if (res.status === 401 && (body.error === "login_required" || body.error === "device_unknown")) {
      cached = null;
      if (body.error === "device_unknown") {
        write(DEVICE_KEY, null);
        markSignedOut();
        throw new LoginRequiredError();
      }
      markExpired();
      throw new AuthExpiredError();
    }
    if (!res.ok || !body.access_token) {
      // 서버·Microsoft 일시 오류: 로그인 상태는 그대로 두고 연결 문제로 던진다.
      throw new Error(`로그인 서버 오류(${res.status}): ${body.message ?? body.error ?? ""}`.trim());
    }
    if (body.email) write(EMAIL_KEY, body.email);
    cached = { token: body.access_token, expiresAt: body.expires_at ?? Date.now() + 30 * 60_000 };
    markAuthOk();
    return body.access_token;
  })();
  inflight = p;
  p.finally(() => {
    if (inflight === p) inflight = null;
  }).catch(() => undefined);
  return p;
}
