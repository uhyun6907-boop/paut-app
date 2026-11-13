/**
 * 앱 전역 인증 상태 하나. msal.ts(토큰)·graph.ts(요청)·session(화면)이 모두 여기를 본다.
 * - 토큰이 죽으면 markExpired() → 진행 중인 Graph 요청을 전부 실제로 abort 하고 구독자에게 "expired" 를 알린다.
 * - 캐시에 계정이 있다는 것만으로는 "ok" 가 아니다. 토큰을 실제로 받아야 "ok".
 * MSAL 을 부르지 않는 순수 모듈이라 테스트에서 msal 을 가짜로 바꿔도 그대로 쓴다.
 */

export type AuthState = "unknown" | "ok" | "expired" | "signedOut";

/** 다시 로그인해야 할 때 동기화 단계에 남는 문구. session 이 이 문구로 '로그인 대기'를 실패와 구분한다. */
export const LOGIN_NEEDED_TEXT = "Microsoft 로그인 필요 — 화면 위 '로그인'을 누르면 대기 중인 기록을 바로 보냅니다.";
/** 만료 문구의 공통 앞부분. GRAPH_LOGIN_TEXT·AuthExpiredError 모두 이것으로 시작한다. */
export const AUTH_EXPIRED_TEXT = "Microsoft 로그인이 만료되었습니다";

export class LoginRequiredError extends Error {
  constructor(message = LOGIN_NEEDED_TEXT) {
    super(message);
    this.name = "LoginRequiredError";
  }
}

/** 토큰이 실제로 죽었다(InteractionRequired·새 토큰으로도 401). LoginRequiredError 를 잡는 기존 코드도 그대로 잡는다. */
export class AuthExpiredError extends LoginRequiredError {
  constructor(message = `${AUTH_EXPIRED_TEXT} — 화면 위 '로그인'을 누르면 대기 중인 기록을 바로 보냅니다.`) {
    super(message);
    this.name = "AuthExpiredError";
  }
}

/** 이름으로도 본다(테스트의 가짜 msal 모듈이 만든 LoginRequiredError 등). */
export function isLoginError(error: unknown): boolean {
  if (error instanceof LoginRequiredError) return true;
  return error instanceof Error && (error.name === "LoginRequiredError" || error.name === "AuthExpiredError");
}

/** 로그인이 끊겨(또는 새 토큰으로도 401) 못 보낸 단계 문구. 실패가 아니라 '로그인 대기'. */
export function isLoginFailureText(message?: string): boolean {
  return !!message && (message.includes(LOGIN_NEEDED_TEXT) || message.includes(AUTH_EXPIRED_TEXT));
}

let authState: AuthState = "unknown";
const listeners = new Set<(s: AuthState) => void>();
const inflight = new Set<AbortController>();

export function getAuthState(): AuthState {
  return authState;
}

export function onAuthChange(fn: (s: AuthState) => void): () => void {
  listeners.add(fn);
  return () => {
    listeners.delete(fn);
  };
}

function setAuth(s: AuthState): void {
  if (authState === s) return;
  authState = s;
  for (const f of [...listeners]) {
    try {
      f(s);
    } catch (error) {
      console.warn("auth 구독자 오류", error);
    }
  }
}

/** 토큰을 실제로 받았다. */
export function markAuthOk(): void {
  setAuth("ok");
}

/** 캐시에 계정이 없다. */
export function markSignedOut(): void {
  abortAll();
  setAuth("signedOut");
}

/** graph.ts 가 요청 시작 때 등록하고, 돌려받은 함수로 끝날 때 해제한다. */
export function trackRequest(c: AbortController): () => void {
  inflight.add(c);
  return () => {
    inflight.delete(c);
  };
}

/** 진행 중 요청 수(테스트·진단용). */
export function inflightCount(): number {
  return inflight.size;
}

function abortAll(): void {
  const err = new AuthExpiredError();
  for (const c of [...inflight]) c.abort(err);
  inflight.clear();
}

/** 토큰이 죽었을 때 부른다. 진행 중 Graph 요청을 전부 실제로 끊고, 구독자에게 알린다. */
export function markExpired(): void {
  if (authState === "expired") return;
  abortAll();
  setAuth("expired");
}

/** 테스트용 */
export function resetAuthForTests(): void {
  authState = "unknown";
  listeners.clear();
  inflight.clear();
}
