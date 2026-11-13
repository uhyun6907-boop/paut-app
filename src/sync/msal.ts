import {
  BrowserAuthError,
  CacheLookupPolicy,
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
} from "@azure/msal-browser";
import { TOKEN_TIMEOUT_TEXT, withTimeout } from "./timeout";
import {
  AuthExpiredError,
  getAuthState,
  LoginRequiredError,
  markAuthOk,
  markExpired,
  markSignedOut,
} from "./authState";

// 전역 인증 상태 API 는 authState.ts 에 있고 여기서 그대로 다시 내보낸다(기존 import 경로 유지).
export {
  AUTH_EXPIRED_TEXT,
  AuthExpiredError,
  getAuthState,
  isLoginError,
  isLoginFailureText,
  LOGIN_NEEDED_TEXT,
  LoginRequiredError,
  markExpired,
  onAuthChange,
  trackRequest,
  type AuthState,
} from "./authState";

/**
 * offline_access 를 명시한다(MSAL 도 기본으로 붙이지만, 빠지면 refresh token 이 안 나와 1시간마다 로그인이 필요해진다).
 * 캐시 조회 때는 MSAL 이 OIDC 범위(offline_access)를 빼고 비교하므로 토큰 캐시 적중에는 영향이 없다.
 */
export const GRAPH_SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"];

/** Public SPA client id. Not a secret. Pages build uses this when settings are empty. */
export const DEFAULT_CLIENT_ID = "52353295-20fe-41c8-873a-a4dae9c92042";

/** Must match the Azure SPA redirect URI, including the trailing slash. */
export const PAGES_REDIRECT_URI = "https://uhyun6907-boop.github.io/paut-app/";

const ACCOUNT_KEY = "paut-msal-account";

let appPromise: Promise<PublicClientApplication> | null = null;
let pca: PublicClientApplication | null = null;
let boundClientId = "";
let redirectOnce: Promise<RedirectOutcome> | null = null;
/** 토큰 단계별 시간 제한(ms). iOS 홈 화면 앱에서 응답 없는 요청·iframe 이 동기화를 영원히 붙잡지 않게 한다. */
export const authTimeouts = {
  redirectMs: 20_000,
  silentMs: 20_000,
  /** 앱을 켤 때(initAuth) 토큰 확인을 기다리는 최대 시간. 넘으면 화면을 먼저 띄우고 확인은 뒤에서 이어 간다. */
  probeMs: 8_000,
};

export function azureAuthority(): string {
  return import.meta.env.VITE_AZURE_AUTHORITY?.trim() || "https://login.microsoftonline.com/consumers";
}

export function resolveClientId(fromSettings: string): string {
  const typed = fromSettings.trim();
  if (typed) return typed;
  const env = import.meta.env.VITE_AZURE_CLIENT_ID?.trim();
  if (env) return env;
  if ((import.meta.env.BASE_URL || "/") === "/paut-app/") return DEFAULT_CLIENT_ID;
  return "";
}

/** Pages build always returns the registered URI. Other builds use this origin plus the Vite base. */
export function buildRedirectUri(baseUrl: string, origin: string, envRedirect?: string): string {
  if (baseUrl === "/paut-app/" || baseUrl === "/paut-app") return PAGES_REDIRECT_URI;
  const fromEnv = envRedirect?.trim();
  if (fromEnv) return fromEnv.endsWith("/") ? fromEnv : `${fromEnv}/`;
  const path = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  const root = origin.replace(/\/$/, "");
  return `${root}${path.startsWith("/") ? path : `/${path}`}`;
}

export function redirectUri(): string {
  return buildRedirectUri(import.meta.env.BASE_URL || "/", window.location.origin, import.meta.env.VITE_AZURE_REDIRECT_URI);
}

/**
 * 캐시는 localStorage(앱을 닫았다 켜도 계정·refresh token 이 남는다). sessionStorage 는 쓰지 않는다.
 *
 * storeAuthStateInCookie 는 false 로 둔다:
 * - 로그인 왕복에 필요한 state/nonce/PKCE 는 이미 temporaryCacheLocation=localStorage 로 남는다.
 * - iOS 에서 홈 화면 앱과 Safari 는 localStorage 뿐 아니라 쿠키 저장소도 따로라, 쿠키가 둘을 이어 주지 못한다.
 * - 이 옵션은 IE/구 Edge 보안 영역 문제용이고 MSAL v4 에서 폐기 예정이다. 켜면 요청 상태가 쿠키로 GitHub Pages 에
 *   매 요청 함께 실려 가고 4KB 한도에 걸릴 수 있다. 얻는 것이 없어 끈다.
 */
export function msalConfig(clientId: string, uri: string): Configuration {
  return {
    auth: {
      clientId,
      authority: azureAuthority(),
      redirectUri: uri,
      postLogoutRedirectUri: uri,
      navigateToLoginRequestUrl: false,
    },
    cache: {
      cacheLocation: "localStorage",
      // iOS Safari and a home-screen PWA drop sessionStorage on the Microsoft round-trip.
      temporaryCacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  };
}

export async function msalApp(clientId: string): Promise<PublicClientApplication> {
  if (!clientId) throw new Error("Azure 앱 클라이언트 ID가 없습니다. 설정에서 입력하거나 빌드 환경 변수에 넣으세요.");
  if (appPromise && boundClientId === clientId) return appPromise;
  boundClientId = clientId;
  redirectOnce = null;
  const app = new PublicClientApplication(msalConfig(clientId, redirectUri()));
  // initialize 가 끝나기 전에 다른 호출이 같은 인스턴스를 쓰지 않도록 약속(promise)을 공유한다.
  appPromise = app.initialize().then(() => {
    pca = app;
    return app;
  });
  appPromise.catch(() => {
    if (boundClientId === clientId) appPromise = null;
  });
  return appPromise;
}

/** 테스트용: 모듈 상태를 비운다. */
export function resetMsalForTests(): void {
  appPromise = null;
  pca = null;
  boundClientId = "";
  redirectOnce = null;
  initPromise = null;
}

function rememberAccount(account: AccountInfo): void {
  try {
    localStorage.setItem(ACCOUNT_KEY, account.homeAccountId);
  } catch {
    /* 저장 불가 */
  }
}

function rememberedAccountId(): string | null {
  try {
    return localStorage.getItem(ACCOUNT_KEY);
  } catch {
    return null;
  }
}

/**
 * 캐시에서 계정을 되살려 활성 계정으로 둔다. 활성 계정 → 마지막으로 쓴 계정 → 첫 계정.
 * 캐시(localStorage)에 계정이 있으면 handleRedirectPromise 가 실패했더라도 로그인은 유지된다.
 */
export function restoreAccount(app: PublicClientApplication): AccountInfo | null {
  const active = app.getActiveAccount();
  if (active) return active;
  const all = app.getAllAccounts();
  if (!all.length) return null;
  const remembered = rememberedAccountId();
  const account = all.find((a) => a.homeAccountId === remembered) ?? all[0];
  app.setActiveAccount(account);
  return account;
}

export interface RedirectOutcome {
  account: AccountInfo | null;
  fromRedirect: boolean;
  /** handleRedirectPromise 가 던진 오류(예: 로그인 취소). 캐시는 지우지 않는다. */
  error?: string;
}

/**
 * 앱을 켤 때마다 맨 먼저: Microsoft 에서 돌아온 응답을 처리하고(handleRedirectPromise), 캐시의 계정을 활성 계정으로 둔다.
 * 응답 처리 실패는 삼키지 않고 error 로 돌려주며, 그래도 캐시된 계정은 그대로 돌려준다.
 * 한 페이지 로드에서 한 번만 실행되고, 토큰 요청(accessToken)은 이것이 끝난 뒤에만 한다.
 */
export async function completeRedirect(clientId: string): Promise<RedirectOutcome> {
  if (!clientId) return { account: null, fromRedirect: false };
  const app = await msalApp(clientId);
  return handleRedirectOnce(app);
}

function handleRedirectOnce(app: PublicClientApplication): Promise<RedirectOutcome> {
  if (!redirectOnce) {
    redirectOnce = (async () => {
      let error: string | undefined;
      let result: Awaited<ReturnType<PublicClientApplication["handleRedirectPromise"]>> = null;
      try {
        result = await app.handleRedirectPromise();
      } catch (e) {
        error = authErrorText(e);
        console.warn("MSAL handleRedirectPromise 실패(캐시는 유지)", e);
      }
      if (result?.account) {
        app.setActiveAccount(result.account);
        // Microsoft 에서 막 받은 토큰: 로그인은 확실히 살아 있다.
        markAuthOk();
      }
      const account = result?.account ?? restoreAccount(app);
      if (account) rememberAccount(account);
      return { account, fromRedirect: !!result, error };
    })();
  }
  return redirectOnce;
}

function authErrorText(e: unknown): string {
  if (e && typeof e === "object" && "errorCode" in e && (e as { errorCode?: string }).errorCode) {
    const code = (e as { errorCode: string }).errorCode;
    if (code === "user_cancelled" || code === "access_denied") return "로그인을 취소했습니다.";
    return `Microsoft 로그인 응답을 처리하지 못했습니다(${code}).`;
  }
  return e instanceof Error && e.message ? e.message : "Microsoft 로그인 응답을 처리하지 못했습니다.";
}

/*
 * 대화형 로그인은 항상 redirect 다(팝업 없음):
 * - iOS 홈 화면 앱(standalone)에서 팝업은 별도 Safari 창으로 열려 결과가 앱으로 돌아오지 않는다.
 * - iPhone Safari 도 팝업 차단·창 전환 문제가 있다. 데스크톱도 같은 흐름으로 통일한다.
 */

export interface LoginOptions {
  /** 사용자가 '다른 계정'을 눌렀을 때만 true. 계정 선택 화면(prompt=select_account)을 띄운다. */
  switchAccount?: boolean;
}

/**
 * 대화형 로그인(redirect). 평소 다시 로그인은 계정 선택 화면 없이 login_hint 로 바로 그 계정을 쓴다.
 * Microsoft 세션 쿠키('로그인 상태 유지')가 살아 있으면 비밀번호 없이 곧바로 돌아온다.
 */
export async function loginRedirect(clientId: string, options: LoginOptions = {}): Promise<void> {
  const app = await msalApp(clientId);
  await handleRedirectOnce(app);
  const uri = redirectUri();
  if (options.switchAccount) {
    await app.loginRedirect({ scopes: GRAPH_SCOPES, redirectUri: uri, prompt: "select_account" });
    return;
  }
  const account = restoreAccount(app);
  await app.loginRedirect({ scopes: GRAPH_SCOPES, redirectUri: uri, ...(account ? { loginHint: account.username } : {}) });
}

export async function logoutRedirect(clientId: string): Promise<void> {
  const app = await msalApp(clientId);
  const account = restoreAccount(app) ?? undefined;
  try {
    localStorage.removeItem(ACCOUNT_KEY);
  } catch {
    /* 저장 불가 */
  }
  await app.logoutRedirect({ account, postLogoutRedirectUri: redirectUri() });
}

export interface TokenOptions {
  /** 캐시를 무시하고 새 토큰을 받는다. 401 뒤에 쓴다. */
  forceRefresh?: boolean;
  /**
   * false 면 다시 로그인이 필요해도 화면을 넘기지 않고 AuthExpiredError(LoginRequiredError)를 던진다.
   * 동기화 도중(리스트는 썼고 작업일지는 아직)에 페이지가 바뀌면 안 되므로 동기화는 false.
   */
  interactive?: boolean;
  /** 이미 만료로 표시돼 있어도 조용한 토큰을 한 번 더 시도한다(앱 시작·화면 복귀 확인). */
  allowWhenExpired?: boolean;
}

/** 확실한 로그인 오류(만료)로 보는 BrowserAuthError 코드. 이 밖의 코드(network_error·timed_out 등)는 연결 문제. */
const EXPIRY_BROWSER_CODES = new Set(["monitor_window_timeout", "block_iframe_reload", "no_tokens_found", "no_account_error"]);

function codeOf(error: unknown): string {
  return error && typeof error === "object" && "errorCode" in error ? String((error as { errorCode?: string }).errorCode ?? "") : "";
}

/**
 * 다시 로그인해야 하는 오류인가. InteractionRequiredAuthError(refresh token 만료 등) 이거나
 * BrowserAuthError 중 위 네 코드만. 네트워크·시간 초과·서버 오류는 false(로그인 상태를 바꾸지 않는다).
 */
export function needsInteraction(error: unknown): boolean {
  if (error instanceof InteractionRequiredAuthError) return true;
  if (error instanceof Error && error.name === "InteractionRequiredAuthError") return true;
  const isBrowser = error instanceof BrowserAuthError || (error instanceof Error && error.name === "BrowserAuthError");
  return isBrowser && EXPIRY_BROWSER_CODES.has(codeOf(error));
}

/**
 * 토큰 받기:
 * 1) acquireTokenSilent: 캐시의 access token → refresh token. cacheLookupPolicy 로 MSAL 이 iframe(ssoSilent)을
 *    시도하지 않게 한다(iOS ITP 가 막아 15초씩 멈추기만 한다). 숨은 iframe 갱신은 쓰지 않는다.
 * 2) InteractionRequired·만료 코드 → markExpired(): 진행 중 Graph 요청을 전부 끊고 '로그인 필요'.
 *    동기화(interactive:false)는 AuthExpiredError, 사용자가 누른 동작(interactive:true)만 redirect.
 * 네트워크·시간 초과·서버 오류는 그대로 던진다(연결 문제, 만료 아님). 어떤 오류에서도 MSAL 캐시는 지우지 않는다.
 */
export async function accessToken(clientId: string, options: TokenOptions = {}): Promise<string> {
  const interactive = options.interactive ?? true;
  const app = await msalApp(clientId);
  // 돌아온 응답 처리가 멈춰도(네트워크) 캐시된 계정으로 이어 간다.
  await withTimeout(handleRedirectOnce(app), authTimeouts.redirectMs, TOKEN_TIMEOUT_TEXT).catch((error) => {
    console.warn("handleRedirectPromise 시간 초과 → 캐시 계정으로 진행", error);
  });
  const account = restoreAccount(app);
  if (!account) {
    markSignedOut();
    if (!interactive) throw new LoginRequiredError();
    await app.loginRedirect({ scopes: GRAPH_SCOPES, redirectUri: redirectUri() });
    throw new Error("로그인을 다시 진행합니다.");
  }
  const skipSilent = getAuthState() === "expired" && !options.allowWhenExpired;
  if (!skipSilent) {
    try {
      const res = await withTimeout(
        app.acquireTokenSilent({
          account,
          scopes: GRAPH_SCOPES,
          forceRefresh: !!options.forceRefresh,
          cacheLookupPolicy: CacheLookupPolicy.AccessTokenAndRefreshToken,
        }),
        authTimeouts.silentMs,
        TOKEN_TIMEOUT_TEXT,
      );
      markAuthOk();
      return res.accessToken;
    } catch (error) {
      // 시간 초과·네트워크 오류는 로그인 문제가 아니다 → 그대로 던져 기록을 대기로 둔다.
      if (!needsInteraction(error)) throw error;
      console.warn("조용한 토큰 실패 → 로그인 만료", error);
      markExpired();
    }
  }
  if (!interactive) throw new AuthExpiredError();
  await app.loginRedirect({ loginHint: account.username, scopes: GRAPH_SCOPES, redirectUri: redirectUri() });
  throw new Error("로그인을 다시 진행합니다.");
}

/** 앱 시작·화면 복귀: 버튼을 누르기 전에 만료를 먼저 알아낸다. 결과는 전역 인증 상태로만 남긴다. */
export async function probeAuth(clientId: string): Promise<void> {
  if (!clientId) return;
  try {
    await accessToken(clientId, { interactive: false, allowWhenExpired: true });
  } catch {
    /* 상태는 accessToken 안에서 바뀐다(네트워크·시간 초과면 그대로) */
  }
}

let initPromise: Promise<void> | null = null;
let visibilityBound = false;

/**
 * main.tsx 가 화면을 그리기 전에 await 한다: initialize → handleRedirectPromise → (캐시 계정이 있으면) 토큰으로 확인.
 * 캐시에 계정이 있어도 '로그인됨'으로 단정하지 않는다. 확인이 probeMs 를 넘으면 화면부터 띄우고 확인은 뒤에서 끝난다.
 */
export function initAuth(clientId: string): Promise<void> {
  if (!clientId) return Promise.resolve();
  if (!initPromise) {
    initPromise = (async () => {
      const out = await withTimeout(completeRedirect(clientId), authTimeouts.redirectMs, TOKEN_TIMEOUT_TEXT).catch(() => null);
      if (out && !out.account) {
        markSignedOut();
        return;
      }
      if (out?.fromRedirect && out.account) return; // 방금 로그인: 이미 ok
      await withTimeout(probeAuth(clientId), authTimeouts.probeMs, TOKEN_TIMEOUT_TEXT).catch(() => undefined);
    })();
    if (!visibilityBound && typeof document !== "undefined" && typeof document.addEventListener === "function") {
      visibilityBound = true;
      document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "visible" && initPromise) void initPromise.then(() => probeAuth(clientId));
      });
    }
  }
  return initPromise;
}

export function currentAccount(clientId: string): string | null {
  if (!pca || boundClientId !== clientId) return null;
  return restoreAccount(pca)?.username ?? null;
}
