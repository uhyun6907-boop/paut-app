/**
 * MSAL 로그인 유지: 실제 login.microsoftonline.com 은 부르지 않는다. PublicClientApplication 을 가짜로 바꾼다.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { BrowserAuthError, CacheLookupPolicy, InteractionRequiredAuthError, type AccountInfo } from "@azure/msal-browser";

const fake = vi.hoisted(() => ({
  instances: [] as unknown[],
  accounts: [] as AccountInfo[],
  redirect: null as null | { account: AccountInfo } | Error,
  silent: [] as (string | Error)[],
  sso: [] as (string | Error)[],
  calls: [] as { fn: string; arg?: unknown }[],
  config: null as unknown,
}));

vi.mock("@azure/msal-browser", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@azure/msal-browser")>();
  class FakePca {
    active: AccountInfo | null = null;
    constructor(config: unknown) {
      fake.config = config;
      fake.instances.push(this);
    }
    async initialize() {
      fake.calls.push({ fn: "initialize" });
    }
    async handleRedirectPromise() {
      fake.calls.push({ fn: "handleRedirectPromise" });
      if (fake.redirect instanceof Error) throw fake.redirect;
      if (fake.redirect && !fake.accounts.includes(fake.redirect.account)) fake.accounts.push(fake.redirect.account);
      return fake.redirect ? { account: fake.redirect.account, accessToken: "redirect-token" } : null;
    }
    getAllAccounts() {
      return [...fake.accounts];
    }
    getActiveAccount() {
      return this.active;
    }
    setActiveAccount(account: AccountInfo | null) {
      fake.calls.push({ fn: "setActiveAccount", arg: account?.username });
      this.active = account;
    }
    async acquireTokenSilent(req: unknown) {
      fake.calls.push({ fn: "acquireTokenSilent", arg: req });
      const next = fake.silent.shift() ?? new InteractionRequiredAuthError("no_tokens_found");
      if (next === "HANG") return new Promise(() => undefined);
      if (next instanceof Error) throw next;
      return { accessToken: next, account: this.active };
    }
    async ssoSilent(req: unknown) {
      fake.calls.push({ fn: "ssoSilent", arg: req });
      const next = fake.sso.shift() ?? new InteractionRequiredAuthError("login_required");
      if (next === "HANG") return new Promise(() => undefined);
      if (next instanceof Error) throw next;
      return { accessToken: next, account: this.active };
    }
    async loginRedirect(req: unknown) {
      fake.calls.push({ fn: "loginRedirect", arg: req });
    }
    async acquireTokenRedirect(req: unknown) {
      fake.calls.push({ fn: "acquireTokenRedirect", arg: req });
    }
    async loginPopup(req: unknown) {
      fake.calls.push({ fn: "loginPopup", arg: req });
    }
    async acquireTokenPopup(req: unknown) {
      fake.calls.push({ fn: "acquireTokenPopup", arg: req });
    }
    async logoutRedirect(req: unknown) {
      fake.calls.push({ fn: "logoutRedirect", arg: req });
    }
    async clearCache() {
      fake.calls.push({ fn: "clearCache" });
    }
  }
  return { ...actual, PublicClientApplication: FakePca };
});

const bag = new Map<string, string>();
vi.stubGlobal("localStorage", {
  getItem: (key: string) => bag.get(key) ?? null,
  setItem: (key: string, value: string) => void bag.set(key, String(value)),
  removeItem: (key: string) => void bag.delete(key),
});
vi.stubGlobal("window", { location: { origin: "https://uhyun6907-boop.github.io" } });

const CLIENT = "52353295-20fe-41c8-873a-a4dae9c92042";
const acct = (name: string): AccountInfo => ({
  homeAccountId: `${name}-home`,
  environment: "login.microsoftonline.com",
  tenantId: "9188040d-6c67-4c5b-b112-36a304b66dad",
  username: `${name}@outlook.com`,
  localAccountId: `${name}-local`,
});

type Msal = typeof import("./msal");
async function load(): Promise<Msal> {
  const m = await import("./msal");
  m.resetMsalForTests();
  (await import("./authState")).resetAuthForTests();
  return m;
}
const fns = () => fake.calls.map((c) => c.fn);
const cacheCleared = () => fake.calls.some((c) => c.fn === "clearCache" || c.fn === "logoutRedirect");

beforeEach(() => {
  fake.instances.length = 0;
  fake.accounts = [];
  fake.redirect = null;
  fake.silent = [];
  fake.sso = [];
  fake.calls = [];
  bag.clear();
  vi.spyOn(console, "warn").mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("MSAL 설정", () => {
  it("범위에 offline_access(refresh token)가 있고 캐시는 localStorage 다(sessionStorage 아님)", async () => {
    const m = await load();
    expect(m.GRAPH_SCOPES).toEqual(expect.arrayContaining(["User.Read", "Files.ReadWrite", "offline_access"]));
    await m.msalApp(CLIENT);
    expect(fake.config).toMatchObject({
      auth: { clientId: CLIENT, authority: "https://login.microsoftonline.com/consumers" },
      cache: { cacheLocation: "localStorage", temporaryCacheLocation: "localStorage" },
    });
  });

  it("동시에 불러도 인스턴스는 하나이고 initialize 가 끝난 뒤에만 쓴다", async () => {
    const m = await load();
    const [a, b] = await Promise.all([m.msalApp(CLIENT), m.msalApp(CLIENT)]);
    expect(a).toBe(b);
    expect(fake.instances).toHaveLength(1);
  });
});

describe("앱 시작: handleRedirectPromise → 캐시 계정 복원", () => {
  it("redirect 응답이 없으면 캐시 계정을 setActiveAccount 로 되살린다", async () => {
    fake.accounts = [acct("woo")];
    const m = await load();
    const out = await m.completeRedirect(CLIENT);
    expect(out).toMatchObject({ fromRedirect: false, account: { username: "woo@outlook.com" } });
    expect(fns().slice(0, 2)).toEqual(["initialize", "handleRedirectPromise"]);
    expect(fake.calls).toContainEqual({ fn: "setActiveAccount", arg: "woo@outlook.com" });
  });

  it("handleRedirectPromise 가 던져도 캐시 계정은 그대로 돌려주고(오류도 알려 줌) 캐시를 지우지 않는다", async () => {
    fake.accounts = [acct("woo")];
    fake.redirect = new BrowserAuthError("no_token_request_cache_error");
    const m = await load();
    const out = await m.completeRedirect(CLIENT);
    expect(out.account?.username).toBe("woo@outlook.com");
    expect(out.error).toContain("no_token_request_cache_error");
    expect(cacheCleared()).toBe(false);
  });

  it("Microsoft 에서 돌아오면 그 계정을 활성 계정으로 두고 기억한다", async () => {
    fake.accounts = [acct("old")];
    fake.redirect = { account: acct("woo") };
    const m = await load();
    const out = await m.completeRedirect(CLIENT);
    expect(out).toMatchObject({ fromRedirect: true, account: { username: "woo@outlook.com" } });
    // 다음에 켤 때는 마지막으로 쓴 계정을 고른다
    fake.redirect = null;
    const again = await load();
    expect((await again.completeRedirect(CLIENT)).account?.username).toBe("woo@outlook.com");
  });

  it("토큰 요청보다 handleRedirectPromise 가 먼저, 한 번만 실행된다", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["t1", "t2"];
    const m = await load();
    await Promise.all([m.accessToken(CLIENT, { interactive: false }), m.completeRedirect(CLIENT), m.accessToken(CLIENT, { interactive: false })]);
    expect(fns().filter((f) => f === "handleRedirectPromise")).toHaveLength(1);
    expect(fns().indexOf("handleRedirectPromise")).toBeLessThan(fns().indexOf("acquireTokenSilent"));
  });
});

describe("토큰: 조용히(캐시·refresh token) → 안 되면 바로 만료(ssoSilent·iframe 없음)", () => {
  it("캐시/refresh token 으로 되면 그대로, 인증 상태 ok. MSAL 이 iframe 을 시도하지 않게 AccessTokenAndRefreshToken", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["at-1"];
    const m = await load();
    expect(await m.accessToken(CLIENT, { interactive: false })).toBe("at-1");
    const req = fake.calls.find((c) => c.fn === "acquireTokenSilent")!.arg as Record<string, unknown>;
    expect(req).toMatchObject({ scopes: m.GRAPH_SCOPES, cacheLookupPolicy: CacheLookupPolicy.AccessTokenAndRefreshToken, forceRefresh: false });
    expect(m.GRAPH_SCOPES).toContain("offline_access");
    expect(fns()).not.toContain("ssoSilent");
    expect(m.getAuthState()).toBe("ok");
  });

  it("계정이 캐시에 있어도 refresh token 만료(InteractionRequired) → ssoSilent 없이 바로 만료: AuthExpiredError(=LoginRequiredError), 진행 중 요청 abort", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = [new InteractionRequiredAuthError("invalid_grant")];
    const m = await load();
    const seen: string[] = [];
    m.onAuthChange((st) => seen.push(st));
    const req = new AbortController();
    const untrack = m.trackRequest(req);
    const err = await m.accessToken(CLIENT, { interactive: false }).catch((e) => e);
    expect(err).toBeInstanceOf(m.AuthExpiredError);
    expect(err).toBeInstanceOf(m.LoginRequiredError);
    expect(m.isLoginFailureText((err as Error).message)).toBe(true);
    expect(m.getAuthState()).toBe("expired");
    expect(seen).toEqual(["expired"]);
    expect(req.signal.aborted).toBe(true);
    expect(req.signal.reason).toBeInstanceOf(m.AuthExpiredError);
    untrack();
    expect(fns()).not.toContain("ssoSilent");
    expect(fns()).not.toContain("acquireTokenRedirect");
    expect(fns()).not.toContain("loginRedirect");
    expect(cacheCleared()).toBe(false);
  });

  it.each(["monitor_window_timeout", "block_iframe_reload", "no_tokens_found", "no_account_error"])(
    "BrowserAuthError %s 도 만료로 본다",
    async (code) => {
      fake.accounts = [acct("woo")];
      fake.silent = [new BrowserAuthError(code)];
      const m = await load();
      await expect(m.accessToken(CLIENT, { interactive: false })).rejects.toBeInstanceOf(m.AuthExpiredError);
      expect(m.getAuthState()).toBe("expired");
    },
  );

  it("만료로 표시된 뒤에는 조용한 토큰을 다시 기다리지 않는다. 화면 복귀 확인(allowWhenExpired)만 다시 시도해 살아나면 ok", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = [new InteractionRequiredAuthError("interaction_required"), "at-back"];
    const m = await load();
    await expect(m.accessToken(CLIENT, { interactive: false })).rejects.toBeInstanceOf(m.AuthExpiredError);
    await expect(m.accessToken(CLIENT, { interactive: false })).rejects.toBeInstanceOf(m.AuthExpiredError);
    expect(fns().filter((f) => f === "acquireTokenSilent")).toHaveLength(1);
    await m.probeAuth(CLIENT);
    expect(m.getAuthState()).toBe("ok");
  });

  it("acquireTokenSilent 가 응답 없이 멈추면 시간 제한 뒤 TimeoutError: 만료도 아니고 '로그인됨'도 아니다. 다음 호출은 정상", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["HANG", "at-2"];
    const m = await load();
    m.authTimeouts.silentMs = 20;
    const err = await m.accessToken(CLIENT, { interactive: false }).catch((e) => e);
    expect(err).not.toBeInstanceOf(m.LoginRequiredError);
    expect((err as Error).name).toBe("TimeoutError");
    expect(m.getAuthState()).toBe("unknown");
    expect(fns()).not.toContain("ssoSilent");
    expect(await m.accessToken(CLIENT, { interactive: false, forceRefresh: true })).toBe("at-2");
    expect(cacheCleared()).toBe(false);
    m.authTimeouts.silentMs = 20_000;
  });

  it("사용자가 누른 동작(interactive)만 redirect(loginRedirect + login_hint) 로 다시 로그인, 계정 선택 화면 없음, 팝업 없음", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = [new InteractionRequiredAuthError("login_required")];
    const m = await load();
    await expect(m.accessToken(CLIENT)).rejects.toThrow("로그인을 다시 진행합니다");
    const req = fake.calls.find((c) => c.fn === "loginRedirect")!.arg as Record<string, unknown>;
    expect(req).toMatchObject({ loginHint: "woo@outlook.com" });
    expect(req.prompt).toBeUndefined();
    expect(fns().some((f) => f.includes("Popup"))).toBe(false);
  });

  it.each([
    ["network_error", new BrowserAuthError("network_error")],
    ["timed_out", new BrowserAuthError("timed_out")],
    ["post_request_failed", new BrowserAuthError("post_request_failed")],
    ["no_network_connectivity", new BrowserAuthError("no_network_connectivity")],
    ["서버 오류", new Error("503 Service Unavailable")],
    ["fetch 실패", new TypeError("Failed to fetch")],
  ])("연결 문제(%s)는 그대로 던진다: 만료로 표시하지 않고, redirect·캐시 삭제·요청 abort 도 없다", async (_name, err) => {
    fake.accounts = [acct("woo")];
    fake.silent = [err];
    const m = await load();
    const req = new AbortController();
    m.trackRequest(req);
    const p = m.accessToken(CLIENT);
    await expect(p).rejects.toBe(err);
    await p.catch((e) => {
      expect(e).not.toBeInstanceOf(m.LoginRequiredError);
      expect(m.isLoginFailureText((e as Error).message)).toBe(false);
    });
    expect(m.getAuthState()).not.toBe("expired");
    expect(req.signal.aborted).toBe(false);
    expect(fns()).not.toContain("ssoSilent");
    expect(fns()).not.toContain("acquireTokenRedirect");
    expect(fns()).not.toContain("loginRedirect");
    expect(cacheCleared()).toBe(false);
    expect(fake.accounts).toHaveLength(1);
  });

  it("401 뒤 forceRefresh 실패(네트워크)도 캐시를 지우지 않고 로그인 상태도 그대로", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["at-1", new BrowserAuthError("network_error")];
    const m = await load();
    await m.accessToken(CLIENT, { interactive: false });
    await expect(m.accessToken(CLIENT, { forceRefresh: true, interactive: false })).rejects.toThrow();
    expect(cacheCleared()).toBe(false);
    expect(m.getAuthState()).toBe("ok");
    expect(m.currentAccount(CLIENT)).toBe("woo@outlook.com");
  });

  it("캐시에 계정이 없으면 LoginRequiredError(화면 이동 없음), 상태 signedOut", async () => {
    const m = await load();
    await expect(m.accessToken(CLIENT, { interactive: false })).rejects.toBeInstanceOf(m.LoginRequiredError);
    expect(fns()).not.toContain("loginRedirect");
    expect(m.getAuthState()).toBe("signedOut");
  });
});

describe("initAuth: 캐시 계정은 '로그인됨'이 아니다 — 토큰으로 확인", () => {
  it("계정 캐시 + 조용한 토큰 성공 → ok", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["at-1"];
    const m = await load();
    await m.initAuth(CLIENT);
    expect(fns().slice(0, 2)).toEqual(["initialize", "handleRedirectPromise"]);
    expect(m.getAuthState()).toBe("ok");
  });

  it("계정 캐시 + InteractionRequired → 렌더 전에 이미 expired", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = [new InteractionRequiredAuthError("interaction_required")];
    const m = await load();
    await m.initAuth(CLIENT);
    expect(m.getAuthState()).toBe("expired");
    expect(fns()).not.toContain("ssoSilent");
  });

  it("계정 캐시 + 토큰이 멈춤 → probeMs 뒤 화면은 뜨고, 상태는 ok 가 아니다", async () => {
    fake.accounts = [acct("woo")];
    fake.silent = ["HANG"];
    const m = await load();
    m.authTimeouts.probeMs = 20;
    const t0 = Date.now();
    await m.initAuth(CLIENT);
    expect(Date.now() - t0).toBeLessThan(2000);
    expect(m.getAuthState()).not.toBe("ok");
    m.authTimeouts.probeMs = 8_000;
  });

  it("계정 없음 → signedOut, 막 돌아온 로그인 → ok(토큰 확인 없이)", async () => {
    const m = await load();
    await m.initAuth(CLIENT);
    expect(m.getAuthState()).toBe("signedOut");
    fake.redirect = { account: acct("woo") };
    const again = await load();
    await again.initAuth(CLIENT);
    expect(again.getAuthState()).toBe("ok");
  });
});

describe("대화형 로그인은 항상 redirect", () => {
  it("홈 화면 앱(standalone)·iPhone 에서도 팝업을 쓰지 않는다. 평소 다시 로그인은 login_hint, 계정 선택 화면 없음", async () => {
    vi.stubGlobal("navigator", { standalone: true, userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X)" });
    fake.accounts = [acct("woo")];
    const m = await load();
    await m.loginRedirect(CLIENT);
    expect(fns().some((f) => f.includes("Popup"))).toBe(false);
    const req = fake.calls.find((c) => c.fn === "loginRedirect")!.arg as Record<string, unknown>;
    expect(req).toMatchObject({ loginHint: "woo@outlook.com", redirectUri: "https://uhyun6907-boop.github.io/" });
    expect(req.prompt).toBeUndefined();
  });

  it("처음 로그인(캐시 계정 없음)도 prompt 없이 loginRedirect, '다른 계정'만 select_account", async () => {
    const m = await load();
    await m.loginRedirect(CLIENT);
    await m.loginRedirect(CLIENT, { switchAccount: true });
    const logins = fake.calls.filter((c) => c.fn === "loginRedirect").map((c) => (c.arg as { prompt?: string }).prompt);
    expect(logins).toEqual([undefined, "select_account"]);
    expect(fake.calls.filter((c) => c.fn === "loginRedirect").every((c) => (c.arg as { scopes: string[] }).scopes.includes("offline_access"))).toBe(true);
  });

  it("Pages 주소: BASE_URL 공식과 등록된 리디렉션 URI 가 끝 슬래시까지 같다", async () => {
    const m = await load();
    expect(new URL("/paut-app/", "https://uhyun6907-boop.github.io").href).toBe(m.PAGES_REDIRECT_URI);
    expect(m.buildRedirectUri("/paut-app/", "https://uhyun6907-boop.github.io")).toBe("https://uhyun6907-boop.github.io/paut-app/");
  });
});
