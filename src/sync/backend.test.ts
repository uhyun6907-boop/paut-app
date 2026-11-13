import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// 브라우저 저장소·주소 흉내
class MemStorage {
  private m = new Map<string, string>();
  getItem(k: string) {
    return this.m.has(k) ? this.m.get(k)! : null;
  }
  setItem(k: string, v: string) {
    this.m.set(k, String(v));
  }
  removeItem(k: string) {
    this.m.delete(k);
  }
}

async function load(hash = "") {
  vi.resetModules();
  vi.stubEnv("VITE_PAUT_API", "https://api.example.dev/");
  const ls = new MemStorage();
  vi.stubGlobal("localStorage", ls);
  vi.stubGlobal("window", { location: { hash, pathname: "/paut-app/", search: "", assign: vi.fn() } });
  vi.stubGlobal("history", { replaceState: vi.fn() });
  const auth = await import("./authState");
  auth.resetAuthForTests();
  const backend = await import("./backend");
  return { ls, auth, backend };
}

describe("로그인 중계 서버 모드", () => {
  beforeEach(() => vi.unstubAllGlobals());
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.unstubAllGlobals();
  });

  it("돌아온 주소의 기기 열쇠를 저장하고 로그인됨", async () => {
    const { ls, auth, backend } = await load("#paut_device=DEV&paut_email=a%40b.com&paut_state=S");
    ls.setItem("paut-login-state", "S");
    const out = backend.completeBackendRedirect();
    expect(out).toMatchObject({ fromRedirect: true, email: "a@b.com" });
    expect(ls.getItem("paut-device-token")).toBe("DEV");
    expect(auth.getAuthState()).toBe("ok");
  });

  it("상태값이 다르면 열쇠를 받지 않는다", async () => {
    const { ls, backend } = await load("#paut_device=DEV&paut_state=X");
    ls.setItem("paut-login-state", "S");
    const out = backend.completeBackendRedirect();
    expect(out.fromRedirect).toBe(false);
    expect(ls.getItem("paut-device-token")).toBeNull();
  });

  it("서버가 login_required 면 만료, 일시 오류면 상태 유지", async () => {
    const { ls, auth, backend } = await load();
    ls.setItem("paut-device-token", "DEV");
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ error: "upstream" }), { status: 502 })));
    await expect(backend.backendToken()).rejects.toThrow(/로그인 서버 오류/);
    expect(auth.getAuthState()).toBe("unknown");
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({ error: "login_required" }), { status: 401 })));
    await expect(backend.backendToken()).rejects.toMatchObject({ name: "AuthExpiredError" });
    expect(auth.getAuthState()).toBe("expired");
  });

  it("토큰을 받아 캐시하고 기기 열쇠를 Bearer 로 보낸다", async () => {
    const { ls, auth, backend } = await load();
    ls.setItem("paut-device-token", "DEV");
    const f = vi.fn(async () => new Response(JSON.stringify({ access_token: "AT", expires_at: Date.now() + 3600_000, email: "a@b.com" }), { status: 200 }));
    vi.stubGlobal("fetch", f);
    expect(await backend.backendToken()).toBe("AT");
    expect(await backend.backendToken()).toBe("AT");
    expect(f).toHaveBeenCalledTimes(1);
    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("https://api.example.dev/api/token");
    expect((init.headers as Record<string, string>).authorization).toBe("Bearer DEV");
    expect(auth.getAuthState()).toBe("ok");
  });

  it("열쇠가 없으면 로그인 필요", async () => {
    const { auth, backend } = await load();
    await expect(backend.backendToken()).rejects.toMatchObject({ name: "LoginRequiredError" });
    expect(auth.getAuthState()).toBe("signedOut");
  });
});
