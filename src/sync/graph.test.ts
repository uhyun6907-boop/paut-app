import { describe, expect, it } from "vitest";
import { listRowFormulas, worksheetRangeUrl } from "./graph";

describe("Graph 엑셀 범위", () => {
  it("워크시트 id와 주소를 경로에 넣는다", () => {
    const url = worksheetRangeUrl("item 1", "{ABC}", "F6:F2500");
    expect(url).toContain("/me/drive/items/item%201/workbook/worksheets/%7BABC%7D/range(address='F6:F2500')");
  });

  it("관경 칸은 RIGHT 수식이고 합부는 값이다", () => {
    const row = listRowFormulas(433, {
      id: "abc",
      date: "2026-11-09",
      line: "HJ-M-099-350A",
      point: "S01",
      result: "합격",
      material: "SPPS380",
    });
    expect(row[5]).toBe("HJ-M-099-350A");
    expect(row[6]).toBe("S01");
    expect(row[11]).toBe("=RIGHT(F433,4)");
    expect(row[15]).toBe("합격");
    expect(row[24]).toBe("abc");
    expect(row[10]).toBe("SPPS380");
  });
});

import { GraphClient, GraphError, graphRetry } from "./graph";
import { toRangeWrites } from "../excel/cellDiff";

function scripted(statuses: number[], seen: { url: string; method: string; auth: string }[], retryAfter?: string) {
  return async (input: RequestInfo | URL, init: RequestInit = {}) => {
    const status = statuses.shift() ?? 200;
    seen.push({ url: String(input), method: init.method ?? "GET", auth: new Headers(init.headers).get("Authorization") ?? "" });
    const headers: Record<string, string> = retryAfter ? { "Retry-After": retryAfter } : {};
    return new Response(status === 200 ? "{}" : JSON.stringify({ error: { message: `s${status}` } }), { status, headers });
  };
}

describe("Graph 재시도", () => {
  it("429 는 Retry-After 만큼 기다렸다 다시 보낸다", async () => {
    const seen: { url: string; method: string; auth: string }[] = [];
    const waits: number[] = [];
    const c = new GraphClient({ getToken: async () => "t", fetchImpl: scripted([429, 200], seen, "3"), sleep: async (ms) => void waits.push(ms) });
    await c.request("https://graph.microsoft.com/v1.0/x");
    expect(seen).toHaveLength(2);
    expect(waits[0]).toBeGreaterThanOrEqual(Math.min(3000, graphRetry.maxWaitMs));
  });

  it("잠김(423)이 계속되면 한국어 이유로 실패한다", async () => {
    const seen: { url: string; method: string; auth: string }[] = [];
    const c = new GraphClient({ getToken: async () => "t", fetchImpl: scripted([423, 423, 423, 423, 423], seen), sleep: async () => undefined });
    const err = await c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    expect(err).toBeInstanceOf(GraphError);
    expect((err as GraphError).status).toBe(423);
    expect((err as GraphError).message).toContain("잠겨");
    expect(seen).toHaveLength(graphRetry.attempts);
  });

  it("401 이면 토큰을 강제로 새로 받아 새 토큰으로 다시 보낸다", async () => {
    const seen: { url: string; method: string; auth: string }[] = [];
    const calls: boolean[] = [];
    const c = new GraphClient({
      getToken: async (force) => {
        calls.push(force);
        return force ? "new" : "old";
      },
      fetchImpl: scripted([401, 200], seen),
      sleep: async () => undefined,
    });
    await c.request("https://graph.microsoft.com/v1.0/x");
    expect(calls).toEqual([false, true]);
    expect(seen.map((s) => s.auth)).toEqual(["Bearer old", "Bearer new"]);
  });

  it("행 삽입은 5xx 에 다시 보내지 않는다(두 번 들어갈 수 있음)", async () => {
    const seen: { url: string; method: string; auth: string }[] = [];
    const c = new GraphClient({ getToken: async () => "t", fetchImpl: scripted([500, 200], seen), sleep: async () => undefined });
    const err = await c.request("https://graph.microsoft.com/v1.0/x", { method: "POST" }, { idempotent: false }).catch((e) => e);
    expect((err as GraphError).status).toBe(500);
    expect(seen).toHaveLength(1);
  });
});

describe("칸 쓰기 묶음", () => {
  it("한 행에서 이어진 칸끼리만 묶고 빈틈은 건드리지 않는다", () => {
    const ranges = toRangeWrites([
      { col: "B", row: 522, value: 46337 },
      { col: "C", row: 522, value: "평일" },
      { col: "J", row: 522, value: "350A" },
      { col: "K", row: 522, value: 7 },
      { col: "P", row: 522, value: "조우현" },
      { col: "S", row: 522, value: "" },
    ]);
    expect(ranges).toEqual([
      { address: "B522:C522", formulas: [[46337, "평일"]] },
      { address: "J522:K522", formulas: [["350A", 7]] },
      { address: "P522", formulas: [["조우현"]] },
      { address: "S522", formulas: [[""]] },
    ]);
  });
});

describe("Graph 요청 ↔ 전역 인증 상태", () => {
  it("새 토큰으로도 401 → 묶음이 끝나길 기다리지 않고 즉시 markExpired: 다른 진행 중 요청도 abort, GRAPH_LOGIN_TEXT", async () => {
    const auth = await import("./authState");
    auth.resetAuthForTests();
    const { GRAPH_LOGIN_TEXT } = await import("./graph");
    const states: string[] = [];
    auth.onAuthChange((st) => states.push(st));
    // 다른 화면(리스트 탭)의 멈춘 요청
    const other = new AbortController();
    auth.trackRequest(other);
    let tokens = 0;
    const fetchImpl = async () => new Response(JSON.stringify({ error: { message: "InvalidAuthenticationToken" } }), { status: 401 });
    const c = new GraphClient({ getToken: async () => `t${++tokens}`, fetchImpl, sleep: async () => undefined });
    const err = await c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    expect(err).toBeInstanceOf(GraphError);
    expect((err as GraphError).message).toBe(GRAPH_LOGIN_TEXT);
    expect(tokens).toBe(2); // 첫 토큰 + 강제 새 토큰 한 번
    expect(c.authFailed).toBe(true);
    expect(auth.getAuthState()).toBe("expired");
    expect(states).toEqual(["expired"]);
    expect(other.signal.aborted).toBe(true);
    expect(auth.inflightCount()).toBe(0);
    // 이후 요청은 보내지 않고 바로 로그인 오류
    const again = await c.request("https://graph.microsoft.com/v1.0/y").catch((e) => e);
    expect((again as GraphError).status).toBe(401);
    auth.resetAuthForTests();
  });

  it("진행 중 fetch 는 전역에 등록되고 끝나면 해제된다. 만료 abort 는 재시도하지 않는다(fetch 가 signal 을 무시해도)", async () => {
    const auth = await import("./authState");
    auth.resetAuthForTests();
    let calls = 0;
    const fetchImpl = () => {
      calls += 1;
      return new Promise<Response>(() => undefined); // signal 무시하고 영원히 멈춤
    };
    const c = new GraphClient({ getToken: async () => "t", fetchImpl, sleep: async () => undefined });
    const p = c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    for (let i = 0; i < 20 && auth.inflightCount() === 0; i++) await new Promise((r) => setTimeout(r, 1));
    expect(auth.inflightCount()).toBe(1);
    auth.markExpired();
    const err = await p;
    expect((err as GraphError).status).toBe(401);
    expect(calls).toBe(1);
    expect(auth.inflightCount()).toBe(0);
    auth.resetAuthForTests();
  });

  it("시간 초과는 그 요청의 AbortController 를 실제로 abort 하고, 로그인 만료로 바꾸지 않는다", async () => {
    const auth = await import("./authState");
    auth.resetAuthForTests();
    const { graphTimeouts } = await import("./graph");
    const signals: AbortSignal[] = [];
    const fetchImpl = (_u: RequestInfo | URL, init?: RequestInit) => {
      signals.push(init!.signal!);
      return new Promise<Response>(() => undefined);
    };
    const c = new GraphClient({ getToken: async () => "t", fetchImpl, sleep: async () => undefined });
    const old = graphTimeouts.requestMs;
    graphTimeouts.requestMs = 20;
    const err = await c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    graphTimeouts.requestMs = old;
    expect((err as GraphError).detail).toBe("timeout");
    expect(signals[0].aborted).toBe(true);
    expect(auth.getAuthState()).toBe("unknown");
    expect(auth.inflightCount()).toBe(0);
  });

  it("토큰이 시간 초과면 같은 클라이언트의 다른 요청도 끊는다(로그인 상태는 그대로)", async () => {
    const auth = await import("./authState");
    auth.resetAuthForTests();
    const { graphTimeouts } = await import("./graph");
    const old = graphTimeouts.tokenMs;
    graphTimeouts.tokenMs = 20;
    const c = new GraphClient({ getToken: () => new Promise<string>(() => undefined), fetchImpl: async () => new Response("{}"), sleep: async () => undefined });
    const err = await c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    graphTimeouts.tokenMs = old;
    expect((err as Error).name).toBe("TimeoutError");
    expect(c.stalled).toBe(true);
    expect(c.isAborted).toBe(true);
    expect(auth.getAuthState()).toBe("unknown");
  });

  it("연결 실패(fetch TypeError)는 만료가 아니다: 재시도 후 '인터넷' 이유, 상태 그대로", async () => {
    const auth = await import("./authState");
    auth.resetAuthForTests();
    const fetchImpl = async () => {
      throw new TypeError("Failed to fetch");
    };
    const c = new GraphClient({ getToken: async () => "t", fetchImpl, sleep: async () => undefined });
    const err = await c.request("https://graph.microsoft.com/v1.0/x").catch((e) => e);
    expect((err as GraphError).status).toBe(0);
    expect((err as GraphError).message).not.toContain("만료");
    expect(auth.getAuthState()).toBe("unknown");
    expect(c.authFailed).toBe(false);
  });
});
