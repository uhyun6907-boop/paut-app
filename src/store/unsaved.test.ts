/**
 * 버그: 아무것도 고치지 않았는데 1번 탭 '저장 · 미저장 10'.
 * 원인: 미저장 수가 현장의 모든 회색(미입력) 카드를 더했다 — 지난날 + 로 만든 S29~S33(5) + S→R 자동이 따라 만든 R29~R33(5).
 * 회색 카드는 날짜가 없는 자리 표시라 저장해도 엑셀에 쓰지 않고 남는다 → 다음 날(11/12) 켜도 계속 '미저장 10'.
 * 실제 드림 합정주공 HJ-M-005-350A(S01~S28·R01~R28 이 이미 엑셀에) 사본으로 다시 만든다. 네트워크 없음(가짜 Graph).
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { DraftCard, InspectionRecord } from "../domain/types";
import type { FakeDrive } from "../test/fakeGraph";
import { fixture, graphConfig, meta, realDrive, realListRecords } from "../test/realBooks";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
let drive: FakeDrive;

vi.mock("./db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => {
    mem.set(key, structuredClone(value));
  },
  kvDel: async (key: string) => {
    mem.delete(key);
  },
}));

const tokenCalls: { forceRefresh?: boolean; interactive?: boolean }[] = [];
/** 가짜 로그인 상태: "ok" 면 토큰을 주고, "lost" 면 LoginRequiredError. */
const auth = vi.hoisted(() => ({
  mode: "ok" as "ok" | "lost" | "lostAfter" | "hang" | "hangAfter" | "expired" | "network",
  /** true 면 캐시(force 아님)는 옛 토큰을 준다(드라이브는 이미 새 토큰만 받음). */
  stale: false,
  /** true 면 force 로 받아도 옛 토큰(실제로 로그인이 끊겨 Graph 가 계속 401). */
  forceBroken: false,
  /** lostAfter: 이 횟수만큼 토큰을 준 뒤 로그인이 끊긴다. */
  okCalls: 0,
  account: "tester" as string | null,
  fromRedirect: false,
  logins: [] as { switchAccount?: boolean }[],
}));
vi.mock("../sync/msal", () => {
  const LOGIN_NEEDED_TEXT = "Microsoft 로그인 필요 — 화면 위 '로그인'을 누르면 대기 중인 기록을 바로 보냅니다.";
  class LoginRequiredError extends Error {
    constructor(message = LOGIN_NEEDED_TEXT) {
      super(message);
      this.name = "LoginRequiredError";
    }
  }
  return {
    LOGIN_NEEDED_TEXT,
    LoginRequiredError,
    accessToken: async (_clientId: string, options: { forceRefresh?: boolean; interactive?: boolean } = {}) => {
      tokenCalls.push(options);
      // 실제 msal.ts 처럼 전역 인증 상태를 바꾼다(같은 모듈 레지스트리의 authState).
      const st = await import("../sync/authState");
      if (auth.mode === "expired") {
        // 계정은 캐시에 있으나 acquireTokenSilent → InteractionRequiredAuthError
        st.markExpired();
        throw new st.AuthExpiredError();
      }
      if (auth.mode === "network") throw Object.assign(new Error("network_error: Network request failed"), { name: "BrowserAuthError", errorCode: "network_error" });
      if (auth.mode === "lost") throw new LoginRequiredError();
      if (auth.mode === "lostAfter") {
        if (auth.okCalls <= 0) throw new LoginRequiredError();
        auth.okCalls -= 1;
      }
      if (auth.mode === "hang" || (auth.mode === "hangAfter" && auth.okCalls-- <= 0)) return new Promise<string>(() => undefined);
      if (auth.forceBroken || (auth.stale && !options.forceRefresh)) return "stale-token";
      st.markAuthOk();
      return drive.validToken;
    },
    completeRedirect: async () => ({ account: auth.account ? { username: auth.account } : null, fromRedirect: auth.fromRedirect }),
    currentAccount: () => auth.account,
    resolveClientId: (id: string) => id || "test-client",
    loginRedirect: async (_clientId: string, options: { switchAccount?: boolean } = {}) => {
      auth.logins.push(options);
    },
    logoutRedirect: async () => undefined,
  };
});

vi.stubGlobal("localStorage", {
  getItem: (key: string) => bag.get(key) ?? null,
  setItem: (key: string, value: string) => void bag.set(key, String(value)),
  removeItem: (key: string) => void bag.delete(key),
});
vi.stubGlobal("fetch", (input: RequestInfo | URL, init?: RequestInit) => drive.fetch(input, init));

let online = true;
vi.stubGlobal("navigator", {
  get onLine() {
    return online;
  },
  vibrate: () => true,
});
vi.stubGlobal("window", { setTimeout, clearTimeout });

type Session = typeof import("./session");

/** 앱을 새로 켠다: 모듈을 다시 읽고(메모리 상태 초기화), IndexedDB(mem)는 그대로. */
async function startApp(tune?: (s: Session, graph: typeof import("../sync/graph")) => void): Promise<Session> {
  vi.resetModules();
  const graph = await import("../sync/graph");
  graph.graphRetry.baseMs = 0;
  graph.graphRetry.maxWaitMs = 0;
  const s = await import("./session");
  tune?.(s, graph);
  await s.initSession();
  await s.syncIdle();
  return s;
}

function seed(records: InspectionRecord[], partialMeta = {}): void {
  mem.set("config", graphConfig());
  mem.set("meta", meta(partialMeta));
  mem.set("records", records);
  mem.set("book:list", fixture("real-list.xlsx"));
  mem.set("book:ilji", fixture("real-worklog.xlsx"));
  mem.set("book:statement", fixture("real-statement.xlsx"));
}


const SITE = "드림 합정주공";
const LINE = "HJ-M-005-350A";
const grey = (prefix: string): DraftCard[] => [29, 30, 31, 32, 33].map((n) => ({ id: `${prefix}${n}-grey`, point: `${prefix}${n}` }));
/** 엑셀 내용을 바꾸는 요청(PATCH·PUT·DELETE·행 삽입·시트 추가). 읽기 세션 열고 닫기(createSession/closeSession)는 내용을 바꾸지 않는다. */
const writes = () => drive.log.filter((e) => e.method !== "GET" && !/\/(createSession|closeSession)$/.test(e.path));

beforeEach(() => {
  mem.clear();
  bag.clear();
  drive = realDrive();
  online = true;
});
afterEach(() => vi.useRealTimers());

function lineRecs(records: InspectionRecord[], prefix: string): InspectionRecord[] {
  return records.filter((r) => r.site === SITE && r.line === LINE && !r.deleted && r.point.startsWith(prefix));
}

describe("1번 탭 '미저장' 은 사용자가 고친 기록만", () => {
  it("S01~S28·R01~R28 이 엑셀에 있고 지난날 만든 회색 S29~S33·R29~R33 만 남은 채 다음 날(11/12)에 켜면 미저장 0, 엑셀 쓰기 없음", async () => {
    const records = realListRecords();
    // 사본에는 이 라인 S 27·R 23 개가 엑셀 행으로 있다(실제 화면은 S01~S28·R01~R28).
    expect(lineRecs(records, "S").length).toBeGreaterThanOrEqual(20);
    expect(lineRecs(records, "R").length).toBeGreaterThanOrEqual(20);
    const drafts = { [`${SITE}||${LINE}||S`]: grey("S"), [`${SITE}||${LINE}||R`]: grey("R") };
    seed(records, { workDate: "2026-11-12", lastSite: SITE, activeLine: LINE, prefix: "S", insp1: "이연행", insp2: "", timeStart: "12:00", timeEnd: "19:00", draftCards: drafts });
    // 지난 세션의 미저장 백업(localStorage)에도 같은 회색 카드만 있다(고친 기록 없음).
    bag.set("paut-unsaved", JSON.stringify({ records: [], draftCards: drafts, stamps: {} }));
    const before = structuredClone(records);

    // 첫 실행은 기존 백필(리스트엔 있는데 작업일지·명세서에 빠진 11/11 7개소)을 한 번 쓴다 — 미저장과 무관, 회색 카드와 무관.
    await startApp();
    drive.log.length = 0;
    // 다시 켠다(사용자가 앱을 연 상황). 이제부터는 아무 쓰기도 없어야 한다.
    const s = await startApp();
    const st = s.getState();
    // 회색 카드는 그대로 보인다(지우지 않음) — 세지만 않는다.
    expect(Object.values(st.meta.draftCards ?? {}).flat()).toHaveLength(10);
    // 예전 계산(미저장 기록 + 현장의 모든 회색 카드) = 0 + 10 → 화면의 '미저장 10'.
    const oldCount = st.records.filter((r) => r.sync === "local" && r.site === SITE && r.date === st.meta.workDate).length
      + Object.entries(st.meta.draftCards ?? {}).filter(([k]) => k.startsWith(`${SITE}||`)).reduce((n, [, l]) => n + l.length, 0);
    expect(oldCount).toBe(10);
    expect(s.unsavedEdits(st.records, st.meta)).toBe(0);
    expect(s.unsavedCount()).toBe(0);
    expect(st.records.filter((r) => r.sync === "local")).toHaveLength(0);
    expect(s.syncQueue(st.records)).toHaveLength(0);
    // 켜기·불러오기는 엑셀에 아무것도 쓰지 않는다.
    expect(writes()).toEqual([]);

    // 날짜·라인·접두어를 바꿔도 미저장이 생기지 않고 쓰지 않는다.
    await s.updateMeta({ workDate: "2026-11-11" });
    await s.updateMeta({ workDate: "2026-11-12", prefix: "R" });
    await s.settled();
    await s.syncIdle();
    expect(s.unsavedEdits(s.getState().records, s.getState().meta)).toBe(0);
    expect(s.syncQueue(s.getState().records)).toHaveLength(0);
    expect(writes()).toEqual([]);
    // 기록 자체(엑셀 행·R 기록)도 그대로: R01~R28 을 다시 대기열에 넣지 않는다.
    const after = s.getState().records;
    expect(lineRecs(after, "R").map((r) => [r.id, r.sync, r.result, r.date])).toEqual(lineRecs(before, "R").map((r) => [r.id, r.sync, r.result, r.date]));
  });

  it("회색 카드를 + 로 더해도 미저장이 아니고, 실제로 카드를 누르면(사용자 편집) 그때만 미저장 1", async () => {
    seed(realListRecords(), { workDate: "2026-11-12", lastSite: SITE, activeLine: LINE, prefix: "S" });
    await startApp(); // 첫 실행: 기존 백필(11/11 작업일지·명세서)만 한 번
    drive.log.length = 0;
    const s = await startApp();
    expect(writes()).toEqual([]);
    drive.log.length = 0;
    s.setSrMirrorEnabled(true);
    await s.setLineDrafts(`${SITE}||${LINE}||S`, [{ id: "g29", point: "S29" }]);
    await s.mirrorGreyDraft("S29"); // S→R 자동: 회색 R29
    await s.settled();
    expect(Object.values(s.getState().meta.draftCards ?? {}).flat().map((c) => c.point).sort()).toEqual(["R29", "S29"]);
    expect(s.unsavedEdits(s.getState().records, s.getState().meta)).toBe(0);
    expect(writes()).toEqual([]);

    await s.savePoint("합격", "S29");
    await s.settled();
    const st = s.getState();
    const local = st.records.filter((r) => r.sync === "local" && r.site === SITE && r.date === "2026-11-12");
    expect(local.length).toBeGreaterThanOrEqual(1);
    expect(s.unsavedEdits(st.records, st.meta)).toBe(local.length);
    // 저장 전에는 엑셀에 쓰지 않는다.
    expect(writes()).toEqual([]);
  });
});
