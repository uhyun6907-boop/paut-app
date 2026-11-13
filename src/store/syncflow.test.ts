/**
 * 세션(앱 상태) 수준 동기화: 파일별 실패 표시, '다시 시도', IndexedDB 에 남은 대기열로 재시작 후 이어 보내기, 오프라인.
 * Graph 는 가짜 드라이브(src/test/fakeGraph.ts)이고 실제 OneDrive 는 부르지 않는다.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { InspectionRecord, SyncReport } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import { parseListWorkbook } from "../excel/listBook";
import { readPautRows, readStatement } from "../excel/statementBook";
import { readSheet } from "../excel/ooxml";
import type { FakeDrive } from "../test/fakeGraph";
import { ILJI_ID, LIST_GM_ID, LIST_GM_NAME, LIST_ID, STMT_ID, fixture, graphConfig, meta, realDrive, realListRecords, rec } from "../test/realBooks";
import { userCreatesMonthSheet } from "../test/userSheet";
import { workbookSheets } from "../excel/ooxml";
import { headerStatus, loginBannerModel, settingsAuthModel, syncPanelModel } from "./syncStatus";

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

function iljiDayRow(iso: string): { J: string; K: string; P: string; Q: string } | null {
  const { cells } = readSheet(drive.files(ILJI_ID), " PAUT FIELD");
  const serial = String(Math.round((Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000));
  const rows = [...cells.values()].filter((c) => c.col === "B" && c.v === serial).map((c) => c.row);
  if (rows.length !== 1) return rows.length ? ({ J: "dup", K: String(rows.length), P: "", Q: "" }) : null;
  const v = (col: string) => cells.get(col + rows[0])?.v ?? "";
  return { J: v("J"), K: v("K"), P: v("P"), Q: v("Q") };
}

function listIds(): string[] {
  return parseListWorkbook(drive.files(LIST_ID)).flatMap((s) => s.records).filter((r) => r.src === "user").map((r) => r.id);
}

/** 화면 위 상태 줄(App.tsx 와 같은 계산). */
function head(s: Session) {
  const st = s.getState();
  return headerStatus({ mode: st.config.syncMode, online: st.online, syncing: st.syncing, loginNeeded: st.loginNeeded, account: st.account, queue: s.syncQueue(st.records), auth: st.auth });
}

function model(s: Session) {
  const st = s.getState();
  return syncPanelModel(st.records, st.syncReport, { online: st.online, syncing: st.syncing, mode: st.config.syncMode, loginNeeded: st.loginNeeded });
}

beforeEach(() => {
  mem.clear();
  bag.clear();
  tokenCalls.length = 0;
  online = true;
  auth.mode = "ok";
  auth.stale = false;
  auth.forceBroken = false;
  auth.okCalls = 0;
  auth.account = "tester";
  auth.fromRedirect = false;
  auth.logins.length = 0;
  drive = realDrive();
});

afterEach(() => {
  drive.clearFaults();
});

describe("앱 동기화 흐름", () => {
  it("리스트만 성공하고 작업일지가 잠겨 실패하면 화면에 '작업일지 · 실패'와 이유가 보이고, 다시 시도하면 중복 없이 끝난다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS), rec("S29", PASS)];
    seed([...realListRecords(), ...today], {
      dayStaff: { "2026-11-12": { dayType: "평일", insp1: "조우현", insp2: "심명보", timeStart: "09:00", timeEnd: "15:30" } },
    });
    drive.fail((m, p) => m === "PATCH" && p.includes(ILJI_ID), 423, Infinity);

    const s = await startApp();
    const st = s.getState();
    const mine = st.records.filter((r) => today.some((t) => t.id === r.id));
    expect(mine.map((r) => r.sync)).toEqual(["error", "error", "error"]);
    for (const r of mine) {
      expect(r.steps?.list?.status).toBe("done");
      expect(r.steps?.ilji?.status).toBe("error");
      expect(r.steps?.statement?.status).toBe("done");
      expect(r.syncError).toMatch(/^작업일지: PC 엑셀/);
    }
    // 11/11 앱 기록(이미 리스트에만 있음)도 작업일지 실패로 대기열에 남는다.
    const nov11 = st.records.filter((r) => r.src === "user" && r.date === "2026-11-11");
    expect(nov11).toHaveLength(7);
    expect(nov11.every((r) => r.sync === "error" && r.steps?.ilji?.status === "error")).toBe(true);
    expect(st.syncReport?.files.ilji.status).toBe("error");
    expect(st.syncReport?.files.list.status).toBe("ok");
    expect(st.toast?.text).toContain("작업일지");

    const view = model(s);
    expect(view.visible).toBe(true);
    expect(view.tone).toBe("bad");
    expect(view.title).toContain("엑셀 반영 실패");
    expect(view.lines.find((l) => l.file === "ilji")?.text).toMatch(/^작업일지 · 실패 — PC 엑셀 등에서 파일을 쓰는 중이라 잠겨 있습니다.*\(10건\)$/);
    expect(view.lines.find((l) => l.file === "list")?.state).toBe("ok");
    expect(view.canRetry).toBe(true);
    expect(s.syncLabel(mine[0])).toBe("작업일지 실패");

    // IndexedDB 에 파일별 상태가 남는다.
    const saved = mem.get("records") as InspectionRecord[];
    expect(saved.find((r) => r.id === today[0].id)?.steps?.ilji?.status).toBe("error");
    expect((mem.get("syncReport") as SyncReport).files.ilji.status).toBe("error");

    // 다시 시도
    drive.clearFaults();
    const listPatches = drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length;
    s.retrySync();
    await s.syncIdle();
    const after = s.getState();
    expect(after.records.filter((r) => r.sync === "error")).toHaveLength(0);
    expect(after.records.filter((r) => today.some((t) => t.id === r.id)).map((r) => r.sync)).toEqual(["synced", "synced", "synced"]);
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length).toBe(listPatches);
    expect(listIds()).toHaveLength(10);
    expect(new Set(listIds()).size).toBe(10);
    expect(iljiDayRow("2026-11-12")).toEqual({ J: "350A", K: "3", P: "조우현", Q: "심명보" });
    expect(iljiDayRow("2026-11-11")).toMatchObject({ J: "350A", K: "7", P: "조우현" });
    expect(model(s).visible).toBe(false);
    // 동기화는 화면을 넘기는 로그인(interactive)을 쓰지 않는다.
    expect(tokenCalls.length).toBeGreaterThan(3);
    expect(tokenCalls.every((c) => c.interactive === false)).toBe(true);
  });

  it("실패한 채로 앱을 껐다 켜도(보내는 중 꺼짐 포함) 대기열이 남아 이어 보내고 11/11 7개소를 채운다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS)];
    seed([...realListRecords(), ...today]);
    drive.fail((_m, p) => p.includes(STMT_ID), 503, Infinity);
    drive.fail((_m, p) => p.includes(ILJI_ID), 503, Infinity);
    const first = await startApp();
    expect(first.getState().records.filter((r) => r.sync === "error")).toHaveLength(9);
    expect(listIds()).toHaveLength(9);

    // 보내는 중 꺼진 기록 하나
    const saved = mem.get("records") as InspectionRecord[];
    saved.find((r) => r.id === today[1].id)!.sync = "syncing";
    mem.set("records", saved);

    drive.clearFaults();
    const s = await startApp();
    const st = s.getState();
    expect(st.records.filter((r) => r.sync !== "synced" && r.result !== "없음")).toHaveLength(0);
    expect(listIds()).toHaveLength(9);
    expect(iljiDayRow("2026-11-11")).toMatchObject({ J: "350A", K: "7" });
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
    const nov = readStatement(drive.files(STMT_ID), "드림합정26.11").paut;
    expect(nov.map((p) => [p.date, p.count])).toEqual([
      ["2026-11-11", 7],
      ["2026-11-12", 2],
    ]);
    expect(st.syncReport?.files.ilji.status).toBe("ok");

    // 한 번 더 켜도 아무것도 늘지 않는다.
    const again = await startApp();
    expect(again.getState().records.filter((r) => r.sync === "error")).toHaveLength(0);
    expect(iljiDayRow("2026-11-11")).toMatchObject({ K: "7" });
    expect(readStatement(drive.files(STMT_ID), "드림합정26.11").paut).toHaveLength(2);
    expect(listIds()).toHaveLength(9);
  });

  it("오프라인이면 조용히 넘기지 않는다: 대기 건수와 '연결되면 자동으로' 안내, 요청은 보내지 않는다", async () => {
    online = false;
    seed([...realListRecords(), rec("S28", PASS)]);
    const s = await startApp();
    const st = s.getState();
    expect(st.records.find((r) => r.point === "S28" && r.date === "2026-11-12")?.sync).toBe("pending");
    expect(st.syncReport?.offline).toBe(true);
    expect(st.toast?.text).toContain("오프라인");
    const view = model(s);
    expect(view.visible).toBe(true);
    expect(view.title).toMatch(/^오프라인 · 1건 대기/);
    expect(view.canRetry).toBe(false);
    expect(drive.log).toHaveLength(0);
  });
});

describe("Microsoft 로그인이 끊겼을 때", () => {
  const banner = (s: Session) => {
    const st = s.getState();
    return loginBannerModel({ mode: st.config.syncMode, loginNeeded: st.loginNeeded, waiting: s.pendingCount(st.records), online: st.online });
  };

  it("시작할 때 로그인이 없으면 '로그인 필요' 배너만 뜨고 대기 기록은 실패로 바뀌지 않는다. 로그인하고 돌아오면 대기열을 자동으로 보낸다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS)];
    seed([...realListRecords(), ...today]);
    auth.mode = "lost";
    const s = await startApp();
    const st = s.getState();
    expect(st.loginNeeded).toBe(true);
    expect(st.records.filter((r) => today.some((t) => t.id === r.id)).map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(st.records.some((r) => r.sync === "error")).toBe(false);
    expect(drive.requests(() => true)).toHaveLength(0);
    expect(banner(s)).toMatchObject({ visible: true, title: "Microsoft 로그인 필요", button: "로그인" });
    expect(banner(s).detail).toContain("대기 2건");
    expect(model(s).visible).toBe(false); // 실패 패널(빨간색)이 아니라 로그인 배너
    // 위 상태 줄: 노란 '동기화 중'이 아니라 '로그인 필요 · 대기 2건'
    expect(head(s)).toEqual({ tone: "login", text: "로그인 필요 · 대기 2건" });

    // 배너를 누르면 계정 선택 화면 없이 redirect 로그인
    await s.signIn();
    expect(auth.logins).toEqual([{}]);

    // Microsoft 에서 돌아옴(redirect) → 계정 복원 + 대기열 자동 전송
    auth.mode = "ok";
    auth.fromRedirect = true;
    const back = await startApp();
    const after = back.getState();
    expect(after.loginNeeded).toBe(false);
    expect(after.account).toBe("tester");
    expect(after.records.filter((r) => today.some((t) => t.id === r.id)).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(listIds()).toEqual(expect.arrayContaining(today.map((r) => r.id)));
    expect(banner(back).visible).toBe(false);
    expect(head(back)).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
  });

  it("보내는 도중 로그인이 끊기면 끝난 파일은 그대로, 남은 파일은 실패가 아니라 대기로 남고 로그인 뒤 그 파일만 보낸다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS)];
    seed([...realListRecords(), ...today]);
    auth.mode = "lostAfter";
    auth.okCalls = 2; // 사전 확인 + 검사리스트 단계까지만
    const s = await startApp();
    const st = s.getState();
    const mine = st.records.filter((r) => today.some((t) => t.id === r.id));
    expect(mine.map((r) => r.steps?.list?.status)).toEqual(["done", "done"]);
    expect(mine.map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(mine.every((r) => r.steps?.ilji?.status !== "error" && r.steps?.statement?.status !== "error")).toBe(true);
    expect(st.loginNeeded).toBe(true);
    expect(model(s).visible).toBe(false);
    const listPatches = drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length;

    auth.mode = "ok";
    auth.fromRedirect = true;
    const back = await startApp();
    expect(back.getState().records.filter((r) => today.some((t) => t.id === r.id)).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(back.getState().loginNeeded).toBe(false);
    // 리스트는 다시 쓰지 않는다(이미 done)
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length).toBe(listPatches);
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
  });

  it("'다른 계정'만 계정 선택 화면을 띄운다", async () => {
    seed(realListRecords());
    const s = await startApp();
    await s.signIn({ switchAccount: true });
    await s.signIn();
    expect(auth.logins).toEqual([{ switchAccount: true }, {}]);
  });

  it("데모 모드에서는 배너가 뜨지 않는다", () => {
    expect(loginBannerModel({ mode: "demo", loginNeeded: true, waiting: 3, online: true }).visible).toBe(false);
    expect(loginBannerModel({ mode: "graph", loginNeeded: false, waiting: 3, online: true }).visible).toBe(false);
    expect(loginBannerModel({ mode: "graph", loginNeeded: true, waiting: 0, online: false }).detail).toContain("오프라인");
  });
});

describe("11/11 HJ-M-005-350A S28 이 리스트에 안 들어간 문제", () => {
  it("보낸 것으로 표시됐지만 리스트 Y열에 앱 ID 가 없는 기록은 다시 켤 때(백필) 찾아 리스트·작업일지·명세서에 채운다", async () => {
    const all = realListRecords();
    const s28 = rec("S28", PASS, {
      date: "2026-11-11",
      sync: "synced",
      steps: { list: { status: "done", at: 1 }, ilji: { status: "done", at: 1 }, statement: { status: "done", at: 1 } },
    });
    seed([...all, s28]);
    const s = await startApp();
    const row = parseListWorkbook(drive.files(LIST_ID))
      .flatMap((x) => x.records)
      .find((r) => r.id === s28.id);
    expect(row).toMatchObject({ excelRow: 440, date: "2026-11-11", line: "HJ-M-005-350A", point: "S28", result: PASS });
    expect(iljiDayRow("2026-11-11")).toMatchObject({ J: "350A", K: "8" });
    expect(readStatement(drive.files(STMT_ID), "드림합정26.11").paut.map((p) => [p.date, p.count])).toEqual([["2026-11-11", 8]]);
    expect(s.getState().records.find((r) => r.id === s28.id)).toMatchObject({ sync: "synced", excelRow: 440 });

    // 다시 켜도 늘지 않는다.
    await startApp();
    expect(listIds().filter((id) => id === s28.id)).toHaveLength(1);
    expect(iljiDayRow("2026-11-11")).toMatchObject({ K: "8" });
  });

  it("지난 날 '미저장'으로 남은 마지막 기록도 다시 켜면 대기열에 넣어 보낸다", async () => {
    const s28 = rec("S28", PASS, { date: "2026-11-11", sync: "local" });
    seed([...realListRecords(), s28], { workDate: "2026-11-12" });
    const s = await startApp();
    expect(listIds()).toContain(s28.id);
    expect(s.getState().records.find((r) => r.id === s28.id)?.sync).toBe("synced");
  });

  it("마지막 카드를 누르자마자 저장해도 그 카드가 저장(대기열)에 들어간다", async () => {
    seed(realListRecords(), { workDate: "2026-11-11" });
    const s = await startApp();
    const tap = s.savePoint(PASS, "S28");
    const save = s.commitDay();
    await Promise.all([tap, save]);
    await s.syncIdle();
    const s28 = s.getState().records.find((r) => r.point === "S28" && r.date === "2026-11-11" && !r.deleted);
    expect(s28?.sync).toBe("synced");
    expect(listIds()).toContain(s28!.id);
  });
});

describe("거래명세서 그 달 시트가 없을 때(앱은 시트를 만들지 않음)", () => {
  it("리스트·작업일지는 반영, 명세서는 시트를 만들어 달라는 요청과 함께 대기. 사용자가 시트를 만든 뒤 다시 켜면 채운다", async () => {
    drive = realDrive({ novSheet: false });
    const sheets = workbookSheets(drive.files(STMT_ID)).map((x) => x.name);
    const today = [rec("S28", PASS), rec("R28", PASS)];
    seed([...realListRecords(), ...today]);
    const s = await startApp();
    const st = s.getState();
    const mine = st.records.filter((r) => today.some((t) => t.id === r.id));
    expect(mine.map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(mine.every((r) => r.steps?.list?.status === "done" && r.steps?.ilji?.status === "done")).toBe(true);
    expect(mine[0].steps?.statement).toMatchObject({ status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (드림합정26.11)" });
    // 11/11 백필 기록도 실패가 아니라 대기열
    const nov11 = st.records.filter((r) => r.src === "user" && r.date === "2026-11-11");
    expect(nov11.every((r) => r.sync === "pending")).toBe(true);
    expect(st.records.filter((r) => r.sync === "error")).toHaveLength(0);
    expect(st.toast?.text ?? "").not.toContain("실패");
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
    expect(workbookSheets(drive.files(STMT_ID)).map((x) => x.name)).toEqual(sheets);
    const view = model(s);
    expect(view.visible).toBe(true);
    expect(view.tone).toBe("ask");
    expect(view.title).toBe("11월 거래명세서 시트를 만들어주세요 (드림합정26.11)");

    // 사용자가 시트를 만들고 앱을 다시 켠다 → 대기 기록을 채운다.
    userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 11);
    const again = await startApp();
    expect(again.getState().records.filter((r) => r.sync === "pending" || r.sync === "error")).toHaveLength(0);
    expect(readStatement(drive.files(STMT_ID), "드림합정26.11").paut.map((p) => [p.date, p.count])).toEqual([
      ["2026-11-11", 7],
      ["2026-11-12", 2],
    ]);
    expect(model(again).visible).toBe(false);
    // 앱이 쓴 줄 기억은 엑셀이 아니라 이 기기(IndexedDB)에 남는다.
    expect(JSON.stringify(mem.get("statementOwned"))).toContain("2026-11-12|350A");
    expect(readSheet(drive.files(STMT_ID), "드림합정26.11").cells.get("I12")?.v ?? "").toBe("");
  });
});


describe("네 번째 파일: PAUT NDT 리스트(군포, 마전).xlsx", () => {
  it("OneDrive 모드로 켜면 드라이브 루트의 군포·마전 리스트를 받아 현장을 넣고 항목 ID 를 기억한다. 대기 중인 군포 기록은 그 파일에만 쓴다", async () => {
    // 설정에 아직 listGm 칸이 없는 예전 설정. 군포 기록은 listBook 없이 시트 이름만 있다.
    const gp = rec("S07", PASS, { site: "지성 군포 PA", sheetName: "지성 군포 PA", line: "군포산본 이원화 750A-5", diameter: "750A" });
    seed([...realListRecords(), gp]);
    const pyXml = () => workbookSheets(drive.files(LIST_ID)).map((x) => readSheet(drive.files(LIST_ID), x.name).xml).join("\n");
    const pyBefore = pyXml();
    const s = await startApp();
    const st = s.getState();
    expect(st.config.files.listGm).toEqual({ id: LIST_GM_ID, name: LIST_GM_NAME });
    expect(mem.get("book:listGm")).toBeTruthy();
    const sites = new Set(st.records.map((r) => r.site));
    for (const site of ["지성 군포 PA", "아전 군포 PA", "지성 마전 PA", "드림 합정주공"]) expect(sites.has(site)).toBe(true);
    expect(st.config.mappings.find((m) => m.site === "지성 군포 PA")).toMatchObject({ iljiColumn: "지성 군포", statementPrefix: "지성군포", unitPrice: 75000 });
    expect(st.config.mappings.find((m) => m.site === "아전 군포 PA")).toMatchObject({ iljiColumn: "아전 군포", statementPrefix: "" });

    const mine = st.records.find((r) => r.id === gp.id)!;
    expect(mine.sync).toBe("synced");
    expect(mine.steps?.list?.status).toBe("done");
    const gmIds = parseListWorkbook(drive.files(LIST_GM_ID), "listGm").flatMap((x) => x.records).filter((r) => r.src === "user").map((r) => r.id);
    expect(gmIds).toEqual([gp.id]);
    expect(listIds()).not.toContain(gp.id);
    // 평택 리스트에는 군포 행이 없다(11/11 백필 확인용 읽기만, 쓰기 없음).
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH")).toHaveLength(0);
    expect(pyXml() === pyBefore).toBe(true);
    expect(st.syncReport?.lists?.listGm).toMatchObject({ status: "ok", count: 1 });

    // 다시 켜도 또 받지 않는다(이미 사본·기록이 있음).
    const downloads = drive.requests((e) => e.item === LIST_GM_ID && e.path.endsWith("/content")).length;
    const again = await startApp();
    expect(drive.requests((e) => e.item === LIST_GM_ID && e.path.endsWith("/content")).length).toBe(downloads);
    expect(again.getState().records.filter((r) => r.site === "지성 군포 PA").length).toBeGreaterThan(100);
  });

  it("군포·마전 리스트를 받지 못해도(파일 없음) 평택·작업일지·명세서 동기화는 그대로 끝난다", async () => {
    drive.items.delete(LIST_GM_ID);
    const today = [rec("S28", PASS)];
    seed([...realListRecords(), ...today]);
    const s = await startApp();
    const st = s.getState();
    expect(st.records.find((r) => r.id === today[0].id)?.sync).toBe("synced");
    expect(st.syncReport?.lists?.listGm?.status).toBe("error");
    expect(st.syncReport?.lists?.listGm?.message).toContain("검사리스트(군포·마전)");
    expect(st.syncReport?.files.list.status).toBe("ok");
  });
});

describe("카드 길게 눌러 날짜 고치기", () => {
  const LINE = "HJ-M-005-350A";
  /** 리스트에서 그 라인의 앱 행(Y 앱 ID)을 행 순서대로: 날짜|포인트 */
  function lineRows(): { id: string; key: string; row: number }[] {
    return parseListWorkbook(drive.files(LIST_ID))
      .flatMap((s) => s.records)
      .filter((r) => r.src === "user" && r.line === LINE)
      .sort((a, b) => (a.excelRow ?? 0) - (b.excelRow ?? 0))
      .map((r) => ({ id: r.id, key: `${r.date}|${r.point}`, row: r.excelRow ?? 0 }));
  }
  function stmtRows(sheet: string) {
    return readPautRows(drive.files(STMT_ID), sheet).map((r) => `${r.date}|${r.diameter}|${r.count}`);
  }
  function sortedKeys(rows: { key: string }[]): string[] {
    return [...rows.map((r) => r.key)].sort((a, b) => {
      const [da, pa] = a.split("|");
      const [db, pb] = b.split("|");
      if (da !== db) return da < db ? -1 : 1;
      const ka = (p: string) => `${p[0] === "S" ? 0 : 1}${p.slice(1).padStart(4, "0")}`;
      return ka(pa).localeCompare(ka(pb));
    });
  }
  function noAppMark(sheet: string): boolean {
    return ![...readSheet(drive.files(STMT_ID), sheet).cells.values()].some((c) => c.col === "I" && c.v === "앱");
  }

  it("같은 달 안에서 날짜를 바꾸면 리스트는 같은 행(Y)을 고쳐 날짜 → 포인트 순으로 다시 놓고, 작업일지·명세서는 옛 날짜에서 빼고 새 날짜에 쓴다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS), rec("S29", PASS), rec("S30", FAIL)];
    seed([...realListRecords(), ...today], {
      dayStaff: { "2026-11-12": { dayType: "평일", insp1: "조우현", insp2: "심명보", timeStart: "09:00", timeEnd: "15:30" } },
    });
    const s = await startApp();
    expect(s.getState().records.filter((r) => today.some((t) => t.id === r.id)).every((r) => r.sync === "synced")).toBe(true);
    const before = lineRows();
    expect(before).toHaveLength(11);
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((x) => x.name);
    expect(stmtRows("드림합정26.11")).toContain("2026-11-12|350A|3");

    // S29(11/12) → 11/10
    await s.changeRecordDate(today[2].id, "2026-11-10");
    await s.settled();
    await s.syncIdle();
    const moved = s.getState().records.find((r) => r.id === today[2].id)!;
    expect(moved.date).toBe("2026-11-10");
    expect(moved.sync).toBe("synced");
    expect(moved.movedFrom).toBeUndefined();

    const after = lineRows();
    // 행이 늘지 않고, 같은 행 번호들 안에서 날짜 → 포인트 순.
    expect(after).toHaveLength(11);
    expect(new Set(after.map((r) => r.id)).size).toBe(11);
    expect(after.map((r) => r.row)).toEqual(before.map((r) => r.row));
    expect(after.map((r) => r.key)).toEqual(sortedKeys(after));
    expect(after[0]).toMatchObject({ id: today[2].id, key: "2026-11-10|S29" });
    // 불합격(S30)은 리스트에는 그대로 있다.
    expect(after.some((r) => r.key === "2026-11-12|S30")).toBe(true);

    // 작업일지: 11/12 는 합격 2개소(S28·R28), 11/10 은 1개소.
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
    expect(iljiDayRow("2026-11-10")).toMatchObject({ J: "350A", K: "1" });
    // 명세서: 같은 11월 시트에서 옮겨진다(중복 없음).
    const nov = stmtRows("드림합정26.11");
    expect(nov).toContain("2026-11-12|350A|2");
    expect(nov).toContain("2026-11-10|350A|1");
    expect(nov.filter((r) => r.startsWith("2026-11-12|"))).toHaveLength(1);
    expect(workbookSheets(drive.files(STMT_ID)).map((x) => x.name)).toEqual(sheetsBefore);
    expect(noAppMark("드림합정26.11")).toBe(true);
  });

  it("다른 달로 옮기면 그 달 기존 시트에 쓰고 옛 달에서는 지운다. 시트가 없으면 만들지 않고 주황 안내로 대기, 시트를 만들면 채운다", async () => {
    const today = [rec("S28", PASS), rec("R28", PASS), rec("S29", PASS)];
    seed([...realListRecords(), ...today]);
    const s = await startApp();
    expect(stmtRows("드림합정26.11")).toContain("2026-11-12|350A|3");
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((x) => x.name);

    // S29 → 10/30: 기존 '드림합정26.10' 시트로 옮긴다.
    await s.changeRecordDate(today[2].id, "2026-10-30");
    await s.settled();
    await s.syncIdle();
    expect(s.getState().records.find((r) => r.id === today[2].id)?.sync).toBe("synced");
    expect(stmtRows("드림합정26.11")).toContain("2026-11-12|350A|2");
    expect(stmtRows("드림합정26.10")).toContain("2026-10-30|350A|1");
    expect(iljiDayRow("2026-10-30")).toMatchObject({ J: "350A", K: "1" });

    // S28 → 12/01: 12월 시트가 없다. S→R 자동으로 같은 번호 R28 도 함께 옮긴다.
    await s.changeRecordDate(today[0].id, "2026-12-01");
    await s.settled();
    await s.syncIdle();
    const st = s.getState();
    const s28 = st.records.find((r) => r.id === today[0].id)!;
    const r28 = st.records.find((r) => r.id === today[1].id)!;
    expect([s28.date, r28.date]).toEqual(["2026-12-01", "2026-12-01"]);
    expect([s28.sync, r28.sync]).toEqual(["pending", "pending"]);
    expect(s28.syncError).toBeUndefined();
    expect(s28.steps?.statement?.status).toBe("wait");
    expect(s28.steps?.statement?.message).toBe("12월 거래명세서 시트를 만들어주세요 (드림합정26.12)");
    expect(st.syncReport?.files.statement.status).not.toBe("error");
    expect(st.syncReport?.files.statement.message).toContain("12월 거래명세서 시트를 만들어주세요 (드림합정26.12)");
    // 옛 달(11월)에서는 이미 빠졌고, 시트는 새로 만들지 않았다.
    expect(stmtRows("드림합정26.11").some((r) => r.startsWith("2026-11-12|"))).toBe(false);
    expect(workbookSheets(drive.files(STMT_ID)).map((x) => x.name)).toEqual(sheetsBefore);
    // 그날 줄(검사자)은 두고 개수만 비운다.
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "", K: "" });
    // 리스트: 행 수 그대로, 정렬 유지.
    const rows = lineRows();
    expect(rows).toHaveLength(10);
    expect(rows.map((r) => r.key)).toEqual(sortedKeys(rows));
    expect(rows.slice(-2).map((r) => r.key)).toEqual(["2026-12-01|S28", "2026-12-01|R28"]);

    // 사용자가 12월 시트를 만들고 다시 켜면 채운다(중복 없음).
    userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 12);
    const again = await startApp();
    const done = again.getState().records.filter((r) => r.id === today[0].id || r.id === today[1].id);
    expect(done.map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(done.every((r) => !r.movedFrom)).toBe(true);
    expect(stmtRows("드림합정26.12")).toEqual(["2026-12-01|350A|2"]);
    expect(noAppMark("드림합정26.12")).toBe(true);
    expect(lineRows()).toHaveLength(10);
  });
});

describe("아이폰: '로그인이 만료되었습니다(1건)' + '보내는 중…'에서 멈춤", () => {
  const LOGIN_EXPIRED = "Microsoft 로그인이 만료되었습니다";
  const today = () => [rec("S28", PASS), rec("R28", PASS)];
  const mineOf = (s: Session, ids: string[]) => s.getState().records.filter((r) => ids.includes(r.id));
  const noDupList = () => {
    const ids = listIds();
    expect(new Set(ids).size).toBe(ids.length);
  };

  it("로그인과 상관없는 Graph 실패(412·409·423·400)는 '만료' 문구가 아니라 실제 이유를 보여 주고 로그인 배너를 띄우지 않는다", async () => {
    const recs = today();
    seed([...realListRecords(), ...recs]);
    drive.fail((m, p) => m === "PATCH" && p.includes(LIST_ID), 400, Infinity);
    const s = await startApp();
    const lines = model(s).lines.map((l) => l.text).join("\n");
    expect(lines).toContain("검사리스트 · 실패 — OneDrive 오류(400)");
    expect(lines).not.toContain(LOGIN_EXPIRED);
    expect(head(s)).toMatchObject({ tone: "bad" });
    expect(head(s).text).toMatch(/실패/);
    expect(s.getState().loginNeeded).toBe(false);
    expect(s.getState().syncing).toBe(false);
    expect(model(s).canRetry).toBe(true);

    const { explainGraphStatus } = await import("../sync/graph");
    for (const status of [0, 400, 403, 404, 409, 412, 423, 429, 500, 503]) expect(explainGraphStatus(status, "x")).not.toContain("만료");
    expect(explainGraphStatus(401)).toContain("만료");
  });

  it("캐시된 옛 토큰으로 401 → 새 토큰(force)으로 다시 보내 성공한다", async () => {
    const recs = today();
    seed([...realListRecords(), ...recs]);
    auth.stale = true;
    const s = await startApp();
    expect(mineOf(s, recs.map((r) => r.id)).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(tokenCalls.some((c) => c.forceRefresh)).toBe(true);
    expect(s.getState().loginNeeded).toBe(false);
    noDupList();
  });

  it("새 토큰으로도 401 이면 실패가 아니라 대기 + 로그인 배너. 로그인하고 돌아오면 첫 토큰부터 새로 받아(force) 그대로 보낸다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.forceBroken = true;
    const s = await startApp();
    const st = s.getState();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(mineOf(s, ids).some((r) => SYNC_ERR(r))).toBe(false);
    expect(st.loginNeeded).toBe(true);
    expect(st.syncing).toBe(false);
    expect(model(s).lines.map((l) => l.text).join("\n")).not.toContain(LOGIN_EXPIRED);
    expect(head(s)).toEqual({ tone: "login", text: "로그인 필요 · 대기 2건" });

    // 로그인하고 돌아옴(redirect). 캐시에는 여전히 옛 토큰이 있어도 첫 동기화는 force 로 새 토큰.
    auth.forceBroken = false;
    auth.stale = true;
    auth.fromRedirect = true;
    tokenCalls.length = 0;
    const back = await startApp();
    expect(mineOf(back, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(back.getState().loginNeeded).toBe(false);
    // 사전 확인 다음, 동기화의 첫 토큰이 force
    expect(tokenCalls[1]?.forceRefresh).toBe(true);
    expect(head(back)).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
    noDupList();
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
  });

  it("옛 판이 남긴 '검사리스트 · 실패 — 로그인이 만료' 기록(작업일지·명세서 대기)도 새로 켜면 대기로 바꿔 새 토큰으로 보낸다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    const at = Date.now();
    const old = "Microsoft 로그인이 만료되었습니다. 설정에서 다시 로그인하세요.";
    recs[0] = { ...recs[0], sync: "error", syncError: `검사리스트: ${old}`, steps: { list: { status: "error", message: old, at } } };
    recs[1] = { ...recs[1], sync: "syncing", steps: { list: { status: "done", at } } };
    seed([...realListRecords(), ...recs]);
    mem.set("syncReport", { at, files: { list: { status: "error", message: `검사리스트: ${old}` }, ilji: { status: "idle" }, statement: { status: "idle" } } } satisfies SyncReport);
    auth.stale = true;
    const s = await startApp();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(tokenCalls.some((c) => c.forceRefresh)).toBe(true);
    expect(model(s).visible).toBe(false);
    noDupList();
  });

  it("토큰 받기가 멈추면(응답 없음) 시간 제한 뒤 '보내는 중'이 풀리고 기록은 대기, 다시 시도 버튼을 누를 수 있다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.mode = "hang";
    const s = await startApp((sess) => {
      sess.syncLimits.loginCheckMs = 30;
    });
    await s.syncIdle();
    expect(s.getState().syncing).toBe(false);
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(model(s).canRetry).toBe(true);
    expect(drive.requests(() => true)).toHaveLength(0);
    // 멈춘 대기열은 '동기화 중'이 아니라 '대기'(주황)
    expect(head(s)).toEqual({ tone: "wait", text: "대기 2건" });

    auth.mode = "ok";
    s.retrySync();
    await s.syncIdle();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(head(s)).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
    noDupList();
  });

  it("동기화 도중 토큰이 멈추면(사전 확인은 통과) 토큰 시간 제한으로 끝나 기록은 대기(실패 아님), 다시 시도로 보낸다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.mode = "hangAfter";
    auth.okCalls = 1; // 사전 확인만 통과
    const s = await startApp((_sess, graph) => {
      graph.graphTimeouts.tokenMs = 30;
    });
    await s.syncIdle();
    expect(s.getState().syncing).toBe(false);
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(model(s).lines.map((l) => l.text).join("\n")).not.toContain("실패");
    expect(model(s).canRetry).toBe(true);

    auth.mode = "ok";
    s.retrySync();
    await s.syncIdle();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    noDupList();
  });

  it("Graph 요청(리스트 칸 PATCH)이 응답 없이 멈추면 요청 시간 제한으로 끊고 대기로 남긴다. 다시 시도하면 중복 없이 채운다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    drive.hangs.push((m, p) => m === "PATCH" && p.includes(LIST_ID));
    const s = await startApp((_sess, graph) => {
      graph.graphTimeouts.requestMs = 30;
    });
    await s.syncIdle();
    expect(s.getState().syncing).toBe(false);
    expect(mineOf(s, ids).every((r) => r.sync === "pending")).toBe(true);
    expect(model(s).canRetry).toBe(true);
    expect(model(s).lines.map((l) => l.text).join("\n")).not.toContain(LOGIN_EXPIRED);

    drive.hangs = [];
    s.retrySync();
    await s.syncIdle();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    noDupList();
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
  });

  it("요청·토큰 시간 제한이 없어도(무한 대기) 묶음 시간 제한이 요청을 끊어 '보내는 중'을 풀고 대기로 남긴다", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    drive.hangs.push((_m, p) => p.includes("/workbook/createSession"));
    const s = await startApp((sess, graph) => {
      graph.graphTimeouts.requestMs = 0;
      sess.syncLimits.runMs = 50;
      sess.syncLimits.graceMs = 20;
    });
    await s.syncIdle();
    expect(s.getState().syncing).toBe(false);
    expect(mineOf(s, ids).every((r) => r.sync === "pending")).toBe(true);
    expect(model(s).canRetry).toBe(true);

    drive.hangs = [];
    s.retrySync();
    await s.syncIdle();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    noDupList();
  });

  it("상태 줄: 실제로 보내는 동안만 '동기화 중'. 지운 현장(보내지 않는 기록)·미저장·없음은 대기로 세지 않아 초록이 된다", async () => {
    const recs = today();
    const removed = { ...rec("S05", PASS), site: "지성 브레인", sheetName: " 지성 브레인", sync: "pending" as const };
    const unsaved = { ...rec("S06", PASS), sync: "local" as const };
    seed([...realListRecords(), ...recs, removed, unsaved]);
    const seen: string[] = [];
    const s = await startApp((sess) => {
      sess.subscribe(() => {
        const st = sess.getState();
        if (st.syncing) seen.push(headerStatus({ mode: "graph", online: true, syncing: true, loginNeeded: st.loginNeeded, account: st.account, queue: sess.syncQueue(st.records) }).text);
      });
    });
    expect(seen.length).toBeGreaterThan(0);
    expect(new Set(seen)).toEqual(new Set(["동기화 중…"]));
    expect(s.getState().syncing).toBe(false);
    expect(s.getState().records.find((r) => r.id === removed.id)?.sync).toBe("pending"); // 보내지 않음(그대로)
    expect(head(s)).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
    expect(headerStatus({ mode: "graph", online: false, syncing: false, loginNeeded: false, account: "tester", queue: recs }).tone).toBe("off");
    expect(headerStatus({ mode: "demo", online: true, syncing: false, loginNeeded: true, account: null, queue: recs })).toEqual({ tone: "on", text: "이 기기에 저장됨" });
    expect(headerStatus({ mode: "graph", online: true, syncing: false, loginNeeded: false, account: null, queue: [] }).tone).toBe("login");
  });
});

describe("로그인 만료(전역 인증 상태): 계정 이름이 보여도 토큰이 죽으면 '로그인 필요'", () => {
  const today = () => [rec("S28", PASS), rec("R28", PASS)];
  const mineOf = (s: Session, ids: string[]) => s.getState().records.filter((r) => ids.includes(r.id));
  const settings = (s: Session) => {
    const st = s.getState();
    return settingsAuthModel({ account: st.account, auth: st.auth, loginNeeded: st.loginNeeded });
  };
  const noDupList = () => {
    const ids = listIds();
    expect(new Set(ids).size).toBe(ids.length);
  };

  it("캐시 계정 + InteractionRequired → expired·로그인 필요·보내는 중 해제, 패널은 대기(실패 아님), 설정은 다시 로그인, 리스트 탭은 로그인 안내. 로그인 뒤 대기열 자동 전송", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.mode = "expired";
    const s = await startApp();
    const st = s.getState();
    expect(st.auth).toBe("expired");
    expect(st.account).toBe("tester"); // 이메일은 보인다
    expect(st.loginNeeded).toBe(true);
    expect(st.syncing).toBe(false);
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["pending", "pending"]);
    expect(drive.requests(() => true)).toHaveLength(0);
    const m = model(s);
    expect(m.visible).toBe(false); // 빨간 실패 패널 없음 — 화면 위 로그인 배너가 안내
    expect(m.tone).toBe("wait");
    expect(m.title).toContain("로그인 필요");
    expect(m.title).not.toContain("실패");
    expect(m.login).toBe(true);
    expect(m.lines.map((l) => l.text).join("\n")).not.toContain("실패");
    expect(m.lines.every((l) => l.state !== "error")).toBe(true);
    expect(head(s)).toEqual({ tone: "login", text: "로그인 필요 · 대기 2건" });
    expect(settings(s)).toMatchObject({ email: "tester", showRelogin: true });

    // 만료 상태에서 다시 시도해도 '보내는 중'으로 바뀌지 않는다
    const seen: boolean[] = [];
    const off = s.subscribe(() => seen.push(s.getState().syncing));
    s.retrySync();
    await s.syncIdle();
    off();
    expect(seen.every((v) => v === false)).toBe(true);
    expect(drive.requests(() => true)).toHaveLength(0);

    // 리스트 탭: '읽는 중…'에 남지 않고 로그인 안내
    const { readLiveList, LIST_LOGIN_TEXT } = await import("../ui/ListScreen");
    const list = await readLiveList("드림 합정주공", { timeoutMs: 2_000 });
    expect(list).toEqual({ ok: false, message: LIST_LOGIN_TEXT });

    // 다시 로그인(redirect) → 돌아오면 새 토큰으로 대기열 자동 전송, 중복 없음
    await s.signIn();
    expect(auth.logins).toEqual([{}]);
    auth.mode = "ok";
    auth.fromRedirect = true;
    tokenCalls.length = 0;
    const back = await startApp();
    expect(back.getState().auth).toBe("ok");
    expect(back.getState().loginNeeded).toBe(false);
    expect(mineOf(back, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    expect(tokenCalls.some((c) => c.forceRefresh)).toBe(true);
    expect(head(back)).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
    expect(settings(back).showRelogin).toBe(false);
    noDupList();
    expect(iljiDayRow("2026-11-12")).toMatchObject({ J: "350A", K: "2" });
  });

  it("캐시 계정 + 새 토큰으로도 Graph 401 → 묶음이 끝나기 전에 즉시 expired: 그 순간 '보내는 중' 해제·로그인 필요, 실패 아님", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.forceBroken = true;
    const atExpiry: { syncing: boolean; loginNeeded: boolean; runActive: boolean }[] = [];
    let runActive = false;
    const s = await startApp((sess) => {
      void import("../sync/authState").then((a) =>
        a.onAuthChange((next) => {
          if (next !== "expired") return;
          queueMicrotask(() => {
            const st = sess.getState();
            atExpiry.push({ syncing: st.syncing, loginNeeded: st.loginNeeded, runActive });
          });
        }),
      );
      sess.subscribe(() => {
        if (sess.getState().syncing) runActive = true;
      });
    });
    expect(atExpiry).toHaveLength(1);
    expect(atExpiry[0]).toMatchObject({ syncing: false, loginNeeded: true, runActive: true });
    const st = s.getState();
    expect(st.auth).toBe("expired");
    expect(st.syncing).toBe(false);
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["pending", "pending"]);
    const m = model(s);
    expect(m.title).not.toContain("실패");
    expect(m.tone).toBe("wait");
    expect(m.lines.every((l) => l.state !== "error")).toBe(true);
    expect(head(s).tone).toBe("login");
    expect(settings(s).showRelogin).toBe(true);
    noDupList();
  });

  it("네트워크 오류(BrowserAuthError network_error)는 만료가 아니다: '로그인 만료'·로그인 필요 없음, 기록은 대기", async () => {
    const recs = today();
    const ids = recs.map((r) => r.id);
    seed([...realListRecords(), ...recs]);
    auth.mode = "network";
    const s = await startApp();
    const st = s.getState();
    expect(st.auth).not.toBe("expired");
    expect(st.loginNeeded).toBe(false);
    expect(st.syncing).toBe(false);
    expect(mineOf(s, ids).every((r) => r.sync === "pending" || r.sync === "error")).toBe(true);
    const text = [model(s).title, ...model(s).lines.map((l) => l.text), head(s).text].join("\n");
    expect(text).not.toContain("만료");
    expect(head(s).tone).not.toBe("login");
    auth.mode = "ok";
    s.retrySync();
    await s.syncIdle();
    expect(mineOf(s, ids).map((r) => r.sync)).toEqual(["synced", "synced"]);
    noDupList();
  });

  it("저장된 '로그인 만료' 실패 단계는 켤 때 대기로 바뀌고, 패널·상태 줄은 빨간 실패가 아니다", async () => {
    const recs = today();
    const at = Date.now();
    const old = "Microsoft 로그인이 만료되었습니다 — 화면 위 '로그인'을 누르면 대기 중인 기록을 바로 보냅니다.";
    recs[0] = { ...recs[0], sync: "error", syncError: `검사리스트: ${old}`, steps: { list: { status: "error", message: old, at } } };
    seed([...realListRecords(), ...recs]);
    auth.mode = "lost";
    const s = await startApp();
    expect(s.getState().records.find((r) => r.id === recs[0].id)?.sync).toBe("pending");
    expect(model(s).tone).not.toBe("bad");
    expect(head(s).tone).toBe("login");
    // 패널 자체도(정리 전 상태라도) 로그인 문구 단계를 실패로 세지 않는다
    const raw = syncPanelModel([{ ...recs[0], result: PASS }], null, { online: true, syncing: false, mode: "graph", loginNeeded: true });
    expect(raw.tone).toBe("wait");
    expect(raw.title).not.toContain("실패");
    expect(headerStatus({ mode: "graph", online: true, syncing: false, loginNeeded: true, account: "tester", queue: [recs[0]] }).tone).toBe("login");
  });

  it("상태 줄 색: 보내는 중(노랑)·대기(주황)·로그인 필요(파랑, auth 만료면 계정이 있어도)·실패(빨강)·동기화됨(초록)", () => {
    const q = today().map((r) => ({ ...r, sync: "pending" as const }));
    const base = { mode: "graph" as const, online: true, syncing: false, loginNeeded: false, account: "tester" };
    expect(headerStatus({ ...base, syncing: true, queue: q }).tone).toBe("send");
    expect(headerStatus({ ...base, queue: q }).tone).toBe("wait");
    expect(headerStatus({ ...base, queue: q, auth: "expired" })).toEqual({ tone: "login", text: "로그인 필요 · 대기 2건" });
    expect(headerStatus({ ...base, queue: q, auth: "signedOut" }).tone).toBe("login");
    expect(headerStatus({ ...base, queue: [{ ...q[0], sync: "error", syncError: "작업일지: 잠겨 있습니다" }] }).tone).toBe("bad");
    expect(headerStatus({ ...base, queue: [], auth: "ok" })).toEqual({ tone: "on", text: "OneDrive 동기화됨" });
    expect(settingsAuthModel({ account: "a@b.c", auth: "unknown", loginNeeded: false }).showRelogin).toBe(true);
    expect(settingsAuthModel({ account: "a@b.c", auth: "ok", loginNeeded: false }).showRelogin).toBe(false);
    expect(settingsAuthModel({ account: null, auth: "signedOut", loginNeeded: true })).toMatchObject({ showLogin: true });
  });
});

function SYNC_ERR(r: InspectionRecord): boolean {
  return r.sync === "error" || Object.values(r.steps ?? {}).some((st) => st?.status === "error");
}
