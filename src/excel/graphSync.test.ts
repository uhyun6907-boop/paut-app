import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../domain/dates";
import type { InspectionRecord } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { aggregateSteps } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import { readIljiDay } from "./iljiBook";
import { ILJI_ID, LIST_ID, STMT_ID, graphConfig, meta, realDrive, realListRecords, rec } from "../test/realBooks";
import { restoreStatementOwnership, syncGraph, type GraphSyncArgs } from "./engine";
import { findIljiSheet, readIljiCell } from "./iljiBook";
import { parseListWorkbook } from "./listBook";
import { readSheet, workbookSheets } from "./ooxml";
import { readPautRows, readStatement } from "./statementBook";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});

beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive, onToken?: (force: boolean) => void): GraphClient {
  let cached = drive.validToken;
  return new GraphClient({
    fetchImpl: drive.fetch,
    sleep: async () => undefined,
    getToken: async (force) => {
      drive.events.push(`token:${force ? "force" : "silent"}`);
      onToken?.(force);
      if (force) cached = drive.validToken;
      return cached;
    },
  });
}

async function sync(drive: FakeDrive, partial: Partial<GraphSyncArgs> & { jobs: InspectionRecord[]; all: InspectionRecord[] }) {
  return syncGraph({ client: client(drive), config: graphConfig(), meta: meta(), sleep: async () => undefined, ...partial });
}

const SHEET = " PAUT FIELD";

/** 작업일지에서 그 날짜(B열)가 적힌 행들 */
function iljiRowsOn(drive: FakeDrive, iso: string): number[] {
  const files = drive.files(ILJI_ID);
  const { cells } = readSheet(files, findIljiSheet(files));
  const serial = String(isoToExcelSerial(iso));
  const rows: number[] = [];
  for (const c of cells.values()) if (c.col === "B" && c.row >= 6 && c.v === serial) rows.push(c.row);
  return rows.sort((a, b) => a - b);
}

function cell(drive: FakeDrive, row: number, col: string): string {
  return readIljiCell(drive.files(ILJI_ID), row, col);
}

function rawCell(drive: FakeDrive, ref: string): string {
  const xml = readSheet(drive.files(ILJI_ID), SHEET).xml;
  return xml.match(new RegExp(`<c r="${ref}"(?:[^>]*/>|[^>]*>[\\s\\S]*?</c>)`))?.[0] ?? "";
}

function listAppRows(drive: FakeDrive): InspectionRecord[] {
  return parseListWorkbook(drive.files(LIST_ID)).flatMap((s) => s.records).filter((r) => r.src === "user");
}

describe("실제 작업일지: 드림 합정주공 열", () => {
  it("관경/개소는 J/K, 검사자는 P/Q, 근무시간은 S/T. 다른 열은 비어 있고 파일 전체 PUT 은 없다", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS), rec("R28", PASS), rec("S29", FAIL)];
    const all = [...realListRecords(), ...jobs];
    const out = await sync(drive, { jobs, all });
    expect(out.report.files).toMatchObject({ list: { status: "ok", count: 3 }, ilji: { status: "ok" }, statement: { status: "ok" } });

    const rows = iljiRowsOn(drive, "2026-11-12");
    // 마지막 기록(521행, 11/05) 바로 아래. 220행처럼 순번·검사자만 있는 중간 행은 쓰지 않는다.
    expect(rows).toEqual([522]);
    expect(cell(drive, 220, "P")).toBe("이연행");
    expect(cell(drive, 220, "B")).toBe("");
    expect(cell(drive, 522, "J")).toBe("350A");
    expect(cell(drive, 522, "K")).toBe("2"); // 불합격 S29 제외
    expect(cell(drive, 522, "P")).toBe("조우현");
    expect(cell(drive, 522, "Q")).toBe("심명보");
    expect(cell(drive, 522, "S")).toBe("09:00");
    expect(cell(drive, 522, "T")).toBe("15:30");
    for (const col of ["D", "E", "F", "G", "H", "I", "L", "M", "N", "O"]) expect(cell(drive, 522, col)).toBe("");
    // 세 파일 모두 셀 쓰기만. 파일 통째 업로드 없음(명세서도 시트를 만들지 않으므로 업로드 없음).
    expect(drive.requests((e) => e.item === ILJI_ID && e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => e.item === STMT_ID && e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => e.item === ILJI_ID && e.method === "PATCH").every((e) => !!e.session)).toBe(true);
  });

  it("근무시간을 비우면 S/T 는 값이 없는 빈칸이다(00:00 아님)", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS)];
    await syncGraph({
      client: client(drive),
      jobs,
      all: [...realListRecords(), ...jobs],
      config: graphConfig(),
      meta: meta({ timeStart: "", timeEnd: "" }),
    });
    expect(iljiRowsOn(drive, "2026-11-12")).toEqual([522]);
    expect(cell(drive, 522, "K")).toBe("1");
    for (const col of ["S", "T"]) expect(rawCell(drive, `${col}522`)).not.toContain("<v>");
  });
});

describe("백필: 리스트엔 있는데 작업일지·명세서에 없는 11/11 합정주공 7개소", () => {
  it("작업일지 J/K 350A×7, 검사자 P/Q, 근무시간 모르면 빈칸 + 사용자가 만든 드림합정26.11 시트에 채운다(비고에 '앱' 없음). 다시 돌려도 늘지 않는다", async () => {
    const drive = realDrive();
    const all = realListRecords();
    const app = all.filter((r) => r.src === "user");
    expect(app.map((r) => r.point)).toEqual(["S22", "S23", "S24", "S25", "S26", "R26", "S27"]);
    expect(new Set(app.map((r) => r.date))).toEqual(new Set(["2026-11-11"]));
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((s) => s.name);

    const m = meta({ insp2ByDate: { "2026-11-11": "김정환" } });
    const out = await syncGraph({ client: client(drive), jobs: [], all, config: graphConfig(), meta: m, backfill: true, sleep: async () => undefined });
    expect(out.report.files.ilji.status).toBe("ok");
    expect(out.report.files.statement.status).toBe("ok");
    for (const r of app) expect(out.steps.get(r.id)).toMatchObject({ ilji: { status: "done" }, statement: { status: "done" } });

    const rows = iljiRowsOn(drive, "2026-11-11");
    expect(rows).toEqual([522]);
    expect(cell(drive, 522, "J")).toBe("350A");
    expect(cell(drive, 522, "K")).toBe("7");
    expect(cell(drive, 522, "P")).toBe("조우현");
    expect(cell(drive, 522, "Q")).toBe("김정환");
    expect(cell(drive, 522, "C")).toBe("평일");
    for (const col of ["S", "T"]) expect(rawCell(drive, `${col}522`)).not.toContain("<v>");

    const stmt = drive.files(STMT_ID);
    // 시트는 하나도 늘지 않는다(앱은 시트를 만들지 않음).
    expect(workbookSheets(stmt).map((s) => s.name)).toEqual(sheetsBefore);
    const view = readStatement(stmt, "드림합정26.11");
    expect(view.paut).toEqual([{ date: "2026-11-11", item: "350A(1.2M)", count: 7, unitPrice: 75000 }]);
    const { cells } = readSheet(stmt, "드림합정26.11");
    expect(cells.get("I12")?.v ?? "").toBe(""); // 비고에 '앱' 같은 표시를 쓰지 않는다
    expect(cells.get("F12")?.f).toBe("D12*1.2");
    expect(cells.get("H12")?.f).toBe("G12*F12");
    expect(cells.get("F29")?.f).toBe("SUM(F12:F28)"); // 합계·부가세 바닥은 그대로
    expect(cells.get("H30")?.f).toBe("ROUNDDOWN(B30+D30,-3)");
    // 10월 시트는 그대로
    expect(readStatement(stmt, "드림합정26.10").paut.map((p) => p.count)).toEqual([8, 10, 10]);
    // 파일 통째 업로드 없음, 명세서에 I열 쓰기 없음
    expect(drive.requests((e) => e.item === STMT_ID && e.method === "PUT")).toHaveLength(0);
    const stmtPatches = drive.requests((e) => e.item === STMT_ID && e.method === "PATCH");
    expect(stmtPatches.length).toBeGreaterThan(0);
    expect(JSON.stringify(stmtPatches.map((e) => e.body ?? null))).not.toContain("앱");

    // 재시도(같은 기록, 상태 초기화): 행·시트·줄이 늘지 않는다.
    for (const r of all) r.steps = undefined;
    await syncGraph({ client: client(drive), jobs: [], all, config: graphConfig(), meta: m, backfill: true, sleep: async () => undefined });
    expect(iljiRowsOn(drive, "2026-11-11")).toEqual([522]);
    expect(cell(drive, 522, "K")).toBe("7");
    expect(workbookSheets(drive.files(STMT_ID)).map((s) => s.name)).toEqual(sheetsBefore);
    expect(readPautRows(drive.files(STMT_ID), "드림합정26.11")).toHaveLength(1);
    expect(drive.requests((e) => e.item === STMT_ID && e.method === "PUT")).toHaveLength(0);

    // 같은 기록을 대기열로 다시 보내도(이미 리스트에 Y 가 있음) 리스트 행은 그대로 7개.
    const again = app.map((r) => ({ ...r, steps: undefined, sync: "pending" as const }));
    await syncGraph({ client: client(drive), jobs: again, all, config: graphConfig(), meta: m, sleep: async () => undefined });
    expect(listAppRows(drive).map((r) => r.excelRow)).toEqual([433, 434, 435, 436, 437, 438, 439]);
    expect(iljiRowsOn(drive, "2026-11-11")).toEqual([522]);
    // 표시가 없어도 이 기기가 쓴 줄(날짜·관경·개수)로 알아보고 같은 행을 다시 쓴다.
    expect(readPautRows(drive.files(STMT_ID), "드림합정26.11")).toEqual([{ row: 12, date: "2026-11-11", diameter: "350A", count: 7, app: false }]);
    expect(readSheet(drive.files(STMT_ID), "드림합정26.11").cells.get("I12")?.v ?? "").toBe("");
  });

  it("이 기기에 기록이 없어도 리스트의 Y열 앱 ID 행을 찾아 채운다(scanList)", async () => {
    const drive = realDrive();
    const out = await syncGraph({ client: client(drive), jobs: [], all: [], config: graphConfig(), meta: meta(), backfill: true, scanList: true, sleep: async () => undefined });
    expect(out.discovered.map((r) => r.point)).toEqual(["S22", "S23", "S24", "S25", "S26", "R26", "S27"]);
    expect(iljiRowsOn(drive, "2026-11-11")).toEqual([522]);
    expect(cell(drive, 522, "K")).toBe("7");
    expect(readStatement(drive.files(STMT_ID), "드림합정26.11").paut[0]).toMatchObject({ count: 7 });
  });

  it("작업일지에 그날 그 열이 이미 적혀 있으면(사람이 적음) 백필은 덮어쓰지 않는다", async () => {
    const drive = realDrive();
    const all = realListRecords();
    await syncGraph({ client: client(drive), jobs: [], all, config: graphConfig(), meta: meta(), backfill: true, sleep: async () => undefined });
    // PC 에서 개수를 6 으로 고쳤다고 하자
    const s = await client(drive).openWorkbook(ILJI_ID);
    const ws = (await s.sheet(SHEET))!;
    await s.writeFormulas(ws.id, "K522", [[6]]);
    for (const r of all) r.steps = undefined;
    await syncGraph({ client: client(drive), jobs: [], all, config: graphConfig(), meta: meta(), backfill: true, sleep: async () => undefined });
    expect(cell(drive, 522, "K")).toBe("6");
  });
});

describe("부분 실패와 다시 시도", () => {
  it("리스트 성공 + 작업일지·명세서 실패 → 파일별로 실패가 남고, 다시 시도하면 리스트 Y 중복 없이 끝난다", async () => {
    const drive = realDrive();
    const ilji = drive.fail((_m, p) => p.includes(ILJI_ID) && p.endsWith("createSession"), 423, Infinity);
    const stmt = drive.fail((m, p) => p.includes(STMT_ID) && p.endsWith("/content") && m === "GET", 503, Infinity);
    const jobs = [rec("S28", PASS), rec("R28", PASS), rec("S29", PASS)];
    const all = [...realListRecords(), ...jobs];
    const first = await sync(drive, { jobs, all });

    expect(first.report.files.list.status).toBe("ok");
    expect(first.report.files.ilji).toMatchObject({ status: "error" });
    expect(first.report.files.ilji.message).toContain("작업일지");
    expect(first.report.files.ilji.message).toContain("잠겨");
    expect(first.report.files.statement.message).toContain("거래명세서");
    for (const r of jobs) {
      expect(r.steps?.list?.status).toBe("done");
      expect(r.steps?.ilji?.status).toBe("error");
      expect(r.steps?.statement?.status).toBe("error");
      const agg = aggregateSteps(r.steps);
      expect(agg.status).toBe("error");
      expect(agg.error).toMatch(/^작업일지: .*잠겨.* · 거래명세서: .*서버 오류/);
      r.sync = agg.status;
    }
    const model = syncPanelModel(all, first.report, { online: true, syncing: false, mode: "graph" });
    expect(model.visible).toBe(true);
    expect(model.canRetry).toBe(true);
    expect(model.lines.map((l) => l.state)).toEqual(["ok", "error", "error"]);
    expect(model.lines[1].text).toMatch(/^작업일지 · 실패 — PC 엑셀.*잠겨.*\(3건\)$/);
    expect(model.lines[2].text).toMatch(/^거래명세서 · 실패 — OneDrive 서버 오류\(503\)/);
    // 재시도 횟수: 잠김은 4번까지 보내고 멈춘다.
    expect(drive.requests((e) => e.item === ILJI_ID && e.path.endsWith("createSession"))).toHaveLength(graphRetry.attempts);
    expect(listAppRows(drive)).toHaveLength(10);

    // 다시 시도: 실패한 파일만. 리스트는 다시 쓰지 않는다.
    ilji.times = 0;
    stmt.times = 0;
    for (const r of jobs) r.steps = { list: r.steps!.list };
    const listPatchesBefore = drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length;
    const second = await sync(drive, { jobs, all });
    expect(second.report.files).toMatchObject({ list: { status: "idle" }, ilji: { status: "ok" }, statement: { status: "ok" } });
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length).toBe(listPatchesBefore);
    for (const r of jobs) expect(aggregateSteps(r.steps).status).toBe("synced");
    const ids = listAppRows(drive).map((r) => r.id);
    expect(ids).toHaveLength(10);
    expect(new Set(ids).size).toBe(10);
    expect(cell(drive, iljiRowsOn(drive, "2026-11-12")[0], "K")).toBe("3");
  });

  it("리스트까지 처음부터 다시 보내도 Y(앱 ID)가 같은 행을 고치므로 행이 늘지 않는다", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS), rec("R28", FAIL)];
    const all = [...realListRecords(), ...jobs];
    await sync(drive, { jobs, all });
    const rows = listAppRows(drive).filter((r) => jobs.some((j) => j.id === r.id)).map((r) => r.excelRow);
    jobs[1].result = PASS;
    for (const r of jobs) r.steps = undefined;
    await sync(drive, { jobs, all });
    const after = listAppRows(drive).filter((r) => jobs.some((j) => j.id === r.id));
    expect(after.map((r) => r.excelRow)).toEqual(rows);
    expect(after.map((r) => r.result)).toEqual([PASS, PASS]);
    expect(cell(drive, iljiRowsOn(drive, "2026-11-12")[0], "K")).toBe("2");
  });
});

describe("토큰", () => {
  it("파일 단계마다 토큰을 새로 받고, 401 이면 강제로 새로 받아 다시 보낸다", async () => {
    const drive = realDrive();
    const forced: boolean[] = [];
    // 작업일지 단계 첫 요청에서 토큰이 만료된다.
    let rotated = false;
    const orig = drive.fetch;
    drive.fetch = async (input, init) => {
      const url = String(input);
      if (!rotated && url.includes(ILJI_ID)) {
        rotated = true;
        drive.rotateToken();
      }
      return orig(input, init);
    };
    const c = client(drive, (force) => forced.push(force));
    const jobs = [rec("S28", PASS)];
    const out = await syncGraph({ client: c, jobs, all: [...realListRecords(), ...jobs], config: graphConfig(), meta: meta(), sleep: async () => undefined });
    expect(out.report.files).toMatchObject({ list: { status: "ok" }, ilji: { status: "ok" }, statement: { status: "ok" } });
    expect(forced.filter(Boolean)).toHaveLength(1);
    expect(drive.log.filter((e) => e.status === 401)).toHaveLength(1);

    // 각 파일의 첫 요청 앞에는 토큰 받기가 있다.
    const ev = drive.events;
    for (const item of [LIST_ID, ILJI_ID, STMT_ID]) {
      const first = ev.findIndex((e) => e.startsWith(`req:${item}`));
      const prevOther = ev.slice(0, first).map((e, i) => ({ e, i })).filter((x) => x.e.startsWith("req:") && !x.e.startsWith(`req:${item}`)).pop()?.i ?? -1;
      expect(ev.slice(prevOther + 1, first).some((e) => e.startsWith("token:"))).toBe(true);
    }
  });

  it("로그인이 필요하면 그 단계만 실패로 남기고 다음 파일도 시도한다(조용히 넘어가지 않음)", async () => {
    const drive = realDrive();
    let calls = 0;
    const c = new GraphClient({
      fetchImpl: drive.fetch,
      sleep: async () => undefined,
      getToken: async () => {
        calls += 1;
        if (calls === 1) return drive.validToken;
        throw new Error("Microsoft 로그인이 만료되었습니다. 설정에서 다시 로그인하세요.");
      },
    });
    const jobs = [rec("S28", PASS)];
    const out = await syncGraph({ client: c, jobs, all: [...realListRecords(), ...jobs], config: graphConfig(), meta: meta(), sleep: async () => undefined });
    expect(out.report.files.list.status).toBe("ok");
    expect(out.report.files.ilji.message).toBe("작업일지: Microsoft 로그인이 만료되었습니다. 설정에서 다시 로그인하세요.");
    expect(out.report.files.statement.status).toBe("error");
  });
});

describe("개소 규칙", () => {
  it("불합격은 작업일지·명세서에서 빠지고, 재검사는 같은 번호의 새 행이다", async () => {
    const drive = realDrive();
    const failS30 = rec("S30", FAIL);
    const reS30 = rec("S30", PASS);
    const jobs = [failS30, reS30, rec("R30", PASS), rec("S31", FAIL)];
    const all = [...realListRecords(), ...jobs];
    await sync(drive, { jobs, all });
    const rows = listAppRows(drive).filter((r) => r.point === "S30");
    expect(rows.map((r) => r.result)).toEqual([FAIL, PASS]);
    expect(rows[1].excelRow).toBe(rows[0].excelRow! + 1);
    const day = iljiRowsOn(drive, "2026-11-12")[0];
    expect(cell(drive, day, "K")).toBe("2");
    // 11월 시트는 그 달 전체: 리스트에 이미 있는 11/11 7개소 + 오늘 합격 2개소(불합격 S30·S31 제외)
    const nov = readStatement(drive.files(STMT_ID), "드림합정26.11").paut;
    expect(nov).toEqual([
      { date: "2026-11-11", item: "350A(1.2M)", count: 7, unitPrice: 75000 },
      { date: "2026-11-12", item: "350A(1.2M)", count: 2, unitPrice: 75000 },
    ]);
  });

  it("앱에서 지운 택지 드림·지성 브레인·드림 화양·브레인대지 인입은 보내지 않고, 다른 현장 동기화를 막지 않으며 '그 외' 열 기존 값도 그대로 둔다", async () => {
    const drive = realDrive();
    const tj = rec("S01", PASS, { site: "택지 드림", sheetName: "택지 드림", line: "TJ-S-001-150A", diameter: "150A" });
    const brain = rec("S01", PASS, { site: "지성 브레인", sheetName: " 지성 브레인", line: "BR-M-001-200A", diameter: "200A", deleted: true });
    const hy = rec("S01", PASS, { site: "드림 화양", sheetName: "드림 화양", line: "HY-S-021-150A", diameter: "150A" });
    const inip = rec("S01", PASS, { site: "브레인대지 인입", sheetName: " 브레인대지 인입", line: "BI-S-001-150A", diameter: "150A" });
    const oldTj = rec("S02", PASS, { site: "택지 드림", sheetName: "택지 드림", line: "TJ-S-001-150A", diameter: "150A", sync: "synced" });
    const hj = rec("S28", PASS);
    const iljiBefore = readIljiDay(drive.files(ILJI_ID), "2026-11-12", "그 외");
    const jobs = [tj, brain, hy, inip, hj];
    const out = await sync(drive, { jobs, all: [...realListRecords(), ...jobs, oldTj] });
    for (const r of [tj, brain, hy, inip]) expect(out.steps.get(r.id)).toBeUndefined();
    expect(out.steps.get(hj.id)?.list?.status).toBe("done");
    expect(out.report.files.list).toMatchObject({ status: "ok", count: 1 });
    expect(listAppRows(drive).some((r) => [tj.id, brain.id, hy.id, inip.id].includes(r.id))).toBe(false);
    // 지운 현장만 쓰는 '그 외' 열은 건드리지 않는다(지우거나 비우지 않음).
    expect(readIljiDay(drive.files(ILJI_ID), "2026-11-12", "그 외") ?? []).toEqual(iljiBefore ?? []);
  });

  it("리스트 시트 이름 앞 공백이 없어도(' 드림 합정주공') 같은 시트를 찾는다", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS, { sheetName: "드림 합정주공" })];
    const out = await sync(drive, { jobs, all: [...realListRecords(), ...jobs] });
    expect(out.report.files.list.status).toBe("ok");
    expect(listAppRows(drive).find((r) => r.id === jobs[0].id)?.sheetName).toBe(" 드림 합정주공");
  });
});

describe("거래명세서 빈 줄이 모자랄 때", () => {
  it("합계 앞에 행을 끼워 넣고(Graph insert) SUM 범위를 늘린다. 다시 보내도 더 끼우지 않는다", async () => {
    const drive = realDrive();
    const jobs: InspectionRecord[] = [];
    for (let d = 1; d <= 20; d++) jobs.push(rec("S01", PASS, { date: `2026-11-${String(d).padStart(2, "0")}`, line: `HJ-M-1${d}-350A` }));
    const all = [...realListRecords(), ...jobs];
    const out = await sync(drive, { jobs, all });
    expect(out.report.files.statement.status).toBe("ok");
    const inserts = drive.requests((e) => e.item === STMT_ID && e.path.endsWith("/insert"));
    expect(inserts).toHaveLength(1);
    // 빈 자리 17줄(12~28행) < 20줄 → 3행 삽입
    expect(inserts[0].path).toContain("range(address='29:31')");
    const view = readStatement(drive.files(STMT_ID), "드림합정26.11");
    expect(view.paut).toHaveLength(20); // 1~20일, 11일은 리스트의 7개소 + 오늘 1개소
    expect(view.paut.find((p) => p.date === "2026-11-11")?.count).toBe(8);
    expect(view.totalRow).toBe(32);
    const { cells } = readSheet(drive.files(STMT_ID), "드림합정26.11");
    expect(cells.get("F32")?.f).toBe("SUM(F12:F31)");
    expect(cells.get("H32")?.f).toBe("SUM(H12:H31)");
    expect(cells.get("B33")?.f).toBe("H32");
    expect(cells.get("H33")?.f).toBe("ROUNDDOWN(B33+D33,-3)");

    for (const r of jobs) r.steps = undefined;
    await sync(drive, { jobs, all });
    expect(drive.requests((e) => e.item === STMT_ID && e.path.endsWith("/insert"))).toHaveLength(1);
    expect(readStatement(drive.files(STMT_ID), "드림합정26.11").paut).toHaveLength(20);
  });
});

describe("검사리스트 새 행 순서", () => {
  const sitePoints = (drive: FakeDrive, fromRow: number) =>
    parseListWorkbook(drive.files(LIST_ID))
      .find((s) => s.site === "드림 합정주공")!
      .records.filter((r) => r.excelRow! >= fromRow)
      .map((r) => `${r.excelRow} ${r.date} ${r.line} ${r.point} ${r.result}`);

  it("S10·S02·R02·S01 순으로 들어와도 마지막 기존 행(439) 아래에 S01·S02·S10·R02 로 붙는다", async () => {
    const drive = realDrive();
    const jobs = [rec("S10", PASS), rec("S02", PASS), rec("R02", PASS), rec("S01", PASS)];
    const out = await sync(drive, { jobs, all: [...realListRecords(), ...jobs] });
    expect(out.report.files.list).toMatchObject({ status: "ok", count: 4 });
    expect(sitePoints(drive, 439)).toEqual([
      "439 2026-11-11 HJ-M-005-350A S27 합격",
      "440 2026-11-12 HJ-M-005-350A S01 합격",
      "441 2026-11-12 HJ-M-005-350A S02 합격",
      "442 2026-11-12 HJ-M-005-350A S10 합격",
      "443 2026-11-12 HJ-M-005-350A R02 합격",
    ]);
  });

  it("백필처럼 날짜·라인이 섞인 묶음도 날짜 → 라인 → 포인트 순, 재검사는 원본 바로 뒤", async () => {
    const drive = realDrive();
    const fail = rec("R05", FAIL, { date: "2026-11-12" });
    const jobs = [
      rec("S03", PASS, { date: "2026-11-13" }),
      rec("R05", PASS, { date: "2026-11-12", ts: fail.ts + 1 }),
      rec("S01", PASS, { date: "2026-11-12", line: "HJ-M-006-350A" }),
      fail,
      rec("S05", PASS, { date: "2026-11-12" }),
      rec("S04", PASS, { date: "2026-11-12" }),
    ];
    await sync(drive, { jobs, all: [...realListRecords(), ...jobs] });
    expect(sitePoints(drive, 440)).toEqual([
      "440 2026-11-12 HJ-M-005-350A S04 합격",
      "441 2026-11-12 HJ-M-005-350A S05 합격",
      "442 2026-11-12 HJ-M-005-350A R05 불합격",
      "443 2026-11-12 HJ-M-005-350A R05 합격",
      "444 2026-11-12 HJ-M-006-350A S01 합격",
      "445 2026-11-13 HJ-M-005-350A S03 합격",
    ]);
  });

  it("이미 엑셀에 있는 Y 앱 ID 행은 옮기지 않고 그 행만 고친다. 새 행만 정렬해 붙는다", async () => {
    const drive = realDrive();
    const all = realListRecords();
    const existing = all.find((r) => r.src === "user" && r.point === "S26")!;
    expect(existing.excelRow).toBe(437);
    existing.result = FAIL;
    existing.sync = "pending";
    const jobs = [rec("S12", PASS), existing, rec("S03", PASS)];
    await sync(drive, { jobs, all: [...all, ...jobs.slice(0, 1), ...jobs.slice(2)] });
    expect(sitePoints(drive, 433)).toEqual([
      "433 2026-11-11 HJ-M-005-350A S22 합격",
      "434 2026-11-11 HJ-M-005-350A S23 합격",
      "435 2026-11-11 HJ-M-005-350A S24 합격",
      "436 2026-11-11 HJ-M-005-350A S25 합격",
      "437 2026-11-11 HJ-M-005-350A S26 불합격",
      "438 2026-11-11 HJ-M-005-350A R26 합격",
      "439 2026-11-11 HJ-M-005-350A S27 합격",
      "440 2026-11-12 HJ-M-005-350A S03 합격",
      "441 2026-11-12 HJ-M-005-350A S12 합격",
    ]);
    expect(listAppRows(drive).filter((r) => r.id === existing.id)).toHaveLength(1);
  });
});
