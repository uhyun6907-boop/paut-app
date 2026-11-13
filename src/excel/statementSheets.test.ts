/**
 * 거래명세서:
 * (A) 비고(I열) 등 엑셀 어디에도 '앱' 같은 표시를 쓰지 않는다. 앱이 쓴 행은 이 기기의 기억(날짜·관경·개수)과
 *     옛 버전이 남긴 I='앱'(읽기만)으로 알아본다. 옛 '앱' 셀은 지우려고 다시 쓰지 않는다.
 * (B) 새 달 시트를 만들지 않는다. 없으면 그 현장·달만 'N월 거래명세서 시트를 만들어주세요 (…)'으로 대기, 시트가 생기면 채운다.
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { aggregateSteps, syncPanelModel } from "../store/syncStatus";
import type { FakeDrive } from "../test/fakeGraph";
import { ILJI_ID, LIST_ID, STMT_ID, fixture, graphConfig, meta, realDrive, realListRecords, rec } from "../test/realBooks";
import { userCreatesMonthSheet } from "../test/userSheet";
import * as engine from "./engine";
import { exportBook, restoreStatementOwnership, statementOwnership, syncGraph, syncPendingDemo, bootstrapFromSamples } from "./engine";
import { readSheet, setCell, setXml, unzipXlsx, workbookSheets, type ZipFiles } from "./ooxml";
import * as statementBook from "./statementBook";
import { LEGACY_APP_MARK, readPautRows, readStatement, writePautLines, type PautLine } from "./statementBook";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});

beforeEach(() => restoreStatementOwnership({}));

const NOV = "드림합정26.11";

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
}

function sync(drive: FakeDrive, jobs: InspectionRecord[], all: InspectionRecord[], extra: Partial<engine.GraphSyncArgs> = {}) {
  return syncGraph({ client: client(drive), jobs, all, config: graphConfig(), meta: meta(), sleep: async () => undefined, ...extra });
}

function line(date: string, count: number, diameter = "350A"): PautLine {
  return { date, diameter, lengthM: 1.2, count, unitPrice: 75000 };
}

function novBook(): ZipFiles {
  const files = unzipXlsx(fixture("real-statement.xlsx"));
  userCreatesMonthSheet(files, "드림합정", 2026, 11);
  return files;
}

function iCell(files: ZipFiles, sheet: string, row: number): string {
  return readSheet(files, sheet).cells.get(`I${row}`)?.v ?? "";
}

/** 옛 버전이 쓴 행처럼: 그 행 I열에 '앱' */
function markLegacy(files: ZipFiles, sheet: string, row: number): void {
  const info = workbookSheets(files).find((s) => s.name === sheet)!;
  setXml(files, info.path, setCell(readSheet(files, sheet).xml, `I${row}`, { kind: "inline", text: LEGACY_APP_MARK }));
}

function setD(files: ZipFiles, sheet: string, row: number, n: number): void {
  const info = workbookSheets(files).find((s) => s.name === sheet)!;
  setXml(files, info.path, setCell(readSheet(files, sheet).xml, `D${row}`, { kind: "number", n }));
}

function setI(files: ZipFiles, sheet: string, row: number, text: string): void {
  const info = workbookSheets(files).find((s) => s.name === sheet)!;
  setXml(files, info.path, setCell(readSheet(files, sheet).xml, `I${row}`, { kind: "inline", text }));
}

/** range(address='D12:H12') 가 I열을 덮는지 */
function touchesColumnI(path: string): boolean {
  const address = /address='([^']+)'/.exec(path)?.[1];
  if (!address) return false;
  const cols = address.split(":").map((a) => a.replace(/\$|\d/g, ""));
  const idx = (c: string) => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const [a, b = a] = cols.map(idx);
  return a <= 9 && 9 <= b;
}

const counts = (files: ZipFiles, sheet = NOV) => readStatement(files, sheet).paut.map((p) => [p.date, p.count]);

describe("(A) 비고에 '앱' 표시를 쓰지 않는다 — 단위", () => {
  it("새로 쓴 행의 I열(비고)은 비어 있고, 쓴 줄은 기억(owned)으로 돌려준다", () => {
    const files = novBook();
    const res = writePautLines(files, NOV, [line("2026-11-11", 7), line("2026-11-12", 2)]);
    expect(counts(files)).toEqual([["2026-11-11", 7], ["2026-11-12", 2]]);
    expect(iCell(files, NOV, 12)).toBe("");
    expect(iCell(files, NOV, 13)).toBe("");
    expect(readSheet(files, NOV).xml).not.toContain(">앱<");
    expect(res.owned).toEqual({ "2026-11-11|350A": 7, "2026-11-12|350A": 2 });
  });

  it("표시 없이도 기억으로 같은 행을 고친다(개수 7→8): 줄이 늘지 않고 I열은 여전히 비어 있다", () => {
    const files = novBook();
    const first = writePautLines(files, NOV, [line("2026-11-11", 7)]);
    const second = writePautLines(files, NOV, [line("2026-11-11", 8), line("2026-11-12", 3)], first.owned);
    expect(counts(files)).toEqual([["2026-11-11", 8], ["2026-11-12", 3]]);
    expect(readPautRows(files, NOV)).toHaveLength(2);
    expect(iCell(files, NOV, 12)).toBe("");
    expect(second.owned).toEqual({ "2026-11-11|350A": 8, "2026-11-12|350A": 3 });
  });

  it("기억이 없으면(다른 기기) 같은 날짜·관경 행은 사람 행으로 보고 건너뛴다: 중복 행이 생기지 않는다", () => {
    const files = novBook();
    writePautLines(files, NOV, [line("2026-11-11", 7)]);
    writePautLines(files, NOV, [line("2026-11-11", 8)]);
    expect(counts(files)).toEqual([["2026-11-11", 7]]);
  });

  it("사람이 개수를 고친 행(기억과 다름)은 사람 행으로 보고 덮어쓰지 않는다", () => {
    const files = novBook();
    const first = writePautLines(files, NOV, [line("2026-11-11", 7)]);
    setD(files, NOV, 12, 6);
    const res = writePautLines(files, NOV, [line("2026-11-11", 8)], first.owned);
    expect(counts(files)).toEqual([["2026-11-11", 6]]);
    expect(res.owned).toEqual({});
  });

  it("기록이 지워져 줄이 없어지면 앱이 쓴 행만 비우고, 사람이 그 행 I열에 적은 비고는 남긴다", () => {
    const files = novBook();
    const first = writePautLines(files, NOV, [line("2026-11-11", 7), line("2026-11-12", 2)]);
    setI(files, NOV, 13, "재검 포함");
    writePautLines(files, NOV, [line("2026-11-11", 7)], first.owned);
    expect(counts(files)).toEqual([["2026-11-11", 7]]);
    expect(iCell(files, NOV, 13)).toBe("재검 포함");
  });

  it("옛 버전 행(I='앱')은 그대로 앱 행으로 알아보고 제자리를 고친다. '앱' 셀은 지우지도 새로 쓰지도 않는다", () => {
    const files = novBook();
    writePautLines(files, NOV, [line("2026-11-11", 7)]);
    markLegacy(files, NOV, 12); // 옛 버전이 쓴 행
    const res = writePautLines(files, NOV, [line("2026-11-11", 8), line("2026-11-13", 1)]); // 기억 없음
    expect(counts(files)).toEqual([["2026-11-11", 8], ["2026-11-13", 1]]);
    expect(iCell(files, NOV, 12)).toBe(LEGACY_APP_MARK); // 옛 값 그대로
    expect(iCell(files, NOV, 13)).toBe(""); // 새 행에는 표시 없음
    expect(readPautRows(files, NOV)[0]).toMatchObject({ row: 12, app: true });
    expect(res.owned).toEqual({ "2026-11-11|350A": 8, "2026-11-13|350A": 1 });
  });

  it("옛 '앱' 행의 줄이 없어지면 그 행 전체를 비운다(그 행을 지울 때만 옛 표시도 함께)", () => {
    const files = novBook();
    writePautLines(files, NOV, [line("2026-11-11", 7)]);
    markLegacy(files, NOV, 12);
    writePautLines(files, NOV, []);
    expect(counts(files)).toEqual([]);
    expect(iCell(files, NOV, 12)).toBe("");
  });
});

describe("(B) 새 달 시트를 만들지 않는다 — 운영 코드", () => {
  it("거래명세서 모듈에 시트 복제·새 달 시트 만들기 함수가 없다", () => {
    for (const name of ["ensureMonthSheet", "cloneWorksheet", "monthSheetPosition"]) {
      expect(name in statementBook).toBe(false);
      expect(name in engine).toBe(false);
    }
  });
});

describe("(B) OneDrive: 그 달 시트가 없으면 건너뛰고 대기, 시트가 생기면 채운다", () => {
  it("시트를 만들지 않고(업로드·시트 추가 없음) '11월 거래명세서 시트를 만들어주세요 (드림합정26.11)' 요청·대기. 리스트·작업일지·다른 현장 명세서는 그대로 반영", async () => {
    const drive = realDrive({ novSheet: false });
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((s) => s.name);
    expect(sheetsBefore).not.toContain(NOV);
    const brain = rec("S01", PASS, { site: "지성 브레인", sheetName: " 지성 브레인", line: "BR-M-001-350A", diameter: "350A" });
    const hap = [rec("S28", PASS), rec("R28", PASS)];
    const jobs = [...hap, brain];
    const all = [...realListRecords(), ...jobs];
    const brainBefore = readStatement(drive.files(STMT_ID), "지성브레인2611").paut;
    const out = await sync(drive, jobs, all);

    // 리스트·작업일지는 정상
    // 지성 브레인은 앱에서 지운 현장: 같이 넘어와도 보내지 않는다.
    expect(out.report.files.list).toMatchObject({ status: "ok", count: 2 });
    expect(out.report.files.ilji.status).toBe("ok");
    // 명세서: 실패가 아니다. 지운 현장(지성브레인2611)은 그대로.
    expect(out.report.files.statement.status).not.toBe("error");
    expect(out.report.files.statement.message).toBe("11월 거래명세서 시트를 만들어주세요 (드림합정26.11)");
    expect(out.steps.get(brain.id)).toBeUndefined();
    expect(readStatement(drive.files(STMT_ID), "지성브레인2611").paut).toEqual(brainBefore);
    for (const r of hap) {
      expect(r.steps?.list?.status).toBe("done");
      expect(r.steps?.ilji?.status).toBe("done");
      expect(r.steps?.statement).toMatchObject({ status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (드림합정26.11)" });
      expect(aggregateSteps(r.steps).status).toBe("pending");
      r.sync = "pending";
    }
    brain.sync = "synced";
    // 시트는 하나도 늘지 않고, 파일 업로드(PUT)도 없다.
    expect(workbookSheets(drive.files(STMT_ID)).map((s) => s.name)).toEqual(sheetsBefore);
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => e.method === "POST" && /worksheets\/add/i.test(e.path))).toHaveLength(0);

    // 화면: 조용한 대기(빨간 실패 아님)
    const model = syncPanelModel(all, out.report, { online: true, syncing: false, mode: "graph" });
    expect(model.visible).toBe(true);
    expect(model.tone).toBe("ask");
    expect(model.title).toBe("11월 거래명세서 시트를 만들어주세요 (드림합정26.11)");
    expect(model.lines.map((l) => l.state)).toEqual(["ok", "ok", "waiting"]);
    expect(model.lines[2].text).toBe("거래명세서 · 11월 거래명세서 시트를 만들어주세요 (드림합정26.11) · 대기 2건");
    expect(model.canRetry).toBe(true);

    // 시트가 아직 없으면 다시 보내도 그대로 대기(오류·시트 추가 없음). 리스트·작업일지는 다시 쓰지 않는다.
    const listPatches = drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length;
    const iljiPatches = drive.requests((e) => e.item === ILJI_ID && e.method === "PATCH").length;
    const again = await sync(drive, hap, all);
    expect(again.report.files.statement.status).not.toBe("error");
    expect(hap.every((r) => r.steps?.statement?.status === "wait")).toBe(true);
    expect(workbookSheets(drive.files(STMT_ID)).map((s) => s.name)).toEqual(sheetsBefore);

    // 사용자가 PC 엑셀에서 드림합정26.11 을 만든다 → 다음 동기화 때 대기 기록을 채운다.
    userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 11);
    const filled = await sync(drive, hap, all);
    expect(filled.report.files.statement).toMatchObject({ status: "ok", count: 1 });
    expect(filled.report.files.statement.message).toBeUndefined();
    for (const r of hap) expect(aggregateSteps(r.steps).status).toBe("synced");
    // 그 달 전체: 리스트의 11/11 7개소 + 오늘 2개소
    expect(counts(drive.files(STMT_ID))).toEqual([["2026-11-11", 7], ["2026-11-12", 2]]);
    expect(iCell(drive.files(STMT_ID), NOV, 12)).toBe("");
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH").length).toBe(listPatches);
    expect(drive.requests((e) => e.item === ILJI_ID && e.method === "PATCH").length).toBe(iljiPatches);
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);
  });

  it("백필(이미 리스트에 있는 11/11 기록)도 시트가 없으면 대기로 두고 실패로 바꾸지 않는다", async () => {
    const drive = realDrive({ novSheet: false });
    const all = realListRecords();
    const app = all.filter((r) => r.src === "user");
    const out = await sync(drive, [], all, { backfill: true });
    expect(out.report.files.ilji.status).toBe("ok");
    expect(out.report.files.statement.status).not.toBe("error");
    for (const r of app) expect(out.steps.get(r.id)?.statement).toMatchObject({ status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (드림합정26.11)" });
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);

    userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 11);
    await sync(drive, [], all, { backfill: true });
    expect(counts(drive.files(STMT_ID))).toEqual([["2026-11-11", 7]]);
    for (const r of app) expect(r.steps?.statement?.status).toBe("done");
  });

  it("이미 있는 시트 채우기·고치기는 그대로: 이 기기가 쓴 줄은 개수가 바뀌면 같은 행을 고친다(표시 없이)", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS)];
    const all = [...realListRecords(), ...jobs];
    await sync(drive, jobs, all);
    expect(counts(drive.files(STMT_ID))).toEqual([["2026-11-11", 7], ["2026-11-12", 1]]);
    expect(Object.values(statementOwnership())).toContainEqual({ "2026-11-11|350A": 7, "2026-11-12|350A": 1 });

    const more = [rec("S29", PASS), rec("S30", PASS)];
    await sync(drive, more, [...all, ...more]);
    expect(counts(drive.files(STMT_ID))).toEqual([["2026-11-11", 7], ["2026-11-12", 3]]);
    expect(readPautRows(drive.files(STMT_ID), NOV)).toHaveLength(2);
    for (const row of [12, 13]) expect(iCell(drive.files(STMT_ID), NOV, row)).toBe("");
    const bodies = drive.requests((e) => e.item === STMT_ID && e.method === "PATCH").map((e) => JSON.stringify(e.body));
    expect(bodies.join("")).not.toContain("앱");
  });

  it("OneDrive 에 옛 버전 행(I='앱')이 있으면 그 행을 제자리에서 고치고, '앱' 셀은 다시 쓰지 않는다", async () => {
    const drive = realDrive();
    const stmt = drive.files(STMT_ID);
    writePautLines(stmt, NOV, [line("2026-11-11", 7)]);
    markLegacy(stmt, NOV, 12); // 옛 버전이 쓴 11/11 행
    const jobs = [rec("S28", PASS, { date: "2026-11-11" })];
    await sync(drive, jobs, [...realListRecords(), ...jobs]);
    expect(counts(drive.files(STMT_ID))).toEqual([["2026-11-11", 8]]);
    expect(iCell(drive.files(STMT_ID), NOV, 12)).toBe(LEGACY_APP_MARK);
    // I열에 쓰는 요청이 없다(옛 '앱'을 지우려고 다시 쓰지도 않음).
    const patches = drive.requests((e) => e.item === STMT_ID && e.method === "PATCH");
    expect(patches.length).toBeGreaterThan(0);
    expect(patches.some((e) => touchesColumnI(e.path))).toBe(false);
  });
});

describe("(B) 데모(이 기기 통합 문서)도 시트를 만들지 않는다", () => {
  it("시트가 없으면 대기 + 시트 요청, 시트 수는 그대로", () => {
    bootstrapFromSamples({ list: fixture("real-list.xlsx"), ilji: fixture("real-worklog.xlsx"), statement: fixture("real-statement.xlsx") });
    const before = workbookSheets(unzipXlsx(exportBook("statement"))).map((s) => s.name);
    const jobs = [rec("S28", PASS)];
    const res = syncPendingDemo(jobs, [...realListRecords(), ...jobs], graphConfig(), meta());
    expect(res.get(jobs[0].id)).toMatchObject({
      status: "pending",
      steps: { list: { status: "done" }, ilji: { status: "done" }, statement: { status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (드림합정26.11)" } },
    });
    expect(workbookSheets(unzipXlsx(exportBook("statement"))).map((s) => s.name)).toEqual(before);
  });
});
