/**
 * 'PAUT NDT 리스트(군포, 마전).xlsx' (드라이브 루트) — 두 번째 검사리스트 파일.
 * 실제 파일 복사본(fixtures/real-list-gunpo-majeon.xlsx)과 가짜 Graph 로만 시험한다(실제 OneDrive 호출 없음).
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../domain/dates";
import type { InspectionRecord } from "../domain/types";
import { FAIL, LIST_GM_FILE, PASS } from "../domain/types";
import { GraphClient, graphRetry, listRowFormulas } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import {
  AJEON_ID,
  ILJI_ID,
  LIST_GM_ID,
  LIST_GM_NAME,
  LIST_ID,
  STMT_ID,
  fixture,
  graphConfig,
  meta,
  realDrive,
  realGmRecords,
  realListRecords,
  rec,
} from "../test/realBooks";
import {
  bootstrapFromSamples,
  downloadListBook,
  exportBook,
  listFileOf,
  restoreStatementOwnership,
  syncGraph,
  syncPendingDemo,
  statementSkipReason,
  type GraphSyncArgs,
} from "./engine";
import { readIljiDay } from "./iljiBook";
import { LIST_LAYOUT, listBookOf, parseListWorkbook, writeListRow } from "./listBook";
import { readSheet, unzipXlsx, workbookSheets, xmlText, type ZipFiles } from "./ooxml";
import { readPautRows } from "./statementBook";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});

beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({
    fetchImpl: drive.fetch,
    sleep: async () => undefined,
    getToken: async (force) => {
      drive.events.push(`token:${force ? "force" : "silent"}`);
      return drive.validToken;
    },
  });
}

async function sync(drive: FakeDrive, partial: Partial<GraphSyncArgs> & { jobs: InspectionRecord[]; all: InspectionRecord[] }) {
  return syncGraph({ client: client(drive), config: graphConfig(), meta: meta(), sleep: async () => undefined, ...partial });
}

const GP_LINE = "군포산본 이원화 750A-5";
const MJ_LINE = "MJDH-D-D-P-58-005-200A";
const AJ_LINE = "GP-001-750A";

/** 군포·마전 기록. listBook 은 일부러 비워 시트 이름으로 길을 찾게 한다(listBookOf). */
function gp(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "지성 군포 PA", sheetName: "지성 군포 PA", line: GP_LINE, diameter: "750A", ...partial });
}
function mj(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "지성 마전 PA", sheetName: "지성 마전 PA", line: MJ_LINE, diameter: "200A", ...partial });
}
function aj(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: AJ_LINE, diameter: "750A", ...partial });
}

function cells(files: ZipFiles, sheet: string) {
  return readSheet(files, sheet).cells;
}

function v(files: ZipFiles, sheet: string, ref: string): string {
  return cells(files, sheet).get(ref)?.v ?? "";
}

function appRows(files: ZipFiles, book: "list" | "listGm" = "listGm"): InspectionRecord[] {
  return parseListWorkbook(files, book).flatMap((s) => s.records).filter((r) => r.src === "user");
}

function allXml(files: ZipFiles): string {
  return Object.keys(files)
    .filter((k) => k.endsWith(".xml"))
    .map((k) => xmlText(files, k))
    .join("\n");
}

describe("군포·마전 리스트 파일 읽기", () => {
  it("PA 시트 3개(지성 군포·아전 군포·지성 마전)를 현장으로 읽고 RT 시트는 건너뛴다. 열 배치는 평택 리스트와 같다", () => {
    const files = unzipXlsx(fixture("real-list-gunpo-majeon.xlsx"));
    expect(workbookSheets(files).map((s) => s.name)).toEqual(["지성 군포 PA", "아전 군포 PA", "지성 마전 PA", "지성마전 RT"]);
    for (const sheet of ["지성 군포 PA", "아전 군포 PA", "지성 마전 PA"]) {
      expect(v(files, sheet, "F4")).toBe("line No.");
      expect(v(files, sheet, "G4")).toBe("Point");
      expect(v(files, sheet, "E4").trim()).toBe("Examination Date");
      expect(v(files, sheet, "P4")).toBe("합부");
      expect(v(files, sheet, "S4")).toBe("비고");
    }
    const sites = parseListWorkbook(files, "listGm");
    expect(sites.map((s) => [s.site, s.records.length])).toEqual([
      ["지성 군포 PA", 128],
      ["아전 군포 PA", 34],
      ["지성 마전 PA", 308],
    ]);
    const recs = sites.flatMap((s) => s.records);
    expect(recs.every((r) => r.listBook === "listGm" && r.src === "xlsx")).toBe(true);
    // 리스트에는 불합격 행도 있다(리스트 규칙). 재검사 R31 은 같은 번호.
    expect(recs.filter((r) => r.result === FAIL).map((r) => `${r.site}:${r.point}`)).toEqual([
      "지성 군포 PA:R31",
      "아전 군포 PA:R10",
      "지성 마전 PA:R03",
      "지성 마전 PA:S25",
    ]);
    const first = sites[0].records[0];
    expect(first).toMatchObject({ date: "2025-11-21", line: "군포산본 이원화 750A-1", point: "S22", diameter: "750A", material: "SPPS380", result: PASS });
  });

  it("마지막 날 합격 개소가 거래명세서와 같다(1 합격 = 1개소): 지성 군포 11/3 4개·11/4 7개", () => {
    const recs = realGmRecords().filter((r) => r.site === "지성 군포 PA" && r.result === PASS);
    expect(recs.filter((r) => r.date === "2026-11-03")).toHaveLength(4);
    expect(recs.filter((r) => r.date === "2026-11-04")).toHaveLength(7);
    const rows = readPautRows(unzipXlsx(fixture("real-statement.xlsx")), "지성군포 2611");
    expect(rows.map((r) => [r.date, r.diameter, r.count])).toEqual([
      ["2026-11-03", "750A", 4],
      ["2026-11-04", "750A", 7],
    ]);
  });

  it("어느 리스트 파일로 갈지: 군포·마전 PA 시트 → 군포·마전 리스트, 평택 현장 → 평택 리스트", () => {
    expect(listBookOf({ sheetName: "지성 군포 PA" })).toBe("listGm");
    expect(listBookOf({ sheetName: " 아전 군포 PA " })).toBe("listGm");
    expect(listBookOf({ sheetName: "지성 마전 PA" })).toBe("listGm");
    expect(listBookOf({ sheetName: " 드림 합정주공" })).toBe("list");
    expect(listBookOf({ sheetName: " 지성 브레인" })).toBe("list");
    expect(realListRecords().every((r) => listBookOf(r) === "list")).toBe(true);
    expect(realGmRecords().every((r) => listBookOf(r) === "listGm")).toBe(true);
    // 예전 설정(listGm 칸 없음) → 드라이브 루트의 기본 파일
    expect(listFileOf(graphConfig(), "listGm")).toEqual({ id: LIST_GM_FILE.id, name: "PAUT NDT 리스트(군포, 마전).xlsx" });
    expect(listFileOf({ ...graphConfig(), files: { ...graphConfig().files, listGm: null } }, "listGm")).toBeNull();
  });

  it("관경(L)은 값으로 쓴다: '군포산본 이원화 750A-5' 의 RIGHT(F,4) 는 '0A-5' 라서 수식이면 틀린다", () => {
    const files = unzipXlsx(fixture("real-list-gunpo-majeon.xlsx"));
    const row = writeListRow(files, "지성 군포 PA", { id: "x-1", date: "2026-11-12", line: GP_LINE, point: "S07", result: PASS, material: "SPPS380" }, LIST_LAYOUT.listGm);
    expect(row).toBe(134);
    const c = cells(files, "지성 군포 PA");
    expect(c.get("L134")?.v).toBe("750A");
    expect(c.get("L134")?.f).toBeUndefined();
    expect(c.get("A134")?.v).toBe("129"); // 미리 적힌 순번은 그대로
    expect(listRowFormulas(134, { id: "x", date: "2026-11-12", line: GP_LINE, point: "S07", result: PASS }, { diameterValue: "750A" })[11]).toBe("750A");
    expect(listRowFormulas(134, { id: "x", date: "2026-11-12", line: GP_LINE, point: "S07", result: PASS })[11]).toBe("=RIGHT(F134,4)");
  });
});

describe("Graph 동기화: 네 번째 파일(군포·마전 리스트)", () => {
  it("현장별로 자기 리스트에만, 날짜→라인→포인트(S 다음 R) 순으로 바뀐 칸만 쓴다. 작업일지·명세서는 합격만, 비고에 '앱' 없음", async () => {
    const drive = realDrive();
    const gpBefore = drive.bytes(LIST_GM_ID);
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((s) => s.name);
    const failS08 = gp("S08", FAIL);
    const jobs = [
      gp("R07", PASS),
      gp("S07", PASS),
      failS08,
      gp("S09", PASS),
      gp("S08", PASS), // 재검사: 같은 번호, 새 행(원본 바로 뒤)
      mj("S04", PASS),
      mj("R04", PASS),
      aj("S17", PASS),
      rec("S28", PASS), // 드림 합정주공 → 평택 리스트
    ];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    const out = await sync(drive, { jobs, all });

    expect(out.report.files.list).toMatchObject({ status: "ok", count: 9 });
    expect(out.report.lists).toMatchObject({ list: { status: "ok", count: 1 }, listGm: { status: "ok", count: 8 } });
    expect(out.report.files.ilji.status).toBe("ok");
    for (const r of jobs) expect(out.steps.get(r.id)?.list?.status).toBe("done");

    // 군포·마전 리스트: 지성 군포 PA 마지막 행(133) 아래 134~138
    const gm = drive.files(LIST_GM_ID);
    const order = [134, 135, 136, 137, 138].map((row) => `${v(gm, "지성 군포 PA", `G${row}`)}:${v(gm, "지성 군포 PA", `P${row}`)}`);
    expect(order).toEqual(["S07:합격", "S08:불합격", "S08:합격", "S09:합격", "R07:합격"]);
    for (const row of [134, 135, 136, 137, 138]) {
      expect(v(gm, "지성 군포 PA", `E${row}`)).toBe(String(isoToExcelSerial("2026-11-12")));
      expect(v(gm, "지성 군포 PA", `F${row}`)).toBe(GP_LINE);
      const l = cells(gm, "지성 군포 PA").get(`L${row}`);
      expect(l?.v).toBe("750A");
      expect(l?.f).toBeUndefined();
      expect(v(gm, "지성 군포 PA", `S${row}`)).toBe(""); // 비고는 건드리지 않음
    }
    // 미리 적힌 순번(A134=129…)은 그대로
    expect([134, 135, 136, 137, 138].map((row) => v(gm, "지성 군포 PA", `A${row}`))).toEqual(["129", "130", "131", "132", "133"]);
    // 지성 마전 PA: 313 아래 314·315, 순번이 없던 행이라 309·310 을 채운다
    expect([314, 315].map((row) => [v(gm, "지성 마전 PA", `A${row}`), v(gm, "지성 마전 PA", `G${row}`), v(gm, "지성 마전 PA", `L${row}`)])).toEqual([
      ["309", "S04", "200A"],
      ["310", "R04", "200A"],
    ]);
    // 아전 군포 PA: 39 아래 40 (순번 35 미리 적힘)
    expect([v(gm, "아전 군포 PA", "A40"), v(gm, "아전 군포 PA", "F40"), v(gm, "아전 군포 PA", "G40")]).toEqual(["35", AJ_LINE, "S17"]);
    // 기존 행은 그대로
    expect(v(gm, "지성 군포 PA", "G133")).toBe("R06");
    expect(v(gm, "지성 군포 PA", "P119")).toBe("불합격");

    // 다른 리스트로 새지 않는다
    const gmIds = new Set(appRows(gm).map((r) => r.id));
    const pyIds = new Set(appRows(drive.files(LIST_ID), "list").map((r) => r.id));
    for (const r of jobs.slice(0, 8)) {
      expect(gmIds.has(r.id)).toBe(true);
      expect(pyIds.has(r.id)).toBe(false);
    }
    expect(gmIds.has(jobs[8].id)).toBe(false);
    expect(pyIds.has(jobs[8].id)).toBe(true);
    expect(gmIds.size).toBe(8);

    // '앱' 은 어디에도 없다. 파일 통째 업로드(PUT)도 없다. 리스트 쓰기는 통합 문서 세션 안의 칸 쓰기.
    expect(allXml(gm)).not.toContain("앱");
    const gmPatches = drive.requests((e) => e.item === LIST_GM_ID && e.method === "PATCH");
    expect(gmPatches).toHaveLength(8);
    expect(gmPatches.every((e) => !!e.session && /range\(address='A\d+:Y\d+'\)/.test(e.path))).toBe(true);
    expect(JSON.stringify(gmPatches.map((e) => e.body))).not.toContain("앱");
    // 통째 업로드(PUT)는 없다. 아전 .xls 명세서는 이번 배포에서 받지도 올리지도 않는다.
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => e.item === AJEON_ID)).toHaveLength(0);
    expect(gpBefore.byteLength).toBeGreaterThan(0);

    // 파일마다 새 토큰: 평택 리스트 세션을 닫은 뒤, 군포·마전 세션을 열기 전에 토큰을 다시 받는다.
    const ev = drive.events;
    const lastPy = ev.lastIndexOf(`req:${LIST_ID}:POST`);
    const firstGm = ev.indexOf(`req:${LIST_GM_ID}:POST`);
    expect(lastPy).toBeGreaterThan(-1);
    expect(firstGm).toBeGreaterThan(lastPy);
    expect(ev.slice(lastPy, firstGm).some((e) => e.startsWith("token:"))).toBe(true);

    // 작업일지: 합격만(불합격 S08 제외). 지성 군포 4, 지성 마전 2, 아전 군포 1.
    const ilji = drive.files(ILJI_ID);
    const sum = (groups: { count: number }[] | null) => (groups ?? []).reduce((n, g) => n + g.count, 0);
    expect(sum(readIljiDay(ilji, "2026-11-12", "지성 군포"))).toBe(4);
    expect(sum(readIljiDay(ilji, "2026-11-12", "지성 마전"))).toBe(2);
    expect(sum(readIljiDay(ilji, "2026-11-12", "아전 군포"))).toBe(1);
    expect(out.steps.get(failS08.id)?.ilji?.status).toBe("done");

    // 거래명세서: 지성군포 2611 에 11/12 750A 4개소(기존 11/3·11/4 줄은 그대로). 시트는 늘지 않는다.
    const rows = readPautRows(drive.files(STMT_ID), "지성군포 2611");
    expect(rows.map((r) => [r.date, r.diameter, r.count])).toEqual([
      ["2026-11-03", "750A", 4],
      ["2026-11-04", "750A", 7],
      ["2026-11-12", "750A", 4],
    ]);
    expect(v(drive.files(STMT_ID), "지성군포 2611", "I14")).toBe("");
    expect(workbookSheets(drive.files(STMT_ID)).map((s) => s.name)).toEqual(sheetsBefore);
    // 지성 마전 26.11 시트는 없다 → 만들지 않고 조용히 대기(리스트·작업일지는 반영)
    const mjStep = out.steps.get(jobs[5].id);
    expect(mjStep?.list?.status).toBe("done");
    expect(mjStep?.ilji?.status).toBe("done");
    expect(mjStep?.statement).toMatchObject({ status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)" });
    // 아전 군포: 명세서는 아직 쓰지 않는다(접두어 미확정 → 건너뜀). .xls 는 건드리지 않는다.
    expect(out.steps.get(jobs[7].id)?.statement).toMatchObject({ status: "skip", message: statementSkipReason(graphConfig(), "아전 군포 PA") });
    expect(out.report.ajeon).toBeUndefined();
    expect(out.report.files.statement.status).not.toBe("error");
  });

  it("군포·마전 기록만 보내면 평택 리스트에는 요청조차 하지 않는다. 다시 보내도 Y(앱 ID)로 같은 행을 고쳐 행이 늘지 않는다", async () => {
    const drive = realDrive();
    // 바이트(zip)는 묶을 때 시각이 들어가 매번 달라질 수 있어, 풀린 파일 내용을 비교한다.
    const pyBefore = { ...drive.files(LIST_ID) };
    const jobs = [gp("S07", PASS), gp("R07", PASS)];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    await sync(drive, { jobs, all });
    expect(drive.requests((e) => e.item === LIST_ID)).toHaveLength(0);
    const pyAfter = drive.files(LIST_ID);
    expect(Object.keys(pyAfter).sort()).toEqual(Object.keys(pyBefore).sort());
    for (const k of Object.keys(pyBefore)) expect(pyAfter[k] === pyBefore[k], k).toBe(true);

    for (const r of jobs) r.steps = undefined;
    const again = await sync(drive, { jobs, all });
    expect(again.report.lists?.listGm).toMatchObject({ status: "ok", count: 2 });
    const mine = appRows(drive.files(LIST_GM_ID));
    expect(mine.map((r) => [r.excelRow, r.point])).toEqual([
      [134, "S07"],
      [135, "R07"],
    ]);
  });

  it("군포·마전 리스트가 잠기면 그 파일만 실패. 평택 리스트·작업일지·명세서는 계속, 화면에 파일별 줄이 따로 보인다", async () => {
    const drive = realDrive();
    drive.fail((m, p) => m === "PATCH" && p.includes(LIST_GM_ID), 423, Infinity);
    const g = gp("S07", PASS);
    const h = rec("S28", PASS);
    const jobs = [g, h];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    const out = await sync(drive, { jobs, all });
    expect(out.steps.get(g.id)?.list?.status).toBe("error");
    expect(out.steps.get(h.id)?.list?.status).toBe("done");
    expect(out.steps.get(g.id)?.ilji?.status).toBe("done");
    expect(out.report.lists?.list).toMatchObject({ status: "ok", count: 1 });
    expect(out.report.lists?.listGm?.status).toBe("error");
    expect(out.report.files.list.message).toMatch(/^검사리스트\(군포·마전\): PC 엑셀/);

    for (const r of jobs) r.sync = r.steps?.list?.status === "error" ? "error" : "synced";
    const model = syncPanelModel(jobs, out.report, { online: true, syncing: false, mode: "graph" });
    expect(model.visible).toBe(true);
    const lines = model.lines.map((l) => [l.book ?? l.file, l.state]);
    expect(lines).toEqual([
      ["list", "ok"],
      ["listGm", "error"],
      ["ilji", "ok"],
      ["statement", "ok"],
    ]);
    const gmLine = model.lines.find((l) => l.book === "listGm")!;
    expect(gmLine.text).toMatch(/^검사리스트\(군포·마전\) · 실패 — PC 엑셀.*\(1건\)$/);
    expect(model.lines.find((l) => l.book === "list")!.label).toBe("검사리스트(평택)");
  });

  it("평택만 있을 때는 예전처럼 검사리스트 한 줄(세 줄)", async () => {
    const drive = realDrive();
    const jobs = [rec("S28", PASS)];
    const out = await sync(drive, { jobs, all: [...realListRecords(), ...jobs] });
    expect(out.report.lists?.listGm).toBeUndefined();
    const model = syncPanelModel(jobs, out.report, { online: true, syncing: false, mode: "graph" });
    expect(model.lines.map((l) => l.label)).toEqual(["검사리스트", "작업일지", "거래명세서"]);
    expect(drive.requests((e) => e.item === LIST_GM_ID)).toHaveLength(0);
  });

  it("저장된 항목 ID 로 못 찾으면 드라이브 루트에서 정확한 이름 'PAUT NDT 리스트(군포, 마전).xlsx' 로 찾는다", async () => {
    const drive = new FakeDrive();
    drive.put("NEW-ID", LIST_GM_NAME, fixture("real-list-gunpo-majeon.xlsx"));
    drive.put("OTHER", "PAUT NDT 리스트(군포).xlsx", fixture("real-list-gunpo-majeon.xlsx"));
    const got = await downloadListBook(client(drive), { id: LIST_GM_ID, name: LIST_GM_NAME });
    expect(got.file).toEqual({ id: "NEW-ID", name: LIST_GM_NAME });
    expect(parseListWorkbook(unzipXlsx(got.bytes), "listGm").map((s) => s.site)).toEqual(["지성 군포 PA", "아전 군포 PA", "지성 마전 PA"]);
    await expect(downloadListBook(client(new FakeDrive()), { id: LIST_GM_ID, name: LIST_GM_NAME })).rejects.toThrow(/찾지 못했습니다/);
  });
});

describe("데모(이 기기 사본)도 같은 규칙", () => {
  it("군포 기록은 군포·마전 사본에만, 합정 기록은 평택 사본에만 쓴다", () => {
    bootstrapFromSamples({
      list: fixture("real-list.xlsx"),
      ilji: fixture("real-worklog.xlsx"),
      statement: fixture("real-statement.xlsx"),
      listGm: fixture("real-list-gunpo-majeon.xlsx"),
    });
    const jobs = [gp("S07", PASS), gp("R07", FAIL), rec("S28", PASS)];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    syncPendingDemo(jobs, all, graphConfig(), meta());
    const gm = unzipXlsx(exportBook("listGm"));
    const py = unzipXlsx(exportBook("list"));
    expect(appRows(gm).map((r) => [r.point, r.result, r.excelRow])).toEqual([
      ["S07", PASS, 134],
      ["R07", FAIL, 135],
    ]);
    expect(appRows(py, "list").some((r) => r.site.includes("군포"))).toBe(false);
    expect(appRows(py, "list").map((r) => r.id)).toContain(jobs[2].id);
    expect(cells(gm, "지성 군포 PA").get("L134")?.f).toBeUndefined();
    expect(allXml(gm)).not.toContain("앱");
  });
});
