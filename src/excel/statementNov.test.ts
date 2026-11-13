/**
 * 거래명세서 11월(26.11) 채우기 — 실제 '02_PAUT 거래명세서 (지성, 드림).xlsx' 복사본(fixtures/real-statement.xlsx).
 * 복사본에는 지성군포 2611·지성브레인2611 만 있다. 사용자가 PC 엑셀에서 직접 만든 드림합정26.11·지성 마전 26.11 은
 * 테스트 도우미(userCreatesMonthSheet: 같은 접두어 최근 시트의 표기를 그대로 따름)로 흉내 낸다. 앱은 시트를 만들지 않는다.
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../domain/dates";
import type { InspectionRecord } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { aggregateSteps, syncPanelModel } from "../store/syncStatus";
import type { FakeDrive } from "../test/fakeGraph";
import { STMT_ID, fixture, graphConfig, meta, realDrive, realGmRecords, realListRecords, rec } from "../test/realBooks";
import { userCreatesMonthSheet } from "../test/userSheet";
import { missingSheetText, missingSheetsText, parseMissingSheetText, restoreStatementOwnership, syncGraph, type GraphSyncArgs } from "./engine";
import { readSheet, unzipXlsx, workbookSheets, xmlText, type ZipFiles } from "./ooxml";
import { expectedMonthSheetName, readPautRows, readStatement } from "./statementBook";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});

beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
}

function sync(drive: FakeDrive, jobs: InspectionRecord[], all: InspectionRecord[], extra: Partial<GraphSyncArgs> = {}) {
  return syncGraph({ client: client(drive), jobs, all, config: graphConfig(), meta: meta(), sleep: async () => undefined, ...extra });
}

const GP_LINE = "군포산본 이원화 750A-5";
const MJ_LINE = "MJDH-D-D-P-58-005-200A";

function gp(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "지성 군포 PA", sheetName: "지성 군포 PA", line: GP_LINE, diameter: "750A", ...partial });
}
function mj(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "지성 마전 PA", sheetName: "지성 마전 PA", line: MJ_LINE, diameter: "200A", ...partial });
}
function br(point: string, result: string, partial: Partial<InspectionRecord> = {}): InspectionRecord {
  return rec(point, result, { site: "지성 브레인", sheetName: " 지성 브레인", line: "BR-M-001-200A", diameter: "200A", ...partial });
}

type Cell = { v?: string; f?: string };
function cell(files: ZipFiles, sheet: string, ref: string): Cell {
  return (readSheet(files, sheet).cells.get(ref) as Cell | undefined) ?? {};
}

/** 한 PAUT 행을 지난달 행과 비교할 수 있는 모양으로: 날짜·공법·품명·개소·단위·수량식·단가·금액식·비고 */
function rowShape(files: ZipFiles, sheet: string, row: number) {
  const f = (ref: string) => (cell(files, sheet, ref).f ?? "").replace(/^=/, "");
  return {
    A: cell(files, sheet, `A${row}`).v,
    B: cell(files, sheet, `B${row}`).v,
    C: cell(files, sheet, `C${row}`).v,
    D: Number(cell(files, sheet, `D${row}`).v),
    E: cell(files, sheet, `E${row}`).v,
    // 지난달 시트는 'D12*1.2' 와 '0.7*D12' 가 섞여 있다(같은 뜻). 곱하는 길이만 비교한다.
    Flen: Number(/(?:D\d+\*([\d.]+))|(?:([\d.]+)\*D\d+)/.exec(f(`F${row}`))?.slice(1).find(Boolean)),
    G: Number(cell(files, sheet, `G${row}`).v),
    H: f(`H${row}`),
    I: cell(files, sheet, `I${row}`).v ?? "",
  };
}

const passes = (all: InspectionRecord[], site: string, date: string) => all.filter((r) => r.site === site && r.date === date && r.result === PASS && !r.deleted).length;

describe("거래명세서 11월 시트 이름 — 이 통합 문서의 표기 그대로", () => {
  const names = workbookSheets(unzipXlsx(fixture("real-statement.xlsx"))).map((s) => s.name);

  it("복사본에 이미 있는 11월: 지성군포 2611, 지성브레인2611 (점 없음). 없는 11월: 드림합정26.11, 지성 마전 26.11 (최근 시트 표기)", () => {
    expect(names).toContain("지성군포 2611");
    expect(names).toContain("지성브레인2611");
    expect(expectedMonthSheetName(names, "지성군포", 2026, 11)).toBe("지성군포 2611");
    expect(expectedMonthSheetName(names, "지성브레인", 2026, 11)).toBe("지성브레인2611");
    expect(expectedMonthSheetName(names, "드림합정", 2026, 11)).toBe("드림합정26.11");
    // 'AM' 시트(단가 다름)는 기준으로 쓰지 않는다: 지성 마전 26.8 → 지성 마전 26.11
    expect(expectedMonthSheetName(names, "지성 마전", 2026, 11)).toBe("지성 마전 26.11");
    // 12월은 각 접두어 최근 표기를 따른다
    expect(expectedMonthSheetName(names, "드림합정", 2026, 12)).toBe("드림합정26.12");
    expect(expectedMonthSheetName(names, "지성군포", 2026, 12)).toBe("지성군포 2612");
  });

  it("사용자가 다른 표기(지성군포 26.12)로 만들어도 그 시트를 찾는다", () => {
    const files = unzipXlsx(fixture("real-statement.xlsx"));
    const all = [...workbookSheets(files).map((s) => s.name), "지성군포 26.12"];
    expect(expectedMonthSheetName(all, "지성군포", 2026, 12)).toBe("지성군포 26.12");
  });
});

describe("거래명세서 11월 채우기 — 사용자가 만든 11월 시트(실제 복사본)", () => {
  it("섞인 묶음(합정·마전·군포·브레인, 불합격 포함): 각 11월 시트에 지난달과 같은 행 모양·단가로 채우고, 시트를 만들지 않고, 비고(I)는 비운다", async () => {
    const drive = realDrive({ novSheet: false });
    const stmt = drive.files(STMT_ID);
    // 사용자가 PC 엑셀에서 만든 11월 시트(실제 접두어 + 최근 표기)
    expect(userCreatesMonthSheet(stmt, "드림합정", 2026, 11)).toBe("드림합정26.11");
    expect(userCreatesMonthSheet(stmt, "지성 마전", 2026, 11)).toBe("지성 마전 26.11");
    const sheetsBefore = workbookSheets(stmt).map((s) => s.name);
    const gpBefore = readPautRows(stmt, "지성군포 2611").map((r) => [r.date, r.diameter, r.count]);
    const brBefore = readPautRows(stmt, "지성브레인2611").map((r) => [r.date, r.diameter, r.count]);

    const jobs = [
      rec("S28", PASS),
      rec("R28", PASS),
      rec("S29", FAIL), // 불합격: 명세서에 세지 않는다
      mj("S04", PASS, { date: "2026-11-12" }),
      mj("S05", PASS, { date: "2026-11-12" }),
      mj("R04", FAIL, { date: "2026-11-12" }),
      mj("S06", PASS, { date: "2026-11-13" }),
      gp("S07", PASS),
      gp("S08", PASS),
      br("S01", PASS), // 앱에서 지운 현장: 보내지 않는다
    ];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    const out = await sync(drive, jobs, all);
    expect(out.report.files.statement.status).toBe("ok");
    expect(out.report.files.statement.message).toBeUndefined();
    for (const r of jobs.filter((x) => x.site !== "지성 브레인")) expect(r.steps?.statement?.status).toBe("done");
    expect(out.steps.get(jobs[jobs.length - 1].id)).toBeUndefined();

    const after = drive.files(STMT_ID);
    // 시트는 사용자가 만든 것 그대로(늘지 않음)
    expect(workbookSheets(after).map((s) => s.name)).toEqual(sheetsBefore);

    // 드림합정26.11: 11/11 앱 행 7개소(리스트) + 11/12 합격 2개소(불합격 S29 제외)
    const hj = readPautRows(after, "드림합정26.11").map((r) => [r.date, r.diameter, r.count]);
    expect(hj).toEqual([
      ["2026-11-11", "350A", passes(all, "드림 합정주공", "2026-11-11")],
      ["2026-11-12", "350A", 2],
    ]);
    expect(hj[0][2]).toBe(7);
    // 지성 마전 26.11: 날짜별 합격만(11/12 2개소, 11/13 1개소). 리스트에 이미 있던 11월 마전 합격도 함께.
    const mjRows = readPautRows(after, "지성 마전 26.11");
    const mjDates = [...new Set(all.filter((r) => r.site === "지성 마전 PA" && r.date.startsWith("2026-11") && r.result === PASS).map((r) => r.date))].sort();
    expect(mjRows.map((r) => [r.date, r.count])).toEqual(mjDates.map((d) => [d, passes(all, "지성 마전 PA", d)]));
    expect(mjRows.find((r) => r.date === "2026-11-12")?.count).toBe(2);
    expect(mjRows.find((r) => r.date === "2026-11-13")?.count).toBe(1);
    // 지성군포 2611: 사람이 적은 11/3·11/4 는 그대로, 11/12 2개소만 더한다
    expect(readPautRows(after, "지성군포 2611").map((r) => [r.date, r.diameter, r.count])).toEqual([...gpBefore, ["2026-11-12", "750A", 2]]);
    // 지성브레인2611: 앱에서 지운 현장이라 그대로(쓰지도 비우지도 않음)
    expect(readPautRows(after, "지성브레인2611").map((r) => [r.date, r.diameter, r.count])).toEqual(brBefore);

    // 행 모양은 지난달 행과 같다(공법·품명·단위·길이·단가·금액식), 날짜는 엑셀 날짜 숫자, 비고(I)는 비어 있음
    const like = (sheet: string, row: number, prevSheet: string, prevRow: number, date: string, count: number) => {
      const got = rowShape(after, sheet, row);
      const prev = rowShape(after, prevSheet, prevRow);
      expect(got).toMatchObject({ B: prev.B, C: prev.C, E: prev.E, Flen: prev.Flen, G: prev.G, I: "" });
      expect(got.A).toBe(String(isoToExcelSerial(date)));
      expect(got.D).toBe(count);
      expect(got.H).toBe(`G${row}*F${row}`);
    };
    like("드림합정26.11", 13, "드림합정26.10", 12, "2026-11-12", 2);
    expect(rowShape(after, "드림합정26.11", 13)).toMatchObject({ C: "350A(1.2M)", G: 75000, Flen: 1.2 });
    const mj12 = mjRows.find((r) => r.date === "2026-11-12")!.row;
    like("지성 마전 26.11", mj12, "지성 마전 26.8", 12, "2026-11-12", 2);
    expect(rowShape(after, "지성 마전 26.11", mj12)).toMatchObject({ C: "200A(0.7M)", G: 65000, Flen: 0.7 });
    like("지성군포 2611", 14, "지성군포 2611", 12, "2026-11-12", 2);
    expect(rowShape(after, "지성군포 2611", 14)).toMatchObject({ C: "750A(2.4M)", G: 75000 });

    // 합계 행·SUM 범위는 지난달과 같은 자리(빈 행 안에서 채움)
    expect(readStatement(after, "드림합정26.11").totalRow).toBe(readStatement(after, "드림합정26.10").totalRow);
    expect(cell(after, "드림합정26.11", "H29").f?.replace(/^=/, "")).toBe("SUM(H12:H28)");

    // '앱' 표시는 어디에도 없다(시트 XML, Graph 요청 본문). I열에 쓰는 요청도 없다. 파일 업로드(PUT)·시트 추가 없음.
    for (const sheet of ["드림합정26.11", "지성 마전 26.11", "지성군포 2611", "지성브레인2611"]) {
      const info = workbookSheets(after).find((s) => s.name === sheet)!;
      expect(xmlText(after, info.path)).not.toContain("앱");
    }
    const patches = drive.requests((e) => e.item === STMT_ID && e.method === "PATCH");
    expect(patches.length).toBeGreaterThan(0);
    expect(JSON.stringify(patches.map((e) => e.body))).not.toContain("앱");
    for (const p of patches) {
      const address = /address='([^']+)'/.exec(p.path)?.[1] ?? "";
      const cols = address.split(":").map((a) => a.replace(/\$|\d/g, ""));
      expect(cols.every((c) => c >= "A" && c <= "H" && c.length === 1)).toBe(true);
    }
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);
    expect(drive.requests((e) => /worksheets\/add/i.test(e.path))).toHaveLength(0);

    // 다시 보내도 줄이 늘지 않는다(이 기기가 쓴 줄은 같은 행을 고침)
    for (const r of jobs) r.steps = undefined;
    await sync(drive, jobs, all);
    expect(readPautRows(drive.files(STMT_ID), "드림합정26.11")).toHaveLength(2);
    expect(readPautRows(drive.files(STMT_ID), "지성 마전 26.11")).toHaveLength(mjDates.length);
    expect(readPautRows(drive.files(STMT_ID), "지성군포 2611")).toHaveLength(3);
  });
});

describe("시트가 없을 때: 'N월 거래명세서 시트를 만들어주세요 (시트 이름)'", () => {
  it("문장 모양: 달은 숫자, 괄호 안은 만들 시트 이름. 여러 시트·여러 달은 달별로 모은다", () => {
    expect(missingSheetText("드림합정26.12", 12)).toBe("12월 거래명세서 시트를 만들어주세요 (드림합정26.12)");
    expect(missingSheetText("지성 마전 26.11", 11)).toBe("11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)");
    expect(parseMissingSheetText("11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)")).toEqual({ month: 11, sheets: ["지성 마전 26.11"] });
    expect(
      missingSheetsText([
        { sheet: "지성 마전 26.11", month: 11 },
        { sheet: "드림합정26.11", month: 11 },
        { sheet: "드림합정26.12", month: 12 },
      ]),
    ).toBe("11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11, 드림합정26.11) · 12월 거래명세서 시트를 만들어주세요 (드림합정26.12)");
  });

  it("11월 마전·12월 합정 시트가 없으면 그 기록만 대기 + 눈에 띄는 요청(ask). 다른 시트는 채우고, 시트를 만들면 자동으로 채운다", async () => {
    const drive = realDrive(); // 드림합정26.11 만 사용자가 만들어 둔 상태
    const novHj = rec("S28", PASS);
    const decHj = rec("S01", PASS, { date: "2026-12-01" });
    const novMj = mj("S04", PASS);
    const jobs = [novHj, decHj, novMj];
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    const sheetsBefore = workbookSheets(drive.files(STMT_ID)).map((s) => s.name);
    const out = await sync(drive, jobs, all);

    expect(out.report.files.statement.status).toBe("ok");
    expect(out.report.files.statement.message).toBe(
      "11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11) · 12월 거래명세서 시트를 만들어주세요 (드림합정26.12)",
    );
    expect(novHj.steps?.statement?.status).toBe("done");
    expect(decHj.steps?.statement).toMatchObject({ status: "wait", message: "12월 거래명세서 시트를 만들어주세요 (드림합정26.12)" });
    expect(novMj.steps?.statement).toMatchObject({ status: "wait", message: "11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)" });
    // 리스트·작업일지는 모두 반영. 실패(error)는 없다.
    for (const r of jobs) {
      expect(r.steps?.list?.status).toBe("done");
      expect(r.steps?.ilji?.status).toBe("done");
    }
    expect(workbookSheets(drive.files(STMT_ID)).map((s) => s.name)).toEqual(sheetsBefore);

    for (const r of jobs) r.sync = aggregateSteps(r.steps).status;
    const model = syncPanelModel(all, out.report, { online: true, syncing: false, mode: "graph" });
    expect(model.visible).toBe(true);
    expect(model.tone).toBe("ask");
    expect(model.title).toBe("11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11) · 12월 거래명세서 시트를 만들어주세요 (드림합정26.12)");
    expect(model.prompts).toEqual([model.title]);
    expect(model.lines.find((l) => l.file === "statement")).toMatchObject({ state: "waiting" });
    expect(model.lines.some((l) => l.state === "error")).toBe(false);

    // 다시 보내도 대기 그대로(오류로 바뀌지 않음)
    const waiting = jobs.filter((r) => r.sync === "pending");
    const again = await sync(drive, waiting, all);
    expect(again.report.files.statement.status).not.toBe("error");
    expect(waiting.every((r) => r.steps?.statement?.status === "wait")).toBe(true);

    // 사용자가 두 시트를 만든다 → 다음 동기화 때 자동으로 채운다
    expect(userCreatesMonthSheet(drive.files(STMT_ID), "지성 마전", 2026, 11)).toBe("지성 마전 26.11");
    expect(userCreatesMonthSheet(drive.files(STMT_ID), "드림합정", 2026, 12)).toBe("드림합정26.12");
    const filled = await sync(drive, waiting, all);
    expect(filled.report.files.statement.message).toBeUndefined();
    for (const r of waiting) expect(aggregateSteps(r.steps).status).toBe("synced");
    expect(readPautRows(drive.files(STMT_ID), "드림합정26.12").map((r) => [r.date, r.count])).toEqual([["2026-12-01", 1]]);
    expect(readPautRows(drive.files(STMT_ID), "지성 마전 26.11").some((r) => r.date === "2026-11-12" && r.count === 1)).toBe(true);
  });

  it("실패와 함께 있어도 시트 요청은 따로 보인다(prompts)", () => {
    const r = rec("S01", PASS, { date: "2026-12-01", sync: "error" });
    r.steps = {
      list: { status: "error", message: "PC 엑셀에서 파일이 잠겨 있습니다" },
      ilji: { status: "done" },
      statement: { status: "wait", message: "12월 거래명세서 시트를 만들어주세요 (드림합정26.12)" },
    };
    const model = syncPanelModel([r], null, { online: true, syncing: false, mode: "graph" });
    expect(model.tone).toBe("bad");
    expect(model.prompts).toEqual(["12월 거래명세서 시트를 만들어주세요 (드림합정26.12)"]);
  });
});
