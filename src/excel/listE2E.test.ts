/**
 * 검사리스트 끝까지 확인 — 실제 두 리스트 복사본(평택 real-list.xlsx, 군포·마전 real-list-gunpo-majeon.xlsx).
 * 앱이 보낸 기록이: 빠짐없이, 중복 없이, 맞는 파일·맞는 시트에, 날짜→라인→포인트(S 다음 R, 재검사는 같은 번호 새 행) 순으로,
 * 맞는 열(E 날짜·F 라인·G 포인트·K 재질·L 관경·P 합부·Y 앱 ID, 비고 S 는 비움)에 들어가는지 본다.
 * 리스트에는 불합격도 들어가고(리스트 규칙), 작업일지·거래명세서에는 합격만 센다.
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../domain/dates";
import { comparePointLabel } from "../domain/point";
import type { InspectionRecord, ListBookKind } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import type { FakeDrive } from "../test/fakeGraph";
import { LIST_GM_ID, LIST_ID, STMT_ID, ILJI_ID, fixture, graphConfig, meta, realDrive, realGmRecords, realListRecords, rec } from "../test/realBooks";
import { bootstrapFromSamples, exportBook, restoreStatementOwnership, syncGraph, syncPendingDemo } from "./engine";
import { readIljiDay } from "./iljiBook";
import { listBookOf, parseListWorkbook } from "./listBook";
import { readSheet, unzipXlsx, workbookSheets, xmlText, type ZipFiles } from "./ooxml";
import { readPautRows } from "./statementBook";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});

beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
}

function sync(drive: FakeDrive, jobs: InspectionRecord[], all: InspectionRecord[]) {
  return syncGraph({ client: client(drive), jobs, all, config: graphConfig(), meta: meta(), sleep: async () => undefined });
}

const HJ5 = "HJ-M-005-350A";
const HJ6 = "HJ-M-006-350A";
const GP = "군포산본 이원화 750A-5";
const MJ4 = "MJDH-D-D-P-58-004-200A";
const MJ5 = "MJDH-D-D-P-58-005-200A";
const AJ = "GP-001-750A";

const hj = (point: string, result: string, date: string, line = HJ5) => rec(point, result, { date, line });
const gp = (point: string, result: string, date: string) => rec(point, result, { site: "지성 군포 PA", sheetName: "지성 군포 PA", line: GP, diameter: "750A", date });
const mj = (point: string, result: string, date: string, line = MJ5) => rec(point, result, { site: "지성 마전 PA", sheetName: "지성 마전 PA", line, diameter: "200A", date });
const aj = (point: string, result: string, date: string) => rec(point, result, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: AJ, diameter: "750A", date });

/** 섞인 묶음(여러 현장·날짜, 불합격·재검사 포함). 입력 순서는 일부러 뒤섞는다. ts(입력 순)는 재검사 원본이 먼저. */
function mixedBatch() {
  const recs = {
    hjS29F: hj("S29", FAIL, "2026-11-12"),
    gpS08F: gp("S08", FAIL, "2026-11-12"),
    mjR04F: mj("R04", FAIL, "2026-11-12"),
  };
  const jobs = [
    hj("S01", PASS, "2026-11-13", HJ6),
    gp("R07", PASS, "2026-11-12"),
    mj("S04", PASS, "2026-11-12"),
    hj("R28", PASS, "2026-11-12"),
    recs.hjS29F,
    aj("S17", PASS, "2026-11-12"),
    gp("S09", PASS, "2026-11-13"),
    recs.gpS08F,
    hj("S02", PASS, "2026-11-12", HJ6),
    hj("S28", PASS, "2026-11-12"),
    gp("S07", PASS, "2026-11-12"),
    mj("S01", PASS, "2026-11-12", MJ4),
    recs.mjR04F,
    hj("S29", PASS, "2026-11-12"), // 재검사: 같은 번호, 새 행(불합격 원본 바로 뒤)
    gp("S08", PASS, "2026-11-12"), // 재검사
  ];
  return { jobs, ...recs };
}

/** 시트별로 기대하는 새 행(행 번호 → 날짜·라인·포인트·합부). 마지막 행: 합정 439, 군포 133, 마전 313, 아전 39. */
const EXPECTED: Record<string, { book: ListBookKind; sheet: string; rows: [number, string, string, string, string][] }> = {
  hj: {
    book: "list",
    sheet: " 드림 합정주공",
    rows: [
      [440, "2026-11-12", HJ5, "S28", PASS],
      [441, "2026-11-12", HJ5, "S29", FAIL],
      [442, "2026-11-12", HJ5, "S29", PASS],
      [443, "2026-11-12", HJ5, "R28", PASS],
      [444, "2026-11-12", HJ6, "S02", PASS],
      [445, "2026-11-13", HJ6, "S01", PASS],
    ],
  },
  gp: {
    book: "listGm",
    sheet: "지성 군포 PA",
    rows: [
      [134, "2026-11-12", GP, "S07", PASS],
      [135, "2026-11-12", GP, "S08", FAIL],
      [136, "2026-11-12", GP, "S08", PASS],
      [137, "2026-11-12", GP, "R07", PASS],
      [138, "2026-11-13", GP, "S09", PASS],
    ],
  },
  mj: {
    book: "listGm",
    sheet: "지성 마전 PA",
    rows: [
      [314, "2026-11-12", MJ4, "S01", PASS],
      [315, "2026-11-12", MJ5, "S04", PASS],
      [316, "2026-11-12", MJ5, "R04", FAIL],
    ],
  },
  aj: { book: "listGm", sheet: "아전 군포 PA", rows: [[40, "2026-11-12", AJ, "S17", PASS]] },
};

type Cell = { v?: string; f?: string };
function cellOf(files: ZipFiles, sheet: string, ref: string): Cell {
  return (readSheet(files, sheet).cells.get(ref) as Cell | undefined) ?? {};
}

function lastLineRow(files: ZipFiles, sheet: string): number {
  let max = 5;
  for (const c of readSheet(files, sheet).cells.values()) if (c.col === "F" && c.row >= 6 && c.v) max = Math.max(max, c.row);
  return max;
}

/** 앱 행(Y=앱 ID)만, 행 순서대로 */
function appRows(files: ZipFiles, book: ListBookKind): InspectionRecord[] {
  return parseListWorkbook(files, book)
    .flatMap((s) => s.records)
    .filter((r) => r.src === "user")
    .sort((a, b) => a.sheetName.localeCompare(b.sheetName) || a.excelRow! - b.excelRow!);
}

/** 두 리스트 파일에 대해 '빠짐없이·중복 없이·맞는 파일·시트·순서·열'을 한꺼번에 확인한다. */
function expectListsExact(books: Record<ListBookKind, ZipFiles>, jobs: InspectionRecord[], before: Record<ListBookKind, InspectionRecord[]>) {
  for (const book of ["list", "listGm"] as const) {
    const files = books[book];
    const rows = appRows(files, book);
    const ids = rows.map((r) => r.id);
    // 중복 없음: 앱 ID 마다 한 행
    expect(new Set(ids).size).toBe(ids.length);
    // 빠짐없음: 이 파일로 갈 기록은 모두 있고, 다른 파일 기록은 하나도 없다
    const mine = jobs.filter((r) => listBookOf(r) === book);
    const other = jobs.filter((r) => listBookOf(r) !== book);
    for (const r of mine) expect(ids).toContain(r.id);
    for (const r of other) expect(ids).not.toContain(r.id);
    // 원래 있던 앱 행은 그대로(개수·자리)
    const prev = before[book].map((r) => `${r.sheetName}:${r.excelRow}:${r.id}`);
    for (const p of prev) expect(rows.map((r) => `${r.sheetName}:${r.excelRow}:${r.id}`)).toContain(p);
    expect(rows).toHaveLength(prev.length + mine.length);
    // '앱' 표시는 어디에도 없다
    for (const info of workbookSheets(files)) expect(xmlText(files, info.path)).not.toContain("앱");
  }
  for (const exp of Object.values(EXPECTED)) {
    const files = books[exp.book];
    expect(lastLineRow(files, exp.sheet)).toBe(exp.rows[exp.rows.length - 1][0]);
    for (const [row, date, line, point, result] of exp.rows) {
      const c = (col: string) => cellOf(files, exp.sheet, `${col}${row}`);
      const job = jobs.find((r) => r.sheetName === exp.sheet && r.date === date && r.line === line && r.point === point && r.result === result)!;
      expect(job, `${exp.sheet} ${row}`).toBeTruthy();
      expect({ row, E: c("E").v, F: c("F").v, G: c("G").v, K: c("K").v, P: c("P").v, S: c("S").v ?? "", Y: c("Y").v }).toEqual({
        row,
        E: String(isoToExcelSerial(date)),
        F: line,
        G: point,
        K: "SPPS380",
        P: result,
        S: "",
        Y: job.id,
      });
      const l = c("L");
      if (exp.book === "listGm") {
        // 군포·마전: 관경은 값, 순번(A)은 숫자
        expect(l.f).toBeUndefined();
        expect(l.v).toBe(/(\d{2,4}A)/.exec(line)![1]);
        expect(c("A").v).toMatch(/^\d+$/);
      } else {
        expect(l.f?.replace(/^=/, "")).toBe(`RIGHT(F${row},4)`);
      }
    }
    // 행 XML: 셀이 열 순서대로(엑셀 '복구' 없음), 검사일(E)은 바로 위 기존 행과 같은 날짜 서식
    const xml = readSheet(files, exp.sheet).xml;
    const styleOf = (ref: string) => new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>)`).exec(xml)?.[1].match(/\bs="(\d+)"/)?.[1];
    const firstNew = exp.rows[0][0];
    for (const [row] of exp.rows) {
      const rowXml = new RegExp(`<row r="${row}"[^>]*>([\\s\\S]*?)</row>`).exec(xml)![1];
      const cols = [...rowXml.matchAll(/<c r="([A-Z]+)\d+"/g)].map((x) => x[1]);
      const idx = (c: string) => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      expect(cols.map(idx), `${exp.sheet} ${row} 열 순서`).toEqual([...cols.map(idx)].sort((a, b) => a - b));
      expect(styleOf(`E${row}`), `${exp.sheet} E${row} 서식`).toBe(styleOf(`E${firstNew - 1}`));
    }
    // 새 행끼리 날짜 → 라인 → 포인트 순
    const seq = exp.rows.map(([, date, line, point]) => ({ date, line, point }));
    for (let i = 1; i < seq.length; i++) {
      const a = seq[i - 1];
      const b = seq[i];
      const cmp = a.date.localeCompare(b.date) || a.line.localeCompare(b.line, "ko", { numeric: true }) || comparePointLabel(a.point, b.point);
      expect(cmp, `${exp.sheet} ${a.point}→${b.point}`).toBeLessThanOrEqual(0);
    }
  }
  // 군포·마전 순번은 이어진다(앞 행 + 1)
  const gm = books.listGm;
  for (const exp of Object.values(EXPECTED).filter((e) => e.book === "listGm")) {
    const first = exp.rows[0][0];
    const seqs = exp.rows.map(([row]) => Number(cellOf(gm, exp.sheet, `A${row}`).v));
    for (let i = 1; i < seqs.length; i++) expect(seqs[i]).toBe(seqs[i - 1] + 1);
    const prevSeq = Number(cellOf(gm, exp.sheet, `A${first - 1}`).v);
    if (Number.isFinite(prevSeq) && prevSeq > 0) expect(seqs[0]).toBe(prevSeq + 1);
  }
}

function snapshotBefore(): Record<ListBookKind, InspectionRecord[]> {
  return {
    list: appRows(unzipXlsx(fixture("real-list.xlsx")), "list"),
    listGm: appRows(unzipXlsx(fixture("real-list-gunpo-majeon.xlsx")), "listGm"),
  };
}

describe("검사리스트 끝까지(OneDrive·Graph): 실제 두 리스트", () => {
  it("섞인 묶음: 빠짐·중복 없이 맞는 파일·시트에, 날짜→라인→포인트 순, 맞는 열로. 다시 보내도 행이 늘지 않는다", async () => {
    const drive = realDrive();
    const before = snapshotBefore();
    const { jobs, hjS29F, gpS08F, mjR04F } = mixedBatch();
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    const out = await sync(drive, jobs, all);

    expect(out.report.lists).toMatchObject({ list: { status: "ok", count: 6 }, listGm: { status: "ok", count: 9 } });
    for (const r of jobs) expect(out.steps.get(r.id)?.list?.status).toBe("done");
    expectListsExact({ list: drive.files(LIST_ID), listGm: drive.files(LIST_GM_ID) }, jobs, before);
    // 엑셀 행 번호를 기록에도 남긴다(다음에 같은 행을 고치는 데 씀)
    expect(hjS29F.excelRow).toBe(441);
    expect(gpS08F.excelRow).toBe(135);
    expect(mjR04F.excelRow).toBe(316);

    // 불합격: 리스트에는 있고, 작업일지·거래명세서에는 세지 않는다
    const sum = (g: { count: number }[] | null) => (g ?? []).reduce((n, x) => n + x.count, 0);
    const ilji = drive.files(ILJI_ID);
    expect(sum(readIljiDay(ilji, "2026-11-12", "드림 합정주공"))).toBe(4); // S28 S29(재검 합격) R28 S02 (S29 불합격 제외)
    expect(sum(readIljiDay(ilji, "2026-11-12", "지성 군포"))).toBe(3); // S07 S08(재검 합격) R07
    expect(sum(readIljiDay(ilji, "2026-11-12", "지성 마전"))).toBe(2); // S01 S04 (R04 불합격 제외)
    const hjNov = readPautRows(drive.files(STMT_ID), "드림합정26.11").find((r) => r.date === "2026-11-12");
    expect(hjNov?.count).toBe(4);
    const gpNov = readPautRows(drive.files(STMT_ID), "지성군포 2611").find((r) => r.date === "2026-11-12");
    expect(gpNov?.count).toBe(3);

    // 리스트 쓰기는 행 하나씩 A:Y 범위, 통합 문서 세션 안. 파일 업로드(PUT) 없음.
    const listPatches = drive.requests((e) => (e.item === LIST_ID || e.item === LIST_GM_ID) && e.method === "PATCH");
    expect(listPatches).toHaveLength(jobs.length);
    expect(listPatches.every((e) => !!e.session && /range\(address='A\d+:Y\d+'\)/.test(e.path))).toBe(true);
    expect(JSON.stringify(listPatches.map((e) => e.body))).not.toContain("앱");
    expect(drive.requests((e) => e.method === "PUT")).toHaveLength(0);

    // 두 번째 동기화(같은 기록을 다시 보냄, 예: 응답을 못 받아 다시 시도) → Y 앱 ID 로 같은 행을 고친다. 행이 늘지 않는다.
    for (const r of jobs) {
      r.steps = undefined;
      r.excelRow = undefined;
    }
    await sync(drive, jobs, all);
    expectListsExact({ list: drive.files(LIST_ID), listGm: drive.files(LIST_GM_ID) }, jobs, before);

    // 다음 날 묶음: 마지막 행 아래에 이어 붙고, 앞 묶음은 그대로
    const next = [gp("S11", PASS, "2026-11-14"), gp("S10", PASS, "2026-11-14"), hj("S03", PASS, "2026-11-14", HJ6), mj("S06", FAIL, "2026-11-14")];
    const all2 = [...all, ...next];
    await sync(drive, next, all2);
    const gm = drive.files(LIST_GM_ID);
    const py = drive.files(LIST_ID);
    expect([139, 140].map((r) => cellOf(gm, "지성 군포 PA", `G${r}`).v)).toEqual(["S10", "S11"]);
    expect(cellOf(gm, "지성 마전 PA", "G317").v).toBe("S06");
    expect(cellOf(gm, "지성 마전 PA", "P317").v).toBe(FAIL);
    expect(cellOf(py, " 드림 합정주공", "G446").v).toBe("S03");
    expect(lastLineRow(gm, "지성 군포 PA")).toBe(140);
    expect(lastLineRow(py, " 드림 합정주공")).toBe(446);
    const ids = [...appRows(gm, "listGm"), ...appRows(py, "list")].map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const r of [...jobs, ...next]) expect(ids.filter((id) => id === r.id)).toHaveLength(1);
  });

  it("이미 리스트에 있는 앱 행(11/11 S27, Y=앱 ID)을 다시 보내면 그 자리(439)를 고친다. 새 행이 생기지 않는다", async () => {
    const drive = realDrive();
    const existing = realListRecords().find((r) => r.src === "user" && r.excelRow === 439)!;
    expect(existing.point).toBe("S27");
    existing.steps = undefined;
    existing.sync = "pending";
    existing.src = "user";
    const fresh = hj("S28", PASS, "2026-11-12");
    const all = [...realListRecords().filter((r) => r.id !== existing.id), existing, fresh];
    await sync(drive, [existing, fresh], all);
    const py = drive.files(LIST_ID);
    expect(cellOf(py, " 드림 합정주공", "Y439").v).toBe(existing.id);
    expect(cellOf(py, " 드림 합정주공", "G440").v).toBe("S28");
    expect(lastLineRow(py, " 드림 합정주공")).toBe(440);
    expect(appRows(py, "list").filter((r) => r.id === existing.id)).toHaveLength(1);
  });
});

describe("검사리스트 끝까지(데모·이 기기 사본): 같은 규칙", () => {
  it("섞인 묶음을 이 기기 사본에 써도 행 번호·순서·열이 OneDrive 와 같다. 다시 써도 행이 늘지 않는다", () => {
    bootstrapFromSamples({
      list: fixture("real-list.xlsx"),
      ilji: fixture("real-worklog.xlsx"),
      statement: fixture("real-statement.xlsx"),
      listGm: fixture("real-list-gunpo-majeon.xlsx"),
    });
    const before = snapshotBefore();
    const { jobs } = mixedBatch();
    const all = [...realListRecords(), ...realGmRecords(), ...jobs];
    syncPendingDemo(jobs, all, graphConfig(), meta());
    expectListsExact({ list: unzipXlsx(exportBook("list")), listGm: unzipXlsx(exportBook("listGm")) }, jobs, before);
    syncPendingDemo(jobs, all, graphConfig(), meta());
    expectListsExact({ list: unzipXlsx(exportBook("list")), listGm: unzipXlsx(exportBook("listGm")) }, jobs, before);
  });
});
