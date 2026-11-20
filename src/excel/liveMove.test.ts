/**
 * 리스트 탭 「수정」 행 순서 바꾸기: 가짜 OneDrive(Graph)에서 옮기는 두 행(B–Y)만 다시 쓰고, A·다른 행은 그대로인지.
 * 실제 OneDrive 는 부르지 않는다.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
vi.mock("../store/db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => void mem.set(key, structuredClone(value)),
  kvDel: async (key: string) => void mem.delete(key),
}));
vi.stubGlobal("window", { setTimeout, clearTimeout });
vi.stubGlobal("localStorage", {
  getItem: (k: string) => bag.get(k) ?? null,
  setItem: (k: string, v: string) => void bag.set(k, String(v)),
  removeItem: (k: string) => void bag.delete(k),
});

import { GraphClient, graphRetry, worksheetRangeUrl } from "../sync/graph";
import type { FakeDrive, LogEntry } from "../test/fakeGraph";
import { LIST_GM_ID, LIST_ID, fixture, graphConfig, realDrive, realGmRecords, realListRecords } from "../test/realBooks";
import { cellText, loadLiveList, matchRecord, type LiveRow, type LiveTable } from "./liveList";
import { LiveMoveRefused, moveLiveRow, moveOrder, restoreLiveRows } from "./liveMove";
import { LiveConflictError } from "./liveList";
import { readSheet, setXml, unzipXlsx, workbookSheets, xmlText } from "./ooxml";
import { exportBook, rememberWork } from "./engine";
import { getState, pullFromList } from "../store/session";
import { applyRowMove, moveConfirmText } from "../ui/ListScreen";
import { summarizeByDate } from "../ui/listLayout";

const MAJEON = "지성 마전 PA";
const HAPJEONG = "드림 합정주공";
let drive: FakeDrive;
const client = () => new GraphClient({ getToken: async () => drive.validToken, fetchImpl: drive.fetch, sleep: async () => undefined });
const writes = () => drive.log.filter((e) => e.method !== "GET");
const text = (row: LiveRow, col: string) => cellText(row.cells.find((c) => c.col === col)!);

function seed(): void {
  getState().records.splice(0, getState().records.length, ...realListRecords(), ...realGmRecords());
}

async function load(site: string): Promise<LiveTable> {
  const res = await loadLiveList(client(), graphConfig(), site, getState().records);
  if (!res.ok) throw new Error(res.message);
  return res.table;
}

/** 쓰기 요청의 범위 주소(배치 안 요청 포함). */
function writtenRanges(log: LogEntry[]): string[] {
  return log.filter((e) => e.method !== "GET" && e.path.includes("/range(")).map((e) => /range\(address='([^']+)'\)/.exec(e.path)![1]);
}

function rowsOf(address: string): [number, number] {
  const [a, b] = address.split(":").map((x) => Number(/\d+$/.exec(x)![0]));
  return [a, b ?? a];
}

beforeEach(() => {
  mem.clear();
  bag.clear();
  graphRetry.baseMs = 0;
  drive = realDrive();
});

describe("moveOrder", () => {
  it("from 을 to 자리로, 사이 행은 한 칸씩", () => {
    expect(moveOrder(152, 151)).toEqual({ lo: 151, hi: 152, sources: [152, 151] });
    expect(moveOrder(151, 152)).toEqual({ lo: 151, hi: 152, sources: [152, 151] });
    expect(moveOrder(150, 153)).toEqual({ lo: 150, hi: 153, sources: [151, 152, 153, 150] });
  });
});

describe("리스트 탭 행 순서 바꾸기 · 엑셀 두 행만 다시 씀", () => {
  it("마전 152 ↔ 151: 확인을 취소하면 아무것도 쓰지 않는다", async () => {
    seed();
    const table = await load(MAJEON);
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    const asked: string[] = [];
    drive.log.length = 0;
    const res = await applyRowMove(client(), table, MAJEON, row, up, (t) => (asked.push(t), false));
    expect(res).toEqual({ done: false, note: "" });
    expect(asked).toEqual([moveConfirmText(row, up)]);
    expect(asked[0]).toMatch(/^행 152 ↔ 151 순서를 바꿀까요\? /);
    expect(writes()).toEqual([]);
  });

  it("위로: B–Y 값·서식이 맞바뀌고 A(순번)는 그대로, 151–152 밖·행 넣기/지우기 요청 없음, 카드 excelRow 도 새 행", async () => {
    seed();
    const table = await load(MAJEON);
    const before = new Map(table.rows.map((r) => [r.row, r.cells.map((c) => c.value)]));
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    const recRow = matchRecord(getState().records, table, row)!;
    const recUp = matchRecord(getState().records, table, up)!;
    expect(recRow.id).toBe(`xlsx:${MAJEON}:152`);
    const sheet = (await client().worksheetsOf(LIST_GM_ID)).find((s) => s.name === MAJEON)!;
    const c = client();
    await c.request(worksheetRangeUrl(LIST_GM_ID, sheet.id, "G152") + "/format/fill", { method: "PATCH", body: JSON.stringify({ color: "#FFFF00" }) });
    await c.request(worksheetRangeUrl(LIST_GM_ID, sheet.id, "B152:Y152") + "/format", { method: "PATCH", body: JSON.stringify({ rowHeight: 31 }) });
    await c.request(worksheetRangeUrl(LIST_GM_ID, sheet.id, "A152") + "/format/fill", { method: "PATCH", body: JSON.stringify({ color: "#00FF00" }) });
    rememberWork("listGm", fixture("real-list-gunpo-majeon.xlsx"));

    drive.log.length = 0;
    const res = await applyRowMove(client(), table, MAJEON, row, up, () => true);
    expect(res.done).toBe(true);
    if (!res.done) return;
    expect(res.snapshot).toBeNull();
    expect(res.note).toBe("행 152 ↔ 151 순서 바꿈 · 엑셀 다시 읽어 확인함 · 카드 2개 행 번호 맞춤");

    const ranges = writtenRanges(drive.log);
    expect(ranges.length).toBeGreaterThan(0);
    for (const address of ranges) {
      const [a, b] = rowsOf(address);
      expect(a).toBeGreaterThanOrEqual(151);
      expect(b).toBeLessThanOrEqual(152);
      expect(address.startsWith("A")).toBe(false);
    }
    expect(drive.log.some((e) => /\/insert|\/delete|\/worksheets\/add|\/copy|\/content/.test(e.path) && e.method !== "GET")).toBe(false);
    expect(drive.log.filter((e) => e.path === "/$batch").length).toBeLessThan(20);

    const after = await load(MAJEON);
    for (const r of after.rows) {
      const was = before.get(r.row)!;
      const now = r.cells.map((x) => x.value);
      if (r.row === 151) expect(now).toEqual([was[0], ...before.get(152)!.slice(1)]);
      else if (r.row === 152) expect(now).toEqual([was[0], ...before.get(151)!.slice(1)]);
      else expect(now).toEqual(was);
    }
    expect(text(after.rows.find((r) => r.row === 151)!, "G")).toBe(text(row, "G"));
    expect(drive.cellFormat(LIST_GM_ID, MAJEON, "G151").fill).toBe("#FFFF00");
    expect(drive.cellFormat(LIST_GM_ID, MAJEON, "G152").fill).not.toBe("#FFFF00");
    expect(drive.rowHeight(LIST_GM_ID, MAJEON, 151)).toBe(31);
    expect(drive.rowHeight(LIST_GM_ID, MAJEON, 152)).not.toBe(31);
    expect(drive.cellFormat(LIST_GM_ID, MAJEON, "A152").fill).toBe("#00FF00");
    expect(drive.cellFormat(LIST_GM_ID, MAJEON, "A151").fill).not.toBe("#00FF00");

    const recs = getState().records;
    expect(recs.find((r) => r.id === recRow.id)!.excelRow).toBe(151);
    expect(recs.find((r) => r.id === recUp.id)!.excelRow).toBe(152);
    expect(matchRecord(recs, after, after.rows.find((r) => r.row === 151)!)!.id).toBe(recRow.id);
    expect(matchRecord(recs, after, after.rows.find((r) => r.row === 152)!)!.id).toBe(recUp.id);

    const local = readSheet(unzipXlsx(exportBook("listGm")), MAJEON).cells;
    expect(local.get("G151")?.v).toBe(text(row, "G"));
    expect(local.get("G152")?.v).toBe(text(up, "G"));
    expect(local.get("A151")?.v).toBe(String(before.get(151)![0]));
  });

  it("수식·앱 ID(Y) 행(합정주공): 수식은 새 행을 가리키게(R1C1), Y 는 행과 함께 옮겨지고 카드는 Y 로 맞는다", async () => {
    seed();
    const table = await load(HAPJEONG);
    const at = table.rows.findIndex((r, i) => i > 0 && text(r, "Y") && table.rows[i - 1].row === r.row - 1 && !text(table.rows[i - 1], "Y"));
    const row = table.rows[at];
    const up = table.rows[at - 1];
    expect(row.cells.find((c) => c.col === "L")!.formula).toBe(true);
    const appId = text(row, "Y");
    const rec = matchRecord(getState().records, table, row)!;
    expect(rec.id).toBe(appId);

    rememberWork("list", fixture("real-list.xlsx"));
    const res = await applyRowMove(client(), table, HAPJEONG, row, up, () => true);
    expect(res.done && res.snapshot === null).toBe(true);
    const local = readSheet(unzipXlsx(exportBook("list")), workbookSheets(drive.files(LIST_ID)).find((s) => s.name.trim() === HAPJEONG)!.name).cells;
    expect(local.get(`Y${up.row}`)?.v).toBe(appId);
    expect(local.get(`G${up.row}`)?.v).toBe(text(row, "G"));
    expect(local.get(`L${up.row}`)?.f).toBe(`RIGHT(F${up.row},4)`);
    expect(local.get(`A${up.row}`)?.v).toBe(String(up.cells[0].value));
    const sheet = workbookSheets(drive.files(LIST_ID)).find((s) => s.name.trim() === HAPJEONG)!;
    const cells = readSheet(drive.files(LIST_ID), sheet.name).cells;
    expect(cells.get(`Y${up.row}`)?.v).toBe(appId);
    expect(cells.get(`Y${row.row}`)?.v ?? "").toBe(text(up, "Y"));
    expect(cells.get(`L${up.row}`)?.f).toBe(`RIGHT(F${up.row},4)`);
    expect(cells.get(`L${row.row}`)?.f).toBe(up.cells.find((c) => c.col === "L")!.formula ? `RIGHT(F${row.row},4)` : undefined);
    expect(cells.get(`L${row.row}`)?.v ?? "").toBe(text(up, "L"));
    expect(cells.get(`A${up.row}`)?.v).toBe(String(up.cells[0].value));
    expect(getState().records.find((r) => r.id === appId)!.excelRow).toBe(up.row);

    getState().records.find((r) => r.id === appId)!.excelRow = row.row;
    await pullFromList(await load(HAPJEONG), HAPJEONG);
    expect(getState().records.find((r) => r.id === appId)!.excelRow).toBe(up.row);
  });

  it("쓰기 직전에 다시 읽어 다른 곳에서 바뀌었으면 쓰지 않는다", async () => {
    seed();
    const table = await load(MAJEON);
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    const sheet = (await client().worksheetsOf(LIST_GM_ID)).find((s) => s.name === MAJEON)!;
    await client().patchValues(LIST_GM_ID, sheet.id, "S151", [["PC에서 고침"]]);
    drive.log.length = 0;
    await expect(applyRowMove(client(), table, MAJEON, row, up, () => true)).rejects.toBeInstanceOf(LiveConflictError);
    expect(writes().filter((e) => e.path !== "/$batch" && !/createSession|closeSession/.test(e.path))).toEqual([]);
    expect(getState().records.find((r) => r.id === `xlsx:${MAJEON}:152`)!.excelRow).toBe(152);
  });

  it("쓰기 뒤 확인이 다르면 다른 칸을 알리고, 되돌리기로 옮기기 전 모습을 다시 쓴다", async () => {
    seed();
    const table = await load(MAJEON);
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    const before = [up, row].map((r) => r.cells.map((c) => c.value));
    drive.fail((m, p) => m === "PATCH" && p.includes("range(address='B151:Y152')") && !p.includes("/format"), 500, 1);
    let calls = 0;
    const res = await applyRowMove(client(), table, MAJEON, row, up, () => (calls++, true));
    expect(calls).toBe(1);
    expect(res.done).toBe(true);
    if (!res.done) return;
    expect(res.snapshot).not.toBeNull();
    expect(res.note).toMatch(/^행 152 ↔ 151 쓰기 뒤 확인이 맞지 않습니다: 행 순서 쓰기 실패 .*, 151행\(.*G.*\), 152행\(.*G.*\)\. 아래 「되돌리기」/);
    expect(getState().records.find((r) => r.id === `xlsx:${MAJEON}:152`)!.excelRow).toBe(152);

    const restored = await restoreLiveRows(client(), table, res.snapshot!);
    expect(restored.mismatches).toEqual([]);
    const again = await load(MAJEON);
    expect([151, 152].map((n) => again.rows.find((r) => r.row === n)!.cells.map((c) => c.value))).toEqual(before);
  });

  it("병합 칸·메모·Z 열 오른쪽 값이 있는 행, 빈 행을 건너는 이동, 제외 현장, 엑셀에 다 안 보낸 카드는 옮기지 않는다", async () => {
    seed();
    const table = await load(MAJEON);
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    const files = drive.files(LIST_GM_ID);
    const sheet = workbookSheets(files).find((s) => s.name === MAJEON)!;
    const original = xmlText(files, sheet.path);

    setXml(files, sheet.path, original.replace("</sheetData>", '</sheetData><mergeCells count="1"><mergeCell ref="T151:W151"/></mergeCells>'));
    drive.log.length = 0;
    await expect(moveLiveRow(client(), table, 152, 151)).rejects.toThrow(/병합 칸 T151:W151/);
    expect(writes()).toEqual([]);
    setXml(files, sheet.path, original);

    const n = sheet.path.match(/sheet(\d+)\.xml$/)![1];
    files[`xl/worksheets/_rels/sheet${n}.xml.rels`] = new TextEncoder().encode(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments9.xml"/></Relationships>',
    );
    files["xl/comments9.xml"] = new TextEncoder().encode('<?xml version="1.0"?><comments><commentList><comment ref="W152"><text><t>메모</t></text></comment></commentList></comments>');
    await expect(moveLiveRow(client(), table, 152, 151)).rejects.toThrow(/메모 W152/);
    await expect(moveLiveRow(client(), table, 160, 161)).resolves.toMatchObject({ mismatches: [] });
    delete files["xl/comments9.xml"];

    const sheetId = (await client().worksheetsOf(LIST_GM_ID)).find((s) => s.name === MAJEON)!.id;
    await client().patchValues(LIST_GM_ID, sheetId, "AB151", [["옆 메모"]]);
    drive.log.length = 0;
    await expect(moveLiveRow(client(), table, 152, 151)).rejects.toThrow(/Y 열 오른쪽 값\(AB151/);
    await expect(moveLiveRow(client(), table, 152, 151)).rejects.toBeInstanceOf(LiveMoveRefused);
    expect(writes()).toEqual([]);

    const gapped = { ...table, rows: table.rows.filter((r) => r.row !== 171) };
    await expect(moveLiveRow(client(), gapped, 170, 172)).rejects.toThrow(/171행이 비어/);

    let asked = 0;
    expect(await applyRowMove(client(), table, "드림 택지", row, up, () => (asked++, true))).toMatchObject({ done: false, note: expect.stringMatching(/제외된 현장/) });
    getState().records.find((r) => r.id === `xlsx:${MAJEON}:151`)!.sync = "pending";
    expect(await applyRowMove(client(), table, MAJEON, row, up, () => (asked++, true))).toMatchObject({ done: false, note: expect.stringMatching(/엑셀에 아직 다 보내지 않은 카드/) });
    expect(asked).toBe(0);
  });
});

describe("리스트 탭 표 · 날짜 안 행 순서", () => {
  it("한 날짜 안의 행은 시트 행 번호 순서 그대로(포인트로 다시 정렬하지 않음), 옮긴 뒤에도 새 행 순서", async () => {
    seed();
    const table = await load(MAJEON);
    const day = summarizeByDate(table.rows).find((g) => g.date === "2026-06-18")!;
    const rows = day.rows.map((r) => r.row);
    expect(rows).toEqual([...rows].sort((a, b) => a - b));
    const row = table.rows.find((r) => r.row === 152)!;
    const up = table.rows.find((r) => r.row === 151)!;
    await applyRowMove(client(), table, MAJEON, row, up, () => true);
    const after = summarizeByDate((await load(MAJEON)).rows).find((g) => g.date === "2026-06-18")!;
    const points = after.rows.map((r) => text(r, "G"));
    expect(points.indexOf(text(row, "G"))).toBe(points.indexOf(text(up, "G")) - 1);
    expect(after.rows.map((r) => r.row)).toEqual(rows);
  });
});
