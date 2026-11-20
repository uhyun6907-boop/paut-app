/**
 * 리스트 탭: 기기 캐시(먼저 보이고 뒤에서 새로 읽기)·빠른 읽기(실제 기록 끝 행만, 묶음 동시 읽기)·열 필터·두 손가락 확대.
 * 가짜 OneDrive(Graph)와 로컬 사본만 쓴다. 실제 네트워크는 부르지 않는다.
 */
import { readFileSync } from "node:fs";
import { beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
vi.mock("../store/db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => void mem.set(key, structuredClone(value)),
  kvDel: async (key: string) => void mem.delete(key),
}));
vi.stubGlobal("localStorage", {
  getItem: (k: string) => bag.get(k) ?? null,
  setItem: (k: string, v: string) => void bag.set(k, String(v)),
  removeItem: (k: string) => void bag.delete(k),
});

import { GraphClient, graphRetry } from "../sync/graph";
import type { FakeDrive } from "../test/fakeGraph";
import { LIST_ID, graphConfig, realDrive, realGmRecords, realListRecords } from "../test/realBooks";
import { cellText, lastDataRow, LIVE_CHUNK_ROWS, LIVE_PARALLEL, loadLiveList, writeLiveCell, type LiveTable } from "../excel/liveList";
import { readSheet } from "../excel/ooxml";
import { packTable, unpackTable } from "../store/listCache";
import { CACHE_STALE_TEXT, loadListCache, mergeTables, openList, type CacheSlot, type ListView } from "./listLoad";
import { BLANK, activeColumns, applyFilters, clearColumn, columnValues, countText, filterableColumns, setText, toggleValue, type ListFilters } from "./listFilter";
import { getState, liveListClient, settled } from "../store/session";
import { KEY_COLS, MODE_LABEL, shortLine, summarizeByDate, toggleDate, toggleEditing, toggleRow, visibleColumns, type DateFocus, type DateGroup } from "./listLayout";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { matchRecord } from "../excel/liveList";
import { DateSummaryCard, applyRowAction, deleteConfirmText } from "./ListScreen";

const SITE = "드림 합정주공";
let drive: FakeDrive;
const client = () => new GraphClient({ getToken: async () => drive.validToken, fetchImpl: drive.fetch, sleep: async () => undefined });
const writes = () => drive.log.filter((e) => e.method !== "GET");
const slot: CacheSlot = { book: "list", fileId: LIST_ID, sheetName: SITE };

async function fresh(): Promise<LiveTable> {
  const res = await loadLiveList(client(), graphConfig(), SITE, realListRecords());
  if (!res.ok) throw new Error(res.message);
  return res.table;
}

/** openList 를 돌리며 화면에 보인 상태를 차례로 모은다. */
async function open(read = () => loadLiveList(client(), graphConfig(), SITE, realListRecords())) {
  const shown: ListView[] = [];
  const end = await openList({ slot, read, show: (v) => shown.push(v) });
  return { shown, end };
}

beforeEach(() => {
  mem.clear();
  bag.clear();
  graphRetry.baseMs = 0;
  drive = realDrive();
});

describe("리스트 캐시 · 먼저 보이고 뒤에서 새로 읽기", () => {
  it("캐시가 없으면 처음만 '읽는 중', 읽은 뒤 그 현장 캐시를 저장한다", async () => {
    const { shown, end } = await open();
    expect(shown[0]).toEqual({ status: "loading" });
    expect(end.status).toBe("ok");
    const cached = await loadListCache(slot);
    expect(cached).not.toBeNull();
    if (end.status !== "ok") return;
    expect(cached!.table.rows.length).toBe(end.table.rows.length);
    expect(writes()).toEqual([]);
  });

  it("캐시가 있으면 먼저 캐시를 보여 주고(읽는 중 없음) 새로 읽으면 바뀐 행만 바꾼다, 캐시도 덮어쓴다", async () => {
    await open();
    // PC 엑셀에서 한 칸을 고친 상황
    const base = await fresh();
    const target = base.rows[10];
    const c = client();
    await c.patchValues(LIST_ID, base.sheetId, `Q${target.row}`, [["PC에서 고침"]]);
    drive.log.length = 0;

    const { shown, end } = await open();
    expect(shown.some((v) => v.status === "loading")).toBe(false);
    const first = shown[0];
    expect(first.status).toBe("ok");
    if (first.status !== "ok" || end.status !== "ok") return;
    expect(first.fromCache).toBe(true);
    expect(first.refreshing).toBe(true);
    expect(cellText(first.table.rows.find((r) => r.row === target.row)!.cells.find((x) => x.col === "Q")!)).not.toBe("PC에서 고침");

    expect(end.refreshing).toBe(false);
    expect(end.fromCache).toBe(false);
    const changed = end.table.rows.find((r) => r.row === target.row)!;
    expect(cellText(changed.cells.find((x) => x.col === "Q")!)).toBe("PC에서 고침");
    // 바뀌지 않은 행은 캐시의 행 객체 그대로(다시 그리지 않음), 바뀐 행만 새 객체.
    const same = end.table.rows.filter((r) => first.table.rows.includes(r)).length;
    expect(same).toBe(end.table.rows.length - 1);
    expect(first.table.rows.includes(changed)).toBe(false);

    const again = await loadListCache(slot);
    expect(cellText(again!.table.rows.find((r) => r.row === target.row)!.cells.find((x) => x.col === "Q")!)).toBe("PC에서 고침");
    expect(writes()).toEqual([]);
  });

  it("새로 읽기에 실패하면 캐시를 그대로 보이고 짧은 안내만(오류 화면 아님)", async () => {
    await open();
    const before = await loadListCache(slot);
    drive.fail((m) => m === "GET", 503, Infinity);
    const { shown, end } = await open();
    expect(shown.every((v) => v.status === "ok")).toBe(true);
    expect(end.status).toBe("ok");
    if (end.status !== "ok") return;
    expect(end.fromCache).toBe(true);
    expect(end.refreshing).toBe(false);
    expect(end.staleError).toBeTruthy();
    expect(end.table.rows.length).toBe(before!.table.rows.length);
    expect(CACHE_STALE_TEXT).toBe("캐시 표시 중 · 갱신 실패");
    // 실패해도 캐시는 지우지 않는다.
    expect((await loadListCache(slot))!.table.rows.length).toBe(before!.table.rows.length);
  });

  it("캐시도 없고 읽기도 실패하면 그때만 오류 화면", async () => {
    const { end } = await open(async () => ({ ok: false as const, message: "검사리스트를 읽지 못했습니다 — 테스트" }));
    expect(end).toEqual({ status: "error", message: "검사리스트를 읽지 못했습니다 — 테스트" });
  });

  it("캐시는 현장(파일·시트)마다 따로, 값·수식 표시가 그대로 돌아온다", async () => {
    const table = await fresh();
    const back = unpackTable(structuredClone(packTable(table)));
    expect(back.rows).toEqual(table.rows);
    expect(back.header).toEqual(table.header);
    expect(back.rows.some((r) => r.cells.some((c) => c.formula))).toBe(true);
    await open();
    expect(await loadListCache({ ...slot, sheetName: "지성 브레인" })).toBeNull();
    expect(await loadListCache({ ...slot, fileId: "OTHER" })).toBeNull();
  });

  it("mergeTables: 행이 늘거나 줄면 새 표의 행 순서·수 그대로", async () => {
    const table = await fresh();
    const fewer = { ...table, rows: table.rows.slice(1) };
    const { table: merged } = mergeTables(table, fewer);
    expect(merged.rows.length).toBe(table.rows.length - 1);
    expect(merged.rows[0]).toBe(table.rows[1]);
  });
});

describe("리스트 빠른 읽기 · 실제 기록 끝 행까지만, 묶음을 동시에", () => {
  it("서식 65536행·순번만 채운 빈 행은 읽지 않고, 묶음 읽기는 실제 끝 행까지만", async () => {
    const c = client();
    const sheet = (await c.worksheetsOf(LIST_ID)).find((s) => s.name.trim() === SITE)!;
    const last = await lastDataRow(c, LIST_ID, sheet.id);
    const recs = realListRecords().filter((r) => r.sheetName.trim() === SITE && r.excelRow);
    expect(last).toBe(Math.max(...recs.map((r) => r.excelRow!)));
    drive.log.length = 0;
    await fresh();
    const chunks = drive.log.filter((e) => /range\(address='A\d+:Y\d+'\)/.test(e.path));
    expect(chunks).toHaveLength(Math.ceil(last / LIVE_CHUNK_ROWS));
    expect(drive.log.every((e) => e.status !== 504)).toBe(true);
    expect(LIVE_PARALLEL).toBeGreaterThan(1);
    expect(LIVE_PARALLEL).toBeLessThanOrEqual(4);
  });

  it("묶음을 동시에 읽어도 행 순서·값은 한 묶음씩 읽은 것과 같다", async () => {
    const c = client();
    const table = await fresh();
    const sheet = (await c.worksheetsOf(LIST_ID)).find((s) => s.name.trim() === SITE)!;
    const { cells } = readSheet(drive.files(LIST_ID), sheet.name);
    for (const row of table.rows.slice(0, 50)) {
      expect(String(row.cells.find((x) => x.col === "F")!.value ?? "")).toBe(String(cells.get(`F${row.row}`)?.v ?? ""));
    }
  });

  it("끝 행 찾기 요청이 안 되면(Graph 오류) 예전처럼 값 범위로 읽는다", async () => {
    drive.fail((m, p) => m === "GET" && /range\(address='E1:G\d+'\)\/usedRange/.test(p), 400, Infinity);
    const table = await fresh();
    expect(table.rows.length).toBeGreaterThan(400);
    expect(drive.log.some((e) => /worksheets\/[^/]+\/usedRange\(valuesOnly=true\)/.test(e.path))).toBe(true);
  });

  it("로그인 중계 서버 모드(VITE_PAUT_API)면 클라이언트 ID 없이도 리스트 클라이언트를 만든다(서버 토큰 경로)", () => {
    vi.stubEnv("VITE_PAUT_API", "https://paut-api.example.test");
    vi.stubEnv("VITE_AZURE_CLIENT_ID", "");
    try {
      Object.assign(getState(), { config: { ...graphConfig(), clientId: "" } });
      expect(liveListClient().client).not.toBeNull();
      vi.stubEnv("VITE_PAUT_API", "");
      expect(liveListClient().client).toBeNull();
    } finally {
      vi.unstubAllEnvs();
    }
  });
});

describe("리스트 열 필터(앱 안에서만)", () => {
  const P = (row: { cells: { col: string; value: unknown }[] }) => cellText(row.cells.find((c) => c.col === "P") as never).trim() || BLANK;

  it("판정 값 고르기·라인 검색·모두 해제, 행 수 표시", async () => {
    const table = await fresh();
    const total = table.rows.length;
    const pv = columnValues(table.rows, "P");
    expect(pv.length).toBeGreaterThan(1);
    const pick = pv[0].value;
    let filters: ListFilters = {};
    // 처음엔 모두 켜짐 → 고른 값 하나만 남기려면 나머지를 끈다.
    for (const v of pv.slice(1)) filters = { ...filters, P: toggleValue(filters.P, v.value, pv.map((x) => x.value)) };
    let shown = applyFilters(table.rows, filters);
    expect(shown.length).toBe(pv[0].count);
    expect(shown.every((r) => P(r) === pick)).toBe(true);

    const line = cellText(shown[0].cells.find((c) => c.col === "F")!);
    const q = line.slice(0, 6).toLowerCase();
    filters = { ...filters, F: setText(filters.F, q) };
    shown = applyFilters(table.rows, filters);
    expect(shown.length).toBeGreaterThan(0);
    expect(shown.every((r) => cellText(r.cells.find((c) => c.col === "F")!).toLowerCase().includes(q) && P(r) === pick)).toBe(true);
    expect(activeColumns(filters).sort()).toEqual(["F", "P"]);
    expect(countText(shown.length, total)).toMatch(/^\d[\d,]* \/ [\d,]+행 표시$/);

    // 날짜 열 값 목록은 다른 열 필터를 따른다(엑셀처럼).
    const dates = columnValues(table.rows, "E", filters);
    expect(dates.reduce((n, d) => n + d.count, 0)).toBe(shown.length);

    expect(applyFilters(table.rows, clearColumn(filters, "F")).length).toBe(pv[0].count);
    expect(applyFilters(table.rows, {})).toBe(table.rows); // 모두 해제
    // 모든 값을 다시 켜면 그 열 필터는 없는 것과 같다.
    let back: ListFilters = {};
    back.P = toggleValue(undefined, pick, pv.map((x) => x.value));
    back.P = toggleValue(back.P, pick, pv.map((x) => x.value));
    expect(back.P.values).toBeNull();
    expect(filterableColumns(table.header).map((h) => h.col)).toEqual(expect.arrayContaining(["E", "I", "P", "F"]));
  });

  it("캐시로 보인 표에도 같은 필터가 걸리고, 필터는 표·캐시·엑셀을 바꾸지 않는다", async () => {
    await open();
    const cached = (await loadListCache(slot))!.table;
    const snapshot = structuredClone(cached.rows);
    const pv = columnValues(cached.rows, "P");
    const filters: ListFilters = { P: { values: [pv[0].value], text: "" } };
    const shown = applyFilters(cached.rows, filters);
    expect(shown.length).toBe(pv[0].count);
    expect(cached.rows).toEqual(snapshot);
    expect(cached.rows.length).toBeGreaterThan(shown.length); // 숨긴 행도 표·캐시에 그대로
    expect((await loadListCache(slot))!.table.rows.length).toBe(cached.rows.length);
    expect(writes()).toEqual([]);
  });

  it("필터로 보이는 행을 고치면 그 칸 하나만 PATCH, 행 추가·중복·필터 기록 없음", async () => {
    const recs = realListRecords();
    const table = await fresh();
    const pv = columnValues(table.rows, "P");
    const shown = applyFilters(table.rows, { P: { values: [pv[0].value], text: "" } });
    const row = shown[0];
    const cell = row.cells.find((c) => c.col === "R")!;
    const { cells: beforeCells } = readSheet(drive.files(LIST_ID), table.sheetName);
    const beforeRows = new Set([...beforeCells.keys()].map((k) => k.replace(/^[A-Z]+/, ""))).size;
    drive.log.length = 0;
    await writeLiveCell(client(), table, cell, "12");
    const patches = writes();
    expect(patches).toHaveLength(1);
    expect(patches[0].path).toContain(`range(address='R${row.row}')`);
    expect(JSON.stringify(patches)).not.toMatch(/autoFilter|filter|앱/i);
    expect(drive.log.some((e) => /autoFilter|\/insert|\/worksheets\/add/.test(e.path))).toBe(false);
    const { cells } = readSheet(drive.files(LIST_ID), table.sheetName);
    expect(String(cells.get(`R${row.row}`)?.v)).toBe("12");
    expect(new Set([...cells.keys()].map((k) => k.replace(/^[A-Z]+/, ""))).size).toBe(beforeRows);
    expect(recs.length).toBeGreaterThan(0);
  });
});

describe("리스트 휴대폰 카드 보기 · 주요 열 / 모든 열", () => {
  it("기본은 주요 열 다섯(날짜·라인·포인트·판정·비고), '모든 열' 은 나머지 열까지", async () => {
    const table = await fresh();
    const key = visibleColumns(table.header, "key");
    expect(key.map((c) => c.col)).toEqual(["E", "F", "G", "P", "S"]);
    expect([...KEY_COLS]).toEqual(["E", "F", "G", "P", "S"]);
    expect(key.find((c) => c.col === "F")!.title).toMatch(/line/i);
    const all = visibleColumns(table.header, "all");
    expect(all.length).toBeGreaterThan(key.length);
    for (const c of key) expect(all.map((x) => x.col)).toContain(c.col);
    expect(all.map((c) => c.col)).toEqual(expect.arrayContaining(["A", "B", "C", "D", "H", "K", "Q", "R"]));
    // 머리글 순서 그대로
    const order = table.header.map((h) => h.col);
    expect(all.map((c) => c.col)).toEqual([...all.map((c) => c.col)].sort((a, b) => order.indexOf(a) - order.indexOf(b)));
    expect(MODE_LABEL).toEqual({ key: "주요 열", all: "모든 열" });
  });

  it("카드에서 주요 열(판정·포인트)을 고쳐도 그 칸 하나만 PATCH, 행 추가·'앱' 없음", async () => {
    const table = await fresh();
    const row = table.rows[3];
    const cols = visibleColumns(table.header, "key").map((c) => c.col);
    const cell = row.cells.find((c) => c.col === "S" && cols.includes(c.col))!;
    drive.log.length = 0;
    await writeLiveCell(client(), table, cell, "카드에서 고침");
    expect(writes()).toHaveLength(1);
    expect(writes()[0].method).toBe("PATCH");
    expect(writes()[0].path).toContain(`range(address='S${row.row}')`);
    expect(JSON.stringify(writes())).not.toContain("앱");
    expect(drive.log.some((e) => /\/insert|\/worksheets\/add/.test(e.path))).toBe(false);
    const again = await fresh();
    expect(again.rows.length).toBe(table.rows.length);
    expect(again.rows.map((r) => r.row)).toEqual(table.rows.map((r) => r.row));
  });
});

describe("리스트 검사 날짜 요약 · 누르면 펼침, 「수정」 전에는 읽기만", () => {
  const SITE_PA = "지성 마전 PA";
  async function maJeon(): Promise<LiveTable> {
    const res = await loadLiveList(client(), graphConfig(), SITE_PA, [...realListRecords(), ...realGmRecords()]);
    if (!res.ok) throw new Error(res.message);
    return res.table;
  }
  function render(group: DateGroup, focus: DateFocus, cols: { col: string; title: string }[], table: LiveTable): string {
    return renderToStaticMarkup(
      createElement(DateSummaryCard, {
        group,
        focus,
        cols,
        busy: false,
        editing: null,
        recordOf: (row) => matchRecord(getState().records, table, row),
        onDate: () => undefined,
        onEditing: () => undefined,
        onRow: () => undefined,
        onCell: () => undefined,
        onAction: () => undefined,
      }),
    );
  }
  const cellButtons = (html: string) => [...html.matchAll(/<button[^>]*class="lc-f[^"]*"[^>]*>/g)].map((m) => m[0]);

  it("날짜마다 요약 하나: 행 수·합격/불합격 수·라인별 포인트 범위. 최근 날짜가 먼저, 날짜 없는 행은 맨 뒤", async () => {
    const table = await maJeon();
    const groups = summarizeByDate(table.rows);
    expect(groups.flatMap((g) => g.rows).length).toBe(table.rows.length);
    const dates = groups.map((g) => g.date);
    expect(dates).toEqual([...dates].sort().reverse());
    for (const g of groups) {
      expect(g.rows.every((r) => cellText(r.cells.find((c) => c.col === "E")!).trim() === g.date)).toBe(true);
      const at = g.rows.map((r) => table.rows.indexOf(r));
      expect(at).toEqual([...at].sort((a, b) => a - b));
      expect(g.pass + g.fail).toBeLessThanOrEqual(g.rows.length);
    }
    const day = groups.find((g) => g.date === "2026-06-18")!;
    expect(day.rows).toHaveLength(44);
    expect(day.pass + day.fail).toBe(44);
    expect(day.lines.find((l) => l.line === "MJDH-D-D-P-58-003-200A")).toEqual({ line: "MJDH-D-D-P-58-003-200A", short: "003", ranges: "S25~S40, R25~R40" });
    expect(day.lines.find((l) => l.short === "006")!.ranges).toBe("S11~S16, R11~R16");

    const undated = { ...table.rows[0], row: 99999, cells: table.rows[0].cells.map((c) => (c.col === "E" ? { ...c, value: "" } : c)) };
    const withUndated = summarizeByDate([undated, ...table.rows]);
    expect(withUndated.at(-1)!.date).toBe("");
    expect(withUndated.at(-1)!.rows).toEqual([undated]);
    expect(withUndated[0].date).toBe(dates[0]);

    expect(shortLine("HY-S-035-1-200A-IN")).toBe("035-1");
    expect(shortLine("GP-001-750A")).toBe("001");
    expect(shortLine("군포산본 이원화 750A-5")).toBe("군포산본 이원화 750A-5");
  });

  it("같은 번호로 줄어드는 라인이 둘이면 라인 전체를 보인다", async () => {
    const table = await maJeon();
    const row = table.rows[0];
    const other = { ...row, row: 99998, cells: row.cells.map((c) => (c.col === "F" ? { ...c, value: "XX-" + cellText(c) } : c)) };
    const lines = summarizeByDate([row, other])[0].lines;
    expect(lines.map((l) => l.short)).toEqual(lines.map((l) => l.line));
  });

  it("한 번에 한 날짜만 펼친다. 다른 날짜를 누르면 앞 날짜는 접히고 「수정」도 풀린다, 같은 날짜를 다시 누르면 접힘", () => {
    let focus: DateFocus = null;
    focus = toggleDate(focus, "2026-06-18");
    expect(focus).toEqual({ date: "2026-06-18", editing: false, row: null });
    expect(toggleRow(focus, 140)).toBe(focus);
    focus = toggleRow(toggleEditing(focus), 140);
    expect(focus).toEqual({ date: "2026-06-18", editing: true, row: 140 });
    expect(toggleRow(focus, 140)).toEqual({ date: "2026-06-18", editing: true, row: null });
    expect(toggleEditing(focus)).toEqual({ date: "2026-06-18", editing: false, row: null });
    focus = toggleDate(focus, "2026-06-10");
    expect(focus).toEqual({ date: "2026-06-10", editing: false, row: null });
    expect(toggleDate(focus, "2026-06-10")).toBeNull();
    expect(toggleEditing(null)).toBeNull();
  });

  it("접힌 요약은 행을 그리지 않고, 펼치면 행이 한 줄씩(칸 버튼·삭제 없음). 「수정」 뒤 연 행만 칸과 리페어 처리·삭제", async () => {
    getState().records.splice(0, getState().records.length, ...realListRecords(), ...realGmRecords());
    const table = await maJeon();
    const group = summarizeByDate(table.rows).find((g) => g.date === "2026-06-18")!;
    const cols = visibleColumns(table.header, "all");

    const closed = render(group, null, cols, table);
    expect(closed).toContain("26.06.18");
    expect(closed).toContain("44행");
    expect(closed).toContain(`합격 ${group.pass}`);
    expect(closed).toContain("003 · S25~S40, R25~R40");
    expect(closed).toContain('aria-expanded="false"');
    expect(closed).not.toContain("dr-row");
    expect(closed).not.toContain(">수정</button>");

    const open = render(group, { date: group.date, editing: false, row: null }, cols, table);
    expect(open.match(/class="dr-row"/g)).toHaveLength(44);
    expect(open).not.toMatch(/<button[^>]*class="dr-row"/);
    expect(open).toContain(">수정</button>");
    expect(cellButtons(open)).toEqual([]);
    expect(open).not.toContain("삭제");
    expect(open).not.toContain("리페어 처리");

    const target = group.rows.find((r) => matchRecord(getState().records, table, r))!;
    const editing = render(group, { date: group.date, editing: true, row: null }, cols, table);
    expect(editing).toContain(">완료</button>");
    expect(editing.match(/<button[^>]*class="dr-row"/g)).toHaveLength(44);
    expect(cellButtons(editing)).toEqual([]);

    const rowOpen = render(group, { date: group.date, editing: true, row: target.row }, cols, table);
    const enabled = cellButtons(rowOpen).filter((b) => !b.includes("disabled"));
    expect(enabled.some((b) => b.includes(`aria-label="${cols.find((c) => c.col === "P")!.title} P${target.row}"`))).toBe(true);
    expect(enabled.some((b) => / Y\d+"/.test(b))).toBe(false);
    expect(rowOpen.match(/>삭제<\/button>/g)).toHaveLength(1);
    expect(rowOpen.match(/>리페어 처리<\/button>/g)).toHaveLength(1);

    const otherDay = summarizeByDate(table.rows).find((g) => g.date === "2026-06-10")!;
    expect(render(otherDay, { date: group.date, editing: true, row: target.row }, cols, table)).not.toContain("dr-row");
  });

  it("「삭제」는 확인을 받아야만 지운다. 취소하면 앱 기록·엑셀 대기 모두 그대로", async () => {
    getState().records.splice(0, getState().records.length, ...realListRecords());
    const table = await fresh();
    const row = table.rows.find((r) => matchRecord(getState().records, table, r) && cellText(r.cells.find((c) => c.col === "G")!))!;
    const rec = matchRecord(getState().records, table, row)!;
    const before = JSON.stringify(rec);
    const asked: string[] = [];
    const res = await applyRowAction("delete", table, row, (text) => (asked.push(text), false));
    await settled();
    expect(res).toEqual({ done: false, note: "" });
    expect(asked).toEqual([deleteConfirmText(row)]);
    expect(asked[0]).toMatch(/^\S+ \(\d{2}\.\d{2}\.\d{2}\) 기록을 지울까요\? 엑셀 행 내용이 지워집니다\.$/);
    expect(JSON.stringify(getState().records.find((r) => r.id === rec.id))).toBe(before);

    const ok = await applyRowAction("delete", table, row, () => true);
    await settled();
    expect(ok.done).toBe(true);
    expect(getState().records.find((r) => r.id === rec.id)).toMatchObject({ deleted: true, sync: "pending" });
  });
});

describe("두 손가락 확대/축소", () => {
  const html = readFileSync(new URL("../../index.html", import.meta.url), "utf8");
  const css = readFileSync(new URL("../styles.css", import.meta.url), "utf8");

  it("viewport 가 확대를 막지 않는다(standalone PWA 설정은 그대로)", () => {
    const meta = /<meta name="viewport" content="([^"]+)"/.exec(html)?.[1] ?? "";
    expect(meta).toContain("width=device-width");
    expect(meta).toContain("initial-scale=1");
    expect(meta).toContain("viewport-fit=cover");
    expect(meta).not.toMatch(/user-scalable\s*=\s*(no|0)/);
    expect(meta).not.toMatch(/maximum-scale\s*=\s*1(\.0)?\b/);
    expect(html).toContain('name="apple-mobile-web-app-capable" content="yes"');
  });

  it("CSS touch-action 이 두 손가락 확대를 막지 않는다(칠하기 카드는 pan-y + pinch-zoom)", () => {
    expect(css).not.toMatch(/touch-action:\s*none/);
    for (const m of css.matchAll(/touch-action:\s*([^;]+);/g)) {
      const v = m[1].trim();
      expect(["manipulation", "auto", "pan-y pinch-zoom", "pan-x pinch-zoom"]).toContain(v);
    }
    expect(css).toMatch(/\.pcard \{[^}]*touch-action: pan-y pinch-zoom;/);
    expect(css).toMatch(/\.point-grid \{[^}]*touch-action: pan-y pinch-zoom;/);
  });
});
