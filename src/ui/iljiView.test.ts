/**
 * 두 번째 탭 「일지」: OneDrive 작업일지(PAUT FIELD)를 달별로 엑셀처럼 보고, 「수정」 중 칸 하나만 고친다.
 * 가짜 OneDrive(Graph)와 실제 작업일지 사본만 쓴다.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
vi.mock("../store/db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => void mem.set(key, structuredClone(value)),
  kvDel: async (key: string) => void mem.delete(key),
}));
vi.stubGlobal("localStorage", { getItem: () => null, setItem: () => undefined, removeItem: () => undefined });

import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { GraphClient, graphRetry } from "../sync/graph";
import type { FakeDrive } from "../test/fakeGraph";
import { ILJI_ID, graphConfig, meta, realDrive } from "../test/realBooks";
import { iljiText, loadIlji, planIljiEdit, type IljiCell, type IljiTable } from "../excel/liveIlji";
import { readSheet, setCell, setXml } from "../excel/ooxml";
import { getState, settled } from "../store/session";
import { IljiGrid, IljiScreen, defaultMonth, saveIljiEdit, type IljiEdit } from "./IljiScreen";
import { TabBar } from "./App";
import { RefScreen } from "./RefScreen";

let drive: FakeDrive;
const client = () => new GraphClient({ getToken: async () => drive.validToken, fetchImpl: drive.fetch, sleep: async () => undefined });
const writes = () => drive.log.filter((e) => e.method === "PATCH");
const SHEET = " PAUT FIELD";
const sheetValue = (ref: string, id = ILJI_ID) => readSheet(drive.files(id), SHEET).cells.get(ref)?.v ?? "";

async function load(): Promise<IljiTable> {
  const res = await loadIlji(client(), getState().config);
  if (!res.ok) throw new Error(res.message);
  return res.table;
}

const cellAt = (table: IljiTable, address: string): IljiCell => table.rows.flatMap((r) => r.cells).find((c) => c.address === address)!;

function render(table: IljiTable, month: string, unlocked: boolean, edit: IljiEdit | null = null): string {
  const noop = () => undefined;
  return renderToStaticMarkup(createElement(IljiGrid, { table, month, unlocked, busy: false, edit, onCell: noop, onEditText: noop, onSave: noop, onCancelEdit: noop }));
}

beforeEach(() => {
  mem.clear();
  graphRetry.baseMs = 0;
  drive = realDrive();
  Object.assign(getState(), { config: graphConfig(), meta: meta({ workDate: "2026-11-12", dayStaff: {} }) });
});

describe("일지 탭 · 작업일지 보기", () => {
  it("시트 행 순서 그대로, 4·5행 머리글(병합 그대로)과 엑셀 행 번호. 쓰기 없음", async () => {
    const table = await load();
    expect(table.cols.join("")).toBe("ABCDEFGHIJKLMNOPQRSTU");
    expect(table.months.at(-1)).toBe("2026-11");
    const html = render(table, "2026-11", false);
    for (const title of ["순번", "일자", "구분", "지성 마전", "드림 합정주공", "그 외", "관경", "개소", "검사자", "근무시간", "비고"]) expect(html).toContain(`>${title}</th>`);
    expect(html).toMatch(/<th scope="col" class="xg-h c-P" rowSpan="2" colSpan="3">검사자<\/th>/);
    const rowNumbers = [...html.matchAll(/<th scope="row" class="xg-rn">(\d+)<\/th>/g)].map((m) => Number(m[1]));
    expect(rowNumbers).toEqual([519, 520, 521]);
    expect(html).toContain(">26.11.03</td>");
    expect(html).toContain(">13:00</td>");
    expect(writes()).toEqual([]);
  });

  it("달 고르기: 작업 날짜의 달이 처음, 고른 달의 행만(시트 순서), 병합된 날은 한 칸으로", async () => {
    const table = await load();
    expect(defaultMonth(table, "2026-11-12")).toBe("2026-11");
    expect(defaultMonth(table, "2026-06-17")).toBe("2026-06");
    expect(defaultMonth(table, "2030-01-01")).toBe("2026-11");
    const html = render(table, "2026-09", false);
    const rowNumbers = [...html.matchAll(/<th scope="row" class="xg-rn">(\d+)<\/th>/g)].map((m) => Number(m[1]));
    expect(rowNumbers).toEqual(Array.from({ length: 12 }, (_, i) => 497 + i));
    expect(html).toMatch(/<td class="c-B k-date" rowSpan="2" colSpan="1">26.09.05<\/td>/);
    expect(html).toMatch(/<tr class="cont"><th scope="row" class="xg-rn">500<\/th><td class="c-D/);
    expect(html).not.toContain(">26.11.03</td>");
  });
});

describe("일지 탭 · 「수정」", () => {
  it("「수정」 전에는 고칠 버튼이 없고, 수정 중에는 고칠 칸만 버튼(순번·제외 현장은 회색)", async () => {
    const table = await load();
    expect(render(table, "2026-11", false)).not.toContain("<button");
    const html = render(table, "2026-11", true);
    expect(html).toContain('aria-label="검사자 P519"');
    expect(html).toContain('aria-label="근무 시작 S519"');
    expect(html).toContain('aria-label="지성 군포 개소 I519"');
    expect(html).not.toContain('aria-label="순번 A519"');
    expect(html).toMatch(/<td class="c-A k-seq num ro"/);
  });

  it("편집 칸: 날짜·시간은 선택기, 구분은 평일/공휴일", async () => {
    const table = await load();
    const edit = (address: string, text: string) => render(table, "2026-11", true, { cell: cellAt(table, address), text });
    expect(edit("B519", "2026-11-03")).toMatch(/type="date"[^>]*value="2026-11-03"/);
    expect(edit("S519", "13:00")).toMatch(/type="time"[^>]*value="13:00"/);
    expect(edit("C519", "평일")).toMatch(/<select[^>]*aria-label="C519 구분 값"/);
  });

  it("확인을 취소하면 쓰지 않는다. 확인하면 그 칸 하나만 쓰고 다시 읽어 확인, 앱의 그날 검사자도 맞춘다", async () => {
    const table = await load();
    const edit = { cell: cellAt(table, "Q519"), text: "심명보" };
    let asked = "";
    const cancelled = await saveIljiEdit(client(), table, edit, (text) => ((asked = text), false));
    expect(cancelled).toEqual({ done: false, note: "" });
    expect(asked).toBe("Q519 (검사자)\n김정환 → 심명보\n작업일지(PAUT 근무일지.xlsx)에 바로 씁니다. 바꿀까요?");
    expect(writes()).toEqual([]);

    drive.log.length = 0;
    const res = await saveIljiEdit(client(), table, edit, () => true);
    expect(res.done).toBe(true);
    if (!res.done) return;
    expect(writes().map((e) => [e.path.match(/address='([^']+)'/)?.[1], e.body])).toEqual([["Q519", { formulas: [["심명보"]] }]]);
    const reads = drive.log.filter((e) => e.method === "GET" && e.path.includes("address='Q519'"));
    expect(reads.length).toBe(2);
    expect(sheetValue("Q519")).toBe("심명보");
    expect(sheetValue("Q520")).toBe("");
    expect(iljiText(cellAt(res.table, "Q519"))).toBe("심명보");
    expect(res.note).toContain("Q519 저장됨 (김정환 → 심명보) · 엑셀 다시 읽어 확인함");
    await settled();
    expect(getState().meta.dayStaff?.["2026-11-03"]).toEqual({ dayType: "평일", insp1: "조우현", insp2: "심명보", timeStart: "13:00", timeEnd: "17:30" });
  });

  it("근무 시간은 하루 비율로, 날짜는 일련번호로 쓴다", async () => {
    const table = await load();
    const res = await saveIljiEdit(client(), table, { cell: cellAt(table, "T520"), text: "16:45" }, () => true);
    expect(res.done).toBe(true);
    expect(Number(sheetValue("T520"))).toBeCloseTo(16.75 / 24, 9);
    await settled();
    expect(getState().meta.dayStaff?.["2026-11-04"]?.timeEnd).toBe("16:45");
    const dated = await saveIljiEdit(client(), table, { cell: cellAt(table, "B520"), text: "2026-11-06" }, () => true);
    expect(dated.done).toBe(true);
    expect(sheetValue("B520")).toBe("46332");
  });

  it("같은 날짜를 다른 행에 또 쓰지 않는다(쓰기 없음)", async () => {
    const table = await load();
    const res = await saveIljiEdit(client(), table, { cell: cellAt(table, "B520"), text: "2026-11-03" }, () => true);
    expect(res).toEqual({ done: false, note: "B520: 519행에 이미 같은 날짜가 있습니다. 같은 날짜 행을 두 개 만들지 않습니다." });
    expect(writes()).toEqual([]);
  });

  it("수식 칸은 고치지 않는다", async () => {
    const files = drive.files(ILJI_ID);
    const { info, xml } = readSheet(files, SHEET);
    setXml(files, info.path, setCell(xml, "U520", { kind: "formula", formula: "I520*2", cached: 14 }));
    const table = await load();
    const cell = cellAt(table, "U520");
    expect(cell.lock).toBe("formula");
    expect(render(table, "2026-11", true)).not.toContain('aria-label="비고 U520"');
    expect(() => planIljiEdit(table, cell, "메모")).toThrow("수식 칸은 고치지 않습니다.");
    const res = await saveIljiEdit(client(), table, { cell, text: "메모" }, () => true);
    expect(res.done).toBe(false);
    expect(writes()).toEqual([]);
  });

  it("제외 현장: 지성 브레인연계 칸과 '그 외'의 드림화양 칸은 잠기고, 제외 현장만 있는 행은 행 전체가 잠긴다", async () => {
    const table = await load();
    expect(["D519", "E519", "N509", "O509"].map((a) => cellAt(table, a).lock)).toEqual(["excluded", "excluded", "excluded", "excluded"]);
    expect(cellAt(table, "P509").lock).toBe("excluded");
    expect(cellAt(table, "P519").lock).toBe(null);
    expect(cellAt(table, "H519").lock).toBe(null);
    expect(table.rows.find((r) => r.row === 521)!.cells.filter((c) => c.col !== "A").every((c) => c.lock === "excluded")).toBe(true);
    const html = render(table, "2026-11", true);
    expect(html).not.toContain('aria-label="검사자 P521"');
    const res = await saveIljiEdit(client(), table, { cell: cellAt(table, "P521"), text: "심명보" }, () => true);
    expect(res.done).toBe(false);
    expect(res.note).toContain("제외된 현장");
    expect(writes()).toEqual([]);
  });

  it("불러온 뒤 다른 곳(PC 엑셀)에서 그 칸이 바뀌었으면 쓰지 않는다", async () => {
    const table = await load();
    const files = drive.files(ILJI_ID);
    const { info, xml } = readSheet(files, SHEET);
    setXml(files, info.path, setCell(xml, "P520", { kind: "inline", text: "이연행" }));
    drive.log.length = 0;
    const res = await saveIljiEdit(client(), table, { cell: cellAt(table, "P520"), text: "심명보" }, () => true);
    expect(res).toMatchObject({ done: false, conflict: true });
    expect(res.note).toContain("P520 칸이");
    expect(res.note).toContain("이연행");
    expect(writes()).toEqual([]);
    expect(sheetValue("P520")).toBe("이연행");
  });

  it("저장된 작업일지 ID 가 404 면 이름으로 찾아 읽고 쓰며, 새 ID 를 설정에 남긴다", async () => {
    drive.reupload(ILJI_ID, "ILJI-NEW");
    const table = await load();
    expect(table.file.id).toBe("ILJI-NEW");
    const stale = { ...table, file: { ...table.file, id: ILJI_ID } };
    const res = await saveIljiEdit(client(), stale, { cell: cellAt(table, "U519"), text: "PC 교체" }, () => true);
    expect(res.done).toBe(true);
    expect(sheetValue("U519", "ILJI-NEW")).toBe("PC 교체");
    await settled();
    expect(getState().config.files.ilji?.id).toBe("ILJI-NEW");
  });
});

describe("5번 탭 · 참고", () => {
  it("참고(관경·스캔·스케줄표)는 5번째 탭이고 일지 탭에는 없다", () => {
    const bar = renderToStaticMarkup(createElement(TabBar, { tab: "ref", onTab: () => undefined }));
    expect([...bar.matchAll(/<button/g)]).toHaveLength(5);
    expect(bar).toMatch(/<button[^>]*class="on"[^>]*><span>⌗<\/span>참고<\/button><\/nav>$/);
    const ref = renderToStaticMarkup(createElement(RefScreen));
    expect(ref).toContain("외경/스캔");
    expect(ref).toContain("이중보온관");
    const ilji = renderToStaticMarkup(createElement(IljiScreen));
    expect(ilji).toContain("OneDrive 에서 작업일지를 읽는 중…");
    expect(ilji).not.toContain("외경/스캔");
    expect(ilji).not.toContain(">참고</button>");
  });
});

describe("입력 탭 날짜별 기록 · 일지·리스트 탭 수정 연동", () => {
  it("일지 탭·리스트 탭에서 맞춘 그날 검사자·시간(dayStaff)이 기기의 작업일지 사본보다 먼저 보인다", async () => {
    const { mergeDayStaff } = await import("./DayHistory");
    const copy = { "2026-11-24": { dayType: "평일", workers: ["조우현", "심명보"], timeStart: "09:00", timeEnd: "18:00" } };
    const merged = mergeDayStaff(copy, {
      "2026-11-24": { dayType: "평일", insp1: "조우현", insp2: "", timeStart: "14:00", timeEnd: "18:30" },
      "2026-11-25": { dayType: "평일", insp1: "김정환", insp2: "", timeStart: "08:00", timeEnd: "12:00" },
    });
    expect(merged["2026-11-24"]).toEqual({ dayType: "평일", workers: ["조우현"], timeStart: "14:00", timeEnd: "18:30" });
    expect(merged["2026-11-25"].workers).toEqual(["김정환"]);
    expect(copy["2026-11-24"].workers).toEqual(["조우현", "심명보"]);
  });
});

describe("일지 탭을 열면 날짜별 기록의 근무자·시간이 엑셀 값으로 맞춰진다", () => {
  it("엑셀(PC 에서 고친 것 포함) 날짜 행의 검사자·시간을 그날 값으로. 보내는 중인 기록이 있는 날은 앱 값을 둔다", async () => {
    const { adoptIljiDays } = await import("../store/session");
    const { rowStaff } = await import("../excel/liveIlji");
    const table = await load();
    Object.assign(getState(), {
      meta: meta({ workDate: "2026-11-12", dayStaff: { "2026-11-03": { dayType: "평일", insp1: "옛값", insp2: "", timeStart: "09:00", timeEnd: "18:00" } } }),
      records: [],
    });
    const rows = table.rows.filter((r) => r.date).map((r) => ({ date: r.date, ...rowStaff(table, r) }));
    expect(await adoptIljiDays(rows)).toBeGreaterThan(0);
    expect(getState().meta.dayStaff?.["2026-11-03"]).toEqual({ dayType: "평일", insp1: "조우현", insp2: "김정환", timeStart: "13:00", timeEnd: "17:30" });
    // 다시 해도 바뀌는 것 없음
    expect(await adoptIljiDays(rows)).toBe(0);
  });
});
