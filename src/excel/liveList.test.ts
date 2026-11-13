/**
 * 세 번째 탭(리스트): 현장 시트를 가짜 OneDrive(Graph)에서 바로 읽고, 칸 하나만 PATCH(If-Match)로 쓴다.
 * 시트를 만들지 않고, 행을 넣지 않고, '앱'을 붙이지 않는다. 실제 OneDrive 는 부르지 않는다.
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
import { LIST_GM_ID, LIST_ID, graphConfig, realDrive, realGmRecords, realListRecords } from "../test/realBooks";
import { cellText, isEditable, normalizeResult, sortBySeq, type LiveRow, liveTarget, loadLiveList, matchRecord, recordPatch, writeLiveCell, LiveConflictError, type LiveTable } from "./liveList";
import { readSheet, workbookSheets } from "./ooxml";
import { applyListCellEdit, getState, liveListClient } from "../store/session";
import { LIST_LOGIN_TEXT, LIST_TIMEOUT_TEXT, listErrorText, readLiveList } from "../ui/ListScreen";
import { AuthExpiredError, getAuthState, inflightCount, markExpired, resetAuthForTests } from "../sync/authState";
import { LoginRequiredError } from "../sync/msal";
import { FAIL, PASS } from "../domain/types";

let drive: FakeDrive;
const client = () => new GraphClient({ getToken: async () => drive.validToken, fetchImpl: drive.fetch, sleep: async () => undefined });
const writes = () => drive.log.filter((e) => e.method !== "GET");

beforeEach(() => {
  mem.clear();
  bag.clear();
  graphRetry.baseMs = 0;
  drive = realDrive();
});

async function loadOk(site: string, records = realListRecords()): Promise<LiveTable> {
  const res = await loadLiveList(client(), graphConfig(), site, records);
  if (!res.ok) throw new Error(res.message);
  return res.table;
}

describe("리스트 탭 · OneDrive 검사리스트 바로 보기·고치기", () => {
  it("현장 시트를 가짜 Graph 에서 읽어 표로 만든다(평택 리스트·군포·마전 리스트)", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    const table = await loadOk(app.site, recs);
    expect(table.itemId).toBe(LIST_ID);
    expect(table.sheetName.trim()).toBe(app.sheetName.trim());
    const row = table.rows.find((r) => r.row === app.excelRow)!;
    expect(cellText(row.cells.find((c) => c.col === "F")!)).toBe(app.line);
    expect(cellText(row.cells.find((c) => c.col === "E")!)).toBe(app.date);
    expect(table.header.find((h) => h.col === "F")?.title).toMatch(/line/i);
    expect(writes()).toEqual([]); // 읽기만

    const gm = realGmRecords();
    const gmTable = await loadOk(gm[0].site, gm);
    expect(gmTable.itemId).toBe(LIST_GM_ID);
    expect(liveTarget(gm[0].site, gm).book).toBe("listGm");
    expect(gmTable.rows.length).toBeGreaterThan(0);
  });

  it("칸 하나를 고치면 그 주소만 If-Match 로 PATCH 하고, 같은 앱 기록(Y 앱 ID)이 바뀐다. 행 추가·'앱' 없음", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    mem.set("records", recs);
    getState().records.splice(0, getState().records.length, ...structuredClone(recs));
    const table = await loadOk(app.site, recs);
    const row = table.rows.find((r) => r.row === app.excelRow)!;
    const cell = row.cells.find((c) => c.col === "Q")!; // 결함 종류
    expect(isEditable(cell)).toBe(true);
    const rec = matchRecord(getState().records, table, row)!;
    expect(rec.id).toBe(app.id);

    drive.log.length = 0;
    const written = await writeLiveCell(client(), table, cell, "SLAG");
    const patches = writes();
    expect(patches).toHaveLength(1);
    expect(patches[0].method).toBe("PATCH");
    expect(patches[0].path).toContain(`range(address='Q${app.excelRow}')`);
    expect(patches[0].ifMatch).toBeTruthy();
    expect(patches[0].body).toEqual({ values: [["SLAG"]] });
    expect(JSON.stringify(patches)).not.toContain("앱");
    expect(drive.log.some((e) => /\/insert|\/worksheets\/add|\/copy/.test(e.path))).toBe(false);
    const { cells } = readSheet(drive.files(LIST_ID), table.sheetName);
    expect(cells.get(`Q${app.excelRow}`)?.v).toBe("SLAG");
    expect(cells.get(`S${app.excelRow}`)?.v ?? "").not.toContain("앱");

    await applyListCellEdit(rec.id, recordPatch("Q", written)!);
    const after = getState().records.find((r) => r.id === app.id)!;
    expect(after.flawType).toBe("SLAG");
    expect(after.sync).toBe(rec.sync); // 엑셀은 이미 그 값: 다시 보내지 않음
  });

  it("엑셀에서 읽은 행(xlsx:시트:행)도 같은 기록으로 찾고, 날짜 칸은 일련번호로 쓴다", async () => {
    const recs = realListRecords();
    const plain = recs.find((r) => r.src === "xlsx")!;
    const table = await loadOk(plain.site, recs);
    const row = table.rows.find((r) => r.row === plain.excelRow)!;
    expect(matchRecord(recs, table, row)?.id).toBe(plain.id);
    const cell = row.cells.find((c) => c.col === "E")!;
    drive.log.length = 0;
    const v = await writeLiveCell(client(), table, cell, "2026-11-13");
    expect(typeof v).toBe("number");
    expect(recordPatch("E", v)).toEqual({ date: "2026-11-13" });
    expect(writes().map((e) => e.path.match(/address='([^']+)'/)?.[1])).toEqual([`E${plain.excelRow}`]);
  });

  it("다른 곳에서 그 칸이 바뀌었으면 쓰지 않는다. 파일만 바뀌었으면(eTag) 새 eTag 로 한 번 다시 보낸다", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    const table = await loadOk(app.site, recs);
    const cell = table.rows.find((r) => r.row === app.excelRow)!.cells.find((c) => c.col === "R")!;
    // PC 엑셀에서 같은 칸을 고침
    await client().patchValues(LIST_ID, table.sheetId, cell.address, [["99"]]);
    drive.log.length = 0;
    await expect(writeLiveCell(client(), table, cell, "12")).rejects.toBeInstanceOf(LiveConflictError);
    expect(writes()).toEqual([]);

    const fresh = await loadOk(app.site, recs);
    const cell2 = fresh.rows.find((r) => r.row === app.excelRow)!.cells.find((c) => c.col === "R")!;
    drive.touch(LIST_ID); // 다른 칸이 바뀜 → eTag 만 다름
    drive.log.length = 0;
    await writeLiveCell(client(), fresh, cell2, "12");
    const p = writes();
    expect(p.map((e) => e.status)).toEqual([412, 200]);
    expect(readSheet(drive.files(LIST_ID), fresh.sheetName).cells.get(cell2.address)?.v).toBe("12");
  });

  it("수식 칸(L = RIGHT)과 Y(앱 ID)는 고치지 않는다", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    const table = await loadOk(app.site, recs);
    const row = table.rows.find((r) => r.row === app.excelRow)!;
    const l = row.cells.find((c) => c.col === "L")!;
    expect(l.formula).toBe(true);
    expect(isEditable(l)).toBe(false);
    expect(isEditable(row.cells.find((c) => c.col === "Y")!)).toBe(false);
    await expect(writeLiveCell(client(), table, l, "100A")).rejects.toThrow("수식");
  });

  it("시트가 없으면 안내만 하고 시트를 만들지 않는다. 파일을 못 찾아도 만들지 않는다", async () => {
    const before = workbookSheets(drive.files(LIST_ID)).length;
    const res = await loadLiveList(client(), graphConfig(), "없는 현장 XYZ", []);
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.message).toContain("시트가 없습니다");
    expect(workbookSheets(drive.files(LIST_ID)).length).toBe(before);
    expect(writes()).toEqual([]);

    drive.items.delete(LIST_ID);
    const gone = await loadLiveList(client(), graphConfig(), "드림 합정주공", realListRecords());
    expect(gone.ok).toBe(false);
    if (!gone.ok) expect(gone.message).toContain("찾지 못했습니다");
    expect(writes()).toEqual([]);
  });

  it("OneDrive 로그인 전이면 Graph 를 부르지 않고 한국어 안내", () => {
    const { client: c, reason } = liveListClient();
    expect(c).toBeNull();
    expect(reason).toMatch(/로그인|OneDrive/);
  });

  it("세 번째 탭은 리스트, 작업일지(일지) 화면은 설정에서 그대로 열 수 있다. 참고·명세서 탭은 그대로", () => {
    const app = readFileSync(new URL("../ui/App.tsx", import.meta.url), "utf8");
    const tabbar = app.slice(app.indexOf('<nav className="tabbar">'), app.indexOf("</nav>"));
    const labels = [...tabbar.matchAll(/<\/span>([^<]+)<\/button>/g)].map((m) => m[1]);
    expect(labels).toEqual(["입력", "참고", "리스트", "명세서"]);
    expect(app).toContain('{tab === "log" && <ExcelScreen part="log" />}');
    expect(app).toMatch(/onClick=\{\(\) => setTab\("log"\)\}>작업일지 미리보기<\/button>/);
  });
});

describe("리스트 탭 · 엑셀 순번(A열) 순서 · 1번 탭 카드 즉시 반영", () => {
  const mk = (row: number, a: unknown): LiveRow => ({ row, cells: [{ address: `A${row}`, col: "A", row, value: a as never, formula: false }] });

  it("물리 행이 3,1,2 순번이면 1,2,3 으로 보이고, 빈·숫자 아닌 순번은 뒤에 시트 행 순서대로", () => {
    const rows = [mk(6, 3), mk(7, ""), mk(8, 1), mk(9, "비고"), mk(10, "2"), mk(11, null)];
    expect(sortBySeq(rows).map((r) => r.row)).toEqual([8, 10, 6, 7, 9, 11]);
    expect(rows.map((r) => r.row)).toEqual([6, 7, 8, 9, 10, 11]); // 원본 배열·값 그대로
  });

  it("OneDrive 시트에서 순번을 3,1,2 로 바꿔 두면 표는 순번 순서. 읽기만 하고 순번은 다시 쓰지 않는다", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    const first = await loadOk(app.site, recs);
    const phys = first.rows.map((r) => r.row).sort((a, b) => a - b).slice(0, 3);
    for (const [i, n] of [3, 1, 2].entries()) await client().patchValues(LIST_ID, first.sheetId, `A${phys[i]}`, [[n]]);
    drive.log.length = 0;
    const table = await loadOk(app.site, recs);
    expect(writes()).toEqual([]);
    const seqs = table.rows.map((r) => r.cells[0].value);
    expect(table.rows.slice(0, 3).map((r) => r.row)).toEqual([phys[1], phys[2], phys[0]]);
    const nums = seqs.filter((v) => typeof v === "number") as number[];
    expect(nums).toEqual([...nums].sort((a, b) => a - b));
    const firstBlank = seqs.findIndex((v) => typeof v !== "number" && !/^\d+$/.test(String(v ?? "").trim()));
    if (firstBlank >= 0) expect(seqs.slice(firstBlank).every((v) => typeof v !== "number")).toBe(true);
  });

  it("리스트에서 E/F/G/P 를 저장하면 같은 앱 기록(1번 탭 카드)의 날짜·라인·포인트·합격/불합격이 바로 바뀐다", async () => {
    const recs = realListRecords();
    const app = recs.find((r) => r.src === "user")!;
    getState().records.splice(0, getState().records.length, ...structuredClone(recs));
    getState().meta.workDate = app.date;
    const table = await loadOk(app.site, recs);
    const row = table.rows.find((r) => r.row === app.excelRow)!;
    const edit = async (col: string, text: string) => {
      const rec = matchRecord(getState().records, table, structuredClone(row))!;
      expect(rec.id).toBe(app.id);
      const cell = row.cells.find((c) => c.col === col)!;
      const written = await writeLiveCell(client(), table, cell, text);
      await applyListCellEdit(rec.id, recordPatch(col, written)!);
      return getState().records.find((r) => r.id === app.id)!;
    };
    const before = getState().records;
    const cardsOn = (date: string) => getState().records.filter((r) => !r.deleted && r.site === app.site && r.date === date && r.id === app.id);

    const wasPass = app.result === PASS;
    let now = await edit("P", wasPass ? "불" : "합");
    expect(now.result).toBe(wasPass ? FAIL : PASS);
    expect(getState().records).not.toBe(before); // 새 배열 → 화면 다시 그림

    now = await edit("F", "NEW-LINE-9");
    expect(now.line).toBe("NEW-LINE-9");

    now = await edit("G", "r2");
    expect(now.point).toBe("R02");
    expect(matchRecord(getState().records, table, row)?.id).toBe(app.id); // 엑셀 R2 ↔ 앱 R02

    now = await edit("E", "2026-11-20");
    expect(now.date).toBe("2026-11-20");
    expect(cardsOn(app.date)).toHaveLength(0); // 옛 날짜 카드 없음
    expect(cardsOn("2026-11-20")).toHaveLength(1); // 새 날짜에 하나만
    expect(getState().records.filter((r) => r.id === app.id)).toHaveLength(1);
    expect(JSON.stringify(writes())).not.toContain("앱");
  });

  it("P 글자를 합격/불합격으로 맞춘다", () => {
    expect(normalizeResult(" 합 ")).toBe(PASS);
    expect(normalizeResult("pass")).toBe(PASS);
    expect(normalizeResult("불합격")).toBe(FAIL);
    expect(normalizeResult("X")).toBe(FAIL);
    expect(recordPatch("P", "불")).toEqual({ result: FAIL });
    expect(recordPatch("E", "")).toBeNull();
  });
});

describe("리스트 탭 · 실제 크기 시트(서식 65536행)와 실패 안내", () => {
  const SITE = "드림 합정주공";
  const useGraphState = () => {
    Object.assign(getState(), { config: graphConfig(), records: realListRecords(), loginNeeded: false });
  };

  it("실제 합정주공 시트는 서식이 A1:AE65536 이라 usedRange 전체(값+수식)는 실패한다 — 값 범위만 확인하고 A–Y 를 나눠 읽어 순번 순서로 보인다", async () => {
    // 예전 방식(usedRange 전체)은 가짜 Graph 에서도 실제처럼 실패한다(실제: 시간 초과/504 → 화면이 비거나 '읽는 중…').
    const recs = realListRecords();
    const c = client();
    const sheets = await c.worksheetsOf(LIST_ID);
    const sheet = sheets.find((x) => x.name.trim() === SITE)!;
    await expect(c.usedRange(LIST_ID, sheet.id)).rejects.toMatchObject({ status: 504 });
    drive.log.length = 0;

    useGraphState();
    const res = await readLiveList(SITE, { client: client() });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const table = res.table;
    expect(table.sheetName.trim()).toBe(SITE);
    expect(table.rows.length).toBeGreaterThan(1000);
    const app = recs.find((r) => r.src === "user" && r.site === SITE)!;
    expect(table.rows.some((r) => r.row === app.excelRow)).toBe(true);
    const seqs = table.rows.map((r) => r.cells[0].value).filter((v) => typeof v === "number") as number[];
    expect(seqs).toEqual([...seqs].sort((a, b) => a - b));
    expect(table.header.find((h) => h.col === "F")?.title).toMatch(/line/i);
    // 수식 칸(L)은 수식으로 알아본다.
    expect(table.rows.some((r) => r.cells.find((x) => x.col === "L")?.formula)).toBe(true);
    // 큰 범위 요청 없음, 쓰기 없음, 65536 행까지 읽지 않음.
    expect(drive.log.every((e) => e.status !== 504)).toBe(true);
    expect(drive.log.some((e) => /usedRange\(valuesOnly=true\)/.test(e.path))).toBe(true);
    expect(drive.log.some((e) => /65536/.test(e.path))).toBe(false);
    expect(writes()).toEqual([]);
  });

  it("읽기 실패는 빈 화면이 아니라 한국어 이유: Graph 오류(실제 이유)·시간 초과·로그인 필요·시트 없음", async () => {
    useGraphState();
    drive.fail((m, p) => m === "GET" && p.includes("/range("), 423, Infinity);
    const locked = await readLiveList(SITE, { client: client() });
    expect(locked.ok).toBe(false);
    if (!locked.ok) expect(locked.message).toMatch(/^검사리스트를 읽지 못했습니다 — .*잠겨/);
    drive.clearFaults();

    drive.hangs.push((m, p) => m === "GET" && p.includes("/range("));
    const slow = await readLiveList(SITE, { client: client(), timeoutMs: 30 });
    expect(slow).toEqual({ ok: false, message: LIST_TIMEOUT_TEXT });
    drive.clearFaults();

    expect(listErrorText(new LoginRequiredError())).toMatch(/로그인 필요/);
    const gone = await readLiveList("없는 현장 XYZ", { client: client() });
    expect(gone.ok).toBe(false);
    if (!gone.ok) expect(gone.message).toMatch(/시트가 없습니다/);
    expect(writes()).toEqual([]);
  });

  it("동기화가 '로그인 필요'를 켜 둔 채여도(401·시간 초과 뒤) 리스트 탭은 막지 않고 읽어 본다", () => {
    Object.assign(getState(), { config: graphConfig(), loginNeeded: true, account: null });
    const { client: c, reason } = liveListClient();
    expect(c).not.toBeNull();
    expect(reason).toBe("");
    Object.assign(getState(), { config: { ...graphConfig(), syncMode: "demo" } });
    expect(liveListClient().client).toBeNull();
  });

  it("로그인 만료: 새 토큰으로도 401 이면 '읽는 중…'에 남지 않고 한국어 '로그인 필요'(다시 로그인 버튼 문구), 전역 만료", async () => {
    resetAuthForTests();
    useGraphState();
    drive.fail((m) => m === "GET", 401, Infinity);
    const res = await readLiveList(SITE, { client: client(), timeoutMs: 5_000 });
    expect(res).toEqual({ ok: false, message: LIST_LOGIN_TEXT });
    expect(getAuthState()).toBe("expired");
    expect(inflightCount()).toBe(0);
    expect(listErrorText(new AuthExpiredError())).toBe(LIST_LOGIN_TEXT);
    resetAuthForTests();
  });

  it("읽는 도중 다른 곳에서 만료(markExpired)되면 멈춘 요청이 실제로 끊기고 바로 로그인 안내", async () => {
    resetAuthForTests();
    useGraphState();
    drive.hangs.push((m, p) => m === "GET" && p.includes("/range("));
    const pending = readLiveList(SITE, { client: client(), timeoutMs: 60_000 });
    for (let i = 0; i < 50 && inflightCount() === 0; i++) await new Promise((r) => setTimeout(r, 5));
    expect(inflightCount()).toBeGreaterThan(0);
    const t0 = Date.now();
    markExpired();
    const res = await pending;
    expect(Date.now() - t0).toBeLessThan(1000);
    expect(res).toEqual({ ok: false, message: LIST_LOGIN_TEXT });
    expect(inflightCount()).toBe(0);
    resetAuthForTests();
  });
});
