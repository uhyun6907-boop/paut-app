/**
 * 세 번째 탭(리스트): OneDrive 검사리스트(.xlsx — 평택 리스트, 군포·마전 리스트)의 현장 시트를 그대로 읽고,
 * 칸 하나를 고치면 그 칸 하나만 Graph range PATCH 로 쓴다.
 * - 시트를 만들거나 복사하지 않는다. 행을 넣거나 지우지 않는다. 비고에 아무것도 덧붙이지 않는다.
 * - 쓰기 전에 그 칸을 다시 읽어 불러온 값과 다르면(다른 곳에서 고침) 쓰지 않는다. PATCH 에는 If-Match(eTag)를 건다.
 */
import { canonicalPoint } from "../domain/point";
import { FAIL, PASS } from "../domain/types";
import { excelSerialToISO, isoToExcelSerial, parseFlexibleDate } from "../domain/dates";
import type { AppConfig, InspectionRecord, ListBookKind } from "../domain/types";
import { LIST_BOOK_LABEL } from "../domain/types";
import { GraphError, type GraphCell, type GraphClient } from "../sync/graph";
import { listFileOf } from "./engine";
import { listBookOf, listSheetOf } from "./listBook";
import { columnIndex, columnLetter } from "./ooxml";

/** 보여 줄 열: A–Y(순번 … 앱 ID). */
export const LIVE_LAST_COL = 25;
/** 한 번에 읽는 행 수(A–Y 25열 × 500행). */
export const LIVE_CHUNK_ROWS = 500;
/** 동시에 읽는 묶음 수(Graph 제한·429 를 피하려고 작게). 429 는 GraphClient 가 Retry-After 만큼 기다렸다 다시 보낸다. */
export const LIVE_PARALLEL = 3;
/** 실제 기록 끝 행을 찾을 때 보는 열(날짜·라인·포인트 — 사람이 적는 칸, 수식 없음). */
export const LIVE_PROBE_COLS = "E:G";
/** 이보다 아래 행은 읽지 않는다(값이 있는 끝 행이 비정상적으로 클 때 아이폰 보호). */
export const LIVE_MAX_ROW = 20000;
/** 데이터는 6행부터(4행이 머리글). */
export const LIVE_FIRST_DATA_ROW = 6;
/** 날짜 칸(엑셀 일련번호로 쓴다). */
const DATE_COLS = new Set(["D", "E"]);
/** 고치지 않는 칸: Y(앱 ID — 앱 기록과 엑셀 행을 잇는 열). */
const LOCKED_COLS = new Set(["Y"]);

export interface LiveCell {
  /** 예: G12 */
  address: string;
  col: string;
  row: number;
  value: GraphCell;
  /** 수식 칸(예: L = RIGHT(F,4))은 고치지 않는다. */
  formula: boolean;
}

export interface LiveRow {
  row: number;
  cells: LiveCell[];
}

export interface LiveTable {
  book: ListBookKind;
  itemId: string;
  fileName: string;
  sheetId: string;
  sheetName: string;
  eTag?: string;
  header: { col: string; title: string }[];
  rows: LiveRow[];
}

export type LiveLoad = { ok: true; table: LiveTable } | { ok: false; message: string };

/** 현장 → 리스트 파일·시트(기존 listBookOf / listSheetOf 규칙 그대로). */
export function liveTarget(site: string, records: InspectionRecord[]): { book: ListBookKind; sheetName: string } {
  const recs = records.filter((r) => r.site === site);
  const probe = { sheetName: recs[0]?.sheetName ?? site, site, listBook: recs.find((r) => r.listBook)?.listBook };
  return { book: listBookOf(probe), sheetName: listSheetOf(probe) };
}

function parseBox(address: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const tail = address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, "");
  const [a, b] = tail.split(":");
  const pa = /^([A-Z]+)(\d+)$/.exec(a ?? "");
  const pb = /^([A-Z]+)(\d+)$/.exec(b ?? a ?? "");
  if (!pa || !pb) return null;
  return { r1: Number(pa[2]), c1: columnIndex(pa[1]), r2: Number(pb[2]), c2: columnIndex(pb[1]) };
}

/** A열 순번 숫자. 비었거나 숫자가 아니면 null. */
function seqOf(row: LiveRow): number | null {
  const v = row.cells.find((c) => c.col === "A")?.value;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const t = typeof v === "string" ? v.trim() : "";
  return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : null;
}

/**
 * 화면 순서: 엑셀 A열 순번 그대로(물리 행·날짜 순 아님). 순번이 없거나 숫자가 아닌 행은 뒤에, 시트 행 순서대로.
 * 엑셀 값·행은 건드리지 않고 보여 주는 순서만 바꾼다.
 */
export function sortBySeq(rows: LiveRow[]): LiveRow[] {
  return rows
    .map((row) => ({ row, seq: seqOf(row) }))
    .sort((a, b) => {
      if (a.seq !== null && b.seq !== null && a.seq !== b.seq) return a.seq - b.seq;
      if (a.seq === null && b.seq !== null) return 1;
      if (a.seq !== null && b.seq === null) return -1;
      return a.row.row - b.row.row;
    })
    .map((x) => x.row);
}

/** 칸 값을 화면 글자로. 날짜 칸의 일련번호는 2026-11-12 로. */
export function cellText(cell: Pick<LiveCell, "col" | "value">): string {
  const v = cell.value;
  if (v === null || v === undefined || v === "") return "";
  if (DATE_COLS.has(cell.col) && typeof v === "number" && v > 20000 && v < 80000) return excelSerialToISO(v);
  return String(v);
}

export function isEditable(cell: Pick<LiveCell, "col" | "row" | "formula">): boolean {
  return !cell.formula && !LOCKED_COLS.has(cell.col) && cell.row >= LIVE_FIRST_DATA_ROW;
}

/** 사용자가 적은 글자 → 엑셀에 쓸 값. 날짜 칸은 일련번호, 숫자였던 칸은 숫자. '=' 로 시작하면 수식이 아니라 글자로. */
export function toCellValue(col: string, text: string, previous: GraphCell): GraphCell {
  const t = text.trim();
  if (!t) return "";
  if (DATE_COLS.has(col)) {
    const iso = parseFlexibleDate(t);
    if (iso) return isoToExcelSerial(iso);
  }
  if (/^-?\d+(\.\d+)?$/.test(t)) return typeof previous === "number" || previous === "" || previous === null ? Number(t) : `'${t}`;
  if (t.startsWith("=")) return `'${t}`;
  return t;
}

function same(a: GraphCell | undefined, b: GraphCell | undefined): boolean {
  const norm = (x: GraphCell | undefined) => (x === null || x === undefined ? "" : String(x));
  return norm(a) === norm(b);
}

/** 현장 시트를 OneDrive 에서 바로 읽는다. 파일·시트가 없으면 이유만 돌려준다(만들지 않는다). */
export async function loadLiveList(client: GraphClient, config: AppConfig, site: string, records: InspectionRecord[]): Promise<LiveLoad> {
  const { book, sheetName } = liveTarget(site, records);
  const label = LIST_BOOK_LABEL[book];
  const file = listFileOf(config, book);
  if (!file?.id) return { ok: false, message: `${label} 파일을 아직 고르지 않았습니다. 설정 → OneDrive 통합 문서에서 고르세요.` };
  let itemId = file.id;
  let meta: { id: string; name: string; eTag?: string };
  // 파일 정보와 시트 목록은 동시에 묻는다(왕복 한 번 절약). 파일이 옮겨졌으면(404) 이름으로 찾아 다시 묻는다.
  let sheetsP = client.worksheetsOf(itemId);
  sheetsP.catch(() => undefined);
  try {
    meta = await client.itemMeta(itemId);
  } catch (error) {
    if (!(error instanceof GraphError) || error.status !== 404) throw error;
    const found = file.name ? await client.findRootFile(file.name) : null;
    if (!found) return { ok: false, message: `OneDrive 에서 ${label} 파일(${file.name || itemId})을 찾지 못했습니다.` };
    itemId = found.id;
    meta = await client.itemMeta(itemId);
    sheetsP = client.worksheetsOf(itemId);
  }
  const sheets = await sheetsP;
  const sheet = sheets.find((s) => s.name === sheetName) ?? sheets.find((s) => s.name.trim() === sheetName.trim());
  if (!sheet) return { ok: false, message: `${meta.name || label} 에 '${sheetName}' 시트가 없습니다. 시트는 앱이 만들지 않습니다(PC 엑셀에서 만들어 주세요).` };
  // 서식만 있는 빈 행(실제 시트는 65536행까지)·미리 채운 순번(A)만 있는 행은 읽지 않는다:
  // 실제 기록이 끝나는 행(날짜·라인·포인트 E–G 의 마지막 값)까지만 A–Y 를 나눠, 몇 개씩 동시에 읽는다.
  const lastRow = await lastDataRow(client, itemId, sheet.id);
  const header: LiveTable["header"] = [];
  const rows: LiveRow[] = [];
  const values: GraphCell[][] = [];
  const formulas: GraphCell[][] = [];
  const chunks: { r1: number; r2: number }[] = [];
  for (let r1 = 1; r1 <= lastRow; r1 += LIVE_CHUNK_ROWS) chunks.push({ r1, r2: Math.min(lastRow, r1 + LIVE_CHUNK_ROWS - 1) });
  await mapLimit(chunks, LIVE_PARALLEL, async ({ r1, r2 }) => {
    const part = await client.readRangeCells(itemId, sheet.id, `A${r1}:${columnLetter(LIVE_LAST_COL)}${r2}`);
    for (let i = 0; i <= r2 - r1; i++) {
      values[r1 - 1 + i] = part.values[i] ?? [];
      formulas[r1 - 1 + i] = part.formulas[i] ?? [];
    }
  });
  const at = (grid: GraphCell[][], r: number, c: number): GraphCell => grid[r - 1]?.[c - 1] ?? "";
  for (let c = 1; c <= LIVE_LAST_COL; c++) {
    const col = columnLetter(c);
    const title = String(at(values, 4, c) || at(values, 5, c) || "").trim();
    header.push({ col, title });
  }
  for (let r = LIVE_FIRST_DATA_ROW; r <= lastRow; r++) {
    const cells: LiveCell[] = [];
    let any = false;
    for (let c = 1; c <= LIVE_LAST_COL; c++) {
      const col = columnLetter(c);
      const value = at(values, r, c);
      const f = at(formulas, r, c);
      if (value !== "" && value !== null) any = true;
      cells.push({ address: `${col}${r}`, col, row: r, value, formula: typeof f === "string" && f.startsWith("=") });
    }
    if (any) rows.push({ row: r, cells });
  }
  return { ok: true, table: { book, itemId, fileName: meta.name, sheetId: sheet.id, sheetName: sheet.name, eTag: meta.eTag, header, rows: sortBySeq(rows) } };
}

/**
 * 실제 기록이 끝나는 행. E–G(날짜·라인·포인트)에 값이 있는 마지막 행을 먼저 묻고,
 * 그 요청이 안 되면(로그인 오류 제외) 예전처럼 시트 전체의 값 범위(usedRange(valuesOnly))로 정한다.
 * 머리글(4–5행)은 언제나 읽는다.
 */
export async function lastDataRow(client: GraphClient, itemId: string, sheetId: string): Promise<number> {
  const [a, b] = LIVE_PROBE_COLS.split(":");
  let last: number | null = null;
  try {
    const box = parseBox(await client.rangeValuesExtent(itemId, sheetId, `${a}1:${b}${LIVE_MAX_ROW}`));
    if (box) last = box.r2;
  } catch (error) {
    // 로그인·연결 끊김·시간 초과는 그대로 알린다. 그 밖의 Graph 오류(지원 안 함·잠김 등)면 예전 방식으로.
    if (!(error instanceof GraphError) || error.status === 0 || error.status === 401 || error.status === 403) throw error;
  }
  if (last === null) {
    const extent = parseBox(await client.valuesExtent(itemId, sheetId));
    last = extent ? extent.r2 : 0;
  }
  return Math.min(Math.max(last, LIVE_FIRST_DATA_ROW - 1), LIVE_MAX_ROW);
}

/** 작업을 최대 limit 개씩 동시에. 하나라도 실패하면 그 오류로 끝낸다. */
async function mapLimit<T>(items: T[], limit: number, run: (item: T) => Promise<void>): Promise<void> {
  let next = 0;
  let failed = false;
  const worker = async () => {
    while (!failed && next < items.length) {
      const item = items[next++];
      try {
        await run(item);
      } catch (error) {
        failed = true;
        throw error;
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
}

export class LiveConflictError extends Error {
  current: GraphCell;
  constructor(current: GraphCell) {
    super(`다른 곳(PC 엑셀 등)에서 이 칸이 이미 바뀌었습니다(지금 값: ${current === "" || current == null ? "빈 칸" : String(current)}). 새로 불러온 뒤 다시 고치세요.`);
    this.name = "LiveConflictError";
    this.current = current;
  }
}

/**
 * 칸 하나만 쓴다. 쓰기 전에 그 칸을 다시 읽어 불러온 값과 같은지 확인하고, PATCH 에 If-Match 를 건다.
 * 412(파일 eTag 가 바뀜)인데 그 칸은 그대로면 eTag 를 새로 받아 한 번만 다시 보낸다.
 * 돌려주는 값: 쓴 값.
 */
export async function writeLiveCell(client: GraphClient, table: LiveTable, cell: LiveCell, text: string): Promise<GraphCell> {
  if (!isEditable(cell)) throw new Error(cell.formula ? "수식 칸은 고치지 않습니다." : "이 칸은 고칠 수 없습니다.");
  const value = toCellValue(cell.col, text, cell.value);
  const check = async () => {
    const now = (await client.readRange(table.itemId, table.sheetId, cell.address))[0]?.[0] ?? "";
    if (!same(now, cell.value)) throw new LiveConflictError(now);
  };
  await check();
  try {
    await client.patchValues(table.itemId, table.sheetId, cell.address, [[value]], table.eTag);
  } catch (error) {
    if (!(error instanceof GraphError) || error.status !== 412) throw error;
    await check();
    table.eTag = (await client.itemMeta(table.itemId)).eTag;
    await client.patchValues(table.itemId, table.sheetId, cell.address, [[value]], table.eTag);
  }
  cell.value = typeof value === "string" && value.startsWith("'") ? value.slice(1) : value;
  try {
    table.eTag = (await client.itemMeta(table.itemId)).eTag;
  } catch {
    /* 다음 쓰기 때 412 면 다시 받는다 */
  }
  return value;
}

/** 열 → 앱 기록 필드. 없으면 앱 기록에는 해당 칸이 없다(순번 등). */
const FIELD_OF: Record<string, keyof InspectionRecord> = {
  B: "issued",
  C: "reportNo",
  D: "reportDate",
  E: "date",
  F: "line",
  G: "point",
  H: "defectNo",
  I: "drawingNo",
  J: "weldNo",
  K: "material",
  L: "diameter",
  M: "sch",
  N: "rtSheets",
  P: "result",
  Q: "flawType",
  R: "flawLen",
  S: "note",
};

/**
 * 엑셀 행과 같은 앱 기록: Y(앱 ID) → 엑셀에서 읽은 행 ID(xlsx:시트:행) → 같은 파일·시트·행 번호이면서 라인·포인트가 같은 기록.
 */
export function matchRecord(records: InspectionRecord[], table: Pick<LiveTable, "book" | "sheetName">, row: LiveRow): InspectionRecord | null {
  const text = (col: string) => {
    const c = row.cells.find((x) => x.col === col);
    return c ? cellText(c).trim() : "";
  };
  const appId = text("Y");
  const alive = records.filter((r) => !r.deleted);
  if (appId) {
    const hit = alive.find((r) => r.id === appId);
    if (hit) return hit;
  }
  const xid = alive.find((r) => r.id === `xlsx:${table.sheetName}:${row.row}` || r.id === `xlsx:${table.sheetName.trim()}:${row.row}`);
  if (xid) return xid;
  const line = text("F").toUpperCase();
  const point = pointKey(text("G"));
  return (
    alive.find(
      (r) =>
        r.excelRow === row.row &&
        listBookOf(r) === table.book &&
        listSheetOf(r) === table.sheetName.trim() &&
        r.line.trim().toUpperCase() === line &&
        pointKey(r.point) === point,
    ) ?? null
  );
}

/** R2 와 R02 는 같은 포인트. */
function pointKey(point: string): string {
  const t = point.trim().toUpperCase();
  return canonicalPoint(t) ?? t;
}

/** P(판정) 글자 → 합격/불합격. 알 수 없는 글자는 그대로. */
export function normalizeResult(text: string): string {
  const t = text.trim().replace(/\s+/g, "").toUpperCase();
  if (["합격", "합", "O", "○", "PASS", "ACC", "ACCEPT"].includes(t)) return PASS;
  if (["불합격", "불", "X", "×", "FAIL", "REJ", "REJECT"].includes(t)) return FAIL;
  return text.trim();
}

/** 고친 칸을 앱 기록 필드 값으로. 필드가 없는 열이면 null. */
export function recordPatch(col: string, value: GraphCell): Partial<InspectionRecord> | null {
  const field = FIELD_OF[col];
  if (!field) return null;
  const raw = typeof value === "string" && value.startsWith("'") ? value.slice(1) : value;
  let text = raw === null || raw === undefined ? "" : String(raw);
  if (DATE_COLS.has(col) && typeof raw === "number") text = excelSerialToISO(raw);
  if (col === "L") text = text.toUpperCase();
  if (col === "G") text = pointKey(text);
  if (col === "F") text = text.trim();
  if (col === "P") text = normalizeResult(text);
  if (DATE_COLS.has(col) && !text) return null; // 날짜를 지우면 카드 날짜는 그대로
  if (DATE_COLS.has(col) && !/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const iso = parseFlexibleDate(text);
    if (!iso) return null; // 날짜가 아닌 글자는 카드 날짜로 옮기지 않는다
    text = iso;
  }
  return { [field]: text } as Partial<InspectionRecord>;
}
