import { excelSerialToISO, excelTimeToHHMM, hhmmToExcelTime, isoToExcelSerial } from "../domain/dates";
import { mentionsRemovedSite } from "../domain/mapping";
import type { AppConfig, DriveFile } from "../domain/types";
import { DriveFileMissing, type GraphCell, type GraphClient } from "../sync/graph";
import { fileOf } from "./engine";
import { detectPairs, detectStaffColumns, findIljiSheet, type IljiPair } from "./iljiBook";
import { LiveConflictError } from "./liveList";
import { columnIndex, columnLetter, readSheet, unzipXlsx, type XCell, type ZipFiles } from "./ooxml";

export const ILJI_HEAD_ROWS = [4, 5] as const;
export const ILJI_FIRST_ROW = 6;

export type IljiKind = "seq" | "date" | "dayType" | "dia" | "count" | "time" | "text";
export type IljiLock = "seq" | "formula" | "merged" | "excluded";

export const LOCK_TEXT: Record<IljiLock, string> = {
  seq: "순번(A)은 고치지 않습니다.",
  formula: "수식 칸은 고치지 않습니다.",
  merged: "병합된 칸입니다. 병합의 첫 칸을 고치세요.",
  excluded: "제외된 현장(지성 브레인·드림 택지·드림 화양·브레인대지 인입) 칸은 고치지 않습니다.",
};

export interface IljiCell {
  address: string;
  col: string;
  row: number;
  value: GraphCell;
  kind: IljiKind;
  lock: IljiLock | null;
  span?: { rows: number; cols: number };
  coveredBy?: string;
}

export interface IljiRow {
  row: number;
  date: string;
  /** 이 행이 속한 날의 달(YYYY-MM). 날짜 없는 이어쓰기 행은 위 날짜 행의 달. */
  month: string;
  cells: IljiCell[];
}

export interface IljiHead {
  col: string;
  text: string;
  rowSpan: number;
  colSpan: number;
}

export interface IljiTable {
  file: DriveFile;
  sheetName: string;
  cols: string[];
  head: IljiHead[][];
  rows: IljiRow[];
  months: string[];
  staff: { insp1: string; insp2: string; start: string; end: string };
}

export type IljiLoad = { ok: true; table: IljiTable } | { ok: false; message: string };

interface Box {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function mergesOf(xml: string): Box[] {
  return [...xml.matchAll(/<mergeCell\b[^>]*\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g)].map((m) => ({
    r1: Number(m[2]),
    c1: columnIndex(m[1]),
    r2: Number(m[4]),
    c2: columnIndex(m[3]),
  }));
}

function valueOf(cell: XCell | undefined): GraphCell {
  const v = cell?.v;
  if (v == null || v === "") return "";
  if (cell!.t === "s" || cell!.t === "inlineStr" || cell!.t === "str") return v;
  const n = Number(v);
  return Number.isFinite(n) ? n : v;
}

function serialDate(value: GraphCell): string {
  return typeof value === "number" && value > 20000 && value < 80000 ? excelSerialToISO(value) : "";
}

export async function loadIlji(client: GraphClient, config: AppConfig): Promise<IljiLoad> {
  const file = fileOf(config, "ilji");
  if (!file?.id) return { ok: false, message: "작업일지 파일을 아직 고르지 않았습니다. 설정 → OneDrive 통합 문서에서 고르세요." };
  const got = await client.withFile(file, (id) => client.download(id)).catch((error: unknown) => {
    if (error instanceof DriveFileMissing) return error;
    throw error;
  });
  if (got instanceof DriveFileMissing) return { ok: false, message: got.message };
  return { ok: true, table: parseIlji(unzipXlsx(got.value), got.file) };
}

/** 받은 작업일지 파일 → 표. 6행부터 B–U 에 값이 있는 마지막 행까지(미리 적힌 순번만 있는 아래 행은 뺀다). */
export function parseIlji(files: ZipFiles, file: DriveFile): IljiTable {
  const sheetName = findIljiSheet(files);
  const { xml, cells } = readSheet(files, sheetName);
  const merges = mergesOf(xml);
  let lastCol = 0;
  for (const c of cells.values()) if (c.row === 4 && c.v?.trim()) lastCol = Math.max(lastCol, columnIndex(c.col));
  for (const m of merges) if (m.r1 >= 4 && m.r2 <= 5) lastCol = Math.max(lastCol, m.c2);
  const cols = Array.from({ length: lastCol }, (_, i) => columnLetter(i + 1));
  const pairs = detectPairs(cells);
  const staff = detectStaffColumns(cells);
  let lastRow = ILJI_FIRST_ROW - 1;
  for (const c of cells.values()) {
    if (c.row >= ILJI_FIRST_ROW && c.col !== "A" && columnIndex(c.col) <= lastCol && c.v?.trim()) lastRow = Math.max(lastRow, c.row);
  }
  const mergeAt = (r: number, c: number) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);

  const head = ILJI_HEAD_ROWS.map((r) =>
    cols.flatMap((col, i) => {
      const m = mergeAt(r, i + 1);
      if (m && (m.r1 !== r || m.c1 !== i + 1)) return [];
      return [{ col, text: (cells.get(col + r)?.v ?? "").trim(), rowSpan: m ? Math.min(m.r2, 5) - r + 1 : 1, colSpan: m ? Math.min(m.c2, lastCol) - m.c1 + 1 : 1 }];
    }),
  );

  const kindOf = new Map<string, IljiKind>([["A", "seq"], ["B", "date"], ["C", "dayType"], [staff.start, "time"], [staff.end, "time"]]);
  for (const p of pairs) {
    kindOf.set(p.diaCol, "dia");
    kindOf.set(p.countCol, "count");
  }

  const rows: IljiRow[] = [];
  let month = "";
  for (let r = ILJI_FIRST_ROW; r <= lastRow; r++) {
    const date = serialDate(valueOf(cells.get(`B${r}`)));
    if (date) month = date.slice(0, 7);
    const row: IljiRow = {
      row: r,
      date,
      month,
      cells: cols.map((col, i) => {
        const c = cells.get(col + r);
        const m = mergeAt(r, i + 1);
        const anchor = m && m.r1 === r && m.c1 === i + 1;
        const cell: IljiCell = { address: col + r, col, row: r, value: valueOf(c), kind: kindOf.get(col) ?? "text", lock: null };
        if (m && anchor && (m.r2 > m.r1 || m.c2 > m.c1)) cell.span = { rows: m.r2 - m.r1 + 1, cols: Math.min(m.c2, lastCol) - m.c1 + 1 };
        if (m && !anchor) cell.coveredBy = columnLetter(m.c1) + m.r1;
        cell.lock = cell.kind === "seq" ? "seq" : c?.f != null ? "formula" : cell.coveredBy ? "merged" : null;
        return cell;
      }),
    };
    lockExcluded(row, pairs);
    rows.push(row);
  }
  const first = rows.find((r) => r.month)?.month ?? "";
  for (const row of rows) if (!row.month) row.month = first;
  return { file, sheetName, cols, head, rows, months: [...new Set(rows.map((r) => r.month).filter(Boolean))].sort(), staff };
}

function lockExcluded(row: IljiRow, pairs: IljiPair[]): void {
  const at = (col: string) => row.cells.find((c) => c.col === col);
  let filled = 0;
  let excluded = 0;
  for (const p of pairs) {
    const dia = at(p.diaCol);
    const count = at(p.countCol);
    if (!dia || !count) continue;
    const hasValue = String(dia.value ?? "") !== "" || String(count.value ?? "") !== "";
    if (hasValue) filled++;
    if (!mentionsRemovedSite(p.column) && !mentionsRemovedSite(String(dia.value ?? ""))) continue;
    if (hasValue) excluded++;
    for (const c of [dia, count]) c.lock ??= "excluded";
  }
  if (filled && filled === excluded) for (const c of row.cells) c.lock ??= "excluded";
}

export function iljiText(cell: Pick<IljiCell, "kind" | "value">): string {
  const v = cell.value;
  if (v === null || v === "") return "";
  if (cell.kind === "date") {
    const iso = serialDate(v);
    return iso ? iso.slice(2).replace(/-/g, ".") : String(v);
  }
  if (cell.kind === "time" && typeof v === "number" && v >= 0 && v < 2) return excelTimeToHHMM(v);
  return String(v);
}

export function iljiInput(cell: Pick<IljiCell, "kind" | "value">): string {
  if (cell.kind === "date") return serialDate(cell.value) || String(cell.value ?? "");
  return iljiText(cell);
}

function roundTime(n: number): number {
  return Math.round(n * 1e10) / 1e10;
}

/**
 * 입력 글자 → 엑셀에 쓸 값. 날짜는 일련번호, 시간은 하루 비율(자정을 넘긴 칸은 그 날짜 차이를 지킨다),
 * 개소는 숫자. '=' '+' '-' 로 시작하는 글자는 수식이 되지 않게 ' 를 붙인다.
 */
export function iljiValue(cell: Pick<IljiCell, "kind" | "value" | "address">, text: string): GraphCell {
  const t = text.trim();
  if (!t) return "";
  if (cell.kind === "date") {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) throw new Error(`${cell.address}: 날짜는 2026-11-03 처럼 고르세요.`);
    return isoToExcelSerial(t);
  }
  if (cell.kind === "time") {
    if (!/^\d{1,2}:\d{2}$/.test(t)) throw new Error(`${cell.address}: 시간은 09:00 처럼 고르세요.`);
    const day = typeof cell.value === "number" && cell.value >= 1 ? Math.floor(cell.value) : 0;
    return roundTime(day + hhmmToExcelTime(t));
  }
  if (/^-?\d+(\.\d+)?$/.test(t) && (cell.kind === "count" || typeof cell.value === "number")) return Number(t);
  return /^[=+-]/.test(t) ? `'${t}` : t;
}

export function sameIlji(a: GraphCell | undefined, b: GraphCell | undefined): boolean {
  const norm = (x: GraphCell | undefined) => (x === null || x === undefined ? "" : typeof x === "string" && x.startsWith("'") ? x.slice(1) : String(x).trim());
  const x = norm(a);
  const y = norm(b);
  if (x !== "" && y !== "" && Number.isFinite(Number(x)) && Number.isFinite(Number(y))) return Math.abs(Number(x) - Number(y)) < 1e-7;
  return x === y;
}

export interface IljiEditPlan {
  cell: IljiCell;
  value: GraphCell;
  before: string;
  after: string;
}

/** 고칠 수 있는지·쓸 값·확인 문구에 넣을 전/후 글자. 고칠 수 없거나 바뀐 게 없으면 이유를 던진다. */
export function planIljiEdit(table: IljiTable, cell: IljiCell, text: string): IljiEditPlan {
  if (cell.lock) throw new Error(`${cell.address}: ${LOCK_TEXT[cell.lock]}`);
  const value = iljiValue(cell, text);
  if (sameIlji(value, cell.value)) throw new Error(`${cell.address}: 바뀐 값이 없습니다.`);
  if (cell.kind === "date" && value !== "") {
    const dup = table.rows.find((r) => r.row !== cell.row && r.cells.some((c) => c.col === "B" && sameIlji(c.value, value)));
    if (dup) throw new Error(`${cell.address}: ${dup.row}행에 이미 같은 날짜가 있습니다. 같은 날짜 행을 두 개 만들지 않습니다.`);
  }
  const shown = (v: GraphCell) => iljiText({ kind: cell.kind, value: typeof v === "string" && v.startsWith("'") ? v.slice(1) : v }) || "빈 칸";
  return { cell, value, before: shown(cell.value), after: shown(value) };
}

/**
 * 칸 하나를 쓴다: 통합 문서 세션을 열고 → 그 칸을 다시 읽어 불러온 값과 같은지 확인(다르면 LiveConflictError, 쓰지 않음) →
 * 쓰고 → 다시 읽어 확인한다. 저장된 항목 ID 가 404 면 이름으로 찾는다(돌려주는 file 을 설정에 남긴다).
 */
export async function writeIljiCell(client: GraphClient, table: IljiTable, plan: IljiEditPlan): Promise<{ value: GraphCell; file: DriveFile }> {
  const { cell, value } = plan;
  const opened = await client.withFile(table.file, (id) => client.openWorkbook(id));
  const wb = opened.value;
  try {
    const ws = await wb.sheet(table.sheetName);
    if (!ws) throw new Error(`작업일지 시트(${table.sheetName.trim()})를 OneDrive 에서 찾지 못했습니다.`);
    const now = (await wb.readValues(ws.id, cell.address))[0]?.[0] ?? "";
    if (!sameIlji(now, cell.value)) throw new LiveConflictError(now, cell.address);
    await wb.writeFormulas(ws.id, cell.address, [[value]]);
    const after = (await wb.readValues(ws.id, cell.address))[0]?.[0] ?? "";
    if (!sameIlji(after, value)) throw new Error(`${cell.address} 에 썼지만 다시 읽은 값이 다릅니다(${after === "" ? "빈 칸" : String(after)}). 새로 불러와 확인하세요.`);
    return { value: after, file: opened.file };
  } finally {
    await wb.close();
  }
}

/** 날짜 행의 검사자·근무시간(P·Q·S·T)과 구분(C). 앱의 그날 값(dayStaff)을 엑셀에 맞출 때 쓴다. */
export function rowStaff(table: IljiTable, row: IljiRow): { dayType: string; insp1: string; insp2: string; timeStart: string; timeEnd: string } {
  const text = (col: string) => {
    const c = row.cells.find((x) => x.col === col);
    return c ? iljiText(c).trim() : "";
  };
  return { dayType: text("C"), insp1: text(table.staff.insp1), insp2: text(table.staff.insp2), timeStart: text(table.staff.start), timeEnd: text(table.staff.end) };
}
