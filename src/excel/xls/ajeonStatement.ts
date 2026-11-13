/**
 * 아전 거래명세서(.xls, OneDrive '00_PAUT 거래명세서/01_PAUT 거래명세서 (아전, 승재).xls') 채우기.
 *
 * 원칙
 * - 파일은 .xls 그대로 둔다. 받은 바이트에서 '그 달 아전 시트'의 셀만 바꾸고(biff.ts), 같은 항목에 eTag(If-Match)로 올린다.
 * - 승재 시트는 읽지도 쓰지도 않는다(바이트 비교로 그대로인지 확인만 한다). 시트를 만들거나 이름을 바꾸지 않는다.
 * - 그 달 시트가 없으면 쓰지 않고 '11월 거래명세서 시트를 만들어주세요 (아전 군포 2611)' 로 기다린다.
 * - 단가는 파일에서 읽는다: 그 시트의 PAUT 행 → 없으면 같은 접두어의 앞 달 시트. 못 찾으면 쓰지 않는다(지어내지 않음).
 * - 합격만, 포인트 1개 = 1개소. 비고(I)나 다른 칸에 앱 표시를 남기지 않는다.
 */
import { findMonthSheet, latestSheet, matchStatementSheet, sheetSortKey, suggestSheetName } from "../../domain/mapping";
import { isoToExcelSerial, excelSerialToISO } from "../../domain/dates";
import { BiffError, cellKey, parseWorkbook, patchSheetCells, readSubstream, readXlsSheet, rgceMulSameRow, sheetBytesWithoutIndexOffsets, type CellEdit, type XlsCell, type XlsSheet, type XlsWorkbook } from "./biff";
import { entryBytes, readAllStreams, readCfb, readStream, replaceStream } from "./cfb";

/** 아전 거래명세서 파일(OneDrive). 앱은 이 항목만 받고, 이 항목에만 (eTag 확인 뒤) 올린다. */
export const AJEON_STATEMENT_FILE = {
  id: "73DF94A06E5230E0!scace5f6140ca47e59840c2863dfd2401",
  name: "01_PAUT 거래명세서 (아전, 승재).xls",
  path: "00_PAUT 거래명세서/01_PAUT 거래명세서 (아전, 승재).xls",
} as const;

export interface XlsRoute {
  /** 시트 이름 앞부분. 예: '아전 군포' → '아전 군포 2611' */
  prefix: string;
  /** true: 그 달 시트가 파일에 있을 때만 쓰고, 없으면 조용히 건너뛴다(만들어 달라고 하지 않음). */
  optional: boolean;
}

/** 현장 → 아전 .xls 명세서. 여기 없는 현장은 지성·드림 .xlsx 명세서로 간다. */
export function xlsStatementRoute(site: string): XlsRoute | null {
  const s = site.replace(/\s+/g, "");
  if (s.includes("아전군포")) return { prefix: "아전 군포", optional: false };
  if (s.includes("아전마곡")) return { prefix: "아전 마곡", optional: true };
  return null;
}

export function isXlsStatementSite(site: string): boolean {
  return xlsStatementRoute(site) != null;
}

/** 그 달 시트 이름(사용자가 만들 이름). 파일의 같은 접두어 최근 시트 표기를 따른다(예: '아전 군포 2610' → '아전 군포 2611'). */
export function expectedXlsSheetName(names: string[], prefix: string, year: number, month: number): string {
  const found = findMonthSheet(names, prefix, year, month);
  if (found) return found.name;
  const latest = latestSheet(names, prefix);
  if (latest) return suggestSheetName(prefix, latest, year, month);
  // 파일에 그 접두어 시트가 하나도 없으면 옛 사본 표기(공백, 점 없음)
  return `${prefix} ${String(year).slice(-2)}${String(month).padStart(2, "0")}`;
}

export interface XlsBook {
  bytes: Uint8Array;
  wb: XlsWorkbook;
}

export function openXlsBook(bytes: Uint8Array): XlsBook {
  const cfb = readCfb(bytes);
  return { bytes, wb: parseWorkbook(readStream(cfb, "Workbook")) };
}

export function xlsSheetNames(book: XlsBook): string[] {
  return book.wb.sheets.filter((s) => s.kind === 0).map((s) => s.name);
}

/* ───────────────────────── 시트 모양 ───────────────────────── */

export interface AjeonLayout {
  headerRow: number;
  totalRow: number;
  col: { date: number; item: number; diameter: number; qty: number; price: number; amount: number; note: number };
}

function text(sheet: XlsSheet, row: number, col: number): string {
  const c = sheet.cells.get(cellKey(row, col));
  if (!c) return "";
  const v = c.value;
  if (v.kind === "string") return v.text.trim();
  if (v.kind === "number") return String(v.n);
  if (v.kind === "formula") return v.cached == null ? "" : String(v.cached);
  return "";
}

function num(sheet: XlsSheet, row: number, col: number): number | null {
  const c = sheet.cells.get(cellKey(row, col));
  if (!c) return null;
  if (c.value.kind === "number") return c.value.n;
  if (c.value.kind === "formula" && typeof c.value.cached === "number") return c.value.cached;
  if (c.value.kind === "string" && /^\d+(\.\d+)?$/.test(c.value.text.replace(/,/g, "").trim())) return Number(c.value.text.replace(/,/g, ""));
  return null;
}

const squash = (s: string) => s.replace(/\s+/g, "");

/** 머리글 행(일자·품명·관경·수량·단가·금액)과 합계 행을 찾는다. 모양이 다르면 BiffError. */
export function ajeonLayout(sheet: XlsSheet): AjeonLayout {
  let maxRow = 0;
  for (const c of sheet.cells.values()) maxRow = Math.max(maxRow, c.row);
  for (let r = 0; r <= Math.min(maxRow, 60); r++) {
    const col: Partial<AjeonLayout["col"]> = {};
    for (let c = 0; c < 16; c++) {
      const t = squash(text(sheet, r, c));
      if (!t) continue;
      if (t === "일자" && col.date == null) col.date = c;
      else if (t === "품명" && col.item == null) col.item = c;
      else if (t === "관경" && col.diameter == null) col.diameter = c;
      else if (t === "수량" && col.qty == null) col.qty = c;
      else if (t === "단가" && col.price == null) col.price = c;
      else if (t === "금액" && col.amount == null) col.amount = c;
      else if (t === "비고" && col.note == null) col.note = c;
    }
    if (col.date == null || col.item == null || col.diameter == null || col.qty == null || col.price == null || col.amount == null) continue;
    let totalRow = -1;
    for (let rr = r + 1; rr <= maxRow; rr++) {
      if (squash(text(sheet, rr, col.date)).includes("합계")) {
        totalRow = rr;
        break;
      }
    }
    if (totalRow < 0) throw new BiffError(`'${sheet.info.name}' 에서 합계 행을 찾지 못했습니다.`);
    return { headerRow: r, totalRow, col: { ...(col as AjeonLayout["col"]), note: col.note ?? -1 } };
  }
  throw new BiffError(`'${sheet.info.name}' 의 머리글(일자·품명·관경·수량·단가·금액)을 찾지 못했습니다.`);
}

export interface AjeonRow {
  row: number;
  kind: "paut" | "blank" | "other";
  date: string;
  diameter: string;
  count: number;
  price: number | null;
}

function isPautItem(item: string): boolean {
  const s = squash(item).toUpperCase();
  return s.includes("위상배열") || s.includes("PAUT");
}

function diameterOf(raw: string): string {
  const m = /(\d{2,4})\s*A/i.exec(raw);
  return m ? `${m[1]}A` : "";
}

export function ajeonRows(sheet: XlsSheet, layout: AjeonLayout): AjeonRow[] {
  const { col } = layout;
  const out: AjeonRow[] = [];
  for (let r = layout.headerRow + 1; r < layout.totalRow; r++) {
    const item = text(sheet, r, col.item);
    const d = num(sheet, r, col.date);
    const dateText = text(sheet, r, col.date);
    const dia = text(sheet, r, col.diameter);
    const qty = text(sheet, r, col.qty);
    const price = num(sheet, r, col.price);
    const base = {
      row: r,
      date: d != null && d > 20000 && d < 80000 ? excelSerialToISO(d) : dateText,
      diameter: diameterOf(dia || item),
      count: num(sheet, r, col.qty) ?? 0,
      price,
    };
    if (!item && !dateText && !dia && !qty && price == null) out.push({ ...base, kind: "blank" });
    else out.push({ ...base, kind: isPautItem(item) ? "paut" : "other" });
  }
  return out;
}

/** 관경 → 단가(그 시트 PAUT 행의 최빈값). */
export function pricesOnSheet(sheet: XlsSheet): Record<string, number> {
  const layout = ajeonLayout(sheet);
  const seen = new Map<string, Map<number, number>>();
  for (const r of ajeonRows(sheet, layout)) {
    if (r.kind !== "paut" || !r.diameter || r.price == null || r.price <= 0) continue;
    const m = seen.get(r.diameter) ?? new Map<number, number>();
    m.set(r.price, (m.get(r.price) ?? 0) + 1);
    seen.set(r.diameter, m);
  }
  const out: Record<string, number> = {};
  for (const [dia, m] of seen) out[dia] = [...m.entries()].sort((a, b) => b[1] - a[1])[0][0];
  return out;
}

/** 단가: 그 시트 → 같은 접두어의 앞 달 시트(가까운 달부터). 출처 시트 이름도 돌려준다. */
export function findPrice(book: XlsBook, prefix: string, target: string, diameter: string): { price: number; from: string } | null {
  const names = xlsSheetNames(book);
  const here = matchStatementSheet(target, prefix);
  const order = names
    .map((n) => matchStatementSheet(n, prefix))
    .filter((x): x is NonNullable<typeof x> => !!x && !/AM/i.test(x.suffix))
    .filter((x) => x.name === target || !here || sheetSortKey(x) <= sheetSortKey(here))
    .sort((a, b) => (a.name === target ? -1 : b.name === target ? 1 : sheetSortKey(b) - sheetSortKey(a)));
  for (const ref of order) {
    try {
      const p = pricesOnSheet(readXlsSheet(book.wb, ref.name))[diameter.toUpperCase()];
      if (p) return { price: p, from: ref.name };
    } catch {
      // 모양이 다른 시트는 건너뛴다.
    }
  }
  return null;
}

/* ───────────────────────── 채우기 계획 ───────────────────────── */

export interface AjeonLine {
  date: string;
  diameter: string;
  count: number;
}

/** 시트에서 앱이 쓴 줄: `날짜|관경` → 쓴 개수. 이 기기에 기억한다(엑셀에는 표시 없음). */
export type AjeonOwnership = Record<string, number>;

const lineKey = (date: string, diameter: string) => `${date}|${diameter.toUpperCase()}`;

export interface AjeonPlan {
  edits: CellEdit[];
  owned: AjeonOwnership;
  /** 단가를 파일에서 못 찾아 쓰지 않은 관경 */
  missingPrice: string[];
  /** 사람이 이미 적어 둔 날짜·관경(건너뜀) */
  manual: string[];
  /** 쓴 줄 수 */
  written: number;
  prices: Record<string, { price: number; from: string }>;
}

const DEFAULT_ITEM = "위상배열초음파탐상검사";

/**
 * 그 달 시트에 쓸 셀. 규칙은 지성·드림 명세서와 같다:
 * - 사람이 같은 날짜·관경을 적어 두었으면 그 줄은 건너뛴다.
 * - 앱이 지난번에 쓴 줄(owned 와 개수가 같음)과 빈 줄에만 쓴다. 빈 줄이 모자라면 쓰지 않고 오류(행을 끼워 넣지 않음).
 * - 더는 없는 앱 줄은 비운다(서식은 그대로).
 * - 서식(XF)은 그 시트의 마지막 PAUT 행을 따른다(사람이 위 줄을 복사해 적는 것과 같게). PAUT 행이 없으면 그 칸 서식 그대로.
 */
export function planAjeonFill(book: XlsBook, prefix: string, sheetName: string, lines: AjeonLine[], owned: AjeonOwnership = {}): AjeonPlan {
  const sheet = readXlsSheet(book.wb, sheetName);
  const layout = ajeonLayout(sheet);
  const { col } = layout;
  const rows = ajeonRows(sheet, layout);
  const isApp = (r: AjeonRow) => r.kind === "paut" && owned[lineKey(r.date, r.diameter)] != null && owned[lineKey(r.date, r.diameter)] === r.count;
  const manualKeys = new Set(rows.filter((r) => r.kind === "paut" && !isApp(r)).map((r) => lineKey(r.date, r.diameter)));
  const sorted = [...lines].sort((a, b) => a.date.localeCompare(b.date) || a.diameter.localeCompare(b.diameter));
  const manual = sorted.filter((l) => manualKeys.has(lineKey(l.date, l.diameter))).map((l) => lineKey(l.date, l.diameter));
  const prices: AjeonPlan["prices"] = {};
  const missingPrice: string[] = [];
  const todo: (AjeonLine & { price: number })[] = [];
  for (const l of sorted) {
    if (manualKeys.has(lineKey(l.date, l.diameter)) || l.count <= 0) continue;
    const dia = l.diameter.toUpperCase();
    if (!prices[dia]) {
      const found = findPrice(book, prefix, sheetName, dia);
      if (found) prices[dia] = found;
    }
    if (!prices[dia]) {
      if (!missingPrice.includes(dia)) missingPrice.push(dia);
      continue;
    }
    todo.push({ ...l, diameter: dia, price: prices[dia].price });
  }
  const writable = rows.filter((r) => isApp(r) || r.kind === "blank");
  if (writable.length < todo.length) {
    throw new BiffError(`'${sheetName}' 에 빈 줄이 모자랍니다(필요 ${todo.length}, 빈 줄 ${writable.length}). 엑셀에서 합계 위에 줄을 넣어 주세요.`);
  }
  // 서식 원본: 마지막 PAUT 행(사람이 적은 줄). 없으면 각 칸 자기 서식.
  const styleRow = [...rows].reverse().find((r) => r.kind === "paut" && !isApp(r))?.row ?? [...rows].reverse().find((r) => r.kind === "paut")?.row;
  const xfOf = (_row: number, c: number): number | undefined => {
    if (styleRow == null) return undefined;
    return sheet.cells.get(cellKey(styleRow, c))?.xf;
  };
  const itemText = (styleRow != null ? text(sheet, styleRow, col.item) : "") || DEFAULT_ITEM;
  const edits: CellEdit[] = [];
  const nextOwned: AjeonOwnership = {};
  let written = 0;
  writable.forEach((slot, i) => {
    const line = todo[i];
    const r = slot.row;
    if (!line) {
      if (slot.kind === "paut") {
        for (const c of [col.date, col.item, col.diameter, col.qty, col.price, col.amount]) edits.push({ row: r, col: c, value: { kind: "blank" } });
      }
      return;
    }
    edits.push({ row: r, col: col.date, value: { kind: "number", n: isoToExcelSerial(line.date) }, xf: xfOf(r, col.date) });
    edits.push({ row: r, col: col.item, value: { kind: "string", text: itemText }, xf: xfOf(r, col.item) });
    edits.push({ row: r, col: col.diameter, value: { kind: "string", text: line.diameter }, xf: xfOf(r, col.diameter) });
    edits.push({ row: r, col: col.qty, value: { kind: "number", n: line.count }, xf: xfOf(r, col.qty) });
    edits.push({ row: r, col: col.price, value: { kind: "number", n: line.price }, xf: xfOf(r, col.price) });
    edits.push({ row: r, col: col.amount, value: { kind: "formula", rgce: rgceMulSameRow(r, col.qty, col.price) }, xf: xfOf(r, col.amount) });
    nextOwned[lineKey(line.date, line.diameter)] = line.count;
    written += 1;
  });
  // 이미 같은 값이면(다시 보냄) 바꾸지 않는다.
  const effective = edits.filter((e) => !sameCell(sheet.cells.get(cellKey(e.row, e.col)), e));
  return { edits: effective, owned: nextOwned, missingPrice, manual, written, prices };
}

function sameCell(c: XlsCell | undefined, e: CellEdit): boolean {
  if (!c) return e.value.kind === "blank" && e.xf == null;
  if (e.xf != null && e.xf !== c.xf) return false;
  const v = e.value;
  if (v.kind === "blank") return c.value.kind === "blank";
  if (v.kind === "number") return c.value.kind === "number" && c.value.n === v.n;
  if (v.kind === "string") return c.value.kind === "string" && c.value.text === v.text;
  if (c.value.kind !== "formula") return false;
  const a = c.value.rgce;
  return a.length === v.rgce.length && a.every((x, i) => x === v.rgce[i]);
}

/* ───────────────────────── 쓰기 + 검증 ───────────────────────── */

export interface SheetEdits {
  sheet: string;
  edits: CellEdit[];
}

/** 받은 .xls 바이트에 시트별 셀 변경을 적용한 새 바이트. 적용 뒤 verifyXlsEdit 로 확인한다(어긋나면 오류, 올리지 않음). */
export function applyXlsEdits(bytes: Uint8Array, plans: SheetEdits[]): Uint8Array {
  const cfb = readCfb(bytes);
  let stream = readStream(cfb, "Workbook");
  for (const p of plans) {
    if (!p.edits.length) continue;
    stream = patchSheetCells(stream, p.sheet, p.edits).stream;
  }
  const out = replaceStream(bytes, "Workbook", stream);
  verifyXlsEdit(bytes, out, plans);
  return out;
}

const PROTECTED = /승재/;

/**
 * 독립 검증. 하나라도 어긋나면 BiffError(올리지 않는다).
 * - 컨테이너: Workbook 말고 다른 스트림·디렉터리 항목(크기·시작 섹터 제외) 바이트가 같다.
 * - 전역: BOUNDSHEET 시작 위치 말고 바이트가 같다. 시트 수·이름·순서·숨김이 같다(새 시트 없음).
 * - 대상이 아닌 시트(승재 포함): INDEX 절대 위치를 빼면 바이트가 같다.
 * - 대상 시트: 바꾸기로 한 칸 말고는 값·서식·수식이 같고, 셀이 아닌 레코드(병합·인쇄·열 너비·그림 등)는 ROW·DIMENSIONS·DBCELL·INDEX 말고 바이트가 같다.
 * - 바뀐 칸에 '앱' 같은 표시가 없다.
 */
export function verifyXlsEdit(before: Uint8Array, after: Uint8Array, plans: SheetEdits[]): void {
  const fail = (m: string): never => {
    throw new BiffError(`쓰기 검증 실패: ${m}`);
  };
  const c0 = readCfb(before);
  const c1 = readCfb(after);
  const s0 = readAllStreams(c0);
  const s1 = readAllStreams(c1);
  if (s0.size !== s1.size) fail("OLE 스트림 수가 바뀌었습니다.");
  for (const [name, data] of s0) {
    if (name.toLowerCase() === "workbook") continue;
    const other = s1.get(name);
    if (!other || !bytesEqual(data, other)) fail(`'${name}' 스트림이 바뀌었습니다.`);
  }
  for (const e of c0.entries) {
    const a = entryBytes(c0, e.index);
    const b = entryBytes(c1, e.index);
    for (let i = 0; i < 128; i++) {
      if (e.name.toLowerCase() === "workbook" && i >= 116 && i < 128) continue;
      if (a[i] !== b[i]) fail(`OLE 디렉터리 항목 '${e.name}' 이 바뀌었습니다.`);
    }
  }
  const w0 = parseWorkbook(readStream(c0, "Workbook"));
  const w1 = parseWorkbook(readStream(c1, "Workbook"));
  if (w0.sheets.length !== w1.sheets.length) fail("시트 수가 바뀌었습니다.");
  w0.sheets.forEach((s, i) => {
    const t = w1.sheets[i];
    if (s.name !== t.name || s.kind !== t.kind || s.visibility !== t.visibility) fail(`시트 '${s.name}' 이름·종류·숨김이 바뀌었습니다.`);
  });
  // 전역: BOUNDSHEET 위치 칸만 다를 수 있다.
  const g0 = w0.globals;
  const g1 = w1.globals;
  if (g0.length !== g1.length) fail("전역 레코드 수가 바뀌었습니다.");
  g0.forEach((r, i) => {
    const q = g1[i];
    if (r.type !== q.type || r.data.length !== q.data.length) fail("전역 레코드가 바뀌었습니다.");
    const from = r.type === 0x0085 ? 4 : 0;
    for (let k = from; k < r.data.length; k++) if (r.data[k] !== q.data[k]) fail("전역 레코드 내용이 바뀌었습니다.");
  });
  const targets = new Map(plans.filter((p) => p.edits.length).map((p) => [p.sheet, p.edits]));
  for (const name of targets.keys()) if (PROTECTED.test(name)) fail("승재 시트에는 쓰지 않습니다.");
  w0.sheets.forEach((s, i) => {
    const t = w1.sheets[i];
    const edits = targets.get(s.name);
    if (!edits) {
      if (!bytesEqual(sheetBytesWithoutIndexOffsets(w0.stream, s.offset), sheetBytesWithoutIndexOffsets(w1.stream, t.offset))) fail(`'${s.name}' 시트가 바뀌었습니다.`);
      return;
    }
    const a = readXlsSheet(w0, s.name);
    const b = readXlsSheet(w1, t.name);
    // 셀이 아닌 레코드
    const strip = (recs: XlsSheet["records"]) => recs.filter((r) => !CELLISH.has(r.type));
    const ra = strip(a.records);
    const rb = strip(b.records);
    if (ra.length !== rb.length) fail(`'${s.name}' 의 셀 밖 레코드 수가 바뀌었습니다.`);
    ra.forEach((r, k) => {
      const q = rb[k];
      if (r.type !== q.type) fail(`'${s.name}' 의 레코드 순서가 바뀌었습니다.`);
      if (VOLATILE.has(r.type)) return;
      if (!bytesEqual(r.data, q.data)) fail(`'${s.name}' 의 셀 밖 레코드(0x${r.type.toString(16)})가 바뀌었습니다.`);
    });
    const ROW = 0x0208;
    const rowsA = a.records.filter((r) => r.type === ROW);
    const rowsB = b.records.filter((r) => r.type === ROW);
    if (rowsA.length !== rowsB.length) fail(`'${s.name}' 의 행 수가 바뀌었습니다.`);
    rowsA.forEach((r, k) => {
      const q = rowsB[k];
      // 행 번호·높이·서식은 같아야 한다(열 범위 2~5 바이트만 바뀔 수 있음).
      for (let j = 0; j < r.data.length; j++) if ((j < 2 || j >= 6) && r.data[j] !== q.data[j]) fail(`'${s.name}' 의 행 높이·서식이 바뀌었습니다.`);
    });
    if (JSON.stringify(a.merges) !== JSON.stringify(b.merges)) fail(`'${s.name}' 의 병합이 바뀌었습니다.`);
    const editMap = new Map(edits.map((e) => [cellKey(e.row, e.col), e]));
    const keys = new Set([...a.cells.keys(), ...b.cells.keys()]);
    for (const key of keys) {
      const x = a.cells.get(key);
      const y = b.cells.get(key);
      const e = editMap.get(key);
      if (!e) {
        if (!x || !y) fail(`'${s.name}' ${key} 칸이 생기거나 없어졌습니다.`);
        if (x!.xf !== y!.xf) fail(`'${s.name}' ${key} 서식이 바뀌었습니다.`);
        if (x!.value.kind !== y!.value.kind) fail(`'${s.name}' ${key} 값 종류가 바뀌었습니다.`);
        if (x!.value.kind === "formula" && y!.value.kind === "formula") {
          if (!bytesEqual(x!.value.rgce, y!.value.rgce)) fail(`'${s.name}' ${key} 수식이 바뀌었습니다.`);
          continue; // 캐시 값은 다시 계산될 수 있다.
        }
        if (JSON.stringify(x!.value) !== JSON.stringify(y!.value)) fail(`'${s.name}' ${key} 값이 바뀌었습니다.`);
        continue;
      }
      if (!y) fail(`'${s.name}' ${key} 칸이 써지지 않았습니다.`);
      const want = e.xf ?? x?.xf;
      if (want != null && y!.xf !== want) fail(`'${s.name}' ${key} 서식이 예상과 다릅니다.`);
      const v = e.value;
      const got = y!.value;
      const ok =
        (v.kind === "blank" && got.kind === "blank") ||
        (v.kind === "number" && got.kind === "number" && got.n === v.n) ||
        (v.kind === "string" && (v.text === "" ? got.kind === "blank" : got.kind === "string" && got.text === v.text)) ||
        (v.kind === "formula" && got.kind === "formula" && bytesEqual(got.rgce, v.rgce));
      if (!ok) fail(`'${s.name}' ${key} 값이 예상과 다릅니다.`);
      if (got.kind === "string" && /앱/.test(got.text)) fail("앱 표시를 쓰려고 했습니다.");
    }
    void readSubstream;
  });
}

const CELLISH = new Set([0x0203, 0x027e, 0x00bd, 0x00fd, 0x0204, 0x00d6, 0x0201, 0x00be, 0x0205, 0x0006, 0x0207, 0x0208, 0x00d7]);
/** 셀 밖 레코드 중 위치·범위만 바뀔 수 있는 것: INDEX, DIMENSIONS */
const VOLATILE = new Set([0x020b, 0x0200]);

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}
