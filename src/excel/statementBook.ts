import { isoToExcelSerial, parseFlexibleDate } from "../domain/dates";
import { findMonthSheet, latestSheet, suggestSheetName } from "../domain/mapping";
import { diameterNumber } from "../domain/size";
import { cellText, extendSumFormulas, readSheet, setCell, setXml, shiftRowNumbers, workbookSheets, xmlText, type ZipFiles } from "./ooxml";

export interface PautLine {
  date: string;
  diameter: string;
  lengthM: number;
  count: number;
  unitPrice: number;
}

export interface StatementView {
  name: string;
  lengths: Record<string, number>;
  /** PAUT 행 단가의 최빈값 */
  unitPrice: number | null;
  paut: { date: string; item: string; count: number; unitPrice: number | null }[];
  totalRow: number;
}

export function readLengths(files: ZipFiles): Record<string, number> {
  const lengths: Record<string, number> = {};
  for (const info of workbookSheets(files)) {
    const view = readStatement(files, info.name);
    for (const [k, v] of Object.entries(view.lengths)) {
      if (lengths[k] == null) lengths[k] = v;
    }
  }
  return lengths;
}

export function readStatement(files: ZipFiles, sheetName: string): StatementView {
  const { cells } = readSheet(files, sheetName);
  const lengths: Record<string, number> = {};
  let maxRow = 11;
  for (const c of cells.values()) maxRow = Math.max(maxRow, c.row);
  for (let r = 1; r <= Math.min(maxRow, 80); r++) {
    const label = cellText(cells, "L", r).toUpperCase().replace(/\s/g, "");
    const m = label.match(/^(\d{2,4})A$/);
    if (!m) continue;
    const raw = cellText(cells, "M", r);
    const num = raw.match(/(\d+(?:\.\d+)?)/);
    if (num) lengths[m[1]] = Number(num[1]);
  }
  let totalRow = 0;
  for (let r = 12; r <= maxRow; r++) {
    const a = cellText(cells, "A", r).replace(/\s/g, "");
    if (a.includes("합계")) {
      totalRow = r;
      break;
    }
  }
  if (!totalRow) totalRow = maxRow + 1;
  const prices: number[] = [];
  const paut: StatementView["paut"] = [];
  for (let r = 12; r < totalRow; r++) {
    const method = cellText(cells, "B", r);
    const item = cellText(cells, "C", r);
    const unit = cellText(cells, "E", r);
    const countRaw = cellText(cells, "D", r);
    const priceRaw = cellText(cells, "G", r);
    if (!method && !item && !countRaw) continue;
    const kind = lineKind(method, item, unit);
    if (kind !== "paut") continue;
    const count = Number(countRaw);
    const price = priceRaw && /^\d+(\.\d+)?$/.test(priceRaw) ? Number(priceRaw) : null;
    if (price) prices.push(price);
    paut.push({
      date: parseFlexibleDate(cellText(cells, "A", r)),
      item,
      count: Number.isFinite(count) ? count : 0,
      unitPrice: price,
    });
  }
  return { name: sheetName, lengths, unitPrice: mode(prices), paut, totalRow };
}

function lineKind(method: string, item: string, unit: string): "paut" | "rt" | "other" {
  const m = method.replace(/\s/g, "");
  if (m.includes("방사선") || m.includes("RT")) return "rt";
  if (unit === "매") return "rt";
  if (m.includes("위상배열") || m.toUpperCase().includes("PAUT")) return "paut";
  if (unit === "M" && /\d+A/.test(item)) return "paut";
  if (!method && !item) return "other";
  return "other";
}

function mode(nums: number[]): number | null {
  if (!nums.length) return null;
  const counts = new Map<number, number>();
  for (const n of nums) counts.set(n, (counts.get(n) || 0) + 1);
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

export interface WritePautResult {
  /** 합계 행 앞에 끼워 넣은 행 수 */
  inserted: number;
  /** 끼워 넣기 전 합계 행 번호. 엑셀에서는 이 행부터 inserted 개 행을 아래로 민다. */
  insertAt: number;
  /** 이번에 앱이 쓴 줄(날짜|관경 → 개수). 다음 동기화 때 '앱이 쓴 행'을 알아보는 데 쓴다(엑셀에는 남기지 않음). */
  owned: SheetOwnership;
}

/**
 * 예전 버전이 앱이 쓴 행 I열(비고)에 남기던 표시. 이제는 쓰지 않는다.
 * 이미 적혀 있는 옛 행은 지우지 않고 그대로 두며, 읽을 때만 '앱이 쓴 행'으로 알아본다.
 */
export const LEGACY_APP_MARK = "앱";

/** 시트 하나에서 앱이 쓴 줄: `날짜|관경` → 그때 쓴 개수(D). 엑셀이 아니라 이 기기에 저장한다. */
export type SheetOwnership = Record<string, number>;

/**
 * 그 달 PAUT 합격 집계를 쓴다. RT·기타 행과 사람이 적은 PAUT 행은 그대로 둔다.
 * 엑셀에는 앱 표시를 남기지 않는다(비고 I열 포함). 앱이 쓴 행은 다음 중 하나로 알아본다.
 * - owned: 이 기기가 지난번에 쓴 날짜·관경과 개수가 그대로인 행(사람이 개수를 고쳤으면 사람 행으로 본다)
 * - 옛 버전이 I열에 '앱'을 적어 둔 행(읽기만, 지우지 않음)
 * 규칙:
 * - 같은 날짜·관경을 사람이 이미 적었으면 그 줄은 건너뛴다(중복 없음).
 * - 앱이 쓴 행과 빈 행에만 쓴다. 다시 보내도 같은 행을 다시 쓰므로 줄이 늘지 않는다.
 * - 빈 칸이 모자라면 합계 행 앞에 행을 끼워 넣고 SUM 범위를 늘린다.
 */
export function writePautLines(files: ZipFiles, sheetName: string, lines: PautLine[], owned: SheetOwnership = {}): WritePautResult {
  lines = [...lines].sort((a, b) => a.date.localeCompare(b.date) || a.diameter.localeCompare(b.diameter));
  let guard = 0;
  let inserted = 0;
  let insertAt = 0;
  while (guard++ < 40) {
    const view = readStatement(files, sheetName);
    const { cells, xml, info } = readSheet(files, sheetName);
    const slots = collectSlots(cells, view.totalRow, owned);
    const manual = new Set(slots.filter((s) => s.kind === "paut" && !s.app).map((s) => s.key));
    const todo = lines.filter((l) => !manual.has(lineKey(l.date, l.diameter)));
    const writable = slots.filter((s) => (s.kind === "paut" && s.app) || s.kind === "blank");
    if (writable.length >= todo.length) {
      let next = xml;
      const wrote: SheetOwnership = {};
      writable.forEach((slot, i) => {
        const line = todo[i];
        if (!line) {
          if (slot.kind === "paut") {
            // 앱이 쓴 줄이 더는 없으면 비운다. I열(비고)은 옛 '앱' 표시일 때만 함께 비우고, 사람이 적은 비고는 둔다.
            for (const col of ["A", "B", "C", "D", "E", "F", "G", "H"]) next = setCell(next, col + slot.row, { kind: "empty" });
            if (slot.legacy) next = setCell(next, "I" + slot.row, { kind: "empty" });
          }
          return;
        }
        next = writeLine(next, slot.row, line);
        wrote[lineKey(line.date, line.diameter)] = line.count;
      });
      setXml(files, info.path, next);
      return { inserted, insertAt, owned: wrote };
    }
    if (!inserted) insertAt = view.totalRow;
    insertRowBeforeTotal(files, sheetName, view.totalRow);
    inserted += 1;
  }
  throw new Error("거래명세서에 행을 더 넣지 못했습니다.");
}

/** 그 달 시트 이름(사용자가 만들 이름). 같은 접두어 최근 시트의 표기를 따른다. 예: 드림합정26.11 */
export function expectedMonthSheetName(names: string[], prefix: string, year: number, month: number): string {
  return findMonthSheet(names, prefix, year, month)?.name ?? suggestSheetName(prefix, latestSheet(names, prefix), year, month);
}

function lineKey(date: string, diameter: string): string {
  return `${date}|${diameter.toUpperCase()}`;
}

export interface StatementRow {
  row: number;
  date: string;
  diameter: string;
  count: number;
  /** 앱이 쓴 행으로 보이는지(owned 또는 옛 I='앱') */
  app: boolean;
}

/** 시트의 PAUT 본문 행. 백필 확인용. */
export function readPautRows(files: ZipFiles, sheetName: string, owned: SheetOwnership = {}): StatementRow[] {
  const view = readStatement(files, sheetName);
  const { cells } = readSheet(files, sheetName);
  return collectSlots(cells, view.totalRow, owned)
    .filter((s) => s.kind === "paut")
    .map((s) => ({ row: s.row, date: s.date, diameter: s.diameter, count: s.count, app: s.app }));
}

interface Slot {
  row: number;
  kind: "paut" | "rt" | "blank" | "other";
  app: boolean;
  /** 옛 버전의 I='앱' 표시가 있는 행 */
  legacy: boolean;
  date: string;
  diameter: string;
  count: number;
  key: string;
}

function collectSlots(cells: Map<string, { col: string; row: number; v?: string }>, totalRow: number, owned: SheetOwnership): Slot[] {
  const out: Slot[] = [];
  for (let r = 12; r < totalRow; r++) {
    const method = cellText(cells, "B", r);
    const item = cellText(cells, "C", r);
    const unit = cellText(cells, "E", r);
    const count = cellText(cells, "D", r);
    const dateRaw = cellText(cells, "A", r);
    const legacy = cellText(cells, "I", r) === LEGACY_APP_MARK;
    const date = parseFlexibleDate(dateRaw);
    const dia = /(\d{2,4})\s*A/i.exec(item)?.[1];
    const diameter = dia ? `${dia}A` : "";
    const key = lineKey(date, diameter);
    const n = Number(count) || 0;
    const app = legacy || (owned[key] != null && owned[key] === n);
    const base = { row: r, app, legacy, date, diameter, count: n, key };
    if (!method && !item && !count && !dateRaw) out.push({ ...base, kind: "blank" });
    else {
      const kind = lineKind(method, item, unit);
      out.push({ ...base, kind: kind === "other" ? "other" : kind });
    }
  }
  return out;
}

function writeLine(xml: string, row: number, line: PautLine): string {
  const len = trimNum(line.lengthM);
  const qty = trimNum(line.count * line.lengthM);
  const amount = line.count * line.lengthM * line.unitPrice;
  let next = xml;
  next = setCell(next, "A" + row, { kind: "number", n: isoToExcelSerial(line.date) });
  next = setCell(next, "B" + row, { kind: "inline", text: "위상배열초음파검사(PAUT)" });
  next = setCell(next, "C" + row, { kind: "inline", text: `${line.diameter}(${len}M)` });
  next = setCell(next, "D" + row, { kind: "number", n: line.count });
  next = setCell(next, "E" + row, { kind: "inline", text: "M" });
  next = setCell(next, "F" + row, { kind: "formula", formula: `D${row}*${len}`, cached: qty, cachedType: "n" });
  next = setCell(next, "G" + row, { kind: "number", n: line.unitPrice });
  next = setCell(next, "H" + row, { kind: "formula", formula: `G${row}*F${row}`, cached: amount, cachedType: "n" });
  // I열(비고)은 쓰지 않는다. 앱 표시를 남기지 않고, 이미 있는 값(옛 '앱' 포함)도 건드리지 않는다.
  return next;
}

function trimNum(n: number): string {
  return String(Math.round(n * 1000) / 1000);
}

function insertRowBeforeTotal(files: ZipFiles, sheetName: string, totalRow: number): void {
  const info = workbookSheets(files).find((s) => s.name === sheetName);
  if (!info) throw new Error(sheetName);
  let xml = xmlText(files, info.path);
  const template = xml.match(new RegExp(`<row r="${totalRow - 1}"[^>]*>[\\s\\S]*?</row>`))?.[0] ?? `<row r="${totalRow - 1}"></row>`;
  xml = shiftRowNumbers(xml, totalRow, 1);
  const blank = template
    .replace(new RegExp(`r="${totalRow - 1}"`, "g"), `r="${totalRow}"`)
    .replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (full, d1: string, col: string, d2: string, row: string) => {
      if (Number(row) === totalRow - 1) return `${d1}${col}${d2}${totalRow}`;
      return full;
    })
    .replace(/<f\b[^>]*>[\s\S]*?<\/f>/g, "")
    .replace(/<v>[\s\S]*?<\/v>/g, "")
    .replace(/<is>[\s\S]*?<\/is>/g, "")
    .replace(/\s+t="[^"]*"/g, "");
  xml = xml.replace(`<row r="${totalRow + 1}"`, blank + `<row r="${totalRow + 1}"`);
  xml = extendSumFormulas(xml, 12, totalRow);
  setXml(files, info.path, xml);
}

export function lengthOf(diameter: string, lengths: Record<string, number>): number | null {
  const n = diameterNumber(diameter);
  if (n == null) return null;
  return lengths[String(n)] ?? null;
}
