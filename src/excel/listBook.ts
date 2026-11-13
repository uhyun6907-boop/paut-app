import { isoToExcelSerial, parseFlexibleDate } from "../domain/dates";
import { comparePointLabel } from "../domain/point";
import { diameterFromLine, excelRight4 } from "../domain/size";
import type { InspectionRecord, ListBookKind } from "../domain/types";
import { LIST_GM_SHEETS } from "../domain/types";
import { cellText, parseCells, readSheet, setCell, setXml, sharedStrings, workbookSheets, xmlText, type ZipFiles } from "./ooxml";

const COL = {
  seq: "A",
  issued: "B",
  report: "C",
  reportDate: "D",
  date: "E",
  line: "F",
  point: "G",
  defect: "H",
  drawing: "I",
  weld: "J",
  material: "K",
  diameter: "L",
  sch: "M",
  rt: "N",
  result: "P",
  flawType: "Q",
  flawLen: "R",
  note: "S",
  appId: "Y",
} as const;

/**
 * 리스트 파일마다 다른 점.
 * - 평택 리스트: 관경(L)은 RIGHT(F,4) 수식(그 파일의 기존 행과 같음). 순번(A)은 Graph 로는 쓰지 않는다(기존 동작).
 * - 군포·마전 리스트: L 은 값('750A')으로 적혀 있다. 라인이 '군포산본 이원화 750A-1' 처럼 관경으로 끝나지 않아
 *   RIGHT(F,4) 는 '0A-1' 이 되므로 값으로 쓴다. 순번(A)이 비어 있으면 다음 번호를 채운다.
 */
export interface ListLayout {
  diameter: "formula" | "value";
  fillSeq: boolean;
}

export const LIST_LAYOUT: Record<ListBookKind, ListLayout> = {
  list: { diameter: "formula", fillSeq: false },
  listGm: { diameter: "value", fillSeq: true },
};

/** 군포·마전 리스트 시트인가(앞뒤 공백 무시). */
export function isGmSheet(sheetName: string): boolean {
  const name = String(sheetName ?? "").trim();
  return LIST_GM_SHEETS.includes(name);
}

const gmKey = (name: string) => String(name ?? "").replace(/\s+/g, "").toUpperCase().replace(/(PAUT|PA)$/, "");

/**
 * 현장·시트 이름이 군포·마전 PA 시트를 가리키면 그 시트의 정확한 이름.
 * '지성 마전'·'지성마전'·'아전 군포' 처럼 직접 추가한 현장(PA 없음, 띄어쓰기 다름)도 같은 시트로 본다.
 * RT 시트('지성마전 RT')나 다른 이름은 null.
 */
export function gmSheetOf(name: string | undefined): string | null {
  if (!name) return null;
  const key = gmKey(name);
  if (!key) return null;
  return LIST_GM_SHEETS.find((s) => gmKey(s) === key) ?? null;
}

/**
 * 기록이 들어갈 리스트 파일. 기록에 적힌 listBook 이 먼저, 없으면 시트 이름으로 정한다.
 * 군포·마전 기록이 평택 리스트에 들어가거나 그 반대가 되지 않게 하는 유일한 기준.
 * 평택 리스트에는 군포·마전 시트가 없으므로, 이름이 군포·마전 PA 시트를 가리키면 군포·마전 리스트다.
 */
export function listBookOf(rec: { listBook?: ListBookKind; sheetName: string; site?: string }): ListBookKind {
  if (rec.listBook === "listGm") return "listGm";
  if (gmSheetOf(rec.sheetName) || gmSheetOf(rec.site)) return "listGm";
  return rec.listBook ?? "list";
}

/** 리스트 파일 안에서 그 기록이 들어갈 시트 이름. 군포·마전이면 정확한 PA 시트 이름. */
export function listSheetOf(rec: { listBook?: ListBookKind; sheetName: string; site?: string }): string {
  if (listBookOf(rec) === "listGm") return gmSheetOf(rec.sheetName) ?? gmSheetOf(rec.site) ?? rec.sheetName.trim();
  return rec.sheetName.trim();
}

/**
 * 군포·마전 리스트가 생기기 전에 '새 현장'으로 직접 추가한 현장('지성 마전'·'아전 군포'·'지성 군포' 등)의 기록은
 * 평택 리스트로 가려다 '시트가 없습니다'로 실패하거나 리스트 단계가 끝난 것으로 남아 군포·마전 파일에 한 번도 쓰이지 않았다.
 * 그런 앱 기록을 군포·마전 리스트의 정확한 PA 시트로 옮기고, 리스트 단계만 다시 하게 한다(작업일지·명세서 단계는 그대로).
 * 실제로 들어갔는지는 동기화 때 Y(앱 ID)·날짜+라인+포인트로 확인하므로 두 번 쓰지 않는다. 돌려주는 값: 바꾼 기록 수.
 */
export function rerouteGmRecords(records: InspectionRecord[]): number {
  let changed = 0;
  for (const rec of records) {
    if (rec.src !== "user" || rec.deleted) continue;
    const sheet = gmSheetOf(rec.sheetName) ?? gmSheetOf(rec.site);
    if (!sheet || (rec.listBook === "listGm" && rec.sheetName === sheet)) continue;
    rec.listBook = "listGm";
    rec.sheetName = sheet;
    if (rec.steps?.list) rec.steps = { ...rec.steps, list: undefined };
    if (rec.sync === "synced") {
      // 이미 끝났던 작업일지·명세서는 끝난 것으로 둔다(리스트만 다시).
      const at = Date.now();
      rec.steps = { ...(rec.steps ?? {}), ilji: rec.steps?.ilji ?? { status: "done", at }, statement: rec.steps?.statement ?? { status: "done", at } };
    }
    if (rec.sync === "error" || rec.sync === "synced") {
      rec.sync = "pending";
      rec.syncError = undefined;
    }
    rec.excelRow = undefined;
    changed += 1;
  }
  return changed;
}

/** 리스트 L열에 쓸 값. 값 모드면 라인에서 관경을 읽고, 못 읽으면 쓰지 않는다. */
export function listDiameterValue(line: string): string | null {
  return diameterFromLine(line);
}

export function isRtSheet(name: string): boolean {
  return name.trim().endsWith(" RT");
}

export interface ParsedSite {
  site: string;
  sheetName: string;
  records: InspectionRecord[];
}

export function parseListWorkbook(files: ZipFiles, book: ListBookKind = "list"): ParsedSite[] {
  const shared = sharedStrings(files);
  const sites: ParsedSite[] = [];
  for (const info of workbookSheets(files)) {
    if (isRtSheet(info.name)) continue;
    const xml = xmlText(files, info.path);
    const cells = parseCells(xml, shared);
    const header = cellText(cells, "F", 4);
    if (header && !/line/i.test(header)) continue;
    const site = info.name.trim();
    const byRow = new Map<number, true>();
    for (const c of cells.values()) {
      if (c.col === "F" && c.row >= 6 && c.v) byRow.set(c.row, true);
    }
    const records: InspectionRecord[] = [];
    const rows = [...byRow.keys()].sort((a, b) => a - b);
    for (const row of rows) {
      const line = cellText(cells, "F", row);
      const point = cellText(cells, "G", row).toUpperCase();
      if (!line || !point) continue;
      const smart = diameterFromLine(line);
      const rawDia = cellText(cells, "L", row).toUpperCase();
      const diameter = smart || (/^\d{2,4}A$/.test(rawDia) ? rawDia : "");
      const result = cellText(cells, "P", row);
      const appId = cellText(cells, "Y", row);
      records.push({
        id: appId || `xlsx:${info.name}:${row}`,
        site,
        sheetName: info.name,
        date: parseFlexibleDate(cellText(cells, "E", row)),
        line,
        point,
        result,
        diameter,
        material: cellText(cells, "K", row),
        sch: cellText(cells, "M", row),
        reportNo: cellText(cells, "C", row),
        reportDate: parseFlexibleDate(cellText(cells, "D", row)),
        defectNo: cellText(cells, "H", row),
        drawingNo: cellText(cells, "I", row),
        weldNo: cellText(cells, "J", row),
        rtSheets: cellText(cells, "N", row),
        flawType: cellText(cells, "Q", row),
        flawLen: cellText(cells, "R", row),
        note: cellText(cells, "S", row),
        issued: cellText(cells, "B", row),
        src: appId ? "user" : "xlsx",
        ...(book !== "list" ? { listBook: book } : {}),
        excelRow: row,
        sync: "synced",
        ts: row,
      });
    }
    sites.push({ site, sheetName: info.name, records });
  }
  return sites;
}

export interface ListWrite {
  id: string;
  date: string;
  line: string;
  point: string;
  result: string;
  material?: string;
  /** true 면 해당 행의 검사값을 비운다. */
  clear?: boolean;
  /** 파일에 앱ID가 아직 없는 기존 행 */
  excelRow?: number;
  /** 이 행에 쓴다(같은 라인 앱 행 다시 정렬). Y 앱 ID 로 찾은 행보다 먼저. */
  atRow?: number;
}

const LINE_COLLATOR = new Intl.Collator("ko", { numeric: true, sensitivity: "base" });

/**
 * 리스트 끝에 새로 붙일 행의 순서: 날짜 → 라인 → 포인트(S 전부 → R 전부, 숫자 순 S1·S2·…·S10).
 * 같은 포인트(재검사)는 먼저 입력한 원본 바로 뒤. 이미 엑셀에 있는 행은 이 순서로 옮기지 않는다
 * (그 행은 Y 앱 ID·excelRow 로 제자리를 고친다). 입력 배열은 바꾸지 않는다.
 */
export function orderForAppend<T extends { date: string; line: string; point: string; ts: number; id: string }>(recs: T[]): T[] {
  return recs
    .map((rec, index) => ({ rec, index }))
    .sort((x, y) => {
      const a = x.rec;
      const b = y.rec;
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      const line = LINE_COLLATOR.compare(a.line.trim(), b.line.trim());
      if (line) return line;
      return comparePointLabel(a.point, b.point) || a.ts - b.ts || x.index - y.index;
    })
    .map((row) => row.rec);
}

/** 시트의 Y(앱 ID) → 행 번호. */
export function appIdRows(files: ZipFiles, sheetName: string): Map<string, number> {
  const { cells } = readSheet(files, sheetName);
  const out = new Map<string, number>();
  for (const c of cells.values()) {
    if (c.col === COL.appId && c.row >= 6 && c.v && c.v.trim()) out.set(c.v.trim(), c.row);
  }
  return out;
}

const sameLineText = (a: string, b: string) => a.trim().toUpperCase() === b.trim().toUpperCase();

/**
 * 날짜를 바꾼 기록(movedFrom)이 있는 라인의 앱 행을 다시 놓는다.
 * 그 시트에서 같은 라인의 앱 행(Y 앱 ID, 이 기기에 있는 합격·불합격 기록)이 차지한 행 번호들 안에서만
 * 날짜 → 라인 → 포인트 순으로 바꿔 앉힌다. 행을 끼우거나 지우지 않으므로 다른 행(사람이 적은 행 포함)은 그대로이고
 * 같은 기록이 두 줄이 되지 않는다. 돌려주는 값: 자리가 바뀐 기록 → 새 행.
 */
export function planLineReorder<T extends InspectionRecord>(rowsById: Map<string, number>, sheetRecs: T[]): Map<string, { rec: T; row: number }> {
  const out = new Map<string, { rec: T; row: number }>();
  const alive = sheetRecs.filter((r) => !r.deleted && (r.result === "합격" || r.result === "불합격") && rowsById.has(r.id));
  const lines = new Set(alive.filter((r) => r.movedFrom?.length).map((r) => r.line.trim().toUpperCase()));
  for (const line of lines) {
    const recs = alive.filter((r) => sameLineText(r.line, line));
    const rows = recs.map((r) => rowsById.get(r.id)!).sort((a, b) => a - b);
    orderForAppend(recs).forEach((rec, i) => {
      if (rowsById.get(rec.id) !== rows[i]) out.set(rec.id, { rec, row: rows[i] });
    });
  }
  return out;
}

export function writeListRow(files: ZipFiles, sheetName: string, input: ListWrite, layout: ListLayout = LIST_LAYOUT.list): number {
  const { xml, cells, info } = readSheet(files, sheetName);
  let row = input.atRow ?? findRowByAppId(cells, input.id) ?? input.excelRow ?? null;
  if (input.clear) {
    if (!row) return -1;
    let next = xml;
    for (const col of [COL.date, COL.line, COL.point, COL.material, COL.diameter, COL.result, COL.appId]) {
      next = setCell(next, col + row, { kind: "empty" });
    }
    setXml(files, info.path, next);
    return row;
  }
  if (!row) row = nextEmptyLineRow(cells);
  const seqCell = cells.get("A" + row);
  const seq = seqCell?.v && /^\d+$/.test(seqCell.v) ? Number(seqCell.v) : maxSeq(cells) + 1;
  let next = xml;
  if (!seqCell?.v) next = setCell(next, "A" + row, { kind: "number", n: seq });
  next = setCell(next, COL.date + row, { kind: "number", n: isoToExcelSerial(input.date) });
  next = setCell(next, COL.line + row, { kind: "inline", text: input.line });
  next = setCell(next, COL.point + row, { kind: "inline", text: input.point });
  if (input.material) next = setCell(next, COL.material + row, { kind: "inline", text: input.material });
  else if (input.atRow) next = setCell(next, COL.material + row, { kind: "empty" });
  if (layout.diameter === "value") {
    const dia = listDiameterValue(input.line);
    if (dia) next = setCell(next, COL.diameter + row, { kind: "inline", text: dia });
  } else {
    const cached = excelRight4(input.line);
    next = setCell(next, COL.diameter + row, {
      kind: "formula",
      formula: `RIGHT(F${row},4)`,
      cached,
      cachedType: "str",
    });
  }
  next = setCell(next, COL.result + row, { kind: "inline", text: input.result });
  next = setCell(next, COL.appId + row, { kind: "inline", text: input.id });
  setXml(files, info.path, next);
  return row;
}

function findRowByAppId(cells: Map<string, { col: string; row: number; v?: string }>, id: string): number | null {
  for (const c of cells.values()) {
    if (c.col === COL.appId && c.v === id) return c.row;
  }
  return null;
}

function nextEmptyLineRow(cells: Map<string, { col: string; row: number; v?: string }>): number {
  let max = 5;
  for (const c of cells.values()) {
    if (c.col === "F" && c.row >= 6 && c.v) max = Math.max(max, c.row);
  }
  return max + 1;
}

function maxSeq(cells: Map<string, { col: string; row: number; v?: string }>): number {
  let max = 0;
  for (const c of cells.values()) {
    if (c.col === "A" && c.row >= 6 && c.v && /^\d+$/.test(c.v)) max = Math.max(max, Number(c.v));
  }
  return max;
}

export function sheetNameForSite(sites: { site: string; sheetName: string }[], site: string): string | null {
  return sites.find((s) => s.site === site)?.sheetName ?? null;
}
