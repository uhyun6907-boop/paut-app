/**
 * 세 번째 탭 「수정」: 검사리스트 행 하나를 위·아래로 옮긴다. 엑셀 행을 끼우거나 지우지 않고,
 * 옮기는 행부터 도착 행까지(lo..hi)만 B–Y 값·수식(R1C1)·서식을 새 순서로 다시 쓴다. A(순번)는 제자리 그대로.
 * - 병합 칸·메모가 그 행들에 있거나, Z 열부터 값이 있으면 옮기지 않는다.
 * - 쓰기 직전에 그 범위를 다시 읽어 화면에 보인 값과 다르면 쓰지 않는다(LiveConflictError).
 * - 쓴 뒤 다시 읽어 행마다 확인하고, 다른 칸을 mismatches 로 돌려준다. snapshot 으로 한 번 되돌릴 수 있다.
 */
import type { GraphCell, GraphClient, RangeBlock, WorkbookSession } from "../sync/graph";
import { worksheetRangeUrl } from "../sync/graph";
import { batchOk, readRowFormats, styleWrites, type RowFormat } from "../sync/rowFormat";
import { LIVE_FIRST_DATA_ROW, LIVE_LAST_COL, LiveConflictError, liveRow, same, type LiveRow, type LiveTable } from "./liveList";
import { columnIndex, columnLetter, unzipXlsx, workbookSheets, xmlText, type ZipFiles } from "./ooxml";

const FIRST_MOVED_COL = 2;
const LAST = columnLetter(LIVE_LAST_COL);

export class LiveMoveRefused extends Error {
  constructor(message: string) {
    super(message);
    this.name = "LiveMoveRefused";
  }
}

interface MovedRow {
  cells: GraphCell[];
  format: RowFormat;
}

/** 옮기기 전 lo..hi 의 B–Y(값·수식·서식). 되돌리기에 쓴다. */
export interface MoveSnapshot {
  lo: number;
  rows: MovedRow[];
  before: RangeBlock;
}

export interface MoveOutcome {
  /** 옛 행 → 새 행 */
  moved: Map<number, number>;
  /** 쓴 뒤 다시 읽은 lo..hi */
  rows: LiveRow[];
  /** 예상과 다른 칸. 비었으면 모두 확인됨. */
  mismatches: string[];
  snapshot: MoveSnapshot;
}

/** from 행을 to 자리로. 사이 행은 한 칸씩 밀린다. 돌려주는 값: lo..hi 각 자리에 올 원래 행. */
export function moveOrder(from: number, to: number): { lo: number; hi: number; sources: number[] } {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  const rest = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i).filter((r) => r !== from);
  rest.splice(to - lo, 0, from);
  return { lo, hi, sources: rest };
}

export async function moveLiveRow(client: GraphClient, table: LiveTable, from: number, to: number): Promise<MoveOutcome> {
  const { lo, hi, sources } = moveOrder(from, to);
  if (from === to) throw new LiveMoveRefused("같은 자리입니다.");
  if (lo < LIVE_FIRST_DATA_ROW) throw new LiveMoveRefused(`${LIVE_FIRST_DATA_ROW}행 위(머리글)는 옮기지 않습니다.`);
  const seen = Array.from({ length: hi - lo + 1 }, (_, i) => table.rows.find((r) => r.row === lo + i));
  const gap = seen.findIndex((r) => !r);
  if (gap >= 0) throw new LiveMoveRefused(`${lo + gap}행이 비어 있어 옮기지 않습니다(빈 행을 건너 옮기지 않음).`);
  const obstacles = [...sheetObstacles(unzipXlsx(await client.download(table.itemId)), table.sheetName, lo, hi), ...(await outsideValues(client, table, lo, hi))];
  if (obstacles.length) throw new LiveMoveRefused(`${lo}–${hi}행은 옮기지 않습니다: ${obstacles.join(", ")}. PC 엑셀에서 옮기세요.`);

  const wb = await client.openWorkbook(table.itemId);
  try {
    const address = `A${lo}:${LAST}${hi}`;
    const before = await wb.readBlock(table.sheetId, address);
    seen.forEach((row, i) => {
      for (const cell of row!.cells) {
        const now = before.values[i]?.[columnIndex(cell.col) - 1] ?? "";
        if (!cell.formula && !same(now, cell.value)) throw new LiveConflictError(now, cell.address);
      }
    });
    const formats = await readRowFormats(wb, table.sheetId, seen.map((r) => r!.row), FIRST_MOVED_COL, LIVE_LAST_COL);
    const current = formats.map((format, i) => ({ cells: movedCells(before, i), format }));
    const next = sources.map((src) => current[src - lo]);
    const writeError = await writeRows(wb, table.sheetId, lo, next, formats).then(
      () => null,
      (error: unknown) => error,
    );
    const after = await wb.readBlock(table.sheetId, address);
    const mismatches = verify(after, before, lo, sources);
    if (writeError) mismatches.unshift(writeError instanceof Error ? writeError.message : String(writeError));
    return {
      moved: new Map(sources.map((src, i) => [src, lo + i])),
      rows: after.values.map((values, i) => liveRow(lo + i, values, after.formulasR1C1[i] ?? [])),
      mismatches,
      snapshot: { lo, rows: current, before },
    };
  } finally {
    await wb.close();
  }
}

/** 옮기기 전 모습으로 lo..hi 를 다시 쓴다(한 번). 돌려주는 값: 다시 읽은 행과 아직 다른 칸. */
export async function restoreLiveRows(client: GraphClient, table: LiveTable, snapshot: MoveSnapshot): Promise<{ rows: LiveRow[]; mismatches: string[] }> {
  const { lo, rows, before } = snapshot;
  const hi = lo + rows.length - 1;
  const wb = await client.openWorkbook(table.itemId);
  try {
    await writeRows(wb, table.sheetId, lo, rows, null);
    const after = await wb.readBlock(table.sheetId, `A${lo}:${LAST}${hi}`);
    return {
      rows: after.values.map((values, i) => liveRow(lo + i, values, after.formulasR1C1[i] ?? [])),
      mismatches: verify(after, before, lo, rows.map((_, i) => lo + i)),
    };
  } finally {
    await wb.close();
  }
}

/** 엑셀은 쓴 글자를 숫자·날짜·수식으로 바꿔 읽으므로, 그럴 수 있는 글자 칸은 ' 를 붙여 글자로 남긴다. */
function movedCells(block: RangeBlock, i: number): GraphCell[] {
  return Array.from({ length: LIVE_LAST_COL - FIRST_MOVED_COL + 1 }, (_, j) => {
    const c = FIRST_MOVED_COL - 1 + j;
    const formula = block.formulasR1C1[i]?.[c] ?? "";
    if (typeof formula === "string" && formula.startsWith("=")) return formula;
    const value = block.values[i]?.[c] ?? "";
    return typeof value === "string" && block.valueTypes[i]?.[c] === "String" && /^\s*[=+\-@'#\d.]|^\s*(true|false)\s*$/i.test(value) ? `'${value}` : value;
  });
}

async function writeRows(wb: WorkbookSession, sheetId: string, lo: number, rows: MovedRow[], current: RowFormat[] | null): Promise<void> {
  const hi = lo + rows.length - 1;
  const url = worksheetRangeUrl(wb.itemId, sheetId, `${columnLetter(FIRST_MOVED_COL)}${lo}:${LAST}${hi}`);
  const key = (f: RowFormat) => JSON.stringify(f.cells);
  await batchOk(
    wb,
    [
      { method: "PATCH", url, body: { numberFormat: rows.map((r) => r.format.numberFormat) } },
      { method: "PATCH", url, body: { formulasR1C1: rows.map((r) => r.cells) } },
      ...rows.flatMap((r, i) => (current && key(current[i]) === key(r.format) ? [] : styleWrites(wb.itemId, sheetId, r.format.cells, FIRST_MOVED_COL, lo + i, lo + i))),
    ],
    "행 순서 쓰기",
  );
}

function verify(after: RangeBlock, before: RangeBlock, lo: number, sources: number[]): string[] {
  const out: string[] = [];
  sources.forEach((src, i) => {
    const s = src - lo;
    const bad: string[] = [];
    if (!same(after.values[i]?.[0], before.values[i]?.[0])) bad.push("A");
    for (let c = FIRST_MOVED_COL - 1; c < LIVE_LAST_COL; c++) {
      const formula = before.formulasR1C1[s]?.[c];
      const ok = typeof formula === "string" && formula.startsWith("=") ? after.formulasR1C1[i]?.[c] === formula : same(after.values[i]?.[c], before.values[s]?.[c]);
      if (!ok) bad.push(columnLetter(c + 1));
    }
    if (bad.length) out.push(`${lo + i}행(${bad.join("·")})`);
  });
  return out;
}

async function outsideValues(client: GraphClient, table: LiveTable, lo: number, hi: number): Promise<string[]> {
  const first = `${columnLetter(LIVE_LAST_COL + 1)}${lo}`;
  const extent = (await client.rangeValuesExtent(table.itemId, table.sheetId, `${first}:XFD${hi}`)).replace(/^.*!/, "").replace(/\$/g, "");
  if (extent === first) {
    const v = (await client.readRange(table.itemId, table.sheetId, first))[0]?.[0];
    if (v === "" || v === null || v === undefined) return [];
  }
  return [`${columnLetter(LIVE_LAST_COL)} 열 오른쪽 값(${extent})`];
}

const rowsOf = (ref: string): [number, number] => {
  const [a, b] = ref.split(":").map((x) => Number(/\d+$/.exec(x)?.[0]));
  return [a, b ?? a];
};

/** 다운로드한 파일에서 lo..hi 행에 걸친 병합 칸·메모. */
export function sheetObstacles(files: ZipFiles, sheetName: string, lo: number, hi: number): string[] {
  const sheet = workbookSheets(files).find((s) => s.name === sheetName) ?? workbookSheets(files).find((s) => s.name.trim() === sheetName.trim());
  if (!sheet) return [`파일에 '${sheetName}' 시트 없음`];
  const touches = (ref: string) => {
    const [a, b] = rowsOf(ref);
    return a <= hi && b >= lo;
  };
  const out = [...xmlText(files, sheet.path).matchAll(/<mergeCell\b[^>]*\bref="([^"]+)"/g)].map((m) => m[1]).filter(touches).map((ref) => `병합 칸 ${ref}`);
  const dir = sheet.path.slice(0, sheet.path.lastIndexOf("/"));
  const rels = files[`${dir}/_rels/${sheet.path.slice(dir.length + 1)}.rels`] ? xmlText(files, `${dir}/_rels/${sheet.path.slice(dir.length + 1)}.rels`) : "";
  for (const m of rels.matchAll(/<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\/>/g)) {
    if (!/comments|threadedComment/i.test(m[1])) continue;
    const path = resolvePath(dir, m[1]);
    if (!files[path]) continue;
    for (const c of xmlText(files, path).matchAll(/<(?:comment|threadedComment)\b[^>]*\bref="([^"]+)"/g)) if (touches(c[1])) out.push(`메모 ${c[1]}`);
  }
  return [...new Set(out)];
}

function resolvePath(dir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = dir.split("/");
  for (const p of target.split("/")) {
    if (p === "..") parts.pop();
    else if (p !== ".") parts.push(p);
  }
  return parts.join("/");
}
