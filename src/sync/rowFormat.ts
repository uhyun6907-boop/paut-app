import { columnLetter } from "../excel/ooxml";
import { worksheetRangeUrl, type BatchRequest, type BatchResponse, type WorkbookSession } from "./graph";

const LIST_ROW_WIDTH = 25;
const LAST_COL = columnLetter(LIST_ROW_WIDTH);
const EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;
/** Graph 는 채우기 없음도 흰색으로 읽는다. 흰색을 칠하면 눈금선이 사라지므로 '채우기 없음'으로 쓴다. */
const NO_FILL = "#FFFFFF";

type Edge = (typeof EDGES)[number];

interface Border {
  style: string;
  color?: string;
  weight?: string;
}

export interface CellFormat {
  format: { horizontalAlignment?: string; verticalAlignment?: string; wrapText?: boolean; rowHeight?: number };
  font: { bold?: boolean; color?: string; italic?: boolean; name?: string; size?: number; underline?: string };
  fill: string;
  borders: Record<Edge, Border>;
}

export interface RowFormat {
  numberFormat: string[];
  cells: CellFormat[];
}

/**
 * fromRow 의 A–Y 서식(맞춤·글꼴·채우기·테두리·표시 형식·행 높이)을 first..last 행에 입힌다. 값은 건드리지 않는다.
 * 읽기는 병렬 배치, 쓰기는 순차 배치. 실패하면 throw(이미 쓴 값은 그대로).
 */
export async function copyRowFormat(wb: WorkbookSession, sheetId: string, fromRow: number, first: number, last: number): Promise<void> {
  const [format] = await readRowFormats(wb, sheetId, [fromRow], 1, LIST_ROW_WIDTH);
  const range = `A${first}:${LAST_COL}${last}`;
  const writes: BatchRequest[] = [
    { method: "PATCH", url: worksheetRangeUrl(wb.itemId, sheetId, range), body: { numberFormat: Array(last - first + 1).fill(format.numberFormat) } },
    ...styleWrites(wb.itemId, sheetId, format.cells, 1, first, last),
  ];
  await batchOk(wb, writes, "서식 쓰기");
}

/** rows 마다 firstCol..lastCol(열 번호) 칸 서식. 한 번의 병렬 배치 묶음으로 읽는다. */
export async function readRowFormats(wb: WorkbookSession, sheetId: string, rows: number[], firstCol: number, lastCol: number): Promise<RowFormat[]> {
  const url = (address: string) => worksheetRangeUrl(wb.itemId, sheetId, address);
  const cols = Array.from({ length: lastCol - firstCol + 1 }, (_, i) => columnLetter(firstCol + i));
  const perRow = 1 + cols.length * 4;
  const reads = rows.flatMap((row): BatchRequest[] => [
    { method: "GET", url: url(`${cols[0]}${row}:${cols[cols.length - 1]}${row}`) + "?$select=numberFormat" },
    ...cols.flatMap((c): BatchRequest[] => [
      { method: "GET", url: url(`${c}${row}`) + "/format?$select=horizontalAlignment,verticalAlignment,wrapText,rowHeight" },
      { method: "GET", url: url(`${c}${row}`) + "/format/font?$select=bold,color,italic,name,size,underline" },
      { method: "GET", url: url(`${c}${row}`) + "/format/fill?$select=color" },
      { method: "GET", url: url(`${c}${row}`) + "/format/borders?$select=sideIndex,style,color,weight" },
    ]),
  ]);
  const got = await wb.batch(reads, false);
  const failed = got.filter((r) => !ok(r));
  if (failed.length) throw new Error(`서식 읽기 실패 ${failed.length}/${got.length} (HTTP ${failed[0].status})`);
  return rows.map((_, r) => {
    const part = got.slice(r * perRow, (r + 1) * perRow);
    return {
      numberFormat: (part[0].body as { numberFormat: string[][] }).numberFormat[0],
      cells: cols.map((_, i) => cellFormat(part.slice(1 + i * 4, 5 + i * 4))),
    };
  });
}

/** 칸 서식(표시 형식 제외)을 first..last 행, firstCol 열부터 입히는 요청. 같은 서식이 이어진 열은 한 범위로 묶는다. */
export function styleWrites(itemId: string, sheetId: string, formats: CellFormat[], firstCol: number, first: number, last: number): BatchRequest[] {
  const rows = last - first + 1;
  const writes: BatchRequest[] = [];
  for (const run of sameFormatRuns(formats)) {
    const f = formats[run.from];
    const range = worksheetRangeUrl(itemId, sheetId, `${columnLetter(firstCol + run.from)}${first}:${columnLetter(firstCol + run.to)}${last}`);
    writes.push({ method: "PATCH", url: range + "/format", body: f.format });
    writes.push({ method: "PATCH", url: range + "/format/font", body: f.font });
    writes.push(f.fill.toUpperCase() === NO_FILL ? { method: "POST", url: range + "/format/fill/clear" } : { method: "PATCH", url: range + "/format/fill", body: { color: f.fill } });
    const sides: [string, Border][] = EDGES.map((e) => [e, f.borders[e]]);
    if (run.to > run.from) sides.push(["InsideVertical", f.borders.EdgeRight]);
    if (rows > 1) sides.push(["InsideHorizontal", f.borders.EdgeBottom.style === "None" ? f.borders.EdgeTop : f.borders.EdgeBottom]);
    for (const [side, border] of sides) writes.push({ method: "PATCH", url: `${range}/format/borders/${side}`, body: borderBody(border) });
  }
  return writes;
}

/** 순차 배치로 보내고, 하나라도 실패하면 throw. */
export async function batchOk(wb: WorkbookSession, writes: BatchRequest[], what: string): Promise<void> {
  const done = await wb.batch(writes, true);
  const failed = done.filter((r) => !ok(r));
  if (failed.length) throw new Error(`${what} 실패 ${failed.length}/${done.length} (HTTP ${failed[0].status})`);
}

function ok(r: BatchResponse): boolean {
  return r.status >= 200 && r.status < 300;
}

function cellFormat([format, font, fill, borders]: BatchResponse[]): CellFormat {
  const sides = (borders.body as { value: (Border & { sideIndex: string })[] }).value;
  const side = (name: Edge): Border => {
    const b = sides.find((s) => s.sideIndex === name);
    return { style: b?.style ?? "None", color: b?.color, weight: b?.weight };
  };
  const { horizontalAlignment, verticalAlignment, wrapText, rowHeight } = format.body as CellFormat["format"];
  const { bold, color, italic, name, size, underline } = font.body as CellFormat["font"];
  return {
    format: { horizontalAlignment, verticalAlignment, wrapText, rowHeight },
    font: { bold, color, italic, name, size, underline },
    fill: (fill.body as { color: string }).color,
    borders: { EdgeTop: side("EdgeTop"), EdgeBottom: side("EdgeBottom"), EdgeLeft: side("EdgeLeft"), EdgeRight: side("EdgeRight") },
  };
}

function sameFormatRuns(formats: CellFormat[]): { from: number; to: number }[] {
  const runs: { from: number; to: number }[] = [];
  const key = (f: CellFormat) => JSON.stringify(f);
  const sameLeftRight = (f: CellFormat) => key({ ...f, borders: { ...f.borders, EdgeLeft: f.borders.EdgeRight } }) === key(f);
  formats.forEach((f, i) => {
    const run = runs[runs.length - 1];
    if (run && key(formats[run.to]) === key(f) && sameLeftRight(f)) run.to = i;
    else runs.push({ from: i, to: i });
  });
  return runs;
}

function borderBody(b: Border): Border {
  return b.style === "None" ? { style: "None" } : b;
}
