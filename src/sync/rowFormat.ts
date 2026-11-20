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

interface CellFormat {
  format: { horizontalAlignment?: string; verticalAlignment?: string; wrapText?: boolean; rowHeight?: number };
  font: { bold?: boolean; color?: string; italic?: boolean; name?: string; size?: number; underline?: string };
  fill: string;
  borders: Record<Edge, Border>;
}

/**
 * fromRow 의 A–Y 서식(맞춤·글꼴·채우기·테두리·표시 형식·행 높이)을 first..last 행에 입힌다. 값은 건드리지 않는다.
 * 읽기는 병렬 배치, 쓰기는 순차 배치. 실패하면 throw(이미 쓴 값은 그대로).
 */
export async function copyRowFormat(wb: WorkbookSession, sheetId: string, fromRow: number, first: number, last: number): Promise<void> {
  const url = (address: string) => worksheetRangeUrl(wb.itemId, sheetId, address);
  const cols = Array.from({ length: LIST_ROW_WIDTH }, (_, i) => columnLetter(i + 1));
  const reads: BatchRequest[] = [
    { method: "GET", url: url(`A${fromRow}:${LAST_COL}${fromRow}`) + "?$select=numberFormat" },
    ...cols.flatMap((c): BatchRequest[] => [
      { method: "GET", url: url(`${c}${fromRow}`) + "/format?$select=horizontalAlignment,verticalAlignment,wrapText,rowHeight" },
      { method: "GET", url: url(`${c}${fromRow}`) + "/format/font?$select=bold,color,italic,name,size,underline" },
      { method: "GET", url: url(`${c}${fromRow}`) + "/format/fill?$select=color" },
      { method: "GET", url: url(`${c}${fromRow}`) + "/format/borders?$select=sideIndex,style,color,weight" },
    ]),
  ];
  const got = await wb.batch(reads, false);
  const failed = got.filter((r) => !ok(r));
  if (failed.length) throw new Error(`서식 읽기 실패 ${failed.length}/${got.length} (HTTP ${failed[0].status})`);
  const numberFormat = (got[0].body as { numberFormat: string[][] }).numberFormat[0];
  const formats = cols.map((_, i) => cellFormat(got.slice(1 + i * 4, 5 + i * 4)));

  const rows = last - first + 1;
  const writes: BatchRequest[] = [{ method: "PATCH", url: url(`A${first}:${LAST_COL}${last}`), body: { numberFormat: Array(rows).fill(numberFormat) } }];
  for (const run of sameFormatRuns(formats)) {
    const f = formats[run.from];
    const range = url(`${cols[run.from]}${first}:${cols[run.to]}${last}`);
    writes.push({ method: "PATCH", url: range + "/format", body: f.format });
    writes.push({ method: "PATCH", url: range + "/format/font", body: f.font });
    writes.push(f.fill.toUpperCase() === NO_FILL ? { method: "POST", url: range + "/format/fill/clear" } : { method: "PATCH", url: range + "/format/fill", body: { color: f.fill } });
    const sides: [string, Border][] = EDGES.map((e) => [e, f.borders[e]]);
    if (run.to > run.from) sides.push(["InsideVertical", f.borders.EdgeRight]);
    if (rows > 1) sides.push(["InsideHorizontal", f.borders.EdgeBottom.style === "None" ? f.borders.EdgeTop : f.borders.EdgeBottom]);
    for (const [side, border] of sides) writes.push({ method: "PATCH", url: `${range}/format/borders/${side}`, body: borderBody(border) });
  }
  const done = await wb.batch(writes, true);
  const writeFailed = done.filter((r) => !ok(r));
  if (writeFailed.length) throw new Error(`서식 쓰기 실패 ${writeFailed.length}/${done.length} (HTTP ${writeFailed[0].status})`);
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
