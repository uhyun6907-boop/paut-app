import { columnIndex, readSheet, type XCell, type ZipFiles } from "./ooxml";

/** Graph range 에 쓸 값. 숫자, 문자열, '=' 로 시작하면 수식, "" 이면 비움. */
export type CellValue = string | number;

export interface CellWrite {
  col: string;
  row: number;
  value: CellValue;
}

export interface RangeWrite {
  address: string;
  formulas: CellValue[][];
}

/**
 * 같은 시트의 전/후 셀을 비교해 바뀐 칸만 돌려준다.
 * 계산은 기존 OOXML 도우미(applyIljiDay, writePautLines …)가 하고, 여기서 나온 칸만 Graph 로 쓴다.
 * 그래서 병합 셀·이어쓰기 행·합계/부가세 바닥은 엑셀 파일 안에서 그대로 유지된다.
 */
export function diffSheet(before: ZipFiles, after: ZipFiles, beforeSheet: string, afterSheet = beforeSheet, beforeXml?: string): CellWrite[] {
  const a = readSheet(after, afterSheet).cells;
  const b = beforeXml != null ? cellsFromXml(before, beforeSheet, beforeXml) : readSheet(before, beforeSheet).cells;
  const refs = new Set<string>([...a.keys(), ...b.keys()]);
  const out: CellWrite[] = [];
  for (const ref of refs) {
    const next = a.get(ref);
    const prev = b.get(ref);
    const value = changedValue(prev, next);
    if (value === undefined) continue;
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) continue;
    out.push({ col: m[1], row: Number(m[2]), value });
  }
  return out.sort((x, y) => x.row - y.row || columnIndex(x.col) - columnIndex(y.col));
}

function cellsFromXml(files: ZipFiles, sheet: string, xml: string): Map<string, XCell> {
  const info = readSheet(files, sheet).info;
  const copy: ZipFiles = { ...files, [info.path]: new TextEncoder().encode(xml) };
  return readSheet(copy, sheet).cells;
}

function changedValue(prev: XCell | undefined, next: XCell | undefined): CellValue | undefined {
  if (next?.f != null) {
    if (prev?.f === next.f) return undefined;
    return `=${next.f}`;
  }
  if (next?.v != null && next.v !== "") {
    if (!prev?.f && prev?.v === next.v) return undefined;
    return typed(next);
  }
  if (prev && ((prev.v != null && prev.v !== "") || prev.f != null)) return "";
  return undefined;
}

function typed(c: XCell): CellValue {
  const v = c.v ?? "";
  if (c.t === "s" || c.t === "inlineStr" || c.t === "str") {
    // '=', '+', '-' 로 시작하는 글자는 엑셀이 수식으로 읽지 않게 ' 를 붙인다.
    return /^[=+-]/.test(v) ? `'${v}` : v;
  }
  if (c.t === "b" || c.t === "e") return v;
  const n = Number(v);
  return Number.isFinite(n) ? n : v;
}

/** 한 행에서 이어진 칸끼리 묶는다. 빈틈은 묶지 않는다(빈틈 칸을 건드리지 않기 위해). */
export function toRangeWrites(writes: CellWrite[]): RangeWrite[] {
  const byRow = new Map<number, CellWrite[]>();
  for (const w of writes) {
    const list = byRow.get(w.row) ?? [];
    list.push(w);
    byRow.set(w.row, list);
  }
  const out: RangeWrite[] = [];
  for (const [row, list] of [...byRow.entries()].sort((x, y) => x[0] - y[0])) {
    list.sort((x, y) => columnIndex(x.col) - columnIndex(y.col));
    let run: CellWrite[] = [];
    const flush = () => {
      if (!run.length) return;
      const first = run[0].col;
      const last = run[run.length - 1].col;
      out.push({ address: first === last ? `${first}${row}` : `${first}${row}:${last}${row}`, formulas: [run.map((w) => w.value)] });
      run = [];
    };
    for (const w of list) {
      const prev = run[run.length - 1];
      if (prev && columnIndex(w.col) !== columnIndex(prev.col) + 1) flush();
      run.push(w);
    }
    flush();
  }
  return out;
}
