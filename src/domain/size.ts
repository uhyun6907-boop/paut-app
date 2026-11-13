export interface SizeInfo {
  a: number;
  inch: number | null;
  label: string;
  scan: number | null;
}

/** 라인 끝(또는 마지막) `350A` 형태. `PT-S-001-300A IN` 도 300A. */
export function diameterFromLine(line: string): string | null {
  const matches = [...String(line ?? "").toUpperCase().matchAll(/(\d{2,4})A/g)];
  if (!matches.length) return null;
  return `${matches[matches.length - 1][1]}A`;
}

export function diameterNumber(diameter: string): number | null {
  const m = String(diameter).toUpperCase().match(/(\d{2,4})A/);
  return m ? Number(m[1]) : null;
}

export function inchLabel(inch: number): string {
  if (inch === 0.375) return '3/8"';
  if (inch === 0.5) return '1/2"';
  if (inch === 0.75) return '3/4"';
  if (inch === 1.25) return '1¼"';
  if (inch === 1.5) return '1½"';
  return `${Number.isInteger(inch) ? inch : inch}"`;
}

export interface OdRow {
  inch: number;
  a: number;
  od: number;
  circ: number;
  scan: number;
  rs?: number;
  rl?: number;
  ds?: number;
  dl?: number;
}

export function sizeFromLine(line: string, table: OdRow[]): SizeInfo | null {
  const dia = diameterFromLine(line);
  if (!dia) return null;
  const a = diameterNumber(dia)!;
  const row = table.find((p) => p.a === a);
  if (!row) return { a, inch: null, label: dia, scan: null };
  return { a, inch: row.inch, label: `${inchLabel(row.inch)} / ${a}A`, scan: row.scan };
}

/** 시트에 적힌 RIGHT(Fn,4) 와 같은 4글자. 끝이 관경이 아니면 청구용 관경과 달라질 수 있다. */
export function excelRight4(line: string): string {
  const s = String(line ?? "");
  return s.length <= 4 ? s : s.slice(-4);
}
