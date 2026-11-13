/**
 * BIFF8(.xls Workbook 스트림) 읽기와 '한 시트의 셀만' 바꾸기.
 *
 * 바꾸는 범위(그 밖의 바이트는 그대로):
 * - 대상 시트: 바뀐 셀 레코드, 그 행의 ROW 열 범위, DIMENSIONS, 그 시트의 DBCELL/INDEX(위치 값), 그 시트 수식의 캐시 값
 * - 전역: BOUNDSHEET 의 시트 시작 위치(뒤 시트가 밀린 만큼)
 * - 뒤 시트: INDEX 의 절대 위치 값(밀린 만큼). 셀·서식·병합·인쇄 설정 등은 한 바이트도 바꾸지 않는다.
 * 문자열은 SST 에 이미 있으면 그 번호(LABELSST), 없으면 LABEL 레코드로 쓴다(전역 SST 를 바꾸지 않기 위해).
 */

export class BiffError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BiffError";
  }
}

export const RT = {
  BOF: 0x0809,
  EOF: 0x000a,
  BOUNDSHEET: 0x0085,
  SST: 0x00fc,
  CONTINUE: 0x003c,
  FILEPASS: 0x002f,
  DATEMODE: 0x0022,
  INDEX: 0x020b,
  DBCELL: 0x00d7,
  ROW: 0x0208,
  DIMENSIONS: 0x0200,
  DEFCOLWIDTH: 0x0055,
  NUMBER: 0x0203,
  RK: 0x027e,
  MULRK: 0x00bd,
  LABELSST: 0x00fd,
  LABEL: 0x0204,
  RSTRING: 0x00d6,
  BLANK: 0x0201,
  MULBLANK: 0x00be,
  BOOLERR: 0x0205,
  FORMULA: 0x0006,
  STRING: 0x0207,
  SHRFMLA: 0x04bc,
  ARRAY: 0x0221,
  TABLE: 0x0236,
  MERGEDCELLS: 0x00e5,
} as const;

const CELL_TYPES = new Set<number>([RT.NUMBER, RT.RK, RT.MULRK, RT.LABELSST, RT.LABEL, RT.RSTRING, RT.BLANK, RT.MULBLANK, RT.BOOLERR, RT.FORMULA]);

export interface BiffRecord {
  type: number;
  /** 스트림 안에서 레코드 머리글 위치 */
  offset: number;
  data: Uint8Array;
}

function u16(b: Uint8Array, o: number): number {
  return b[o] | (b[o + 1] << 8);
}
function u32(b: Uint8Array, o: number): number {
  return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + b[o + 3] * 0x1000000;
}
function put16(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
}
function put32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
}
function f64(b: Uint8Array, o: number): number {
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
}
function putF64(b: Uint8Array, o: number, v: number): void {
  new DataView(b.buffer, b.byteOffset + o, 8).setFloat64(0, v, true);
}

/** start 의 BOF 부터 짝이 맞는 EOF 까지(포함). */
export function readSubstream(stream: Uint8Array, start: number): BiffRecord[] {
  const out: BiffRecord[] = [];
  let p = start;
  let depth = 0;
  while (true) {
    if (p + 4 > stream.length) throw new BiffError("BIFF 레코드가 스트림 끝을 넘습니다.");
    const type = u16(stream, p);
    const len = u16(stream, p + 2);
    if (len > 8224 || p + 4 + len > stream.length) throw new BiffError("BIFF 레코드 길이가 이상합니다.");
    out.push({ type, offset: p, data: stream.subarray(p + 4, p + 4 + len) });
    p += 4 + len;
    if (type === RT.BOF) depth += 1;
    if (type === RT.EOF) {
      depth -= 1;
      if (depth === 0) return out;
    }
    if (out.length === 1 && type !== RT.BOF) throw new BiffError("BIFF 시트가 BOF 로 시작하지 않습니다.");
  }
}

export interface SheetInfo {
  name: string;
  /** BOF 위치 */
  offset: number;
  /** 0 보임, 1 숨김, 2 아주 숨김 */
  visibility: number;
  /** 0 워크시트, 2 차트, 6 VB */
  kind: number;
  /** BOUNDSHEET 레코드 머리글 위치(전역) */
  boundsheetAt: number;
}

export interface XlsWorkbook {
  stream: Uint8Array;
  globals: BiffRecord[];
  sheets: SheetInfo[];
  sst: string[];
  date1904: boolean;
}

export function parseWorkbook(stream: Uint8Array): XlsWorkbook {
  const globals = readSubstream(stream, 0);
  const bof = globals[0].data;
  if (u16(bof, 0) !== 0x0600 || u16(bof, 2) !== 0x0005) throw new BiffError("BIFF8(엑셀 97-2003) 통합 문서가 아닙니다.");
  if (globals.some((r) => r.type === RT.FILEPASS)) throw new BiffError("암호가 걸린 .xls 는 쓸 수 없습니다.");
  const sheets: SheetInfo[] = [];
  for (const r of globals) {
    if (r.type !== RT.BOUNDSHEET) continue;
    const d = r.data;
    const cch = d[6];
    const high = d[7] & 1;
    let name = "";
    for (let i = 0; i < cch; i++) name += String.fromCharCode(high ? u16(d, 8 + i * 2) : d[8 + i]);
    sheets.push({ name, offset: u32(d, 0), visibility: d[4], kind: d[5], boundsheetAt: r.offset });
  }
  for (const s of sheets) {
    if (s.offset + 8 > stream.length || u16(stream, s.offset) !== RT.BOF) throw new BiffError(`시트 '${s.name}' 시작 위치가 이상합니다.`);
  }
  const date1904 = globals.some((r) => r.type === RT.DATEMODE && u16(r.data, 0) === 1);
  return { stream, globals, sheets, sst: readSst(globals), date1904 };
}

/** SST(+CONTINUE)의 문자열. 이어지는 CONTINUE 에서 글자 폭 표시가 다시 나오는 규칙을 따른다. */
export function readSst(globals: BiffRecord[]): string[] {
  const at = globals.findIndex((r) => r.type === RT.SST);
  if (at < 0) return [];
  const segs: Uint8Array[] = [globals[at].data];
  for (let i = at + 1; i < globals.length && globals[i].type === RT.CONTINUE; i++) segs.push(globals[i].data);
  let seg = 0;
  let pos = 8;
  const unique = u32(segs[0], 4);
  const out: string[] = [];
  const ensure = () => {
    while (seg < segs.length && pos >= segs[seg].length) {
      seg += 1;
      pos = 0;
    }
    if (seg >= segs.length) throw new BiffError("SST 가 중간에 끝납니다.");
  };
  const byte = () => {
    ensure();
    return segs[seg][pos++];
  };
  const skip = (n: number) => {
    while (n > 0) {
      ensure();
      const k = Math.min(n, segs[seg].length - pos);
      pos += k;
      n -= k;
    }
  };
  for (let i = 0; i < unique; i++) {
    const cch = byte() | (byte() << 8);
    const flags = byte();
    let high = flags & 1;
    const rich = flags & 8 ? byte() | (byte() << 8) : 0;
    const ext = flags & 4 ? (byte() | (byte() << 8) | (byte() << 16)) + byte() * 0x1000000 : 0;
    let s = "";
    let n = 0;
    while (n < cch) {
      if (pos >= segs[seg].length) {
        seg += 1;
        pos = 0;
        if (seg >= segs.length) throw new BiffError("SST 문자열이 중간에 끝납니다.");
        high = segs[seg][pos++] & 1;
      }
      if (high) {
        s += String.fromCharCode(segs[seg][pos] | (segs[seg][pos + 1] << 8));
        pos += 2;
      } else {
        s += String.fromCharCode(segs[seg][pos]);
        pos += 1;
      }
      n += 1;
    }
    skip(rich * 4 + ext);
    out.push(s);
  }
  return out;
}

export type CellValue =
  | { kind: "number"; n: number }
  | { kind: "string"; text: string }
  | { kind: "blank" }
  | { kind: "bool"; b: boolean }
  | { kind: "error"; code: number }
  | { kind: "formula"; cached: number | string | boolean | null; rgce: Uint8Array; flags: number };

export interface XlsCell {
  row: number;
  col: number;
  xf: number;
  value: CellValue;
}

export function rkToNumber(rk: number): number {
  let v: number;
  if (rk & 2) v = (rk | 0) >> 2;
  else {
    const b = new Uint8Array(8);
    put32(b, 4, (rk & 0xfffffffc) >>> 0);
    v = f64(b, 0);
  }
  return rk & 1 ? v / 100 : v;
}

function readXlUnicode(d: Uint8Array, o: number): string {
  const cch = u16(d, o);
  const high = d[o + 2] & 1;
  let s = "";
  for (let i = 0; i < cch; i++) s += String.fromCharCode(high ? u16(d, o + 3 + i * 2) : d[o + 3 + i]);
  return s;
}

/** 셀 레코드 하나 → 셀들. 셀 레코드가 아니면 빈 배열. */
export function cellsOf(r: BiffRecord, sst: string[], next?: BiffRecord): XlsCell[] {
  const d = r.data;
  switch (r.type) {
    case RT.NUMBER:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "number", n: f64(d, 6) } }];
    case RT.RK:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "number", n: rkToNumber(u32(d, 6)) } }];
    case RT.MULRK: {
      const row = u16(d, 0);
      const first = u16(d, 2);
      const n = (d.length - 6) / 6;
      const out: XlsCell[] = [];
      for (let i = 0; i < n; i++) out.push({ row, col: first + i, xf: u16(d, 4 + i * 6), value: { kind: "number", n: rkToNumber(u32(d, 6 + i * 6)) } });
      return out;
    }
    case RT.LABELSST:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "string", text: sst[u32(d, 6)] ?? "" } }];
    case RT.LABEL:
    case RT.RSTRING:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "string", text: readXlUnicode(d, 6) } }];
    case RT.BLANK:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "blank" } }];
    case RT.MULBLANK: {
      const row = u16(d, 0);
      const first = u16(d, 2);
      const n = (d.length - 6) / 2;
      const out: XlsCell[] = [];
      for (let i = 0; i < n; i++) out.push({ row, col: first + i, xf: u16(d, 4 + i * 2), value: { kind: "blank" } });
      return out;
    }
    case RT.BOOLERR:
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: d[7] ? { kind: "error", code: d[6] } : { kind: "bool", b: !!d[6] } }];
    case RT.FORMULA: {
      let cached: number | string | boolean | null;
      if (d[12] === 0xff && d[13] === 0xff) {
        if (d[6] === 0) cached = next?.type === RT.STRING ? readXlUnicode(next.data, 0) : "";
        else if (d[6] === 1) cached = !!d[8];
        else if (d[6] === 3) cached = "";
        else cached = null;
      } else cached = f64(d, 6);
      const cce = u16(d, 20);
      return [{ row: u16(d, 0), col: u16(d, 2), xf: u16(d, 4), value: { kind: "formula", cached, rgce: d.subarray(22, 22 + cce), flags: u16(d, 14) } }];
    }
    default:
      return [];
  }
}

export interface XlsSheet {
  info: SheetInfo;
  records: BiffRecord[];
  cells: Map<string, XlsCell>;
  merges: { r1: number; r2: number; c1: number; c2: number }[];
}

export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

export function readXlsSheet(wb: XlsWorkbook, name: string): XlsSheet {
  const info = wb.sheets.find((s) => s.name === name);
  if (!info) throw new BiffError(`시트 '${name}' 가 없습니다.`);
  const records = readSubstream(wb.stream, info.offset);
  const cells = new Map<string, XlsCell>();
  const merges: XlsSheet["merges"] = [];
  records.forEach((r, i) => {
    for (const c of cellsOf(r, wb.sst, records[i + 1])) cells.set(cellKey(c.row, c.col), c);
    if (r.type === RT.MERGEDCELLS) {
      const n = u16(r.data, 0);
      for (let k = 0; k < n; k++) merges.push({ r1: u16(r.data, 2 + k * 8), r2: u16(r.data, 4 + k * 8), c1: u16(r.data, 6 + k * 8), c2: u16(r.data, 8 + k * 8) });
    }
  });
  return { info, records, cells, merges };
}

/* ───────────────────────── 수식 토큰(쓰기·간단 계산) ───────────────────────── */

const PTG_REF = 0x44; // 값 클래스 ptgRef
const PTG_MUL = 0x05;

/** `=<a열><행>*<b열><행>` (같은 행 두 칸 곱). 기존 금액 수식(F*G)과 같은 토큰 모양. */
export function rgceMulSameRow(row: number, colA: number, colB: number): Uint8Array {
  const out = new Uint8Array(11);
  out[0] = PTG_REF;
  put16(out, 1, row);
  put16(out, 3, colA | 0xc000);
  out[5] = PTG_REF;
  put16(out, 6, row);
  put16(out, 8, colB | 0xc000);
  out[10] = PTG_MUL;
  return out;
}

type Lookup = (row: number, col: number) => number | null;

/**
 * 시트 안 참조·사칙연산·SUM/ROUND/ROUNDDOWN/ROUNDUP 만 계산한다. 다른 토큰이 있으면 null(계산 못함).
 */
export function evalRgce(rgce: Uint8Array, at: { row: number; col: number }, get: Lookup): number | null {
  const st: (number | number[])[] = [];
  let p = 0;
  const num = (v: number | number[] | undefined): number | null => {
    if (v == null) return null;
    if (Array.isArray(v)) return v.length === 1 ? v[0] : null;
    return v;
  };
  const refRow = (raw: number, rel: boolean) => (rel ? raw : raw);
  void at;
  while (p < rgce.length) {
    const t = rgce[p++];
    const base = t & 0x1f | (t & 0x60 ? 0x20 : 0); // 클래스 비트를 정규화
    const kind = t >= 0x20 ? (t & 0x1f) | 0x20 : t;
    void base;
    switch (kind) {
      case 0x24: {
        // ptgRef
        const r = u16(rgce, p);
        const c = u16(rgce, p + 2) & 0x00ff;
        p += 4;
        const v = get(refRow(r, false), c);
        if (v == null) return null;
        st.push(v);
        break;
      }
      case 0x25: {
        // ptgArea
        const r1 = u16(rgce, p);
        const r2 = u16(rgce, p + 2);
        const c1 = u16(rgce, p + 4) & 0x00ff;
        const c2 = u16(rgce, p + 6) & 0x00ff;
        p += 8;
        const vals: number[] = [];
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            const v = get(r, c);
            if (v == null) return null;
            vals.push(v);
          }
        }
        st.push(vals);
        break;
      }
      case 0x1e:
        st.push(u16(rgce, p));
        p += 2;
        break;
      case 0x1f:
        st.push(f64(rgce, p));
        p += 8;
        break;
      case 0x03:
      case 0x04:
      case 0x05:
      case 0x06: {
        const b = num(st.pop());
        const a = num(st.pop());
        if (a == null || b == null) return null;
        st.push(t === 3 ? a + b : t === 4 ? a - b : t === 5 ? a * b : b === 0 ? NaN : a / b);
        break;
      }
      case 0x12:
        break;
      case 0x13: {
        const a = num(st.pop());
        if (a == null) return null;
        st.push(-a);
        break;
      }
      case 0x15:
        break;
      case 0x19: {
        const grbit = rgce[p];
        p += 3;
        if (grbit & 0x10) {
          // tAttrSum: 위 값 하나의 합
          const v = st.pop();
          if (v == null) return null;
          st.push(Array.isArray(v) ? v.reduce((s, x) => s + x, 0) : v);
        } else if (grbit & 0x40 || grbit & 0x01) {
          // 공백·volatile 표시: 값에 영향 없음
        } else return null;
        break;
      }
      case 0x21:
      case 0x22: {
        const argc = kind === 0x22 ? rgce[p++] & 0x7f : -1;
        const fn = u16(rgce, p) & 0x7fff;
        p += 2;
        const n = argc >= 0 ? argc : fn === 27 || fn === 212 || fn === 213 ? 2 : fn === 4 ? -1 : -1;
        if (fn === 4) {
          if (n < 0) return null;
          let sum = 0;
          for (let i = 0; i < n; i++) {
            const v = st.pop();
            if (v == null) return null;
            sum += Array.isArray(v) ? v.reduce((s, x) => s + x, 0) : v;
          }
          st.push(sum);
        } else if (fn === 27 || fn === 212 || fn === 213) {
          const digits = num(st.pop());
          const v = num(st.pop());
          if (digits == null || v == null) return null;
          const f = 10 ** Math.trunc(digits);
          const x = v * f;
          const r = fn === 27 ? Math.sign(x) * Math.round(Math.abs(x) + 1e-9) : fn === 213 ? Math.trunc(x + Math.sign(x) * 1e-9) : Math.sign(x) * Math.ceil(Math.abs(x) - 1e-9);
          st.push(r / f);
        } else return null;
        break;
      }
      default:
        return null;
    }
  }
  if (st.length !== 1) return null;
  const v = num(st[0]);
  return v == null || !Number.isFinite(v) ? null : v;
}

/* ───────────────────────── 시트 셀 바꾸기 ───────────────────────── */

export type CellWrite =
  | { kind: "number"; n: number }
  | { kind: "string"; text: string }
  | { kind: "blank" }
  | { kind: "formula"; rgce: Uint8Array };

export interface CellEdit {
  row: number;
  col: number;
  value: CellWrite;
  /** 서식(XF). 없으면 그 칸에 원래 있던 서식. */
  xf?: number;
}

export interface Item {
  /** 원래 레코드(그대로 두면 바이트 그대로) 또는 새 레코드 */
  type: number;
  data: Uint8Array;
  /** 원래 레코드면 그 머리글 위치 */
  orig?: number;
}

function rec(type: number, data: Uint8Array): Item {
  return { type, data };
}

function numberRecord(row: number, col: number, xf: number, n: number): Item {
  const d = new Uint8Array(14);
  put16(d, 0, row);
  put16(d, 2, col);
  put16(d, 4, xf);
  putF64(d, 6, n);
  return rec(RT.NUMBER, d);
}

function blankRecord(row: number, col: number, xf: number): Item {
  const d = new Uint8Array(6);
  put16(d, 0, row);
  put16(d, 2, col);
  put16(d, 4, xf);
  return rec(RT.BLANK, d);
}

function stringRecord(row: number, col: number, xf: number, text: string, sst: string[]): Item {
  const idx = sst.indexOf(text);
  if (idx >= 0) {
    const d = new Uint8Array(10);
    put16(d, 0, row);
    put16(d, 2, col);
    put16(d, 4, xf);
    put32(d, 6, idx);
    return rec(RT.LABELSST, d);
  }
  if (text.length > 255) throw new BiffError("문자열이 너무 깁니다.");
  const high = [...text].some((ch) => ch.charCodeAt(0) > 0xff);
  const d = new Uint8Array(9 + text.length * (high ? 2 : 1));
  put16(d, 0, row);
  put16(d, 2, col);
  put16(d, 4, xf);
  put16(d, 6, text.length);
  d[8] = high ? 1 : 0;
  for (let i = 0; i < text.length; i++) {
    if (high) put16(d, 9 + i * 2, text.charCodeAt(i));
    else d[9 + i] = text.charCodeAt(i);
  }
  return rec(RT.LABEL, d);
}

function formulaRecord(row: number, col: number, xf: number, rgce: Uint8Array, cached: number): Item {
  const d = new Uint8Array(22 + rgce.length);
  put16(d, 0, row);
  put16(d, 2, col);
  put16(d, 4, xf);
  putF64(d, 6, cached);
  put16(d, 14, 0);
  put32(d, 16, 0);
  put16(d, 20, rgce.length);
  d.set(rgce, 22);
  return rec(RT.FORMULA, d);
}

function multiRecord(type: number, row: number, parts: { col: number; xf: number; rk?: number }[]): Item {
  if (type === RT.MULBLANK) {
    if (parts.length === 1) return blankRecord(row, parts[0].col, parts[0].xf);
    const d = new Uint8Array(6 + parts.length * 2);
    put16(d, 0, row);
    put16(d, 2, parts[0].col);
    parts.forEach((p, i) => put16(d, 4 + i * 2, p.xf));
    put16(d, 4 + parts.length * 2, parts[parts.length - 1].col);
    return rec(RT.MULBLANK, d);
  }
  if (parts.length === 1) {
    const d = new Uint8Array(10);
    put16(d, 0, row);
    put16(d, 2, parts[0].col);
    put16(d, 4, parts[0].xf);
    put32(d, 6, parts[0].rk!);
    return rec(RT.RK, d);
  }
  const d = new Uint8Array(6 + parts.length * 6);
  put16(d, 0, row);
  put16(d, 2, parts[0].col);
  parts.forEach((p, i) => {
    put16(d, 4 + i * 6, p.xf);
    put32(d, 6 + i * 6, p.rk!);
  });
  put16(d, 4 + parts.length * 6, parts[parts.length - 1].col);
  return rec(RT.MULRK, d);
}

export interface PatchResult {
  stream: Uint8Array;
  /** 대상 시트가 늘어난(줄어든) 바이트 수 */
  delta: number;
  /** 캐시 값을 다시 계산하지 못해 '열 때 다시 계산' 표시를 붙인 수식 칸 */
  forcedRecalc: string[];
}

/**
 * 시트 하나의 셀만 바꾼 새 Workbook 스트림.
 * - 공유 수식·배열 수식 칸, 행(ROW)이 없는 칸은 바꾸지 않고 오류를 낸다.
 * - 그 시트 수식의 캐시 값은 다시 계산한다(못 하면 열 때 다시 계산하도록 표시).
 */
export function patchSheetCells(stream: Uint8Array, sheetName: string, edits: CellEdit[]): PatchResult {
  const wb = parseWorkbook(stream);
  const sheet = readXlsSheet(wb, sheetName);
  if (sheet.info.kind !== 0) throw new BiffError("워크시트가 아닙니다.");
  const byRow = new Map<number, Map<number, CellEdit>>();
  for (const e of edits) {
    if (!byRow.has(e.row)) byRow.set(e.row, new Map());
    byRow.get(e.row)!.set(e.col, e);
  }
  const rowRecs = new Map<number, number>();
  sheet.records.forEach((r, i) => {
    if (r.type === RT.ROW) rowRecs.set(u16(r.data, 0), i);
  });
  for (const row of byRow.keys()) if (!rowRecs.has(row)) throw new BiffError(`${row + 1}행이 시트에 없습니다(행을 새로 만들지 않습니다).`);

  const editedKeys = new Set(edits.map((e) => cellKey(e.row, e.col)));
  for (const c of sheet.cells.values()) {
    if (!editedKeys.has(cellKey(c.row, c.col)) || c.value.kind !== "formula") continue;
    const t = c.value.rgce[0];
    if (t === 0x01 || t === 0x02) throw new BiffError("공유·배열 수식 칸은 바꾸지 않습니다.");
  }

  // 1) 레코드 목록을 만들며 대상 행의 셀 레코드만 쪼개고 바꾼다.
  const items: Item[] = [];
  const placed = new Set<string>();
  const recs = sheet.records;
  let i = 0;
  while (i < recs.length) {
    const r = recs[i];
    const cells = CELL_TYPES.has(r.type) ? cellsOf(r, wb.sst, recs[i + 1]) : [];
    const row = cells[0]?.row;
    const rowEdits = row != null ? byRow.get(row) : undefined;
    const touches = rowEdits ? cells.some((c) => rowEdits.has(c.col)) : false;
    // 이 행에 새로 들어갈(기존 레코드가 없는) 칸은 열 순서에 맞게 앞에 끼운다.
    if (rowEdits && row != null) {
      for (const [col, e] of [...rowEdits.entries()].sort((a, b) => a[0] - b[0])) {
        const key = cellKey(row, col);
        if (placed.has(key) || sheet.cells.has(key)) continue;
        if (col < cells[0].col) {
          items.push(newCell(e, 15));
          placed.add(key);
        }
      }
    }
    if (!touches) {
      items.push({ type: r.type, data: r.data, orig: r.offset });
      // 수식 뒤 STRING 등은 그대로 따라간다.
      i += 1;
      continue;
    }
    if (r.type === RT.MULBLANK || r.type === RT.MULRK) {
      const run: { col: number; xf: number; rk?: number }[] = [];
      const flush = () => {
        if (run.length) items.push(multiRecord(r.type, row!, run.splice(0)));
      };
      cells.forEach((c, k) => {
        const e = rowEdits!.get(c.col);
        if (e) {
          flush();
          items.push(newCell(e, c.xf));
          placed.add(cellKey(c.row, c.col));
        } else run.push({ col: c.col, xf: c.xf, rk: r.type === RT.MULRK ? u32(r.data, 6 + k * 6) : undefined });
      });
      flush();
      i += 1;
      continue;
    }
    // 단일 셀 레코드
    const c = cells[0];
    items.push(newCell(rowEdits!.get(c.col)!, c.xf));
    placed.add(cellKey(c.row, c.col));
    i += 1;
    // 문자열 결과 수식의 STRING 레코드는 함께 버린다.
    if (r.type === RT.FORMULA && recs[i]?.type === RT.STRING) i += 1;
    if (r.type === RT.FORMULA && (recs[i]?.type === RT.SHRFMLA || recs[i]?.type === RT.ARRAY || recs[i]?.type === RT.TABLE)) {
      throw new BiffError("공유·배열 수식 칸은 바꾸지 않습니다.");
    }
  }
  // 그 행의 마지막 셀 뒤(또는 셀이 하나도 없는 행)에 들어갈 칸
  const leftovers = edits.filter((e) => !placed.has(cellKey(e.row, e.col)) && !sheet.cells.has(cellKey(e.row, e.col)));
  for (const e of leftovers.sort((a, b) => a.row - b.row || a.col - b.col)) {
    const at = insertionIndex(items, e.row, e.col, wb.sst);
    items.splice(at, 0, newCell(e, 15));
    placed.add(cellKey(e.row, e.col));
  }

  function newCell(e: CellEdit, xfWas: number): Item {
    const xf = e.xf ?? xfWas;
    const v = e.value;
    if (v.kind === "number") return numberRecord(e.row, e.col, xf, v.n);
    if (v.kind === "string") return v.text === "" ? blankRecord(e.row, e.col, xf) : stringRecord(e.row, e.col, xf, v.text, wb.sst);
    if (v.kind === "blank") return blankRecord(e.row, e.col, xf);
    return formulaRecord(e.row, e.col, xf, v.rgce, 0);
  }

  // 2) ROW 열 범위·DIMENSIONS
  for (const e of edits) {
    const idx = items.findIndex((it) => it.type === RT.ROW && u16(it.data, 0) === e.row);
    const d = items[idx].data;
    const lo = u16(d, 2);
    const hi = u16(d, 4);
    if (e.col < lo || e.col + 1 > hi) {
      const nd = d.slice();
      put16(nd, 2, Math.min(lo, e.col));
      put16(nd, 4, Math.max(hi, e.col + 1));
      items[idx] = { type: RT.ROW, data: nd };
    }
  }
  const dimIdx = items.findIndex((it) => it.type === RT.DIMENSIONS);
  if (dimIdx >= 0) {
    const d = items[dimIdx].data;
    let r1 = u32(d, 0);
    let r2 = u32(d, 4);
    let c1 = u16(d, 8);
    let c2 = u16(d, 10);
    for (const e of edits) {
      r1 = Math.min(r1, e.row);
      r2 = Math.max(r2, e.row + 1);
      c1 = Math.min(c1, e.col);
      c2 = Math.max(c2, e.col + 1);
    }
    if (r1 !== u32(d, 0) || r2 !== u32(d, 4) || c1 !== u16(d, 8) || c2 !== u16(d, 10)) {
      const nd = d.slice();
      put32(nd, 0, r1);
      put32(nd, 4, r2);
      put16(nd, 8, c1);
      put16(nd, 10, c2);
      items[dimIdx] = { type: RT.DIMENSIONS, data: nd };
    }
  }

  // 3) 수식 캐시 값 다시 계산
  const forcedRecalc = recomputeFormulas(items, wb.sst);

  // 4) 직렬화 + DBCELL/INDEX 다시 계산
  const body = serializeSheet(items, sheet.info.offset);
  const oldLen = sheetLength(sheet.records);
  const delta = body.length - oldLen;
  const out = new Uint8Array(stream.length + delta);
  out.set(stream.subarray(0, sheet.info.offset), 0);
  out.set(body, sheet.info.offset);
  out.set(stream.subarray(sheet.info.offset + oldLen), sheet.info.offset + body.length);
  if (delta !== 0) shiftLaterSheets(out, wb, sheet.info.offset, delta);
  return { stream: out, delta, forcedRecalc };
}

function sheetLength(records: BiffRecord[]): number {
  const last = records[records.length - 1];
  return last.offset + 4 + last.data.length - records[0].offset;
}

function rowColOf(it: Item): { row: number; col: number } | null {
  if (!CELL_TYPES.has(it.type)) return null;
  return { row: u16(it.data, 0), col: u16(it.data, 2) };
}

function insertionIndex(items: Item[], row: number, col: number, sst: string[]): number {
  void sst;
  let lastSameRowOrBefore = -1;
  let rowRec = -1;
  for (let k = 0; k < items.length; k++) {
    const it = items[k];
    if (it.type === RT.ROW && u16(it.data, 0) === row) rowRec = k;
    const rc = rowColOf(it);
    if (!rc) continue;
    if (rc.row < row || (rc.row === row && rc.col < col)) {
      lastSameRowOrBefore = k;
      // 수식 뒤의 STRING 레코드는 건너뛴다.
      while (items[lastSameRowOrBefore + 1]?.type === RT.STRING) lastSameRowOrBefore += 1;
    }
  }
  if (lastSameRowOrBefore >= 0) {
    // 같은 행 블록 안이어야 한다(다음 블록으로 넘어가지 않게): 앞 칸이 같은 블록에 있으면 그 뒤.
    return lastSameRowOrBefore + 1;
  }
  // 이 행 블록의 셀 레코드 시작(마지막 ROW 뒤)
  let k = rowRec + 1;
  while (items[k]?.type === RT.ROW) k += 1;
  return k;
}

/** 시트 안 수식의 캐시 값을 다시 계산한다. 계산하지 못한 수식은 fAlwaysCalc 를 켠다. */
function recomputeFormulas(items: Item[], sst: string[]): string[] {
  const values = new Map<string, number>();
  const formulas: { idx: number; row: number; col: number; rgce: Uint8Array }[] = [];
  items.forEach((it, k) => {
    for (const c of cellsOf({ type: it.type, offset: 0, data: it.data }, sst, items[k + 1] ? { type: items[k + 1].type, offset: 0, data: items[k + 1].data } : undefined)) {
      if (c.value.kind === "number") values.set(cellKey(c.row, c.col), c.value.n);
      else if (c.value.kind === "blank") values.set(cellKey(c.row, c.col), 0);
      else if (c.value.kind === "formula") {
        formulas.push({ idx: k, row: c.row, col: c.col, rgce: c.value.rgce });
        if (typeof c.value.cached === "number") values.set(cellKey(c.row, c.col), c.value.cached);
      }
    }
  });
  const get: Lookup = (r, c) => {
    const v = values.get(cellKey(r, c));
    return v === undefined ? 0 : v;
  };
  const isString = new Set<string>();
  items.forEach((it) => {
    if (it.type === RT.LABELSST || it.type === RT.LABEL || it.type === RT.RSTRING) isString.add(cellKey(u16(it.data, 0), u16(it.data, 2)));
  });
  const getStrict: Lookup = (r, c) => (isString.has(cellKey(r, c)) ? null : get(r, c));
  const results = new Map<number, number | null>();
  for (let pass = 0; pass < 6; pass++) {
    let changed = false;
    for (const f of formulas) {
      const v = evalRgce(f.rgce, f, getStrict);
      const key = cellKey(f.row, f.col);
      if (v != null && values.get(key) !== v) {
        values.set(key, v);
        changed = true;
      }
      results.set(f.idx, v);
    }
    if (!changed) break;
  }
  const forced: string[] = [];
  for (const f of formulas) {
    const it = items[f.idx];
    const v = results.get(f.idx);
    const d = it.data;
    if (v == null) {
      // 계산 못 함: 원래 레코드면 그대로 두되, 바뀐 칸에 기대는지 모르므로 열 때 다시 계산하게 한다.
      if (!(u16(d, 14) & 0x0001)) {
        const nd = d.slice();
        put16(nd, 14, u16(d, 14) | 0x0001);
        items[f.idx] = { type: it.type, data: nd };
        forced.push(`${f.row},${f.col}`);
      }
      continue;
    }
    const isNum = !(d[12] === 0xff && d[13] === 0xff);
    if (isNum && f64(d, 6) === v) continue;
    if (!isNum) continue; // 문자열·논리 결과 수식은 손대지 않는다.
    const nd = d.slice();
    putF64(nd, 6, v);
    items[f.idx] = { type: it.type, data: nd };
  }
  return forced;
}

/**
 * 시트 레코드를 바이트로 만들고 그 시트의 DBCELL·INDEX 위치 값을 다시 계산한다.
 * DBCELL: [첫 ROW 까지 거리(4)] + 행마다 [첫 셀 레코드 위치 차이(2)]
 *   첫 값은 블록의 두 번째 ROW 레코드 시작 → 첫 행 첫 셀, 다음부터는 앞 행 첫 셀 → 이 행 첫 셀.
 * INDEX: DEFCOLWIDTH 위치, DBCELL 위치들(스트림 절대 위치).
 */
export function serializeSheet(items: Item[], base: number): Uint8Array {
  // 위치 계산
  const pos: number[] = [];
  let p = base;
  for (const it of items) {
    pos.push(p);
    p += 4 + it.data.length;
  }
  // DBCELL 다시 계산
  let blockStart = -1;
  const dbcellIdx: number[] = [];
  for (let k = 0; k < items.length; k++) {
    const it = items[k];
    if (it.type === RT.ROW && blockStart < 0) blockStart = k;
    if (it.type === RT.DBCELL) {
      dbcellIdx.push(k);
      if (blockStart < 0) continue;
      const rowIdx: number[] = [];
      for (let j = blockStart; j < k; j++) if (items[j].type === RT.ROW) rowIdx.push(j);
      const rows = rowIdx.map((j) => u16(items[j].data, 0));
      // 행마다 '그 행 자리의 첫 셀 레코드'(그 행 이후의 첫 셀). 셀이 없는 행은 다음 셀 레코드(없으면 DBCELL)를 가리킨다.
      // 엑셀이 쓴 실제 파일(승재·아전 시트 11개)의 DBCELL 과 바이트가 같은 규칙.
      const firstCell = rows.map((row) => {
        for (let j = blockStart; j < k; j++) {
          const rc = rowColOf(items[j]);
          if (rc && rc.row >= row) return j;
        }
        return k;
      });
      const old = it.data;
      const nRows = (old.length - 4) / 2;
      if (nRows !== rows.length) throw new BiffError("DBCELL 의 행 수가 행 블록과 맞지 않습니다.");
      const nd = new Uint8Array(old.length);
      put32(nd, 0, pos[k] - pos[rowIdx[0]]);
      let prev = rowIdx.length > 1 ? pos[rowIdx[1]] : pos[rowIdx[0]] + 4 + items[rowIdx[0]].data.length;
      for (let n = 0; n < rows.length; n++) {
        const fc = firstCell[n];
        put16(nd, 4 + n * 2, pos[fc] - prev);
        prev = pos[fc];
      }
      items[k] = { type: RT.DBCELL, data: nd, orig: it.orig };
      blockStart = -1;
    }
  }
  const idxAt = items.findIndex((it) => it.type === RT.INDEX);
  if (idxAt >= 0) {
    const old = items[idxAt].data;
    const nd = old.slice();
    const defcol = items.findIndex((it) => it.type === RT.DEFCOLWIDTH);
    if (defcol >= 0 && u32(old, 12) !== 0) put32(nd, 12, pos[defcol]);
    const n = (old.length - 16) / 4;
    if (n !== dbcellIdx.length) throw new BiffError("INDEX 와 DBCELL 수가 맞지 않습니다.");
    dbcellIdx.forEach((k, j) => put32(nd, 16 + j * 4, pos[k]));
    items[idxAt] = { type: RT.INDEX, data: nd };
  }
  const out = new Uint8Array(p - base);
  let o = 0;
  for (const it of items) {
    put16(out, o, it.type);
    put16(out, o + 2, it.data.length);
    out.set(it.data, o + 4);
    o += 4 + it.data.length;
  }
  return out;
}

/** 대상 시트 뒤의 시트들이 delta 만큼 밀렸다: BOUNDSHEET 시작 위치와 그 시트 INDEX 절대 위치만 고친다. */
function shiftLaterSheets(out: Uint8Array, wb: XlsWorkbook, targetOffset: number, delta: number): void {
  for (const s of wb.sheets) {
    if (s.offset <= targetOffset) continue;
    put32(out, s.boundsheetAt + 4, s.offset + delta);
    const recs = readSubstream(out, s.offset + delta);
    for (const r of recs) {
      if (r.type !== RT.INDEX) continue;
      const at = r.offset + 4;
      if (u32(out, at + 12) !== 0) put32(out, at + 12, u32(out, at + 12) + delta);
      const n = (r.data.length - 16) / 4;
      for (let j = 0; j < n; j++) put32(out, at + 16 + j * 4, u32(out, at + 16 + j * 4) + delta);
    }
  }
}

/** 검증용: INDEX 의 절대 위치 칸을 0 으로 지운 시트 바이트. */
export function sheetBytesWithoutIndexOffsets(stream: Uint8Array, offset: number): Uint8Array {
  const recs = readSubstream(stream, offset);
  const end = recs[recs.length - 1].offset + 4 + recs[recs.length - 1].data.length;
  const out = stream.slice(offset, end);
  for (const r of recs) {
    if (r.type !== RT.INDEX) continue;
    const at = r.offset - offset + 4;
    put32(out, at + 12, 0);
    for (let j = 0; j < (r.data.length - 16) / 4; j++) put32(out, at + 16 + j * 4, 0);
  }
  return out;
}

/** 검증용: 시트의 DBCELL·INDEX 를 지금 레코드 배치로 다시 계산해 원래 값과 같은지(계산 규칙 확인). */
export function dbcellConsistent(stream: Uint8Array, offset: number): boolean {
  const recs = readSubstream(stream, offset);
  const items: Item[] = recs.map((r) => ({ type: r.type, data: r.data.slice(), orig: r.offset }));
  const bytes = serializeSheet(items, offset);
  const end = recs[recs.length - 1].offset + 4 + recs[recs.length - 1].data.length;
  const orig = stream.subarray(offset, end);
  if (bytes.length !== orig.length) return false;
  for (let k = 0; k < bytes.length; k++) if (bytes[k] !== orig[k]) return false;
  return true;
}
