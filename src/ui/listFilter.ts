/**
 * 리스트 탭 열별 필터(엑셀 자동 필터 비슷하게, 앱 안에서만).
 * - 보이는 행만 고른다. 엑셀의 필터·행 숨김은 건드리지 않고, 워크북에 아무것도 쓰지 않는다.
 * - 표(캐시든 새로 읽은 것이든)를 바꾸지 않고 새 배열만 돌려준다.
 */
import { cellText, type LiveRow } from "../excel/liveList";

/** 빈 칸을 목록에서 고를 때 쓰는 이름. */
export const BLANK = "(빈 칸)";

export interface ColumnFilter {
  /** 고른 값(화면 글자). null 이면 모두. */
  values: string[] | null;
  /** 이 글자가 들어간 칸만(대소문자 무시). */
  text: string;
}

export type ListFilters = Record<string, ColumnFilter>;

export function cellKey(row: LiveRow, col: string): string {
  const cell = row.cells.find((c) => c.col === col);
  const t = cell ? cellText(cell).trim() : "";
  return t || BLANK;
}

export function isActive(f: ColumnFilter | undefined): boolean {
  return !!f && (f.values !== null || f.text.trim() !== "");
}

export function activeColumns(filters: ListFilters): string[] {
  return Object.keys(filters).filter((col) => isActive(filters[col]));
}

function passes(row: LiveRow, col: string, f: ColumnFilter): boolean {
  const key = cellKey(row, col);
  if (f.values && !f.values.includes(key)) return false;
  const q = f.text.trim().toLowerCase();
  if (q && (key === BLANK || !key.toLowerCase().includes(q))) return false;
  return true;
}

/** 모든 열 필터를 통과한 행(순서 그대로). except 열은 빼고 본다(그 열의 값 목록을 만들 때). */
export function applyFilters(rows: LiveRow[], filters: ListFilters, except?: string): LiveRow[] {
  const cols = activeColumns(filters).filter((c) => c !== except);
  if (!cols.length) return rows;
  return rows.filter((row) => cols.every((col) => passes(row, col, filters[col])));
}

/**
 * 열의 값 목록(엑셀 필터 목록처럼): 다른 열 필터를 통과한 행에서 값마다 개수.
 * 날짜·숫자처럼 보이면 그 순서로, 아니면 가나다순. 빈 칸은 맨 뒤.
 */
export function columnValues(rows: LiveRow[], col: string, filters: ListFilters = {}): { value: string; count: number }[] {
  const counts = new Map<string, number>();
  for (const row of applyFilters(rows, filters, col)) {
    const key = cellKey(row, col);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  const num = (s: string) => (/^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null);
  return [...counts.entries()]
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => {
      if (a.value === BLANK) return 1;
      if (b.value === BLANK) return -1;
      const na = num(a.value);
      const nb = num(b.value);
      if (na !== null && nb !== null) return na - nb;
      return a.value.localeCompare(b.value, "ko");
    });
}

/** 값 하나를 켜고/끈다. all = 지금 목록의 모든 값. 모두 켜지면 null(필터 없음)로. */
export function toggleValue(f: ColumnFilter | undefined, value: string, all: string[]): ColumnFilter {
  const base = f ?? { values: null, text: "" };
  const now = new Set(base.values ?? all);
  if (now.has(value)) now.delete(value);
  else now.add(value);
  const values = all.every((v) => now.has(v)) && [...now].every((v) => all.includes(v)) ? null : all.filter((v) => now.has(v));
  return { ...base, values };
}

export function setText(f: ColumnFilter | undefined, text: string): ColumnFilter {
  return { ...(f ?? { values: null, text: "" }), text };
}

/** 열 필터 하나 지우기. */
export function clearColumn(filters: ListFilters, col: string): ListFilters {
  const next = { ...filters };
  delete next[col];
  return next;
}

/** 표시용 열: 머리글이 있는 열(+ 날짜·도면·판정은 머리글이 비어도 넣는다). */
export function filterableColumns(header: { col: string; title: string }[]): { col: string; title: string }[] {
  const always = new Set(["E", "I", "P"]);
  return header.filter((h) => h.title.trim() || always.has(h.col)).map((h) => ({ col: h.col, title: h.title.trim() || h.col }));
}

/** '120 / 1,601행' */
export function countText(shown: number, total: number): string {
  const n = (x: number) => x.toLocaleString("ko-KR");
  return `${n(shown)} / ${n(total)}행 표시`;
}
