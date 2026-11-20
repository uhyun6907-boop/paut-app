/**
 * 리스트 탭 휴대폰 보기: 검사 날짜마다 요약 한 줄, 누르면 그 날짜의 행이 엑셀 같은 표로 펼쳐진다(시트 행 순서 그대로).
 * 표의 열은 시트 열 순서: 주요 열(날짜·라인·포인트·판정·비고) 또는 모든 열(A–S).
 */
import { rangesFromPoints } from "../domain/history";
import { cellText, isDateCol, type LiveCell, type LiveRow } from "../excel/liveList";
import { columnIndex } from "../excel/ooxml";

/** 주요 열: E 날짜 · F 라인(도면) · G 포인트 · P 판정 · S 비고 */
export const KEY_COLS = ["E", "F", "G", "P", "S"] as const;
/** '모든 열' 의 마지막 열(비고). 그 오른쪽(T–Y, Y 앱 ID)은 보이지 않는다. */
export const ALL_LAST_COL = "S";
export type ColumnMode = "key" | "all";
export const MODE_LABEL: Record<ColumnMode, string> = { key: "주요 열", all: "모든 열" };
/** 한 번에 그리는 날짜 요약 수. '더 보기' 로 늘린다. */
export const DATE_PAGE = 30;

const FALLBACK_TITLE: Record<string, string> = { E: "날짜", F: "라인", G: "포인트", P: "판정", S: "비고" };

/** 엑셀 머리글(여러 줄) → 표 머리글 한 줄. 영어·한글이 섞인 머리글은 한글 줄만. */
export function gridTitle(title: string, col: string): string {
  const lines = title.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  if (!lines.length) return FALLBACK_TITLE[col] ?? col;
  const korean = lines.filter((l) => /[가-힣]/.test(l));
  return (korean.length && korean.length < lines.length ? korean : lines).join(" ");
}

/** 보여 줄 열(시트 열 순서). key: 주요 다섯, all: A–S. */
export function visibleColumns(header: { col: string; title: string }[], mode: ColumnMode): { col: string; title: string }[] {
  const keys = new Set<string>(KEY_COLS);
  const last = columnIndex(ALL_LAST_COL);
  return header
    .filter((h) => (mode === "key" ? keys.has(h.col) : columnIndex(h.col) <= last))
    .map((h) => ({ col: h.col, title: gridTitle(h.title, h.col) }));
}

/** 표 칸 글자: 날짜 칸은 26.06.18, 나머지는 엑셀 값 그대로. */
export function gridText(cell: Pick<LiveCell, "col" | "value">): string {
  const text = cellText(cell);
  return isDateCol(cell.col) ? dotDate(text) : text;
}

/** 표에서 row 바로 위·아래 행. 시트에서도 바로 붙은 행일 때만(사이에 다른 날짜·빈 행이 끼면 null). */
export function adjacentRows(rows: LiveRow[], row: LiveRow): { up: LiveRow | null; down: LiveRow | null } {
  const at = rows.indexOf(row);
  const up = rows[at - 1];
  const down = rows[at + 1];
  return { up: up && up.row === row.row - 1 ? up : null, down: down && down.row === row.row + 1 ? down : null };
}

const textOf = (row: LiveRow, col: string): string => {
  const cell = row.cells.find((c) => c.col === col);
  return cell ? cellText(cell).trim() : "";
};

export type Verdict = "ok" | "bad" | "";

/** 합부(P): 불합격이면 bad, 다른 값이면 ok, 비었으면 "". */
export function verdictOf(row: LiveRow): Verdict {
  const text = textOf(row, "P");
  return /불|FAIL|REJ/i.test(text) ? "bad" : text ? "ok" : "";
}

/** 2026-06-18 → 26.06.18. 날짜 모양이 아니면 그대로. */
export function dotDate(text: string): string {
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text.slice(2).replace(/-/g, ".") : text;
}

/** MJDH-D-D-P-58-003-200A → 003, HY-S-035-1-200A → 035-1. 번호를 못 찾으면 라인 전체. */
export function shortLine(line: string): string {
  return /(?:^|[-\s])(\d{3}(?:-\d)?)-\d+A\b/.exec(line)?.[1] ?? line;
}

export interface LinePoints {
  line: string;
  short: string;
  /** S25~S40, R25~R40 */
  ranges: string;
}

/** 검사 날짜(E) 하나의 요약. date 가 "" 이면 날짜 없는 행. */
export interface DateGroup {
  date: string;
  rows: LiveRow[];
  pass: number;
  fail: number;
  lines: LinePoints[];
}

function bucket(rows: LiveRow[], col: string): Map<string, LiveRow[]> {
  const out = new Map<string, LiveRow[]>();
  for (const row of rows) {
    const key = textOf(row, col);
    const list = out.get(key);
    if (list) list.push(row);
    else out.set(key, [row]);
  }
  return out;
}

function linePoints(rows: LiveRow[]): LinePoints[] {
  const lines = [...bucket(rows, "F")].map(([line, list]) => ({
    line,
    short: shortLine(line),
    ranges: rangesFromPoints(list.map((r) => ({ point: textOf(r, "G"), result: textOf(r, "P") }))).map((r) => r.text).join(", "),
  }));
  const shorts = new Set(lines.map((l) => l.short));
  return shorts.size === lines.length ? lines : lines.map((l) => ({ ...l, short: l.line }));
}

/** 행을 검사 날짜(E)별로 요약한다. 최근 날짜가 먼저, 날짜 없는 행은 맨 뒤. 묶음 안은 시트 행 순서 그대로. */
export function summarizeByDate(rows: LiveRow[]): DateGroup[] {
  return [...bucket(rows, "E")]
    .sort(([a], [b]) => (!a ? 1 : !b ? -1 : b.localeCompare(a)))
    .map(([date, list]) => ({
      date,
      rows: list,
      pass: list.filter((r) => verdictOf(r) === "ok").length,
      fail: list.filter((r) => verdictOf(r) === "bad").length,
      lines: linePoints(list),
    }));
}

/** 펼친 날짜(한 번에 하나)와 그 날짜의 「수정」 상태, 수정 중 메뉴를 연 행(행 번호). */
export type DateFocus = { date: string; editing: boolean; row: number | null } | null;

/** 요약을 누르면 그 날짜만 펼친다(다른 날짜는 접힘). 펼친 날짜를 다시 누르면 접는다. */
export function toggleDate(focus: DateFocus, date: string): DateFocus {
  return focus?.date === date ? null : { date, editing: false, row: null };
}

/** 펼친 날짜의 「수정」/「완료」. */
export function toggleEditing(focus: DateFocus): DateFocus {
  return focus && { date: focus.date, editing: !focus.editing, row: null };
}

/** 「수정」 중에만 행 번호를 눌러 메뉴를 연다. 연 행을 다시 누르면 닫는다. */
export function toggleRow(focus: DateFocus, row: number): DateFocus {
  return focus?.editing ? { ...focus, row: focus.row === row ? null : row } : focus;
}
