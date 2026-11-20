/**
 * 리스트 탭 휴대폰 보기: 검사 날짜마다 요약 한 줄, 누르면 그 날짜의 행이 짧게 펼쳐진다(가로 스크롤 없음).
 * 「수정」 중 행을 누르면 칸이 열린다: 기본은 주요 열 다섯(날짜·라인·포인트·판정·비고), '모든 열' 이면 나머지 열도.
 * 보기만 바꾼다 — 엑셀·행 순서(순번)는 그대로.
 */
import { rangesFromPoints } from "../domain/history";
import { cellText, type LiveRow } from "../excel/liveList";

/** 주요 열: E 날짜 · F 라인(도면) · G 포인트 · P 판정 · S 비고 */
export const KEY_COLS = ["E", "F", "G", "P", "S"] as const;
export type ColumnMode = "key" | "all";
export const MODE_LABEL: Record<ColumnMode, string> = { key: "주요 열", all: "모든 열" };
/** 한 번에 그리는 날짜 요약 수. '더 보기' 로 늘린다. */
export const DATE_PAGE = 30;

const FALLBACK_TITLE: Record<string, string> = { E: "날짜", F: "라인", G: "포인트", P: "판정", S: "비고" };

/** 보여 줄 열(머리글 순서 그대로). key: 주요 다섯, all: 머리글이 있는 모든 열(+주요 열). */
export function visibleColumns(header: { col: string; title: string }[], mode: ColumnMode): { col: string; title: string }[] {
  const keys = new Set<string>(KEY_COLS);
  return header
    .filter((h) => (mode === "key" ? keys.has(h.col) : keys.has(h.col) || h.title.trim() !== ""))
    .map((h) => ({ col: h.col, title: h.title.trim() || FALLBACK_TITLE[h.col] || h.col }));
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

/** 행을 검사 날짜(E)별로 요약한다. 최근 날짜가 먼저, 날짜 없는 행은 맨 뒤. 묶음 안은 행 순서(순번) 그대로. */
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

/** 펼친 날짜(한 번에 하나)와 그 날짜의 「수정」 상태, 수정 중 연 행. */
export type DateFocus = { date: string; editing: boolean; row: number | null } | null;

/** 요약을 누르면 그 날짜만 펼친다(다른 날짜는 접힘). 펼친 날짜를 다시 누르면 접는다. */
export function toggleDate(focus: DateFocus, date: string): DateFocus {
  return focus?.date === date ? null : { date, editing: false, row: null };
}

/** 펼친 날짜의 「수정」/「완료」. */
export function toggleEditing(focus: DateFocus): DateFocus {
  return focus && { date: focus.date, editing: !focus.editing, row: null };
}

/** 「수정」 중에만 행을 눌러 칸을 연다. 연 행을 다시 누르면 닫는다. */
export function toggleRow(focus: DateFocus, row: number): DateFocus {
  return focus?.editing ? { ...focus, row: focus.row === row ? null : row } : focus;
}
