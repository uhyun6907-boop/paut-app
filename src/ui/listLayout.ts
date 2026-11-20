/**
 * 리스트 탭 휴대폰 보기: 행마다 카드 한 장(가로 스크롤 없음).
 * 기본은 주요 열 다섯(날짜·라인·포인트·판정·비고), '모든 열' 을 누르면 나머지 열도 카드 안에 이름:값으로.
 * 보기만 바꾼다 — 엑셀·행 순서(순번)는 그대로.
 */
import { cellText, type LiveRow } from "../excel/liveList";

/** 주요 열: E 날짜 · F 라인(도면) · G 포인트 · P 판정 · S 비고 */
export const KEY_COLS = ["E", "F", "G", "P", "S"] as const;
export type ColumnMode = "key" | "all";
export const MODE_LABEL: Record<ColumnMode, string> = { key: "주요 열", all: "모든 열" };
/** 한 번에 그리는 카드 수(아이폰에서 수천 장을 한 번에 그리지 않게). '더 보기' 로 늘린다. */
export const CARD_PAGE = 200;

const FALLBACK_TITLE: Record<string, string> = { E: "날짜", F: "라인", G: "포인트", P: "판정", S: "비고" };

/** 보여 줄 열(머리글 순서 그대로). key: 주요 다섯, all: 머리글이 있는 모든 열(+주요 열). */
export function visibleColumns(header: { col: string; title: string }[], mode: ColumnMode): { col: string; title: string }[] {
  const keys = new Set<string>(KEY_COLS);
  return header
    .filter((h) => (mode === "key" ? keys.has(h.col) : keys.has(h.col) || h.title.trim() !== ""))
    .map((h) => ({ col: h.col, title: h.title.trim() || FALLBACK_TITLE[h.col] || h.col }));
}

/** 카드 목록: 거른 행의 순서를 그대로 두고 앞에서부터 limit 장만. */
export function pageRows(rows: LiveRow[], limit: number): LiveRow[] {
  return rows.length > limit ? rows.slice(0, limit) : rows;
}

/** 검사 날짜(E) 묶음. date 가 "" 이면 날짜 없는 행. */
export interface DateGroup {
  date: string;
  rows: LiveRow[];
}

/** 행을 검사 날짜(E)별로 묶는다. 묶음은 행 순서(순번)에서 그 날짜가 처음 나온 차례, 묶음 안은 행 순서 그대로. */
export function groupByDate(rows: LiveRow[]): DateGroup[] {
  const groups = new Map<string, LiveRow[]>();
  for (const row of rows) {
    const cell = row.cells.find((c) => c.col === "E");
    const date = cell ? cellText(cell).trim() : "";
    const list = groups.get(date);
    if (list) list.push(row);
    else groups.set(date, [row]);
  }
  return [...groups].map(([date, list]) => ({ date, rows: list }));
}
