/**
 * 리스트 탭 휴대폰 보기: 행마다 카드 한 장(가로 스크롤 없음).
 * 기본은 주요 열 다섯(날짜·라인·포인트·판정·비고), '모든 열' 을 누르면 나머지 열도 카드 안에 이름:값으로.
 * 보기만 바꾼다 — 엑셀·행 순서(순번)는 그대로.
 */
import type { LiveRow } from "../excel/liveList";

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
