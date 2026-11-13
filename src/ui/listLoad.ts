/**
 * 리스트 탭 읽기 순서(stale-while-revalidate):
 * 1) 기기 캐시가 있으면 바로 보여 준다(작은 '업데이트 중' 표시만).
 * 2) 뒤에서 OneDrive 를 새로 읽어, 바뀐 행만 새 객체로 바꾼다(같은 행은 그대로 — 화면이 깜빡이지 않음).
 * 3) 새로 읽기에 실패하면 캐시를 그대로 두고 짧은 안내('캐시 표시 중 · 갱신 실패').
 * 4) 잘 읽었으면 그 현장 캐시를 덮어쓴다.
 * 엑셀에는 쓰지 않는다.
 */
import type { ListBookKind } from "../domain/types";
import type { LiveRow, LiveTable } from "../excel/liveList";
import { readListCache, writeListCache, type CachedTable } from "../store/listCache";

export const CACHE_STALE_TEXT = "캐시 표시 중 · 갱신 실패";
export const REFRESHING_TEXT = "업데이트 중…";

export type ListView =
  | { status: "loading" }
  | { status: "error"; message: string }
  | {
      status: "ok";
      table: LiveTable;
      /** 화면의 표가 기기 캐시에서 왔고 아직 새로 읽지 못했다. */
      fromCache: boolean;
      /** 뒤에서 새로 읽는 중(목록은 그대로 보인다). */
      refreshing: boolean;
      /** 새로 읽기 실패 시 이유(캐시는 그대로). */
      staleError?: string;
      savedAt?: number;
    };

/** 캐시 이름표: 파일 종류 + 설정에서 고른 파일 ID + 시트 이름. 파일을 바꾸면 다른 캐시. */
export interface CacheSlot {
  book: ListBookKind;
  fileId: string;
  sheetName: string;
}

function slotSheet(slot: CacheSlot): string {
  return `${slot.fileId}:${slot.sheetName.trim()}`;
}

export function loadListCache(slot: CacheSlot): Promise<CachedTable | null> {
  return readListCache(slot.book, slotSheet(slot));
}

export function saveListCache(slot: CacheSlot, table: LiveTable): Promise<void> {
  return writeListCache(slot.book, slotSheet(slot), table);
}

function sameRow(a: LiveRow, b: LiveRow): boolean {
  if (a.row !== b.row || a.cells.length !== b.cells.length) return false;
  for (let i = 0; i < a.cells.length; i++) {
    const x = a.cells[i];
    const y = b.cells[i];
    if (x.col !== y.col || x.formula !== y.formula || (x.value ?? "") !== (y.value ?? "")) return false;
  }
  return true;
}

/**
 * 새로 읽은 표에서 바뀌지 않은 행은 예전 행 객체를 그대로 쓴다(React 가 그 행을 다시 그리지 않는다).
 * 순서·행 수·값은 새 표 그대로.
 */
export function mergeTables(prev: LiveTable | null, fresh: LiveTable): { table: LiveTable; changed: number } {
  if (!prev) return { table: fresh, changed: fresh.rows.length };
  const old = new Map(prev.rows.map((r) => [r.row, r]));
  let changed = 0;
  const rows = fresh.rows.map((row) => {
    const before = old.get(row.row);
    if (before && sameRow(before, row)) return before;
    changed++;
    return row;
  });
  if (rows.length !== prev.rows.length) changed = Math.max(changed, 1);
  return { table: { ...fresh, rows }, changed };
}

export interface OpenListDeps {
  slot: CacheSlot | null;
  read: () => Promise<{ ok: true; table: LiveTable } | { ok: false; message: string }>;
  /** 화면에 보여 줄 상태. 같은 열기 안에서 여러 번 불린다. */
  show: (view: ListView) => void;
  /** 이 열기가 아직 유효한지(현장이 바뀌면 false). */
  alive?: () => boolean;
  /** 지금 화면에 있는 표(같은 현장 새로 불러오기 때 캐시 대신 그대로 두고 덮어쓴다). */
  current?: LiveTable | null;
}

/** 캐시 먼저 → 뒤에서 새로 읽기. 끝난 뒤의 상태를 돌려준다. */
export async function openList(deps: OpenListDeps): Promise<ListView> {
  const alive = deps.alive ?? (() => true);
  let shown: LiveTable | null = deps.current ?? null;
  let savedAt: number | undefined;
  let fromCache = false;
  if (shown) {
    deps.show({ status: "ok", table: shown, fromCache: false, refreshing: true });
  } else {
    const cached = deps.slot ? await loadListCache(deps.slot) : null;
    if (!alive()) return { status: "loading" };
    if (cached) {
      shown = cached.table;
      savedAt = cached.savedAt;
      fromCache = true;
      deps.show({ status: "ok", table: shown, fromCache: true, refreshing: true, savedAt });
    } else {
      deps.show({ status: "loading" });
    }
  }
  let res: Awaited<ReturnType<OpenListDeps["read"]>>;
  try {
    res = await deps.read();
  } catch (error) {
    res = { ok: false, message: error instanceof Error && error.message ? error.message : "검사리스트를 읽지 못했습니다." };
  }
  if (!alive()) return { status: "loading" };
  let view: ListView;
  if (res.ok) {
    const { table } = mergeTables(shown, res.table);
    view = { status: "ok", table, fromCache: false, refreshing: false };
    deps.show(view);
    if (deps.slot) await saveListCache(deps.slot, table);
  } else if (shown) {
    view = { status: "ok", table: shown, fromCache, refreshing: false, staleError: res.message, savedAt };
    deps.show(view);
  } else {
    view = { status: "error", message: res.message };
    deps.show(view);
  }
  return view;
}

/** 캐시 저장 시각 → 짧은 글자(예: 11/14 09:05). */
export function savedAtText(ms: number | undefined): string {
  if (!ms) return "";
  const d = new Date(ms);
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getMonth() + 1}/${d.getDate()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
