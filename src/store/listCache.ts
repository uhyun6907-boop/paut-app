/**
 * 리스트 탭(세 번째 탭) 기기 캐시: 마지막으로 OneDrive 에서 잘 읽은 현장 시트를 IndexedDB(kv)에 둔다.
 * 탭을 열면 이 사본을 바로 보여 주고, 뒤에서 새로 읽어 바뀐 행만 바꾼다.
 * - 엑셀에는 아무것도 쓰지 않는다(보기용 사본).
 * - 크기: 칸 객체 대신 행마다 값 배열 + 수식 칸 비트 하나로 줄여 둔다(A–Y 25열 → 32비트 안).
 */
import type { ListBookKind } from "../domain/types";
import type { LiveRow, LiveTable } from "../excel/liveList";
import { columnLetter } from "../excel/ooxml";
import type { GraphCell } from "../sync/graph";
import { kvDel, kvGet, kvSet } from "./db";

const VERSION = 1;
const PREFIX = "live-list:v1:";

interface CachedList {
  v: number;
  savedAt: number;
  book: ListBookKind;
  itemId: string;
  fileName: string;
  sheetId: string;
  sheetName: string;
  eTag?: string;
  header: { col: string; title: string }[];
  /** [행 번호, 값(A부터), 수식 칸 비트] */
  rows: [number, GraphCell[], number][];
}

export interface CachedTable {
  table: LiveTable;
  savedAt: number;
}

/** 현장(파일 종류 + 시트 이름)마다 하나. */
export function listCacheKey(book: ListBookKind, sheetName: string): string {
  return `${PREFIX}${book}:${sheetName.trim()}`;
}

export function packTable(table: LiveTable, savedAt = Date.now()): CachedList {
  return {
    v: VERSION,
    savedAt,
    book: table.book,
    itemId: table.itemId,
    fileName: table.fileName,
    sheetId: table.sheetId,
    sheetName: table.sheetName,
    eTag: table.eTag,
    header: table.header.map((h) => ({ col: h.col, title: h.title })),
    rows: table.rows.map((row) => {
      let mask = 0;
      const values = row.cells.map((cell, i) => {
        if (cell.formula) mask |= 1 << i;
        return cell.value ?? "";
      });
      return [row.row, values, mask];
    }),
  };
}

export function unpackTable(data: CachedList): LiveTable {
  const rows: LiveRow[] = data.rows.map(([r, values, mask]) => ({
    row: r,
    cells: values.map((value, i) => {
      const col = columnLetter(i + 1);
      return { address: `${col}${r}`, col, row: r, value, formula: (mask & (1 << i)) !== 0 };
    }),
  }));
  return {
    book: data.book,
    itemId: data.itemId,
    fileName: data.fileName,
    sheetId: data.sheetId,
    sheetName: data.sheetName,
    eTag: data.eTag,
    header: data.header,
    rows,
  };
}

/** 저장된 사본. 없거나 형식이 다르면 null(캐시는 없어도 된다 — 오류를 내지 않는다). */
export async function readListCache(book: ListBookKind, sheetName: string): Promise<CachedTable | null> {
  try {
    const data = await kvGet<CachedList>(listCacheKey(book, sheetName));
    if (!data || data.v !== VERSION || !Array.isArray(data.rows)) return null;
    return { table: unpackTable(data), savedAt: data.savedAt };
  } catch {
    return null;
  }
}

/** 잘 읽은 표로 덮어쓴다. 저장 실패(용량 등)는 화면을 막지 않는다. */
export async function writeListCache(book: ListBookKind, sheetName: string, table: LiveTable): Promise<void> {
  try {
    await kvSet(listCacheKey(book, sheetName), packTable(table));
  } catch {
    /* 기기 저장소가 가득 차도 리스트는 그대로 보인다 */
  }
}

export async function clearListCache(book: ListBookKind, sheetName: string): Promise<void> {
  try {
    await kvDel(listCacheKey(book, sheetName));
  } catch {
    /* 무시 */
  }
}
