import { excelTimeToHHMM, hhmmToExcelTime, isoToExcelSerial, parseFlexibleDate } from "../domain/dates";
import { ILJI_COLUMNS } from "../domain/types";
import { cellText, columnLetter, readSheet, setCell, setXml, workbookSheets, type ZipFiles } from "./ooxml";

export interface IljiPair {
  column: string;
  diaCol: string;
  countCol: string;
}

export interface IljiGroup {
  diameter: string;
  count: number;
}

/**
 * 검사자·근무시간이 undefined 면 "모름": 이미 있는 행은 그 칸을 건드리지 않고, 새 행은 비워 둔다.
 * 빈 문자열은 "비움": 그 칸을 비운다. (근무시간을 비우면 S/T 도 빈칸. 0·00:00 을 쓰지 않는다.)
 */
export interface IljiUpdate {
  date: string;
  column: string;
  groups: IljiGroup[];
  dayType?: string;
  insp1?: string;
  insp2?: string;
  timeStart?: string;
  timeEnd?: string;
}

interface RowSnap {
  row: number;
  seq: number | null;
  date: string;
  dayType: string;
  values: Map<string, string>;
}

const SKIP_HEADERS = new Set(["순번", "일자", "구분", "검사자", "근무시간", "비고", "관경", "개소"]);

export function findIljiSheet(files: ZipFiles): string {
  const sheets = workbookSheets(files);
  const hit = sheets.find((s) => s.name.replace(/\s+/g, "") === "PAUTFIELD" || s.name.includes("PAUT FIELD"));
  if (!hit) throw new Error("작업일지 시트(PAUT FIELD)를 찾지 못했습니다.");
  return hit.name;
}

export function detectPairs(cells: Map<string, { col: string; row: number; v?: string }>): IljiPair[] {
  const labels: { col: string; index: number; text: string }[] = [];
  for (const c of cells.values()) {
    if (c.row !== 4 || !c.v) continue;
    const text = c.v.replace(/\s+/g, " ").trim();
    if (!text || SKIP_HEADERS.has(text)) continue;
    labels.push({ col: c.col, index: colIndex(c.col), text });
  }
  labels.sort((a, b) => a.index - b.index);
  const pairs: IljiPair[] = [];
  for (const label of labels) {
    if (label.index < 4 || label.index > 15) continue;
    pairs.push({
      column: label.text,
      diaCol: label.col,
      countCol: columnLetter(label.index + 1),
    });
  }
  if (pairs.length >= 4) return pairs;
  return ILJI_COLUMNS.map((name, i) => ({
    column: name,
    diaCol: columnLetter(4 + i * 2),
    countCol: columnLetter(5 + i * 2),
  }));
}

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function snapshotRows(cells: Map<string, { col: string; row: number; v?: string }>): RowSnap[] {
  const rows = new Map<number, Map<string, string>>();
  for (const c of cells.values()) {
    if (c.row < 6 || c.v == null) continue;
    let m = rows.get(c.row);
    if (!m) {
      m = new Map();
      rows.set(c.row, m);
    }
    m.set(c.col, c.v);
  }
  return [...rows.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([row, values]) => {
      const seqRaw = values.get("A");
      const dateRaw = values.get("B");
      const date = dateRaw && /^\d+(\.\d+)?$/.test(dateRaw) ? parseFlexibleDate(dateRaw) : dateRaw || "";
      return {
        row,
        seq: seqRaw && /^\d+$/.test(seqRaw) ? Number(seqRaw) : null,
        date: /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : "",
        dayType: values.get("C") || "",
        values,
      };
    });
}

export interface StaffColumns {
  /** 검사자 머리글의 첫 칸. 샘플은 P (조우현). */
  insp1: string;
  /** 바로 옆 칸. 샘플은 Q (둘째 검사자). 한 칸에 이름을 붙이지 않는다. */
  insp2: string;
  start: string;
  end: string;
}

/** 4행 머리글에서 검사자·근무시간 칸을 찾는다. 샘플은 검사자 P4:R5, 값은 P·Q, 근무시간은 S·T. */
export function detectStaffColumns(cells: Map<string, { col: string; row: number; v?: string }>): StaffColumns {
  let insp = "";
  let time = "";
  for (const c of cells.values()) {
    if (c.row !== 4 || !c.v) continue;
    const text = c.v.replace(/\s+/g, "");
    if (!insp && text.includes("검사자")) insp = c.col;
    if (!time && text.includes("근무시간")) time = c.col;
  }
  const inspCol = insp || "P";
  const timeCol = time || "S";
  return {
    insp1: inspCol,
    insp2: columnLetter(colIndex(inspCol) + 1),
    start: timeCol,
    end: columnLetter(colIndex(timeCol) + 1),
  };
}

/** 셀 하나에는 이름 하나만. 슬래시로 이어 붙인 값은 첫 이름만 남긴다. */
export function singleInspectorName(value: string): string {
  return value.split(/[/／,\n]/)[0]?.trim() ?? "";
}

/** 작업일지에서 날짜별 검사자. 같은 날짜가 여러 행이면 이름이 있는 칸을 남긴다. */
export function readDayStaff(files: ZipFiles): Record<string, { insp1: string; insp2: string }> {
  const name = findIljiSheet(files);
  const { cells } = readSheet(files, name);
  const staff = detectStaffColumns(cells);
  const out: Record<string, { insp1: string; insp2: string }> = {};
  for (const snap of snapshotRows(cells)) {
    if (!snap.date) continue;
    const insp1 = singleInspectorName(snap.values.get(staff.insp1) || "");
    const insp2 = singleInspectorName(snap.values.get(staff.insp2) || "");
    if (!insp1 && !insp2) continue;
    const prev = out[snap.date];
    out[snap.date] = { insp1: insp1 || prev?.insp1 || "", insp2: insp2 || prev?.insp2 || "" };
  }
  return out;
}

/** 작업일지 하루: 검사자(근무자)·근무시간·구분. 읽기만 한다. */
export interface WorkLogDay {
  date: string;
  dayType: string;
  workers: string[];
  timeStart: string;
  timeEnd: string;
}

/** 근무시간 칸 → 'HH:MM'. 엑셀 시간(0.375)·'9:00' 글자 모두. 모르면 빈 글자. */
function timeText(v: string | undefined): string {
  const t = (v ?? "").trim();
  if (!t) return "";
  if (/^\d+(\.\d+)?$/.test(t)) {
    const n = Number(t);
    return n > 0 && n < 1 ? excelTimeToHHMM(n) : "";
  }
  const m = /^(\d{1,2}):(\d{2})/.exec(t);
  return m ? `${m[1].padStart(2, "0")}:${m[2]}` : "";
}

/**
 * 작업일지에서 날짜별 근무자·근무시간(읽기만). 같은 날짜가 여러 행이면 비어 있지 않은 값을 모은다.
 * 근무자도 시간도 없는 날은 넣지 않는다.
 */
export function readWorkLog(files: ZipFiles): Record<string, WorkLogDay> {
  const { cells } = readSheet(files, findIljiSheet(files));
  const staff = detectStaffColumns(cells);
  const out: Record<string, WorkLogDay> = {};
  for (const snap of snapshotRows(cells)) {
    if (!snap.date) continue;
    const names = [staff.insp1, staff.insp2].map((c) => singleInspectorName(snap.values.get(c) || "")).filter((n) => n && !/^\d/.test(n));
    const start = timeText(snap.values.get(staff.start));
    const end = timeText(snap.values.get(staff.end));
    if (!names.length && !start && !end) continue;
    const prev = out[snap.date];
    out[snap.date] = {
      date: snap.date,
      dayType: prev?.dayType || snap.dayType,
      workers: [...new Set([...(prev?.workers ?? []), ...names])],
      timeStart: prev?.timeStart || start,
      timeEnd: prev?.timeEnd || end,
    };
  }
  return out;
}

export function readInspectors(files: ZipFiles): string[] {
  const name = findIljiSheet(files);
  const { cells } = readSheet(files, name);
  const staff = detectStaffColumns(cells);
  const counts = new Map<string, number>();
  for (const c of cells.values()) {
    if (c.row < 6) continue;
    if (c.col !== staff.insp1 && c.col !== staff.insp2) continue;
    const v = singleInspectorName(c.v || "");
    if (!v || /^\d/.test(v)) continue;
    counts.set(v, (counts.get(v) || 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1]).map(([n]) => n);
}

/**
 * 하루는 날짜가 있는 행에서 시작한다.
 * 바로 아래, 날짜도 순번도 없는 행은 관경이 더 있을 때의 이어쓰기 행이다.
 * 순번만 있고 날짜가 없는 행은 아직 쓰지 않은 다음 날 칸이다.
 */
export function applyIljiDay(files: ZipFiles, update: IljiUpdate): number {
  const sheet = findIljiSheet(files);
  const { xml, cells, info } = readSheet(files, sheet);
  const pairs = detectPairs(cells);
  const staff = detectStaffColumns(cells);
  const pair = pairs.find((p) => p.column === update.column) ?? pairs.find((p) => p.column === "그 외");
  if (!pair) throw new Error(`작업일지 열을 찾지 못했습니다: ${update.column}`);
  const snaps = snapshotRows(cells);
  const located = locateDay(snaps, update.date);
  const groups = update.groups.filter((g) => g.count > 0);
  if (!groups.length && !located.block.length) return -1;
  const rows = [...located.block];
  while (rows.length < groups.length && located.blanks.length) {
    rows.push(located.blanks.shift()!);
  }
  if (!rows.length) throw new Error("작업일지에 빈 행이 없습니다.");

  let next = xml;
  const primary = rows[0];
  const isNew = primary.date !== update.date;
  if (isNew) {
    next = setCell(next, "B" + primary.row, { kind: "number", n: isoToExcelSerial(update.date) });
    next = setCell(next, "C" + primary.row, { kind: "inline", text: update.dayType || defaultDayType(update.date) });
  } else if (update.dayType) {
    next = setCell(next, "C" + primary.row, { kind: "inline", text: update.dayType });
  }
  if (update.insp1 !== undefined) {
    const lead = singleInspectorName(update.insp1);
    if (lead) next = setCell(next, staff.insp1 + primary.row, { kind: "inline", text: lead });
  }
  if (update.insp2 !== undefined) {
    const second = singleInspectorName(update.insp2);
    next = setCell(next, staff.insp2 + primary.row, second ? { kind: "inline", text: second } : { kind: "empty" });
  }
  if (update.timeStart !== undefined) next = setCell(next, staff.start + primary.row, timeCell(update.timeStart));
  if (update.timeEnd !== undefined) next = setCell(next, staff.end + primary.row, timeCell(update.timeEnd));

  rows.forEach((snap, i) => {
    const g = groups[i];
    if (i > 0) {
      next = setCell(next, "A" + snap.row, { kind: "empty" });
      next = setCell(next, "B" + snap.row, { kind: "empty" });
      next = setCell(next, "C" + snap.row, { kind: "empty" });
    }
    if (g) {
      next = setCell(next, pair.diaCol + snap.row, { kind: "inline", text: g.diameter });
      next = setCell(next, pair.countCol + snap.row, { kind: "number", n: g.count });
    } else if (rows.length > 1 || located.block.includes(snap)) {
      next = setCell(next, pair.diaCol + snap.row, { kind: "empty" });
      next = setCell(next, pair.countCol + snap.row, { kind: "empty" });
    }
  });

  if (groups.length > rows.length) {
    const extra = groups
      .slice(rows.length)
      .map((g) => `${g.diameter}×${g.count}`)
      .join(", ");
    const prev = cellText(cells, "U", primary.row);
    const note = prev && !prev.includes("추가 ") ? `${prev} / 추가 ${extra}` : `추가 ${extra}`;
    next = setCell(next, "U" + primary.row, { kind: "inline", text: note });
  }

  setXml(files, info.path, next);
  return primary.row;
}

/** 토·일은 공휴일, 나머지는 평일. 사용자가 고른 값이 있으면 그 값을 쓴다. */
function defaultDayType(iso: string): string {
  const [y, m, d] = iso.split("-").map(Number);
  const dow = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  return dow === 0 || dow === 6 ? "공휴일" : "평일";
}

function roundTime(n: number): number {
  return Math.round(n * 1e10) / 1e10;
}

/** 시간을 모르면 셀을 비운다. 0이나 00:00으로 채우지 않는다. */
function timeCell(hhmm: string): { kind: "empty" } | { kind: "number"; n: number } {
  const value = hhmm.trim();
  if (!/^\d{1,2}:\d{2}$/.test(value)) return { kind: "empty" };
  return { kind: "number", n: roundTime(hhmmToExcelTime(value)) };
}

function locateDay(snaps: RowSnap[], date: string): { block: RowSnap[]; blanks: RowSnap[] } {
  const primaryIdx = snaps.findIndex((s) => s.date === date);
  if (primaryIdx >= 0) {
    const block = [snaps[primaryIdx]];
    for (let i = primaryIdx + 1; i < snaps.length; i++) {
      const s = snaps[i];
      if (s.date) break;
      if (s.seq != null && !hasSiteValue(s)) break;
      block.push(s);
    }
    const blanks: RowSnap[] = [];
    for (let i = primaryIdx + block.length; i < snaps.length; i++) {
      const s = snaps[i];
      if (s.date || hasSiteValue(s)) break;
      blanks.push(s);
    }
    return { block, blanks };
  }
  // 새 날짜는 마지막으로 쓴 행 아래 빈 행에만 쓴다. 실제 작업일지 220행처럼 순번·검사자만 있고
  // 날짜가 빠진 중간 행을 "빈 행"으로 잡아 지난 기록을 덮어쓰면 안 된다.
  let lastUsed = 0;
  for (const s of snaps) if (s.date || hasSiteValue(s)) lastUsed = Math.max(lastUsed, s.row);
  const blanks = snaps.filter((s) => s.row > lastUsed && !s.date && !hasSiteValue(s) && s.seq != null);
  return { block: [], blanks };
}

function hasSiteValue(s: RowSnap): boolean {
  for (const [col, v] of s.values) {
    if (col >= "D" && col <= "O" && v) return true;
  }
  return false;
}

/**
 * 작업일지에 그날·그 열로 적힌 관경/개소. 날짜 행이 없으면 null.
 * 백필: 리스트엔 있는데 작업일지엔 빠진 날을 찾을 때 쓴다.
 */
export function readIljiDay(files: ZipFiles, date: string, column: string): IljiGroup[] | null {
  const { cells } = readSheet(files, findIljiSheet(files));
  const pairs = detectPairs(cells);
  const pair = pairs.find((p) => p.column === column) ?? pairs.find((p) => p.column === "그 외");
  if (!pair) return null;
  const located = locateDay(snapshotRows(cells), date);
  if (!located.block.length) return null;
  const out: IljiGroup[] = [];
  for (const snap of located.block) {
    const diameter = snap.values.get(pair.diaCol)?.trim() ?? "";
    const count = Number(snap.values.get(pair.countCol) ?? "");
    if (diameter && Number.isFinite(count) && count > 0) out.push({ diameter, count });
  }
  return out;
}

export function readIljiCell(files: ZipFiles, row: number, col: string): string {
  const sheet = findIljiSheet(files);
  const { cells } = readSheet(files, sheet);
  const v = cellText(cells, col, row);
  if (/^\d+(\.\d+)?$/.test(v) && col === "B") return parseFlexibleDate(v);
  if (/^\d+\.\d+$/.test(v) && (col === "S" || col === "T")) return excelTimeToHHMM(Number(v));
  return v;
}

export { detectPairs as detectIljiPairs };
