/**
 * 1번 탭 카드 아래 '날짜별 기록': 날짜마다 그날 검사한 내역(도면·개수·판정)과 작업일지의 근무자·근무시간.
 * 앱에 이미 있는 기록과 기기에 있는 작업일지 사본만 읽는다. 아무것도 쓰지 않는다(엑셀·기록 그대로).
 * 기존 '검사 기록'(라인별) 블록은 그대로 두고, 이것은 그 아래에 더하는 보기다.
 */
import { isPassing } from "./billing";
import type { InspectionRecord } from "./types";
import { FAIL } from "./types";

export interface WorkLogInfo {
  dayType: string;
  workers: string[];
  timeStart: string;
  timeEnd: string;
}

export interface DayDrawing {
  /** 도면(라인). 도면번호 칸이 있으면 drawingNo 에. */
  line: string;
  drawingNo: string;
  count: number;
  pass: number;
  fail: number;
  other: number;
  failPoints: string[];
}

export interface DayEntry {
  date: string;
  total: number;
  pass: number;
  fail: number;
  other: number;
  drawings: DayDrawing[];
  /** 작업일지에 그날 근무자·시간이 없으면 null. */
  work: WorkLogInfo | null;
}

export const NO_WORKLOG_TEXT = "근무일지에 근무자·근무시간이 없습니다";

/**
 * 날짜별로 묶는다(최근 날짜가 먼저). 검사 기록이 있거나 작업일지에 근무자·시간이 있는 날만.
 * records 는 보여 줄 현장의 기록(지운 기록은 뺀다). 입력 배열·객체는 바꾸지 않는다.
 */
export function dayHistory(records: InspectionRecord[], workLog: Record<string, WorkLogInfo>, opts: { workOnlyDays?: boolean } = {}): DayEntry[] {
  const days = new Map<string, DayEntry & { byLine: Map<string, DayDrawing> }>();
  const day = (date: string) => {
    let d = days.get(date);
    if (!d) {
      const w = workLog[date];
      d = { date, total: 0, pass: 0, fail: 0, other: 0, drawings: [], work: w ? { dayType: w.dayType, workers: [...w.workers], timeStart: w.timeStart, timeEnd: w.timeEnd } : null, byLine: new Map() };
      days.set(date, d);
    }
    return d;
  };
  for (const rec of records) {
    if (rec.deleted || !/^\d{4}-\d{2}-\d{2}$/.test(rec.date)) continue;
    const d = day(rec.date);
    const line = rec.line.trim() || "(라인 없음)";
    const key = line.toUpperCase();
    let g = d.byLine.get(key);
    if (!g) {
      g = { line, drawingNo: "", count: 0, pass: 0, fail: 0, other: 0, failPoints: [] };
      d.byLine.set(key, g);
    }
    if (!g.drawingNo && rec.drawingNo?.trim()) g.drawingNo = rec.drawingNo.trim();
    g.count++;
    d.total++;
    if (rec.result === FAIL) {
      g.fail++;
      d.fail++;
      if (rec.point) g.failPoints.push(rec.point);
    } else if (isPassing(rec.result)) {
      g.pass++;
      d.pass++;
    } else {
      g.other++;
      d.other++;
    }
  }
  if (opts.workOnlyDays) {
    for (const [date, w] of Object.entries(workLog)) {
      if (/^\d{4}-\d{2}-\d{2}$/.test(date) && (w.workers.length || w.timeStart || w.timeEnd)) day(date);
    }
  }
  return [...days.values()]
    .map(({ byLine, ...d }) => ({
      ...d,
      drawings: [...byLine.values()].sort((a, b) => a.line.localeCompare(b.line, "ko", { numeric: true })),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));
}

/** '09:00–15:30' · 하나만 있으면 그것만. */
export function workTimeText(w: WorkLogInfo | null): string {
  if (!w || (!w.timeStart && !w.timeEnd)) return "";
  return `${w.timeStart || "?"}–${w.timeEnd || "?"}`;
}

/** '조우현 · 심명보' */
export function workersText(w: WorkLogInfo | null): string {
  return w ? w.workers.join(" · ") : "";
}
