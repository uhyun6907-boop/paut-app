import type { InspectionRecord } from "./types";
import { isPassing } from "./billing";
import { canonicalPoint, fmtPoint, parsePoint } from "./point";
import { FAIL, PASS } from "./types";

export interface DaySummary {
  date: string;
  ok: number;
  ng: number;
  other: number;
  lines: string[];
}

export function visibleRecords(records: InspectionRecord[]): InspectionRecord[] {
  return records.filter((r) => !r.deleted);
}

export function historyByDate(records: InspectionRecord[]): DaySummary[] {
  const map = new Map<string, DaySummary & { lineSet: Set<string> }>();
  for (const r of visibleRecords(records)) {
    const date = r.date || "날짜없음";
    let row = map.get(date);
    if (!row) {
      row = { date, ok: 0, ng: 0, other: 0, lines: [], lineSet: new Set() };
      map.set(date, row);
    }
    if (isPassing(r.result)) row.ok += 1;
    else if (r.result === "불합격") row.ng += 1;
    else row.other += 1;
    if (r.line) row.lineSet.add(r.line);
  }
  return [...map.values()]
    .map(({ lineSet, ...row }) => ({ ...row, lines: [...lineSet] }))
    .sort((a, b) => b.date.localeCompare(a.date));
}

export interface RecentLine {
  line: string;
  date: string;
}

/** 현장 라인을 마지막 검사일 내림차순으로. 엑셀 기록과 앱 기록을 함께 본다. */
export function linesByLatestDate(records: InspectionRecord[]): RecentLine[] {
  const last = new Map<string, string>();
  for (const rec of visibleRecords(records)) {
    if (!rec.line) continue;
    const prev = last.get(rec.line);
    if (!prev || rec.date > prev) last.set(rec.line, rec.date);
  }
  return [...last.entries()]
    .sort((a, b) => b[1].localeCompare(a[1]) || a[0].localeCompare(b[0]))
    .map(([line, date]) => ({ line, date }));
}

export function latestLine(records: InspectionRecord[]): string {
  return linesByLatestDate(records)[0]?.line ?? "";
}

export type LineSort = "recent" | "name";

/** 최신순은 검사일, 라인순은 이름 자연 정렬(숫자가 있는 번호 순). */
export function sortLineRows(rows: RecentLine[], mode: LineSort): RecentLine[] {
  const copy = [...rows];
  if (mode === "name") {
    copy.sort((a, b) => a.line.localeCompare(b.line, "ko", { numeric: true, sensitivity: "base" }));
  }
  return copy;
}

export function recentLines(records: InspectionRecord[], limit = 16): string[] {
  return linesByLatestDate(records).slice(0, limit).map((row) => row.line);
}

export function pointsOnLine(records: InspectionRecord[], line: string): string[] {
  return visibleRecords(records)
    .filter((r) => r.line === line)
    .map((r) => r.point);
}

const PREFIX_ORDER = ["S", "R", "P"];

export interface PrefixRange {
  prefix: string;
  /** S01~S05 or R01~R08, R10~R21. A single number stays S05. */
  text: string;
  fails: string[];
}

export interface LineProgress {
  line: string;
  lastDate: string;
  ranges: PrefixRange[];
}

export interface DatedRanges {
  date: string;
  ranges: PrefixRange[];
}

function prefixRank(prefix: string): number {
  const i = PREFIX_ORDER.indexOf(prefix);
  return i === -1 ? 100 : i;
}

function formatRanges(prefix: string, nums: number[]): string {
  const sorted = [...new Set(nums)].sort((a, b) => a - b);
  if (!sorted.length) return "";
  const parts: string[] = [];
  let start = sorted[0];
  let prev = sorted[0];
  const flush = () => {
    parts.push(start === prev ? fmtPoint(prefix, start) : `${fmtPoint(prefix, start)}~${fmtPoint(prefix, prev)}`);
  };
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === prev + 1) prev = sorted[i];
    else {
      flush();
      start = prev = sorted[i];
    }
  }
  flush();
  return parts.join(", ");
}

/** Overall (or any subset) of points: contiguous runs, holes left as a break, fails listed separately.
 * 같은 번호의 나중 행이 합격이면 재검사가 끝난 것으로 보고 불합격 목록에서 뺀다.
 */
export function rangesFromPoints(records: { point: string; result: string; ts?: number }[]): PrefixRange[] {
  const nums = new Map<string, Set<number>>();
  const latest = new Map<string, { result: string; ts: number; order: number; label: string }>();
  let order = 0;
  for (const rec of records) {
    order += 1;
    if (rec.result === "없음") continue;
    const parsed = parsePoint(rec.point);
    if (!parsed || parsed.suffix) continue;
    let set = nums.get(parsed.prefix);
    if (!set) {
      set = new Set();
      nums.set(parsed.prefix, set);
    }
    set.add(parsed.num);
    const key = `${parsed.prefix}|${parsed.num}`;
    const ts = rec.ts ?? order;
    const prev = latest.get(key);
    if (!prev || ts > prev.ts || (ts === prev.ts && order >= prev.order)) {
      latest.set(key, { result: rec.result, ts, order, label: fmtPoint(parsed.prefix, parsed.num) });
    }
  }
  const fails = new Map<string, Set<string>>();
  for (const [key, row] of latest) {
    if (row.result !== "불합격") continue;
    const prefix = key.split("|")[0];
    let bag = fails.get(prefix);
    if (!bag) {
      bag = new Set();
      fails.set(prefix, bag);
    }
    bag.add(row.label);
  }
  return [...nums.keys()]
    .sort((a, b) => prefixRank(a) - prefixRank(b) || a.localeCompare(b))
    .map((prefix) => ({
      prefix,
      text: formatRanges(prefix, [...(nums.get(prefix) ?? [])]),
      fails: [...(fails.get(prefix) ?? [])].sort(),
    }));
}

export function shortDate(iso: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!match) return iso;
  return `${match[2]}/${match[3]}`;
}

export function recentLineProgress(records: InspectionRecord[], limit = 16): LineProgress[] {
  const byLine = new Map<string, InspectionRecord[]>();
  for (const rec of visibleRecords(records)) {
    if (!rec.line) continue;
    const list = byLine.get(rec.line) ?? [];
    list.push(rec);
    byLine.set(rec.line, list);
  }
  const rows: LineProgress[] = [];
  for (const [line, list] of byLine) {
    const lastDate = list.reduce((max, rec) => (rec.date > max ? rec.date : max), "");
    rows.push({ line, lastDate, ranges: rangesFromPoints(list) });
  }
  rows.sort((a, b) => b.lastDate.localeCompare(a.lastDate) || a.line.localeCompare(b.line));
  return rows.slice(0, limit);
}

export function lineDayProgress(records: InspectionRecord[], line: string, limit = 3): DatedRanges[] {
  const list = visibleRecords(records).filter((rec) => rec.line === line);
  const dates = [...new Set(list.map((rec) => rec.date))].filter(Boolean).sort((a, b) => b.localeCompare(a)).slice(0, limit);
  return dates.map((date) => ({
    date,
    ranges: rangesFromPoints(list.filter((rec) => rec.date === date)),
  }));
}

const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];

export function formatInspectionDate(iso: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!match) return iso || "날짜없음";
  const day = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]))).getUTCDay();
  return `${match[1]}-${match[2]}-${match[3]} (${WEEKDAYS[day]})`;
}

export function lineSummary(records: InspectionRecord[], line: string): string {
  const rows = linePointRows(records, line);
  if (!rows.length) return "";
  const ranges = rangesFromPoints(rows);
  const last = rows.reduce((max, row) => (row.date > max ? row.date : max), "");
  const fails = ranges.flatMap((range) => range.fails);
  const failText = fails.length ? ` · 불합격 ${fails.join(", ")}` : "";
  return `${ranges.map((range) => range.text).join(", ")}${failText} · 마지막 ${shortDate(last)}`;
}

export interface LinePointRow {
  id: string;
  point: string;
  prefix: string;
  num: number;
  date: string;
  dateLabel: string;
  result: string;
  fail: boolean;
}

/** Every recorded point on a line, S then R then P, each with its own date. Includes Excel imports. */
export function linePointRows(records: InspectionRecord[], line: string): LinePointRow[] {
  const needle = line.trim().toUpperCase();
  if (!needle) return [];
  const rows: LinePointRow[] = [];
  for (const rec of visibleRecords(records)) {
    if (rec.line.trim().toUpperCase() !== needle) continue;
    const parsed = parsePoint(rec.point);
    rows.push({
      id: rec.id,
      point: parsed ? (parsed.suffix ? parsed.raw : fmtPoint(parsed.prefix, parsed.num)) : rec.point,
      prefix: parsed?.prefix ?? "Z",
      num: parsed?.num ?? 9999,
      date: rec.date,
      dateLabel: formatInspectionDate(rec.date),
      result: rec.result || "—",
      fail: rec.result === "불합격",
    });
  }
  rows.sort((a, b) => {
    const prefixOrder = prefixRank(a.prefix) - prefixRank(b.prefix) || a.prefix.localeCompare(b.prefix);
    if (prefixOrder) return prefixOrder;
    if (a.num !== b.num) return a.num - b.num;
    const as = a.point.includes("-") ? a.point.split("-").slice(1).join("-") : "";
    const bs = b.point.includes("-") ? b.point.split("-").slice(1).join("-") : "";
    if (!as && bs) return -1;
    if (as && !bs) return 1;
    return as.localeCompare(bs, undefined, { numeric: true }) || a.date.localeCompare(b.date);
  });
  return rows;
}

export interface DefectStats {
  pass: number;
  fail: number;
  /** 합격이 0이면 '–'. 그 외에는 '10%' 또는 '12.5%'. */
  rate: string;
}

/** 라인+포인트. S1 과 S01 은 같고, S01-1 과 다른 라인의 S01 은 별도 개소다. */
function pointKey(rec: InspectionRecord): string {
  const line = rec.line.trim().toUpperCase();
  const point = canonicalPoint(rec.point) ?? rec.point.trim().toUpperCase();
  return `${line}||${point}`;
}

function laterRecord(next: InspectionRecord, prev: InspectionRecord): boolean {
  if (next.date !== prev.date) return next.date > prev.date;
  return next.ts >= prev.ts;
}

/**
 * 넘긴 기록의 불량률. 화면에서는 선택 현장의 모든 라인을 넣는다.
 * 합격은 지금 합격인 포인트를 한 번만 센다. 재검사 합격은 그 포인트의 합격이고, 하위 번호와 다른 라인은 별도다.
 * 불합격은 기록된 불합격 행마다 센다. 불량률은 불합격 ÷ 합격.
 */
export function defectStats(records: InspectionRecord[]): DefectStats {
  const latest = new Map<string, InspectionRecord>();
  let fail = 0;
  for (const rec of records) {
    if (rec.deleted || (rec.result !== PASS && rec.result !== FAIL)) continue;
    if (rec.result === FAIL) fail += 1;
    const key = pointKey(rec);
    const prev = latest.get(key);
    if (!prev || laterRecord(rec, prev)) latest.set(key, rec);
  }
  let pass = 0;
  for (const rec of latest.values()) if (rec.result === PASS) pass += 1;
  return { pass, fail, rate: formatDefectRate(fail, pass) };
}

export function formatDefectRate(fail: number, pass: number): string {
  if (pass <= 0) return "–";
  const rounded = Math.round((fail / pass) * 1000) / 10;
  const text = Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
  return `${text}%`;
}

export function siteLastDay(records: InspectionRecord[]): { date: string; lines: { line: string; ranges: PrefixRange[] }[] } | null {
  const visible = visibleRecords(records).filter((rec) => rec.date && rec.line);
  if (!visible.length) return null;
  const date = visible.reduce((max, rec) => (rec.date > max ? rec.date : max), "");
  const day = visible.filter((rec) => rec.date === date);
  const lines = [...new Set(day.map((rec) => rec.line))];
  return {
    date,
    lines: lines.map((line) => ({ line, ranges: rangesFromPoints(day.filter((rec) => rec.line === line)) })),
  };
}
