import { buildBill, iljiDiameterLabel, passingRecords } from "../domain/billing";
import { isoToExcelSerial, parseFlexibleDate } from "../domain/dates";
import type { AppConfig, DayMeta, DayStaff, FileSteps, InspectionRecord, ListBookKind, SyncFile, SyncReport, SyncStatus } from "../domain/types";
import { FAIL, LIST_BOOKS, LIST_BOOK_LABEL, LIST_GM_FILE, PASS, SYNC_FILE_LABEL } from "../domain/types";
import { findMonthSheet, isRemovedSite, mappingFor, REMOVED_SITE_MAPPINGS } from "../domain/mapping";
import { readDayStaff, readInspectors, readWorkLog, applyIljiDay, findIljiSheet, readIljiDay, type IljiGroup, type WorkLogDay } from "./iljiBook";
import { appIdRows, LIST_LAYOUT, listBookOf, listDiameterValue, listSheetOf, orderForAppend, parseListWorkbook, planLineReorder, writeListRow, type ListWrite } from "./listBook";
import { dropCalcChain, readSheet, shiftRowNumbers, unzipXlsx, workbookSheets, zipXlsx, type ZipFiles } from "./ooxml";
import { expectedMonthSheetName, readLengths, readPautRows, writePautLines, type PautLine, type SheetOwnership } from "./statementBook";
import { diffSheet, toRangeWrites } from "./cellDiff";
import { GraphError, listRowFormulas, type GraphClient, type WorkbookSession } from "../sync/graph";

/** 이 기기에 사본을 두는 통합 문서. 검사리스트는 두 파일(평택 list, 군포·마전 listGm). */
export type BookKind = SyncFile | "listGm";

/** 리스트 파일 설정. 군포·마전은 예전 설정에 칸이 없으면(undefined) 드라이브 루트의 기본 파일을 쓴다. */
export function listFileOf(config: AppConfig, book: ListBookKind): { id: string; name: string } | null {
  if (book === "listGm") return config.files.listGm === undefined ? { ...LIST_GM_FILE } : config.files.listGm;
  return config.files.list;
}

/** 화면·상태에 보일 리스트 이름. 평택만 쓰면 예전처럼 '검사리스트'. */
export function listLabel(book: ListBookKind): string {
  return book === "list" ? SYNC_FILE_LABEL.list : LIST_BOOK_LABEL[book];
}

const work: Partial<Record<BookKind, ZipFiles>> = {};
const baseline: Partial<Record<BookKind, Uint8Array>> = {};

/**
 * 거래명세서에서 앱이 쓴 줄. 엑셀(비고 등)에는 표시를 남기지 않으므로 이 기기에 기억한다.
 * 키는 `통합 문서|시트 이름`, 값은 그 시트의 `날짜|관경` → 쓴 개수. 세션이 IndexedDB 에 저장·복원한다.
 */
export type StatementOwnership = Record<string, SheetOwnership>;

let ownership: StatementOwnership = {};

export function statementOwnership(): StatementOwnership {
  return ownership;
}

export function restoreStatementOwnership(saved: StatementOwnership | undefined | null): void {
  ownership = saved && typeof saved === "object" ? { ...saved } : {};
}

const LOCAL_BOOK = "local";

function ownershipKey(book: string, sheet: string): string {
  return `${book}|${sheet}`;
}

/**
 * 거래명세서에 그 달 시트가 없을 때 사용자에게 보이는 요청. 앱은 시트를 만들지 않는다.
 * 예: "11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)". 기록은 대기로 두고, 시트가 생기면 다음 동기화 때 채운다.
 */
export function missingSheetText(sheet: string, month: number): string {
  return `${month}월 거래명세서 시트를 만들어주세요 (${sheet})`;
}

/** 위 요청 문장을 다시 읽는다(여러 시트를 한 줄로 모을 때). */
export function parseMissingSheetText(message: string): { month: number; sheets: string[] } | null {
  const m = /^(\d{1,2})월 거래명세서 시트를 만들어주세요 \((.+)\)$/.exec(message);
  return m ? { month: Number(m[1]), sheets: m[2].split(", ") } : null;
}

/** 여러 시트 요청을 달별로 모은다. 예: "11월 거래명세서 시트를 만들어주세요 (A, B) · 12월 … (C)" */
export function missingSheetsText(items: { sheet: string; month: number }[]): string {
  const byMonth = new Map<number, string[]>();
  for (const { sheet, month } of items) {
    const list = byMonth.get(month) ?? [];
    if (!list.includes(sheet)) list.push(sheet);
    byMonth.set(month, list);
  }
  return [...byMonth.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([month, sheets]) => missingSheetText(sheets.join(", "), month)).join(" · ");
}

export function rememberBaseline(kind: BookKind, bytes: Uint8Array): void {
  baseline[kind] = bytes;
  work[kind] = unzipXlsx(bytes);
}

export function rememberWork(kind: BookKind, bytes: Uint8Array): void {
  work[kind] = unzipXlsx(bytes);
}

export function hasBook(kind: BookKind): boolean {
  return !!work[kind];
}

export function exportBook(kind: BookKind): Uint8Array {
  const files = work[kind];
  if (!files) throw new Error("통합 문서가 아직 없습니다.");
  return zipXlsx(files);
}

export function bootstrapFromSamples(bytes: Record<SyncFile, Uint8Array> & { listGm?: Uint8Array }): {
  records: InspectionRecord[];
  inspectors: string[];
  lengths: Record<string, number>;
} {
  rememberBaseline("list", bytes.list);
  rememberBaseline("ilji", bytes.ilji);
  rememberBaseline("statement", bytes.statement);
  if (bytes.listGm) rememberBaseline("listGm", bytes.listGm);
  const records = parseListWorkbook(work.list!).flatMap((s) => s.records);
  if (work.listGm) records.push(...parseListWorkbook(work.listGm, "listGm").flatMap((s) => s.records));
  const inspectors = readInspectors(work.ilji!);
  const lengths = readLengths(work.statement!);
  return { records, inspectors, lengths };
}

let staffCache: { files: ZipFiles; map: Record<string, { insp1: string; insp2: string }> } | null = null;

export function inspectorsOnDate(date: string): { insp1: string; insp2: string } | null {
  if (!work.ilji) return null;
  if (!staffCache || staffCache.files !== work.ilji) staffCache = { files: work.ilji, map: readDayStaff(work.ilji) };
  return staffCache.map[date] ?? null;
}

let workLogCache: { files: ZipFiles; map: Record<string, WorkLogDay> } | null = null;

/**
 * 기기에 있는 작업일지(동기화가 쓰고 받는 그 사본)에서 날짜별 근무자·근무시간. 네트워크를 부르지 않고 쓰지도 않는다.
 * 작업일지가 아직 없으면 빈 객체.
 */
export function workLogByDate(): Record<string, WorkLogDay> {
  if (!work.ilji) return {};
  if (!workLogCache || workLogCache.files !== work.ilji) {
    try {
      workLogCache = { files: work.ilji, map: readWorkLog(work.ilji) };
    } catch {
      workLogCache = { files: work.ilji, map: {} };
    }
  }
  return workLogCache.map;
}

export function sheetNameFor(records: InspectionRecord[], site: string): string | null {
  return records.find((r) => r.site === site)?.sheetName ?? null;
}

function toWrite(rec: InspectionRecord): ListWrite {
  return {
    id: rec.id,
    date: rec.date,
    line: rec.line,
    point: rec.point,
    result: rec.result,
    material: rec.material || undefined,
    clear: !!rec.deleted,
    excelRow: rec.src === "xlsx" ? rec.excelRow : undefined,
  };
}

/** 그날 작업일지에 쓸 검사자·근무시간. 저장 때 남긴 값 → 오늘이면 지금 값 → 모르면 비움/유지. */
export function staffFor(meta: DayMeta, date: string): Partial<DayStaff> {
  const saved = meta.dayStaff?.[date];
  if (saved) return saved;
  if (date === meta.workDate) {
    return { dayType: meta.dayType, insp1: meta.insp1, insp2: meta.insp2, timeStart: meta.timeStart, timeEnd: meta.timeEnd };
  }
  // 나중에 보내는 지난 날(백필). 검사자 1은 앱 사용자, 2는 그날 고른 값이 있으면 그 값. 근무시간은 모르므로 쓰지 않는다.
  const insp2 = meta.insp2ByDate?.[date];
  return { insp1: meta.insp1, insp2 };
}

/** 하루·현장 → 작업일지 열과 관경별 합격 개소. 같은 열을 쓰는 다른 현장도 같이 센다(그 외). */
export function iljiPlan(site: string, date: string, records: InspectionRecord[], config: AppConfig): { column: string; groups: IljiGroup[] } {
  const mapping = mappingFor(config.mappings, site);
  const column = mapping.iljiColumn || "그 외";
  // 지운 현장(택지 드림 등)도 같은 열의 기존 개소를 지우지 않도록 함께 센다.
  const known = [...config.mappings, ...REMOVED_SITE_MAPPINGS.filter((h) => !config.mappings.some((m) => m.site === h.site))];
  const sitesInColumn = known.filter((m) => (m.iljiColumn || "그 외") === column).map((m) => m.site);
  if (!sitesInColumn.includes(site)) sitesInColumn.push(site);
  return { column, groups: groupsFor(records, sitesInColumn, column, date) };
}

/** staff=false: 날짜를 옮겨 옛 날짜의 개수만 다시 쓸 때. 그날 검사자·시간은 건드리지 않는다. */
function applyDayIlji(files: ZipFiles, site: string, date: string, records: InspectionRecord[], config: AppConfig, meta: DayMeta, staff = true): void {
  const plan = iljiPlan(site, date, records, config);
  applyIljiDay(files, { date, column: plan.column, groups: plan.groups, ...(staff ? staffFor(meta, date) : {}) });
}

/** 같은 시트에 있는 리스트 기록(같은 파일·같은 시트). */
function sheetMates(all: InspectionRecord[], book: ListBookKind, sheet: string): InspectionRecord[] {
  return all.filter((r) => r.result !== "없음" && listBookOf(r) === book && listSheetOf(r) === sheet && !isRemovedSite(r.site));
}

/** 로컬 사본: 날짜를 바꾼 기록이 있는 라인의 앱 행을 그 자리들 안에서 다시 정렬한다. */
function reorderMovedLocally(pending: InspectionRecord[], all: InspectionRecord[]): void {
  const seen = new Set<string>();
  for (const rec of pending) {
    if (!rec.movedFrom?.length || rec.deleted || !rec.sheetName) continue;
    const book = listBookOf(rec);
    const target = listSheetOf(rec);
    const key = `${book}|${target}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const files = work[book];
    if (!files) continue;
    const sheet = workbookSheets(files).find((x) => x.name === rec.sheetName) ?? workbookSheets(files).find((x) => x.name.trim() === target);
    if (!sheet) continue;
    const moves = planLineReorder(appIdRows(files, sheet.name), sheetMates(all, book, target));
    for (const { rec: mate, row } of moves.values()) {
      writeListRow(files, sheet.name, { ...toWrite(mate), atRow: row }, LIST_LAYOUT[book]);
      mate.excelRow = row;
    }
  }
}

/** 그 달 거래명세서 줄(합격만, 길이를 아는 관경만). */
export function statementLines(site: string, ym: string, records: InspectionRecord[], config: AppConfig): PautLine[] {
  const mapping = mappingFor(config.mappings, site);
  if (mapping.unitPrice == null) return [];
  const monthRecs = records.filter((r) => !r.deleted && r.site === site && r.date.startsWith(ym));
  const bill = buildBill(monthRecs, mapping.unitPrice, config.weldLengths);
  return bill.lines
    .filter((l) => l.lengthM != null && l.amount != null)
    .map((l) => ({ date: l.date, diameter: l.diameter, lengthM: l.lengthM!, count: l.count, unitPrice: mapping.unitPrice! }));
}

/**
 * 그 달 시트에 쓴다. 시트가 없으면 만들지 않고(사용자가 직접 만든다) 쓸 줄이 있을 때만 없는 시트 이름을 돌려준다.
 */
function applyMonthStatement(files: ZipFiles, site: string, ym: string, records: InspectionRecord[], config: AppConfig): string | null {
  const mapping = mappingFor(config.mappings, site);
  const lines = statementLines(site, ym, records, config);
  const [year, month] = ym.split("-").map(Number);
  const names = workbookSheets(files).map((x) => x.name);
  const found = findMonthSheet(names, mapping.statementPrefix, year, month);
  if (!found) return lines.length ? expectedMonthSheetName(names, mapping.statementPrefix, year, month) : null;
  const key = ownershipKey(LOCAL_BOOK, found.name);
  const res = writePautLines(files, found.name, lines, ownership[key]);
  ownership[key] = res.owned;
  return null;
}

/** 로컬 통합 문서에 쓴다. 돌려주는 값: 거래명세서 시트가 없어 못 쓴 `현장|YYYY-MM` → 없는 시트 이름. */
export function applyRecordsLocally(pending: InspectionRecord[], all: InspectionRecord[], config: AppConfig, meta: DayMeta): Map<string, string> {
  if (!work.list || !work.ilji || !work.statement) throw new Error("로컬 통합 문서가 없습니다.");
  // 새 행은 날짜 → 라인 → 포인트 순으로 붙는다. 이미 있는 행(Y 앱 ID·excelRow)은 writeListRow 가 제자리를 고친다.
  // 기록마다 자기 리스트 파일에만 쓴다(군포·마전 ↔ 평택 섞이지 않음). 그 파일 사본이 없으면 건너뛴다.
  for (const rec of orderForAppend(pending)) {
    if (!rec.sheetName || rec.result === "없음") continue;
    const book = listBookOf(rec);
    const files = work[book];
    if (!files) continue;
    const target = listSheetOf(rec);
    const sheet = workbookSheets(files).find((x) => x.name === rec.sheetName) ?? workbookSheets(files).find((x) => x.name.trim() === target);
    if (!sheet) continue;
    const row = writeListRow(files, sheet.name, toWrite(rec), LIST_LAYOUT[book]);
    if (row > 0) rec.excelRow = row;
  }
  reorderMovedLocally(pending, all);
  const days = new Set<string>();
  const months = new Set<string>();
  const oldDays = new Set<string>();
  const oldMonths = new Set<string>();
  for (const rec of pending) {
    if (rec.result === "없음") continue;
    days.add(`${rec.site}|${rec.date}`);
    if (!statementSkipReason(config, rec.site)) months.add(`${rec.site}|${rec.date.slice(0, 7)}`);
    // 날짜를 바꾼 기록: 옛 날짜·옛 달도 다시 계산해 옛 자리의 개수를 뺀다.
    for (const from of rec.movedFrom ?? []) {
      oldDays.add(`${rec.site}|${from}`);
      if (!statementSkipReason(config, rec.site)) oldMonths.add(`${rec.site}|${from.slice(0, 7)}`);
    }
  }
  for (const key of days) {
    const [site, date] = key.split("|");
    applyDayIlji(work.ilji, site, date, all, config, meta);
  }
  for (const key of oldDays) {
    if (days.has(key)) continue;
    const [site, date] = key.split("|");
    applyDayIlji(work.ilji, site, date, all, config, meta, false);
  }
  const missing = new Map<string, string>();
  for (const key of months) {
    const [site, ym] = key.split("|");
    const sheet = applyMonthStatement(work.statement, site, ym, all, config);
    if (sheet) missing.set(key, sheet);
  }
  for (const key of oldMonths) {
    if (months.has(key)) continue;
    const [site, ym] = key.split("|");
    // 옛 달 시트가 없으면 지울 것도 없다(안내도 띄우지 않음).
    applyMonthStatement(work.statement, site, ym, all, config);
  }
  dropCalcChain(work.list);
  if (work.listGm) dropCalcChain(work.listGm);
  dropCalcChain(work.ilji);
  dropCalcChain(work.statement);
  return missing;
}

function groupsFor(records: InspectionRecord[], sites: string[], column: string, date: string): IljiGroup[] {
  const map = new Map<string, number>();
  for (const rec of passingRecords(records)) {
    if (rec.date !== date || !sites.includes(rec.site)) continue;
    const label = iljiDiameterLabel(rec.diameter || "미상", column, rec.site);
    map.set(label, (map.get(label) || 0) + 1);
  }
  return [...map.entries()].map(([diameter, count]) => ({ diameter, count }));
}

export function statementSkipReason(config: AppConfig, site: string): string | undefined {
  const mapping = mappingFor(config.mappings, site);
  if (!mapping.prefixConfirmed || !mapping.statementPrefix) return "거래명세서 접두어가 확정되지 않아 그 파일은 건너뛰었습니다.";
  if (mapping.unitPrice == null) return "단가가 없어 거래명세서 금액은 쓰지 않았습니다.";
  return undefined;
}

export interface SyncResult {
  status: SyncStatus;
  note?: string;
  error?: string;
  excelRow?: number;
  steps?: FileSteps;
}

/* ───────────────────────── Graph 동기화 (파일별 단계) ───────────────────────── */

export interface GraphSyncArgs {
  client: GraphClient;
  /** 보낼 기록(대기·실패). 파일별 steps 에서 done/skip 인 파일은 다시 보내지 않는다. */
  jobs: InspectionRecord[];
  /** 개소 집계에 쓰는 전체 기록 */
  all: InspectionRecord[];
  config: AppConfig;
  meta: DayMeta;
  /** 이미 보낸 앱 기록이 작업일지·명세서에 빠졌는지 확인해 채운다. */
  backfill?: boolean;
  /** 백필 때 검사리스트 파일도 받아 Y열 앱 ID 행을 후보에 넣는다. */
  scanList?: boolean;
  /** 시트가 Graph 에 보일 때까지 기다리는 함수(테스트에서 0). */
  sleep?: (ms: number) => Promise<void>;
  /** 거래명세서에서 앱이 쓴 줄 기억(생략하면 이 모듈의 기억 = 세션이 저장하는 것). */
  ownership?: StatementOwnership;
}

export interface GraphSyncOutcome {
  /** 이번에 바뀐 파일별 상태. 키는 기록 id. */
  steps: Map<string, FileSteps>;
  excelRows: Map<string, number>;
  report: SyncReport;
  /** 리스트에는 앱 ID 로 있는데 이 기기에 없던 기록 */
  discovered: InspectionRecord[];
}

function needs(rec: InspectionRecord, file: SyncFile): boolean {
  const st = rec.steps?.[file];
  return st?.status !== "done" && st?.status !== "skip";
}

export function errorText(error: unknown): string {
  if (error instanceof GraphError) return error.message;
  if (error instanceof Error && error.message) return error.message;
  return "알 수 없는 오류";
}

const BACKFILL_DAYS = 62;

class Run {
  steps = new Map<string, FileSteps>();
  excelRows = new Map<string, number>();
  report: SyncReport;
  constructor(at: number) {
    this.report = {
      at,
      files: { list: { status: "idle" }, ilji: { status: "idle" }, statement: { status: "idle" } },
    };
  }
  mark(recs: Iterable<InspectionRecord>, file: SyncFile, status: "done" | "error" | "skip" | "wait", message?: string): void {
    for (const rec of recs) {
      const prev = this.steps.get(rec.id) ?? {};
      prev[file] = { status, message, at: this.report.at };
      this.steps.set(rec.id, prev);
      rec.steps = { ...(rec.steps ?? {}), [file]: prev[file] };
    }
  }
  fail(file: SyncFile, recs: Iterable<InspectionRecord>, error: unknown): void {
    const message = errorText(error);
    this.mark(recs, file, "error", message);
    this.report.files[file] = { status: "error", message: `${SYNC_FILE_LABEL[file]}: ${message}` };
  }
  ok(file: SyncFile, count: number): void {
    if (this.report.files[file].status === "error") return;
    this.report.files[file] = { status: "ok", count };
  }
  /** 리스트 파일 하나의 실패. 기록의 steps.list 와 그 파일 줄, 합친 files.list 에 남긴다. */
  listFail(book: ListBookKind, recs: Iterable<InspectionRecord>, error: unknown): void {
    const message = errorText(error);
    this.mark(recs, "list", "error", message);
    this.report.lists = { ...(this.report.lists ?? {}), [book]: { status: "error", message: `${listLabel(book)}: ${message}` } };
    this.report.files.list = { status: "error", message: `${listLabel(book)}: ${message}` };
  }
  listOk(book: ListBookKind, count: number): void {
    const lists = (this.report.lists ??= {});
    if (lists[book]?.status !== "error") lists[book] = { status: "ok", count };
    const all = this.report.files.list;
    if (all.status === "error") return;
    this.report.files.list = { status: "ok", count: (all.status === "ok" ? all.count ?? 0 : 0) + count };
  }
}

/**
 * 검사리스트 → 작업일지 → 거래명세서를 따로따로 보낸다.
 * - 단계마다 토큰을 새로 받는다(중간 만료 방지). 401·잠김·429·5xx 는 GraphClient 가 다시 보낸다.
 * - 리스트가 성공하고 작업일지가 실패하면, 그 기록은 작업일지만 실패로 남아 다음에 작업일지만 다시 보낸다.
 * - 셋 다 Graph 통합 문서 세션으로 바뀐 칸만 쓴다. 파일 전체 업로드는 하지 않는다.
 * - 거래명세서 새 달 시트는 만들지 않는다. 없으면 그 현장·달만 '대기'로 두고 다음 동기화 때 채운다.
 */
export async function syncGraph(args: GraphSyncArgs): Promise<GraphSyncOutcome> {
  const run = new Run(Date.now());
  const discovered: InspectionRecord[] = [];
  // 앱에서 지운 현장은 보내지 않는다(엑셀은 그대로).
  const jobs = args.jobs.filter((r) => r.result !== "없음" && !isRemovedSite(r.site));
  let all = args.all;

  // 백필: 이미 보냈다고 표시된 앱 기록도 검사리스트 Y열(앱 ID)로 확인한다. 없으면 리스트 단계에서 처음부터 다시 보낸다.
  const listChecked = await listStep(args, run, jobs, args.backfill ? sentAppRecords(all, jobs) : []);


  if (args.backfill && args.scanList) {
    for (const book of LIST_BOOKS) {
      const file = listFileOf(args.config, book);
      if (!file?.id) continue;
      // 군포·마전은 그 파일의 기록이 이 기기에 있을 때만 받는다(아직 안 쓰면 요청하지 않음).
      if (book !== "list" && !all.some((r) => listBookOf(r) === book)) continue;
      try {
        await args.client.refreshToken();
        const files = unzipXlsx(await args.client.download(file.id));
        const known = new Set([...all, ...discovered].map((r) => r.id));
        for (const site of parseListWorkbook(files, book)) {
          for (const rec of site.records) {
            if (rec.src !== "user" || known.has(rec.id) || isRemovedSite(rec.site)) continue;
            rec.steps = { list: { status: "done", at: run.report.at } };
            discovered.push(rec);
          }
        }
      } catch (error) {
        console.warn(`${listLabel(book)} 백필 확인 실패`, error);
      }
    }
    all = [...all, ...discovered];
  }

  const candidates = args.backfill ? backfillCandidates(all, jobs) : [];
  // 후보는 이미 리스트에 있는 행(동기화됨 또는 리스트에서 읽은 Y 앱 ID). 리스트 단계는 끝난 것으로 둔다.
  // Y열 앱 ID 로 확인한 것만 리스트 완료로 표시한다(확인 못 했으면 표시하지 않음).
  if (listChecked) run.mark(candidates.filter((r) => !r.steps?.list && listChecked.present.has(r.id)), "list", "done");
  await iljiStep(args, run, jobs, candidates, all);
  await statementStep(args, run, jobs, candidates, all);
  return { steps: run.steps, excelRows: run.excelRows, report: run.report, discovered };
}

/**
 * 백필 확인 대상: 이 기기에서 이미 보낸(저장된) 최근 앱 기록.
 * 군포·마전 리스트는 기간 제한 없이 그 파일로 갈 앱 기록을 전부 확인한다(새로 연결된 파일이라
 * 예전에 평택 리스트로 보내려다 실패·건너뛴 기록, 리스트 단계가 끝난 것으로 표시된 대기 기록도 포함).
 */
function sentAppRecords(all: InspectionRecord[], jobs: InspectionRecord[]): InspectionRecord[] {
  const jobIds = new Set(jobs.map((r) => r.id));
  const app = (r: InspectionRecord) => r.src === "user" && r.sync !== "local" && !r.deleted && (r.result === PASS || r.result === FAIL) && !isRemovedSite(r.site);
  const gm = all.filter((r) => app(r) && listBookOf(r) === "listGm" && (!jobIds.has(r.id) || !needs(r, "list")));
  const recs = all.filter((r) => app(r) && listBookOf(r) !== "listGm" && !jobIds.has(r.id));
  if (!recs.length) return gm;
  const latest = recs.reduce((max, r) => (r.date > max ? r.date : max), "");
  const [y, m, d] = latest.split("-").map(Number);
  const from = new Date(Date.UTC(y, m - 1, d - BACKFILL_DAYS)).toISOString().slice(0, 10);
  return [...recs.filter((r) => r.date >= from), ...gm];
}

/** 이미 리스트에 들어간 앱 기록(Y=앱 ID) 중 최근 것. 작업일지·명세서 반영이 확인되지 않은 것만. */
function backfillCandidates(all: InspectionRecord[], jobs: InspectionRecord[]): InspectionRecord[] {
  const jobIds = new Set(jobs.map((r) => r.id));
  // 저장 전(local)·보내는 중 기록은 빼고, 이미 동기화된 앱 기록만.
  const appRecs = all.filter((r) => r.src === "user" && r.sync === "synced" && !r.deleted && (r.result === PASS || r.result === FAIL) && !jobIds.has(r.id) && !isRemovedSite(r.site));
  if (!appRecs.length) return [];
  const latest = appRecs.reduce((max, r) => (r.date > max ? r.date : max), "");
  const [y, m, d] = latest.split("-").map(Number);
  const from = new Date(Date.UTC(y, m - 1, d - BACKFILL_DAYS)).toISOString().slice(0, 10);
  return appRecs.filter((r) => r.date >= from && (needs(r, "ilji") || needs(r, "statement")));
}

interface ListCheck {
  /** Y열에서 확인한 앱 ID */
  present: Set<string>;
  /** 보낸 것으로 표시됐지만 Y열에 없어 이번에 다시 쓴 기록 */
  missing: InspectionRecord[];
}

/**
 * 검사리스트 단계. 기록마다 자기 리스트 파일(평택 / 군포·마전)로 나눠 파일마다 따로 보낸다.
 * 파일마다 토큰을 새로 받고 통합 문서 세션을 따로 연다. 한 파일이 실패해도 다른 파일은 그대로 보낸다.
 * verify(이미 보낸 앱 기록)가 있으면 같은 세션에서 Y열 앱 ID 를 확인해,
 * 없는 기록은 steps 를 지우고 이번 묶음에 넣어 다시 쓴다(작업일지·명세서도 다시 확인).
 * 확인하지 못했으면(파일 없음·시트 없음·오류) 그 기록은 present 에 들어가지 않는다.
 */
async function listStep(args: GraphSyncArgs, run: Run, jobs: InspectionRecord[], verify: InspectionRecord[] = []): Promise<ListCheck | null> {
  const recs = jobs.filter((r) => needs(r, "list"));
  if (!recs.length && !verify.length) return null;
  let check: ListCheck | null = null;
  for (const book of LIST_BOOKS) {
    const bookRecs = recs.filter((r) => listBookOf(r) === book);
    const bookVerify = verify.filter((r) => listBookOf(r) === book);
    if (!bookRecs.length && !bookVerify.length) continue;
    const res = await listBookStep(args, run, book, jobs, bookRecs, bookVerify);
    if (!res) continue;
    check ??= { present: new Set(), missing: [] };
    for (const id of res.present) check.present.add(id);
    check.missing.push(...res.missing);
  }
  return check;
}

async function listBookStep(
  args: GraphSyncArgs,
  run: Run,
  book: ListBookKind,
  jobs: InspectionRecord[],
  recs: InspectionRecord[],
  verify: InspectionRecord[],
): Promise<ListCheck | null> {
  const itemId = listFileOf(args.config, book)?.id;
  if (!itemId) {
    if (recs.length) run.listFail(book, recs, new Error(`설정에서 ${listLabel(book)} 파일을 고르세요.`));
    return null;
  }
  const layout = LIST_LAYOUT[book];
  const check: ListCheck = { present: new Set(), missing: [] };
  const remaining = new Set(recs);
  let wb: WorkbookSession | null = null;
  let written = 0;
  try {
    // 파일마다 새 토큰(중간 만료 방지).
    await args.client.refreshToken();
    wb = await openListWorkbook(args.client, itemId, listFileOf(args.config, book)?.name);
    const bySheet = new Map<string, InspectionRecord[]>();
    const verifyBySheet = new Map<string, InspectionRecord[]>();
    for (const rec of recs) {
      const key = listSheetOf(rec);
      const list = bySheet.get(key) ?? [];
      list.push(rec);
      bySheet.set(key, list);
    }
    for (const rec of verify) {
      const key = listSheetOf(rec);
      verifyBySheet.set(key, [...(verifyBySheet.get(key) ?? []), rec]);
      if (!bySheet.has(key)) bySheet.set(key, []);
    }
    for (const [sheetName, group] of bySheet) {
      const ws = await wb.sheet(sheetName);
      if (!ws) {
        if (group.length) run.listFail(book, group, new Error(`${listLabel(book)}에 '${sheetName}' 시트가 없습니다.`));
        for (const rec of group) remaining.delete(rec);
        continue;
      }
      const LAST = 6000;
      const colF = await wb.readValues(ws.id, `F6:F${LAST}`);
      const colY = await wb.readValues(ws.id, `Y6:Y${LAST}`);
      // 군포·마전: Y(앱 ID)가 없는 기존 행도 날짜+라인+포인트가 같으면 같은 검사로 보고 다시 쓰지 않는다.
      const colE = book === "listGm" ? await wb.readValues(ws.id, `E6:E${LAST}`) : [];
      const colG = book === "listGm" ? await wb.readValues(ws.id, `G6:G${LAST}`) : [];
      // 군포·마전: 순번(A)이 비어 있는 새 행에 다음 번호를 쓴다(미리 적힌 번호는 그대로).
      const colA = layout.fillSeq ? await wb.readValues(ws.id, `A6:A${LAST}`) : [];
      const seqAt = (row: number): number | null => {
        const v = colA[row - 6]?.[0];
        if (typeof v === "number") return v;
        if (typeof v === "string" && /^\d+$/.test(v.trim())) return Number(v.trim());
        return null;
      };
      const nextSeq = (row: number): number | undefined => {
        if (!layout.fillSeq || seqAt(row) !== null) return undefined;
        let max = 0;
        for (let r = 6; r < row; r++) max = Math.max(max, seqAt(r) ?? 0);
        return max + 1;
      };
      let last = 5;
      colF.forEach((row, i) => {
        const v = row?.[0];
        if (v !== null && v !== undefined && String(v).trim() !== "") last = 6 + i;
      });
      const byId = new Map<string, number>();
      colY.forEach((row, i) => {
        const v = row?.[0];
        if (typeof v === "string" && v.trim()) byId.set(v.trim(), 6 + i);
      });
      for (const id of byId.keys()) check.present.add(id);
      const plain = new Map<string, number[]>();
      colF.forEach((row, i) => {
        const y = colY[i]?.[0];
        if (y !== null && y !== undefined && String(y).trim() !== "") return;
        const key = rowKey(colE[i]?.[0], row?.[0], colG[i]?.[0]);
        if (key) plain.set(key, [...(plain.get(key) ?? []), 6 + i]);
      });
      const claimPlain = (rec: InspectionRecord): number | undefined => {
        if (book !== "listGm" || rec.src !== "user" || rec.deleted) return undefined;
        const rows = plain.get(rowKey(isoToExcelSerial(rec.date), rec.line, rec.point) ?? "");
        return rows?.shift();
      };
      // 보낸 것으로 표시됐는데 Y열에 앱 ID 가 없는 기록(예: 11/11 S28) → 새 행으로 다시 쓴다.
      for (const rec of verifyBySheet.get(sheetName) ?? []) {
        if (byId.has(rec.id)) continue;
        // 군포·마전: 리스트 단계만 다시 한다(작업일지·명세서에 이미 반영된 것은 그대로 — 두 번 세지 않음).
        if (book === "listGm") rec.steps = rec.steps ? { ...rec.steps, list: undefined } : undefined;
        else rec.steps = undefined;
        rec.excelRow = undefined;
        check.missing.push(rec);
        if (!group.includes(rec)) group.push(rec);
        remaining.add(rec);
        if (!jobs.includes(rec)) jobs.push(rec);
      }
      // 이미 엑셀에 있는 행(Y 앱 ID 또는 앱 ID 없는 기존 행)은 제자리를 고치고 옮기지 않는다.
      // 새 행만 날짜 → 라인 → 포인트(S1·S2·…·S10 → R…, 재검사는 원본 바로 뒤) 순으로 마지막 행 아래에 붙인다.
      const known = (rec: InspectionRecord) => byId.get(rec.id) ?? (rec.src === "xlsx" && rec.excelRow ? rec.excelRow : undefined);
      const fresh: InspectionRecord[] = [];
      for (const rec of orderForAppend(group.filter((rec) => !known(rec)))) {
        const row = claimPlain(rec);
        if (row === undefined) {
          fresh.push(rec);
          continue;
        }
        // 이미 손으로 적힌 같은 검사 행: 그 행은 건드리지 않고 완료로 둔다.
        check.present.add(rec.id);
        run.excelRows.set(rec.id, row);
        rec.excelRow = row;
        run.mark([rec], "list", "done");
        remaining.delete(rec);
      }
      // 날짜를 바꾼 기록이 있으면 같은 라인의 앱 행(Y)끼리 자리만 바꿔 날짜 → 라인 → 포인트 순으로 둔다.
      // 행을 끼우거나 지우지 않는다. 이번 묶음이 아닌 같은 라인 앱 행도 자리가 바뀌면 함께 다시 쓴다.
      const moves = planLineReorder(byId, sheetMates(args.all, book, sheetName));
      const mates = [...moves.values()].map((m) => m.rec).filter((r) => !group.includes(r));
      const ordered = [...group.filter((rec) => known(rec)), ...fresh];
      const writeAt = async (rec: InspectionRecord, target: number, reorder: boolean) => {
        const input = toWrite(rec);
        const options =
          layout.diameter === "value"
            ? { diameterValue: input.clear ? undefined : listDiameterValue(rec.line), seq: input.clear ? undefined : nextSeq(target) }
            : {};
        const row = listRowFormulas(target, input, options);
        // 자리를 바꿔 쓰는 행: 재질(K)이 없으면 그 자리에 있던 다른 기록의 재질을 지운다.
        if (reorder && row[10] === null) row[10] = "";
        await wb!.writeFormulas(ws.id, `A${target}:Y${target}`, [row]);
        if (typeof row[0] === "number" && colA.length) {
          while (colA.length <= target - 6) colA.push([null]);
          colA[target - 6] = [row[0]];
        }
        run.excelRows.set(rec.id, target);
        rec.excelRow = target;
        mirrorListLocally(rec, target);
      };
      for (const rec of mates) await writeAt(rec, moves.get(rec.id)!.row, true);
      for (const rec of mates) byId.set(rec.id, moves.get(rec.id)!.row);
      for (const rec of ordered) {
        // 같은 앱 ID(Y)가 이미 있으면 그 행을 고친다. 다시 보내도 행이 늘지 않는다.
        let target = moves.get(rec.id)?.row ?? known(rec);
        if (rec.deleted && !target) {
          run.mark([rec], "list", "done");
          remaining.delete(rec);
          continue;
        }
        if (!target) {
          last += 1;
          target = last;
        }
        await writeAt(rec, target, moves.has(rec.id));
        byId.set(rec.id, target);
        run.mark([rec], "list", "done");
        remaining.delete(rec);
        written += 1;
      }
    }
    run.listOk(book, written);
  } catch (error) {
    run.listFail(book, remaining, error);
    if (!recs.length && !check.missing.length) return null;
  } finally {
    await wb?.close();
  }
  return check;
}

/** 리스트 행의 같은 검사 키(날짜 일련번호|라인|포인트). 셋 중 하나라도 비면 null. */
function rowKey(date: unknown, line: unknown, point: unknown): string | null {
  let serial: number | null = null;
  if (typeof date === "number") serial = Math.floor(date);
  else if (typeof date === "string" && date.trim()) {
    const t = date.trim();
    if (/^\d+(\.\d+)?$/.test(t)) serial = Math.floor(Number(t));
    else {
      const iso = parseFlexibleDate(t);
      if (iso) serial = isoToExcelSerial(iso);
    }
  }
  const l = String(line ?? "").trim();
  const p = String(point ?? "").trim().toUpperCase();
  if (serial === null || !l || !p) return null;
  return `${serial}|${l}|${p}`;
}

/** 리스트 통합 문서 세션. 저장된 항목 ID 가 404 면 드라이브 루트에서 같은 이름의 파일을 찾아 연다. */
async function openListWorkbook(client: GraphClient, itemId: string, name?: string): Promise<WorkbookSession> {
  try {
    return await client.openWorkbook(itemId);
  } catch (error) {
    if (!(error instanceof GraphError) || error.status !== 404 || !name) throw error;
    const found = await client.findRootFile(name);
    if (!found) throw error;
    return client.openWorkbook(found.id);
  }
}

/** 앱 화면의 엑셀 미리보기용 사본에도 같은 행을 쓴다. 실패해도 동기화와는 상관없다. */
function mirrorListLocally(rec: InspectionRecord, row: number): void {
  const book = listBookOf(rec);
  const files = work[book];
  if (!files) return;
  try {
    const sheet = workbookSheets(files).find((x) => x.name.trim() === listSheetOf(rec));
    // OneDrive 에 쓴 행 번호 그대로(미리보기 사본의 마지막 행이 달라도 같은 자리).
    if (sheet) writeListRow(files, sheet.name, { ...toWrite(rec), excelRow: row, atRow: row }, LIST_LAYOUT[book]);
  } catch {
    /* 미리보기 사본 */
  }
}

function cloneFiles(files: ZipFiles): ZipFiles {
  return { ...files };
}

async function iljiStep(args: GraphSyncArgs, run: Run, jobs: InspectionRecord[], candidates: InspectionRecord[], all: InspectionRecord[]): Promise<void> {
  const { config, meta, client } = args;
  const days = new Map<string, InspectionRecord[]>();
  /** 날짜를 바꿔 개수만 다시 쓸 옛 날짜(검사자·시간은 그대로). */
  const oldDays = new Set<string>();
  for (const rec of jobs) {
    if (!needs(rec, "ilji")) continue;
    const key = `${rec.site}|${rec.date}`;
    days.set(key, [...(days.get(key) ?? []), rec]);
    for (const from of rec.movedFrom ?? []) {
      const old = `${rec.site}|${from}`;
      if (old !== key) oldDays.add(old);
      days.set(old, [...(days.get(old) ?? []), rec]);
    }
  }
  for (const key of days.keys()) if (oldDays.has(key) && days.get(key)!.some((r) => `${r.site}|${r.date}` === key)) oldDays.delete(key);
  const verify = new Map<string, InspectionRecord[]>();
  for (const rec of candidates) {
    if (!needs(rec, "ilji")) continue;
    const key = `${rec.site}|${rec.date}`;
    if (days.has(key)) days.get(key)!.push(rec);
    else verify.set(key, [...(verify.get(key) ?? []), rec]);
  }
  if (!days.size && !verify.size) return;
  const allRecs = [...days.values(), ...verify.values()].flat();
  const itemId = config.files.ilji?.id;
  if (!itemId) {
    run.fail("ilji", allRecs, new Error("설정에서 작업일지 파일을 고르세요."));
    return;
  }
  try {
    await client.refreshToken();
    const before = unzipXlsx(await client.download(itemId));
    const sheet = findIljiSheet(before);
    // 백필: 작업일지에 그날 그 열이 비어 있을 때만 채운다. 이미 뭔가 적혀 있으면(사람이 적었을 수 있음) 건드리지 않는다.
    for (const [key, recs] of verify) {
      const [site, date] = key.split("|");
      const plan = iljiPlan(site, date, all, config);
      const actual = readIljiDay(before, date, plan.column) ?? [];
      if (!plan.groups.length || actual.length) run.mark(recs, "ilji", "done");
      else days.set(key, recs);
    }
    if (!days.size) {
      run.ok("ilji", 0);
      return;
    }
    const after = cloneFiles(before);
    const okDays: InspectionRecord[][] = [];
    const failed = new Set<InspectionRecord>();
    for (const [key, recs] of days) {
      const [site, date] = key.split("|");
      try {
        applyDayIlji(after, site, date, all, config, meta, !oldDays.has(key));
        okDays.push(recs);
      } catch (error) {
        run.fail("ilji", recs, error);
        for (const rec of recs) failed.add(rec);
      }
    }
    const writes = toRangeWrites(diffSheet(before, after, sheet));
    if (writes.length) {
      await client.refreshToken();
      const wb = await client.openWorkbook(itemId);
      try {
        const ws = await wb.sheet(sheet);
        if (!ws) throw new Error(`작업일지 시트(${sheet.trim()})를 OneDrive에서 찾지 못했습니다.`);
        for (const w of writes) await wb.writeFormulas(ws.id, w.address, w.formulas);
      } finally {
        await wb.close();
      }
    }
    for (const recs of okDays) run.mark(recs.filter((r) => !failed.has(r)), "ilji", "done");
    run.ok("ilji", okDays.length);
    work.ilji = after;
  } catch (error) {
    run.fail("ilji", allRecs.filter((r) => r.steps?.ilji?.status !== "done"), error);
  }
}

interface MonthGroup {
  site: string;
  ym: string;
  prefix: string;
  recs: InspectionRecord[];
  lines: PautLine[];
}

async function statementStep(args: GraphSyncArgs, run: Run, jobs: InspectionRecord[], candidates: InspectionRecord[], all: InspectionRecord[]): Promise<void> {
  const { config, client } = args;
  const owned = args.ownership ?? ownership;
  const groups = new Map<string, MonthGroup>();
  const verify = new Map<string, MonthGroup>();
  const skipped: InspectionRecord[] = [];
  let skipNote = "";
  const add = (target: Map<string, MonthGroup>, rec: InspectionRecord) => {
    const reason = statementSkipReason(config, rec.site);
    if (reason) {
      run.mark([rec], "statement", "skip", reason);
      skipped.push(rec);
      skipNote = reason;
      return;
    }
    const ym = rec.date.slice(0, 7);
    const key = `${rec.site}|${ym}`;
    const existing = groups.get(key) ?? target.get(key);
    if (existing) existing.recs.push(rec);
    else target.set(key, { site: rec.site, ym, prefix: mappingFor(config.mappings, rec.site).statementPrefix, recs: [rec], lines: [] });
  };
  for (const rec of jobs) if (needs(rec, "statement")) add(groups, rec);
  // 날짜를 바꿔 달이 바뀐 기록: 옛 달 시트도 다시 계산해 그 줄을 지운다(시트가 없으면 지울 것도 없음, 안내 없음).
  const cleanup = new Set<string>();
  for (const rec of jobs) {
    if (!needs(rec, "statement") || statementSkipReason(config, rec.site)) continue;
    for (const from of rec.movedFrom ?? []) {
      const ym = from.slice(0, 7);
      const key = `${rec.site}|${ym}`;
      if (groups.has(key)) continue;
      cleanup.add(key);
      groups.set(key, { site: rec.site, ym, prefix: mappingFor(config.mappings, rec.site).statementPrefix, recs: [], lines: [] });
    }
  }
  for (const rec of candidates) if (needs(rec, "statement")) add(verify, rec);
  if (!groups.size && !verify.size) {
    if (skipped.length) run.report.files.statement = { status: "skip", message: skipNote };
    return;
  }
  const allRecs = [...groups.values(), ...verify.values()].flatMap((g) => g.recs);
  const itemId = config.files.statement?.id;
  if (!itemId) {
    run.fail("statement", allRecs, new Error("설정에서 거래명세서 파일을 고르세요."));
    return;
  }
  try {
    for (const g of [...groups.values(), ...verify.values()]) g.lines = statementLines(g.site, g.ym, all, config);
    await client.refreshToken();
    const before = unzipXlsx(await client.download(itemId));

    // 백필: 그 달 시트에 이 날짜·관경 줄이 없을 때만 채운다.
    for (const [key, g] of verify) {
      const [year, month] = g.ym.split("-").map(Number);
      const names = workbookSheets(before).map((x) => x.name);
      const found = findMonthSheet(names, g.prefix, year, month);
      const rows = found ? readPautRows(before, found.name) : [];
      const missing = g.lines.some((l) => !rows.some((r) => r.date === l.date && r.diameter === l.diameter.toUpperCase()));
      if (missing) groups.set(key, g);
      else run.mark(g.recs, "statement", "done");
    }
    const todo = [...groups.values()];
    if (!todo.length) {
      run.ok("statement", 0);
      return;
    }

    // 새 달 시트는 만들지 않는다(사용자가 직접 만든다). 시트가 없으면 그 현장·달은 건너뛰고 기록은 대기로 둔다.
    // 실패가 아니므로 다른 현장·다른 파일(리스트·작업일지)에는 영향이 없고, 시트가 생기면 다음 동기화 때 채운다.
    const after = cloneFiles(before);
    const names = workbookSheets(after).map((x) => x.name);
    const plans: { sheet: string; inserted: number; insertAt: number; group: MonthGroup; owned: SheetOwnership }[] = [];
    const missingSheets: { sheet: string; month: number }[] = [];
    for (const g of todo) {
      const [year, month] = g.ym.split("-").map(Number);
      const found = findMonthSheet(names, g.prefix, year, month);
      if (!found) {
        if (cleanup.has(`${g.site}|${g.ym}`) && !g.recs.length) continue;
        if (!g.lines.length) {
          // 쓸 줄이 없고 시트도 없으면 할 일이 없다.
          run.mark(g.recs, "statement", "done");
          continue;
        }
        const sheet = expectedMonthSheetName(names, g.prefix, year, month);
        missingSheets.push({ sheet, month });
        run.mark(g.recs, "statement", "wait", missingSheetText(sheet, month));
        continue;
      }
      const res = writePautLines(after, found.name, g.lines, owned[ownershipKey(itemId, found.name)]);
      plans.push({ sheet: found.name, inserted: res.inserted, insertAt: res.insertAt, group: g, owned: res.owned });
    }

    if (plans.length) {
      await client.refreshToken();
      const wb = await client.openWorkbook(itemId);
      try {
        for (const plan of plans) {
          const ws = await waitForSheet(wb, plan.sheet, args.sleep);
          let baseXml: string | undefined;
          if (plan.inserted > 0) {
            await wb.insertRows(ws.id, plan.insertAt, plan.inserted);
            baseXml = shiftRowNumbers(readSheet(before, plan.sheet).xml, plan.insertAt, plan.inserted);
          }
          const writes = toRangeWrites(diffSheet(before, after, plan.sheet, plan.sheet, baseXml));
          for (const w of writes) await wb.writeFormulas(ws.id, w.address, w.formulas);
          owned[ownershipKey(itemId, plan.sheet)] = plan.owned;
          run.mark(plan.group.recs, "statement", "done");
        }
      } finally {
        await wb.close();
      }
    }
    run.ok("statement", plans.length);
    if (missingSheets.length && run.report.files.statement.status !== "error") {
      run.report.files.statement = { ...run.report.files.statement, message: missingSheetsText(missingSheets) };
    }
    work.statement = after;
  } catch (error) {
    run.fail("statement", allRecs.filter((r) => r.steps?.statement?.status !== "done"), error);
  }
}

/** 시트가 Excel 서비스(세션 목록)에 보일 때까지 잠깐 기다린다. */
async function waitForSheet(wb: WorkbookSession, name: string, sleep?: (ms: number) => Promise<void>) {
  for (let i = 0; i < 4; i++) {
    const ws = await (i === 0 ? wb.sheet(name) : wb.worksheets(true).then(() => wb.sheet(name)));
    if (ws) return ws;
    await (sleep ? sleep(1500 * (i + 1)) : new Promise((r) => setTimeout(r, 1500 * (i + 1))));
  }
  throw new GraphError(404, `거래명세서에 '${name}' 시트가 OneDrive에 아직 보이지 않습니다. 잠시 뒤 다시 시도하세요.`);
}

/**
 * 리스트 파일을 받는다. 저장된 항목 ID 로 못 찾으면(404: 이동·다시 올림) 드라이브 루트에서 정확히 같은 이름을 찾는다.
 * 돌려주는 file 의 id 가 바뀌었으면 설정에 저장해야 한다.
 */
export async function downloadListBook(
  client: GraphClient,
  file: { id: string; name: string },
): Promise<{ bytes: Uint8Array; file: { id: string; name: string } }> {
  await client.refreshToken();
  try {
    return { bytes: await client.download(file.id), file };
  } catch (error) {
    if (!(error instanceof GraphError) || error.status !== 404 || !file.name) throw error;
    const found = await client.findRootFile(file.name);
    if (!found) throw error;
    return { bytes: await client.download(found.id), file: { id: found.id, name: found.name } };
  }
}

/** 데모(로그인 없음): 이 기기 안의 통합 문서에만 쓴다. */
export function syncPendingDemo(pending: InspectionRecord[], all: InspectionRecord[], config: AppConfig, meta: DayMeta): Map<string, SyncResult> {
  const missing = applyRecordsLocally(pending, all, config, meta);
  const results = new Map<string, SyncResult>();
  const at = Date.now();
  for (const rec of pending) {
    const skip = statementSkipReason(config, rec.site);
    const sheet = missing.get(`${rec.site}|${rec.date.slice(0, 7)}`);
    const statement: FileSteps["statement"] = skip
      ? { status: "skip", message: skip, at }
      : sheet && rec.result !== "없음"
        ? { status: "wait", message: missingSheetText(sheet, Number(rec.date.slice(5, 7))), at }
        : { status: "done", at };
    results.set(rec.id, {
      // 명세서 시트가 없으면 그 기록은 대기로 남아 시트가 생긴 뒤 다음 동기화 때 채운다.
      status: statement.status === "wait" ? "pending" : "synced",
      note: skip,
      excelRow: rec.excelRow,
      steps: {
        list: { status: "done", at },
        ilji: { status: "done", at },
        statement,
      },
    });
  }
  return results;
}
