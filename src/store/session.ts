import { useSyncExternalStore } from "react";
import { latestLine } from "../domain/history";
import { guessIljiColumn, defaultMappings, defaultWeldLengths, isRemovedSite, mappingFor } from "../domain/mapping";
import { canonicalPoint, fmtPoint, mirrorPoint, parsePoint, planPrefixCopy } from "../domain/point";
import { normalizeResult } from "../excel/liveList";
import { diameterFromLine } from "../domain/size";
import { todayISO } from "../domain/dates";
import type { AppConfig, DayMeta, DraftCard, FileSteps, InspectionRecord, ListBookKind, ResultKind, SiteMapping, SyncFile, SyncReport } from "../domain/types";
import { FAIL, PASS, SYNC_FILES, SYNC_FILE_LABEL } from "../domain/types";
import {
  bootstrapFromSamples,
  downloadListBook,
  errorText,
  exportBook,
  hasBook,
  inspectorsOnDate,
  listFileOf,
  listLabel,
  rememberBaseline,
  rememberWork,
  restoreStatementOwnership,
  sheetNameFor,
  statementOwnership,
  statementSkipReason,
  syncGraph,
  syncPendingDemo,
  type BookKind,
  type GraphSyncOutcome,
  type StatementOwnership,
} from "../excel/engine";
import { backendEnabled } from "../sync/backend";
import { accessToken, completeRedirect, currentAccount, LOGIN_NEEDED_TEXT, loginRedirect, logoutRedirect, resolveClientId } from "../sync/msal";
import { downloadItem, GRAPH_LOGIN_TEXT, GRAPH_TIMEOUT_TEXT, GraphClient } from "../sync/graph";
import { getAuthState, isLoginError, isLoginFailureText as authLoginText, onAuthChange, type AuthState } from "../sync/authState";
import { ABORTED_TEXT, isTimeoutLike, TOKEN_TIMEOUT_TEXT, withTimeout } from "../sync/timeout";
import { listBookOf, parseListWorkbook, rerouteGmRecords } from "../excel/listBook";
import { unzipXlsx } from "../excel/ooxml";
import { kvGet, kvSet } from "./db";
import { aggregateSteps } from "./syncStatus";

export { aggregateSteps };

export interface Toast {
  id: number;
  text: string;
  undoId?: string;
}

export interface SessionState {
  ready: boolean;
  error: string | null;
  records: InspectionRecord[];
  config: AppConfig;
  meta: DayMeta;
  account: string | null;
  /**
   * OneDrive 모드에서 조용한 토큰 갱신이 안 되어 사용자가 한 번 로그인해야 한다.
   * 화면 위 '로그인 필요' 배너가 뜨고, 대기 기록은 실패로 바꾸지 않고 남겨 두었다가 로그인 직후 보낸다.
   */
  loginNeeded: boolean;
  /** 전역 인증 상태(토큰으로 확인한 것). 계정 이름이 보여도 "ok" 가 아니면 로그인된 것이 아니다. */
  auth: AuthState;
  online: boolean;
  syncing: boolean;
  /** 마지막 동기화의 파일별 결과. 앱을 다시 켜도 남는다. */
  syncReport: SyncReport | null;
  toast: Toast | null;
}

const emptyConfig = (): AppConfig => ({
  clientId: "",
  syncMode: "demo",
  files: { list: null, ilji: null, statement: null },
  mappings: defaultMappings(),
  weldLengths: defaultWeldLengths(),
  inspectors: ["조우현", "심명보"],
});

function emptyMeta(): DayMeta {
  return {
    workDate: todayISO(),
    dayType: "평일",
    insp1: "조우현",
    insp2: "심명보",
    timeStart: "09:00",
    timeEnd: "18:00",
    lastSite: "",
    activeLine: "",
    prefix: "S",
  };
}

let state: SessionState = {
  ready: false,
  error: null,
  records: [],
  config: emptyConfig(),
  meta: emptyMeta(),
  account: null,
  loginNeeded: false,
  auth: getAuthState(),
  online: typeof navigator === "undefined" ? true : navigator.onLine,
  syncing: false,
  syncReport: null,
  toast: null,
};

const listeners = new Set<() => void>();

function emit(): void {
  for (const fn of listeners) fn();
}

function patch(partial: Partial<SessionState>): void {
  state = { ...state, ...partial };
  if (partial.records || partial.meta) trackOriginals();
  emit();
}

/* ───────────── 초기화(저장 안 한 카드 표시 되돌리기) ───────────── */

const ORIGINALS_KEY = "paut-unsaved-originals-v1";

interface Originals {
  /** 기록 id → 처음 미저장(local)이 되기 직전 모습. rec=null 이면 저장 뒤 새로 생긴 기록. */
  recs: Record<string, { site: string; date: string; rec: InspectionRecord | null }>;
  /** 현장별 마지막 저장 때의 회색 카드(키 `현장||라인||접두어`). 현장 키가 없으면 아직 기준 없음. */
  drafts: Record<string, Record<string, DraftCard[]>>;
}

/** 지금 현장·작업일의 저장된(미저장 아님) 기록 사본. 다음 편집 직전 모습을 알기 위해 patch 때마다 갱신한다. */
const cleanCopy = new Map<string, InspectionRecord>();

function readOriginals(): Originals {
  try {
    const raw = localStorage.getItem(ORIGINALS_KEY);
    const parsed = raw ? (JSON.parse(raw) as Partial<Originals>) : {};
    return { recs: parsed.recs ?? {}, drafts: parsed.drafts ?? {} };
  } catch {
    return { recs: {}, drafts: {} };
  }
}

function writeOriginals(o: Originals): void {
  try { localStorage.setItem(ORIGINALS_KEY, JSON.stringify(o)); } catch { /* 저장 불가 */ }
}

function siteDrafts(draftCards: DayMeta["draftCards"], site: string): Record<string, DraftCard[]> {
  const out: Record<string, DraftCard[]> = {};
  for (const [key, list] of Object.entries(draftCards ?? {})) if (key.startsWith(`${site}||`) && list.length) out[key] = list;
  return out;
}

function sameDrafts(a: Record<string, DraftCard[]>, b: Record<string, DraftCard[]>): boolean {
  const norm = (x: Record<string, DraftCard[]>) => JSON.stringify(Object.keys(x).sort().map((k) => [k, x[k]]));
  return norm(a) === norm(b);
}

/**
 * 편집으로 기록이 처음 미저장이 될 때 그 직전 모습(cleanCopy)을 기억한다. 회색 카드는 현장별 기준이 없으면 지금 모습을 기준으로.
 * 저장(commitDay)·초기화 때 그 현장·날짜 기억을 지운다.
 */
function trackOriginals(): void {
  const { lastSite, workDate } = state.meta ?? ({} as DayMeta);
  if (!lastSite || !state.records) return;
  const o = readOriginals();
  let changed = false;
  for (const rec of state.records) {
    if (rec.site !== lastSite || rec.date !== workDate) continue;
    if (rec.sync === "local") {
      if (!(rec.id in o.recs)) {
        const prev = cleanCopy.get(rec.id);
        // 저장 때 모습을 모르는데 이미 엑셀에 들어간 기록(초기화 배포 전부터 미저장으로 남았던 R02 등)은
        // '새 카드'가 아니다 → 기억하지 않는다(초기화가 지우지 않음).
        if (!prev && wasSent(rec)) continue;
        o.recs[rec.id] = { site: rec.site, date: rec.date, rec: prev ? structuredClone(prev) : null };
        changed = true;
      }
    } else {
      cleanCopy.set(rec.id, structuredClone(rec));
    }
  }
  if (!(lastSite in o.drafts)) {
    o.drafts[lastSite] = structuredClone(siteDrafts(state.meta.draftCards, lastSite));
    changed = true;
  }
  if (changed) writeOriginals(o);
}

/** 한 번이라도 엑셀(또는 엑셀에서 읽은 행)에 있는 기록. 초기화가 '새 카드'로 보고 지우면 안 된다. */
function wasSent(rec: InspectionRecord): boolean {
  return !!rec.excelRow || !!rec.steps || rec.src !== "user" || !!rec.movedFrom?.length;
}

/** 지금 현장·작업일에 되돌릴 수 있는 미저장 카드 표시 수(초기화 버튼을 켤지). */
export function resettableCount(records: InspectionRecord[], meta: DayMeta): number {
  const { lastSite, workDate } = meta;
  const o = readOriginals();
  const recs = Object.entries(o.recs).filter(([, v]) => v.site === lastSite && v.date === workDate);
  const ids = new Set(records.filter((r) => r.site === lastSite && r.date === workDate && r.sync === "local").map((r) => r.id));
  const n = recs.filter(([id, v]) => ids.has(id) || v.rec !== null).length;
  const base = o.drafts[lastSite];
  const draftsChanged = base !== undefined && !sameDrafts(base, siteDrafts(meta.draftCards, lastSite));
  return n + (draftsChanged ? 1 : 0);
}

function forgetOriginals(site: string, date: string, drafts: DayMeta["draftCards"]): void {
  const o = readOriginals();
  for (const [id, v] of Object.entries(o.recs)) if (v.site === site && v.date === date) delete o.recs[id];
  o.drafts[site] = structuredClone(siteDrafts(drafts, site));
  writeOriginals(o);
}

/**
 * 초기화: 지금 현장·작업일에서 마지막 저장 뒤 바꾼 카드 표시(합격/불합격·새 카드·회색 카드)를 저장 때 모습으로 되돌린다.
 * 저장하지 않고 엑셀에도 보내지 않는다.
 */
export function resetUnsaved(): Promise<void> {
  return enqueue(resetUnsavedNow);
}

async function resetUnsavedNow(): Promise<void> {
  const { lastSite, workDate } = state.meta;
  const o = readOriginals();
  let records = [...state.records];
  let n = 0;
  for (const [id, v] of Object.entries(o.recs)) {
    if (v.site !== lastSite || v.date !== workDate) continue;
    const i = records.findIndex((r) => r.id === id);
    if (v.rec === null) {
      if (i >= 0 && records[i].sync === "local" && !wasSent(records[i])) {
        records = records.filter((r) => r.id !== id);
        n += 1;
      }
    } else {
      const back = structuredClone(v.rec);
      // 보내는 중에 고친 기록이었다면 다시 보낼 대기로 둔다(같은 내용이라 엑셀은 그대로).
      if (back.sync === "syncing" || back.sync === "local") back.sync = "pending";
      if (i >= 0) records[i] = back;
      else records.push(back);
      n += 1;
    }
  }
  const base = o.drafts[lastSite];
  let draftCards = state.meta.draftCards ?? {};
  if (base !== undefined && !sameDrafts(base, siteDrafts(draftCards, lastSite))) {
    draftCards = Object.fromEntries(Object.entries(draftCards).filter(([key]) => !key.startsWith(`${lastSite}||`)));
    Object.assign(draftCards, structuredClone(base));
    n += 1;
  }
  state.records = records;
  state.meta = { ...state.meta, draftCards };
  forgetOriginals(lastSite, workDate, draftCards);
  resetUnsavedDirectR(lastSite, workDate);
  patch({ records: [...state.records], meta: state.meta });
  await persistRecords();
  await persistMeta();
  mirrorUnsaved();
  toast(n ? "저장 안 한 표시를 되돌렸습니다" : "되돌릴 표시가 없습니다");
}

export function getState(): SessionState {
  return state;
}

export function subscribe(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

export function useSession(): SessionState {
  return useSyncExternalStore(subscribe, getState, getState);
}

function toast(text: string, undoId?: string): void {
  patch({ toast: { id: Date.now(), text, undoId } });
}

export function dismissToast(): void {
  if (state.toast) patch({ toast: null });
}

const SAMPLE: Record<SyncFile, string> = {
  list: `${import.meta.env.BASE_URL}samples/ndt-list.xlsx`,
  ilji: `${import.meta.env.BASE_URL}samples/work-log.xlsx`,
  statement: `${import.meta.env.BASE_URL}samples/statement.xlsx`,
};

async function fetchBytes(url: string): Promise<Uint8Array> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${url} 를 열지 못했습니다 (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

function sitesOf(records: InspectionRecord[]): string[] {
  const names: string[] = [];
  for (const rec of records) {
    if (!names.includes(rec.site) && !isRemovedSite(rec.site)) names.push(rec.site);
  }
  for (const preset of defaultMappings()) {
    if (!names.includes(preset.site)) names.push(preset.site);
  }
  return names;
}

function mergeConfig(saved: AppConfig | undefined, inspectors: string[], lengths: Record<string, number>, sites: string[]): AppConfig {
  const base = saved ?? emptyConfig();
  const presets = defaultMappings();
  const mappings: SiteMapping[] = [];
  const seen = new Set<string>();
  for (const site of sites) {
    seen.add(site);
    const prev = base.mappings.find((m) => m.site === site);
    const preset = presets.find((m) => m.site === site);
    mappings.push(
      prev ??
        preset ?? {
          site,
          iljiColumn: guessIljiColumn(site),
          statementPrefix: "",
          prefixConfirmed: false,
          unitPrice: null,
          priceConfirmed: false,
          priceNote: "새 현장입니다. 작업일지 열, 거래명세서 접두어, 단가를 지정하세요.",
        },
    );
  }
  for (const extra of base.mappings) {
    if (!seen.has(extra.site) && !isRemovedSite(extra.site)) mappings.push(extra);
  }
  return {
    ...base,
    mappings,
    inspectors: inspectors.length ? inspectors : base.inspectors,
    weldLengths: { ...defaultWeldLengths(), ...lengths, ...base.weldLengths },
  };
}

async function loadBooks(): Promise<void> {
  const stored: Partial<Record<BookKind, Uint8Array>> = {};
  for (const kind of ["list", "ilji", "statement", "listGm"] as const) {
    stored[kind] = await kvGet<Uint8Array>(`book:${kind}`);
  }
  // 군포·마전 리스트 사본은 OneDrive 에서 처음 받은 뒤에만 있다(샘플 없음).
  if (stored.listGm) rememberWork("listGm", stored.listGm);
  if (stored.list && stored.ilji && stored.statement) {
    rememberWork("list", stored.list);
    rememberWork("ilji", stored.ilji);
    rememberWork("statement", stored.statement);
    return;
  }
  const bytes = {
    list: await fetchBytes(SAMPLE.list),
    ilji: await fetchBytes(SAMPLE.ilji),
    statement: await fetchBytes(SAMPLE.statement),
  };
  rememberBaseline("list", bytes.list);
  rememberBaseline("ilji", bytes.ilji);
  rememberBaseline("statement", bytes.statement);
  for (const kind of ["list", "ilji", "statement"] as const) {
    await kvSet(`book:${kind}`, bytes[kind]);
  }
}

let bookTimer = 0;
function scheduleBookSave(): void {
  window.clearTimeout(bookTimer);
  bookTimer = window.setTimeout(() => {
    void persistBooks();
  }, 600);
}

async function persistBooks(): Promise<void> {
  for (const kind of ["list", "ilji", "statement", "listGm"] as const) {
    if (!hasBook(kind)) continue;
    await kvSet(`book:${kind}`, exportBook(kind));
  }
}

/**
 * 리스트 파일 하나에서 읽은 기록으로 그 파일의 기록만 바꾼다. 다른 리스트 파일의 기록은 그대로 둔다.
 * requeue: 그 파일에 Y 앱 ID 가 없는 앱 기록을 다시 보낸다(파일을 새로 고른 경우).
 */
function replaceListRecords(book: ListBookKind, parsed: InspectionRecord[], requeue: boolean): InspectionRecord[] {
  const others = state.records.filter((r) => listBookOf(r) !== book);
  const mine = state.records.filter((r) => listBookOf(r) === book);
  const parsedIds = new Set(parsed.map((r) => r.id));
  const user = mine.filter((r) => r.src === "user" && !parsedIds.has(r.id));
  if (requeue) {
    for (const rec of user) {
      if (rec.sync === "local") continue;
      rec.sync = "pending";
      rec.steps = undefined;
    }
  }
  return book === "list" ? [...parsed, ...user, ...others] : [...others, ...parsed, ...user];
}

/**
 * 군포·마전 리스트를 아직 받지 않았으면 OneDrive 에서 받아 현장·기록을 넣는다(동기화 때 한 번).
 * 저장된 항목 ID 로 못 찾으면 드라이브 루트에서 'PAUT NDT 리스트(군포, 마전).xlsx' 를 이름으로 찾는다.
 * 실패해도 다른 파일 동기화는 계속한다. 돌려주는 값: 실패 이유.
 */
async function ensureListGm(client: GraphClient): Promise<string | null> {
  const file = listFileOf(state.config, "listGm");
  if (!file?.id) return null;
  if (hasBook("listGm") && state.records.some((r) => r.src === "xlsx" && listBookOf(r) === "listGm")) return null;
  try {
    const got = await downloadListBook(client, file);
    rememberBaseline("listGm", got.bytes);
    await kvSet("book:listGm", got.bytes);
    const known = new Set(state.records.map((r) => r.id));
    const parsed = parseListWorkbook(unzipXlsx(got.bytes), "listGm")
      .flatMap((x) => x.records)
      .filter((r) => !known.has(r.id));
    state.records = replaceListRecords("listGm", [...parsed, ...state.records.filter((r) => listBookOf(r) === "listGm" && r.src === "xlsx")], false);
    let config = mergeConfig(state.config, state.config.inspectors, {}, sitesOf(state.records));
    if (got.file.id !== file.id || state.config.files.listGm === undefined) config = { ...config, files: { ...config.files, listGm: got.file } };
    state.config = config;
    patch({ records: [...state.records], config });
    await persistRecords();
    await persistConfig();
    return null;
  } catch (error) {
    console.warn("군포·마전 리스트를 받지 못했습니다", error);
    return `${listLabel("listGm")}: 파일을 받지 못했습니다 — ${errorText(error)}`;
  }
}

async function persistRecords(): Promise<void> {
  await kvSet("records", state.records);
}

async function persistMeta(): Promise<void> {
  await kvSet("meta", state.meta);
}

async function persistConfig(): Promise<void> {
  await kvSet("config", state.config);
}

let syncing = false;
let syncAgain: SyncOptions | null = null;
let currentRun: Promise<void> = Promise.resolve();
/** 동기화 묶음 번호. 시간 초과로 놓아 준 묶음이 늦게 끝나도 새 묶음 결과를 덮지 않게 한다. */
let runGen = 0;
/** 다음 동기화의 첫 토큰은 캐시 없이 새로 받는다(로그인 직후·로그인 끊김·401 뒤). */
let freshTokenNext = false;

/** 동기화 한 묶음의 시간 제한(ms). 넘으면 Graph 요청을 끊고, 그래도 안 끝나면 grace 뒤 '보내는 중'을 풀어 준다. */
export const syncLimits = {
  runMs: 5 * 60_000,
  graceMs: 5_000,
  loginCheckMs: 70_000,
};

interface RunCtl {
  gen: number;
  client: GraphClient | null;
  /** 시간 초과로 끊었다 */
  cancelled: boolean;
  /** 시간 초과로 놓아 줬다(결과는 버린다) */
  released: boolean;
}

/** 로그인이 끊겨(또는 새 토큰으로도 401) 못 보낸 단계 문구. 실패가 아니라 '로그인 대기'. */
export function isLoginFailureText(message?: string): boolean {
  return !!message && (authLoginText(message) || message.includes(LOGIN_NEEDED_TEXT) || message.includes(GRAPH_LOGIN_TEXT));
}

/** 지금 돌고 있는 동기화 묶음(만료 때 바로 끊는다). */
let activeCtl: RunCtl | null = null;

/**
 * 전역 인증 상태가 바뀌면 화면을 바로 맞춘다.
 * - expired/signedOut: '로그인 필요'를 켜고 '보내는 중'을 즉시 끈다. 진행 중 묶음의 Graph 요청은 끊는다
 *   (끝난 파일은 그대로, 남은 파일은 대기). 동기화가 끝나기를 기다리지 않는다.
 * - ok(화면 복귀 확인 등으로 다시 살아남): '로그인 필요'를 끄고 대기 기록을 새 토큰으로 보낸다.
 */
function handleAuthChange(next: AuthState): void {
  const graph = state.config.syncMode === "graph";
  if (next === "expired" || next === "signedOut") {
    freshTokenNext = true;
    activeCtl?.client?.abort();
    patch({ auth: next, loginNeeded: graph ? true : state.loginNeeded, syncing: false });
    return;
  }
  if (next === "ok" && graph && state.loginNeeded && state.ready) {
    freshTokenNext = true;
    patch({ auth: next, loginNeeded: false });
    if (jobs().length) requestSync();
    return;
  }
  patch({ auth: next });
}
onAuthChange(handleAuthChange);

/** 응답이 없어(시간 초과·중단) 못 보낸 단계 문구. 실패가 아니라 대기. */
export function isStallFailureText(message?: string): boolean {
  return !!message && (message.includes(GRAPH_TIMEOUT_TEXT) || message.includes(TOKEN_TIMEOUT_TEXT) || message.includes(ABORTED_TEXT));
}

export interface SyncOptions {
  /** 이미 보낸 앱 기록이 작업일지·거래명세서에 빠졌는지 확인해 채운다. 앱 시작·다시 시도 때. */
  backfill?: boolean;
  /** 백필 때 검사리스트 파일도 받아 Y열 앱 ID 행을 찾는다(다시 시도 버튼). */
  scanList?: boolean;
}

function mergeOptions(a: SyncOptions | null, b: SyncOptions): SyncOptions {
  return { backfill: !!(a?.backfill || b.backfill), scanList: !!(a?.scanList || b.scanList) };
}

export function requestSync(options: SyncOptions = {}): void {
  if (state.config.syncMode === "graph" && !state.online) {
    // 오프라인: 조용히 넘기지 않는다. 기록은 대기로 남고, 연결되면 bindOnline 이 다시 부른다.
    const waiting = jobs().length;
    const report: SyncReport = {
      ...(state.syncReport ?? { files: emptyFiles() }),
      at: Date.now(),
      offline: true,
    };
    patch({ syncReport: report });
    void kvSet("syncReport", report);
    if (waiting) toast(`오프라인 · ${waiting}건 대기. 연결되면 자동으로 엑셀에 보냅니다`);
    return;
  }
  if (state.config.syncMode === "graph" && getAuthState() === "expired") {
    // 토큰이 죽었다: 보내기를 시작하지 않는다('보내는 중' 없이 바로 '로그인 필요'). 기록은 대기로 남는다.
    freshTokenNext = true;
    if (!state.loginNeeded || state.syncing) patch({ loginNeeded: true, syncing: false });
    return;
  }
  if (syncing) {
    syncAgain = mergeOptions(syncAgain, options);
    return;
  }
  syncing = true;
  patch({ syncing: true });
  const ctl: RunCtl = { gen: ++runGen, client: null, cancelled: false, released: false };
  activeCtl = ctl;
  let abortTimer: ReturnType<typeof setTimeout> | undefined;
  let releaseTimer: ReturnType<typeof setTimeout> | undefined;
  // 시간 초과: 먼저 Graph 요청·토큰 기다림을 끊어 정상적으로 끝나게 하고(끝난 파일은 그대로),
  // 그래도 안 끝나면(멈춘 토큰 요청 등) grace 뒤 놓아 준다. 어느 쪽이든 '보내는 중'은 반드시 풀린다.
  const watchdog = new Promise<"released">((resolve) => {
    abortTimer = setTimeout(() => {
      ctl.cancelled = true;
      ctl.client?.abort();
      releaseTimer = setTimeout(() => resolve("released"), syncLimits.graceMs);
    }, syncLimits.runMs);
  });
  const run = runSync(options, ctl).catch((error) => console.warn("동기화 오류", error));
  currentRun = Promise.race([run.then(() => "done" as const), watchdog]).then((how) => {
    clearTimeout(abortTimer);
    clearTimeout(releaseTimer);
    try {
      if (how === "released") releaseStuckRun(ctl);
      if (ctl.cancelled) toast(ABORTED_TEXT);
    } catch (error) {
      console.warn("동기화 정리 오류", error);
    } finally {
      // 끝남·시간 초과·중단 어느 경우든 '동기화 중'은 반드시 푼다(정리 중 오류가 나도).
      if (ctl.gen === runGen) {
        syncing = false;
        activeCtl = null;
        if (state.syncing) patch({ syncing: false });
      }
    }
    if (ctl.gen !== runGen) return;
    if (syncAgain) {
      const next = syncAgain;
      syncAgain = null;
      requestSync(next);
    }
  });
}

/** 끝나지 않는 묶음을 놓아 준다: '보내는 중' 기록을 대기로 되돌리고 저장한다. 늦게 끝난 결과는 버린다. */
function releaseStuckRun(ctl: RunCtl): void {
  ctl.released = true;
  ctl.client?.abort();
  let n = 0;
  for (const rec of state.records) {
    if (rec.sync !== "syncing") continue;
    rec.sync = "pending";
    rec.steps = deferSoftSteps(rec.steps);
    n += 1;
  }
  if (n) {
    patch({ records: [...state.records] });
    void persistRecords();
  }
}

/** 진행 중인 동기화가 끝날 때까지 기다린다(테스트용). */
export async function syncIdle(): Promise<void> {
  for (let i = 0; i < 10 && syncing; i++) await currentRun;
}

function emptyFiles(): SyncReport["files"] {
  return { list: { status: "idle" }, ilji: { status: "idle" }, statement: { status: "idle" } };
}

/**
 * 다시 시도: 실패한 파일만 다시 보낸다(성공한 파일은 그대로).
 * 리스트 파일도 받아 이미 들어간 앱 행 중 작업일지·명세서에 빠진 것을 채운다.
 */
export function retrySync(): void {
  for (const rec of state.records) {
    if (rec.sync !== "error") continue;
    rec.sync = "pending";
    rec.syncError = undefined;
    rec.steps = dropErrorSteps(rec.steps);
  }
  patch({ records: [...state.records] });
  void persistRecords();
  if (state.config.syncMode === "graph" && !state.online) {
    requestSync();
    return;
  }
  requestSync({ backfill: true, scanList: true });
}

function dropErrorSteps(steps: FileSteps | undefined): FileSteps | undefined {
  if (!steps) return undefined;
  const next: FileSteps = {};
  for (const file of SYNC_FILES) {
    const st = steps[file];
    if (st && st.status !== "error") next[file] = st;
  }
  return next;
}

const UNSAVED_KEY = "paut-unsaved-v1";
/** 거래명세서에서 앱이 쓴 줄(엑셀에는 표시를 남기지 않으므로 이 기기에 기억). */
const OWNERSHIP_KEY = "statementOwned";

/** 앱에서 지운 현장(지성 브레인, 택지 드림)의 기록은 숨기기만 하고 어디에도 보내지 않는다(삭제도 보내지 않음). */
function isJob(r: InspectionRecord): boolean {
  if (isRemovedSite(r.site)) return false;
  return r.sync !== "local" && r.result !== "없음" && (r.sync === "pending" || r.sync === "error" || (!!r.deleted && r.sync !== "synced"));
}

function jobs(): InspectionRecord[] {
  return state.records.filter(isJob);
}

/** 엑셀로 실제 보낼 대기·실패 기록(화면 위 상태 줄). 지운 현장·미저장·없음은 넣지 않는다. */
export function syncQueue(records: InspectionRecord[]): InspectionRecord[] {
  return records.filter(isJob);
}

/** 엑셀로 보낼 대기·실패 기록 수(로그인 배너 안내용). */
export function pendingCount(records: InspectionRecord[]): number {
  return records.filter(isJob).length;
}

function mirrorUnsaved(): void {
  try {
    const records = state.records.filter((rec) => rec.sync === "local");
    const prev = readUnsaved();
    localStorage.setItem(UNSAVED_KEY, JSON.stringify({
      records,
      draftCards: state.meta.draftCards ?? {},
      stamps: prev.stamps,
    }));
  } catch { /* 사생활 보호 모드 */ }
}

function readUnsaved(): { records: InspectionRecord[]; draftCards: DayMeta["draftCards"]; stamps: Record<string, string> } {
  try {
    const raw = localStorage.getItem(UNSAVED_KEY);
    if (!raw) return { records: [], draftCards: {}, stamps: {} };
    const parsed = JSON.parse(raw) as { records?: InspectionRecord[]; draftCards?: DayMeta["draftCards"]; stamps?: Record<string, string> };
    return { records: parsed.records ?? [], draftCards: parsed.draftCards ?? {}, stamps: parsed.stamps ?? {} };
  } catch {
    return { records: [], draftCards: {}, stamps: {} };
  }
}

function restoreUnsaved(records: InspectionRecord[], meta: DayMeta): void {
  const saved = readUnsaved();
  const byId = new Map(records.map((rec) => [rec.id, rec]));
  for (const rec of saved.records) {
    if (rec.sync !== "local") continue;
    const prev = byId.get(rec.id);
    if (!prev) records.push(rec);
    else Object.assign(prev, rec);
  }
  meta.draftCards = { ...(meta.draftCards ?? {}), ...(saved.draftCards ?? {}) };
}

function stampOf(meta: DayMeta): string {
  return [meta.insp1, meta.insp2, meta.dayType, meta.timeStart, meta.timeEnd].join("|");
}

export function unsavedCount(): number {
  const { lastSite, workDate } = state.meta;
  const locals = state.records.filter((rec) => rec.sync === "local" && rec.site === lastSite && rec.date === workDate).length;
  const grey = Object.entries(state.meta.draftCards ?? {})
    .filter(([key]) => key.startsWith(`${lastSite}||`))
    .reduce((sum, [, list]) => sum + list.length, 0);
  const saved = readUnsaved().stamps[`${lastSite}|${workDate}`];
  const metaDirty = saved !== undefined && saved !== stampOf(state.meta) ? 1 : 0;
  return locals + grey + metaDirty;
}

/**
 * 저장 버튼. 카드 누르기·칠하기와 같은 편집 큐에 넣는다.
 * 마지막 카드(예: S28)를 누르자마자 저장해도, 그 카드 기록이 들어간 뒤에 저장한다.
 */
export function commitDay(): Promise<void> {
  return enqueue(commitDayNow);
}

/**
 * 지난 날(작업일 이전)에 '미저장'으로 남은 기록. 화면은 오늘만 보여 저장 버튼으로는 다시 못 보내므로
 * 시작할 때 대기열에 넣는다(내용이 바뀐 기록이라 세 파일 모두 다시).
 */
function promoteStaleLocals(records: InspectionRecord[], workDate: string): number {
  let n = 0;
  for (const rec of records) {
    if (rec.sync !== "local" || rec.date >= workDate) continue;
    if (rec.result === "없음") {
      rec.sync = "synced";
      continue;
    }
    if (rec.deleted && !rec.excelRow) continue;
    rec.sync = "pending";
    rec.steps = undefined;
    rec.syncError = undefined;
    n += 1;
  }
  return n;
}

async function commitDayNow(): Promise<void> {
  const { lastSite, workDate } = state.meta;
  const day = state.records.filter((rec) => rec.site === lastSite && rec.date === workDate);
  const stampKey = `${lastSite}|${workDate}`;
  const saved = readUnsaved();
  const metaDirty = saved.stamps[stampKey] !== undefined && saved.stamps[stampKey] !== stampOf(state.meta);
  const locals = day.filter((rec) => rec.sync === "local");
  if (!locals.length && !metaDirty && !day.some((rec) => !rec.deleted && rec.result !== "없음")) {
    toast("저장할 내용이 없습니다");
    return;
  }
  for (const rec of day) {
    if (rec.result === "없음") {
      rec.sync = "synced";
      continue;
    }
    if (rec.deleted && rec.sync === "local" && !rec.excelRow) continue;
    if (rec.sync === "local") {
      // 내용이 바뀌었으니 세 파일 모두 다시 쓴다(리스트는 Y 앱 ID 로 같은 행).
      rec.sync = "pending";
      rec.steps = undefined;
      rec.syncError = undefined;
    } else if (metaDirty) {
      // 검사자·근무시간만 바뀜 → 작업일지만 다시.
      rec.sync = "pending";
      if (rec.steps) rec.steps = { ...rec.steps, ilji: undefined };
    }
  }
  // 그날 검사자·근무시간을 날짜별로 남긴다. 나중에 다시 보내도 오늘 값이 아니라 그날 값이 들어간다.
  state.meta = {
    ...state.meta,
    dayStaff: {
      ...(state.meta.dayStaff ?? {}),
      [workDate]: {
        dayType: state.meta.dayType,
        insp1: state.meta.insp1,
        insp2: state.meta.insp2,
        timeStart: state.meta.timeStart,
        timeEnd: state.meta.timeEnd,
      },
    },
  };
  void persistMeta();
  state.records = state.records.filter((rec) => !(rec.deleted && rec.sync === "local" && !rec.excelRow));
  const places = state.records.filter((rec) => !rec.deleted && rec.site === lastSite && rec.date === workDate && rec.result === PASS).length;
  patch({ records: [...state.records] });
  await persistRecords();
  saved.stamps[stampKey] = stampOf(state.meta);
  forgetOriginals(lastSite, workDate, state.meta.draftCards);
  commitDirectR(lastSite, workDate);
  try {
    localStorage.setItem(UNSAVED_KEY, JSON.stringify({ records: state.records.filter((rec) => rec.sync === "local"), draftCards: state.meta.draftCards ?? {}, stamps: saved.stamps }));
  } catch { /* 저장 불가 */ }
  toast(`저장됨 · ${places}개소`);
  requestSync();
}

export { inspectorsOnDate };

async function runSync(options: SyncOptions = {}, ctl: RunCtl = { gen: runGen, client: null, cancelled: false, released: false }): Promise<void> {
  const pending = jobs();
  const graph = state.config.syncMode === "graph";
  if (state.syncReport?.offline) patch({ syncReport: { ...state.syncReport, offline: false } });
  if (!pending.length && !(graph && options.backfill && hasAnyFile())) return;
  const clientId = graph ? resolveClientId(state.config.clientId) : "";
  if (graph && !(await loginReady(clientId))) return;
  if (ctl.cancelled) return;
  for (const rec of pending) rec.sync = "syncing";
  patch({ records: [...state.records] });
  try {
    const alive = state.records.filter((r) => !r.deleted || pending.includes(r));
    if (graph) {
      let loginLost = false;
      const fresh = freshTokenNext;
      freshTokenNext = false;
      const client = new GraphClient({
        // 동기화 중에는 화면을 넘기지 않는다(interactive:false). 도중에 로그인이 끊기면 그 기록은 대기로 되돌린다.
        getToken: (force) =>
          accessToken(clientId, { forceRefresh: force, interactive: false }).catch((error) => {
            if (isLoginError(error)) loginLost = true;
            throw error;
          }),
        // 로그인 직후 첫 동기화는 캐시된 옛 토큰을 쓰지 않는다.
        freshToken: fresh,
      });
      ctl.client = client;
      const outcome = await syncGraph({
        client,
        jobs: pending,
        all: alive,
        config: state.config,
        meta: state.meta,
        backfill: options.backfill,
        scanList: options.scanList,
      });
      // 새 토큰으로도 401 → 로그인이 실제로 끊긴 것. 실패가 아니라 '로그인 대기'.
      if (client.authFailed) loginLost = true;
      if (ctl.released) return; // 시간 초과로 이미 놓아 준 묶음: 결과를 버린다(다음 묶음이 Y 앱 ID 로 다시 확인).
      deferSoftFailures(outcome);
      applyOutcome(pending, outcome);
      if (loginLost) {
        freshTokenNext = true;
        patch({ loginNeeded: true });
      }
      // 군포·마전 리스트를 아직 안 받았으면 대기 기록을 다 보낸 뒤 받아 현장·기록을 채운다(다른 파일을 막지 않음).
      const listGmError = loginLost || client.stalled || client.isAborted ? null : await ensureListGm(client);
      if (listGmError && state.syncReport && !state.syncReport.lists?.listGm) {
        const report: SyncReport = { ...state.syncReport, lists: { ...(state.syncReport.lists ?? {}), listGm: { status: "error", message: listGmError } } };
        state.syncReport = report;
        patch({ syncReport: report });
      }
    } else {
      const results = syncPendingDemo(pending, alive, state.config, state.meta);
      for (const rec of pending) {
        const res = results.get(rec.id);
        if (!res) continue;
        if (res.excelRow && res.excelRow > 0) rec.excelRow = res.excelRow;
        // 전송 중에 다시 고쳐진 기록(예: S를 눌러 따라 바뀐 R)은 "local" 로 남겨 다음 저장 때 다시 보낸다.
        if (rec.sync !== "syncing") continue;
        rec.sync = res.status;
        if (rec.sync === "synced") delete rec.movedFrom;
        rec.syncNote = res.note;
        rec.syncError = res.error;
        rec.steps = res.steps;
      }
    }
    state.records = state.records.filter((r) => !(r.deleted && r.sync === "synced"));
    patch({ records: [...state.records] });
    await persistRecords();
    if (state.syncReport) await kvSet("syncReport", state.syncReport);
    await kvSet(OWNERSHIP_KEY, statementOwnership());
    scheduleBookSave();
  } catch (error) {
    if (ctl.released) return;
    const message = error instanceof Error ? error.message : "동기화에 실패했습니다.";
    // 로그인 끊김·시간 초과는 실패가 아니라 대기(다시 시도·로그인 뒤 그대로 보낸다).
    const soft = isLoginError(error) || isTimeoutLike(error) || isLoginFailureText(message) || isStallFailureText(message);
    for (const rec of pending) {
      if (rec.sync === "syncing") {
        rec.sync = soft ? "pending" : "error";
        rec.syncError = soft ? undefined : message;
      }
    }
    patch({ records: [...state.records] });
    await persistRecords();
    toast(message);
  }
}

/**
 * 동기화 전에 조용히 토큰을 받아 본다. 로그인이 필요하면 기록을 실패로 바꾸지 않고(대기 유지)
 * '로그인 필요' 배너만 띄운다. 로그인하고 돌아오면 doInit 이 대기열을 바로 보낸다.
 * 네트워크 등 다른 오류는 그대로 진행해 파일별 이유를 보여 준다.
 */
async function loginReady(clientId: string): Promise<boolean> {
  try {
    await withTimeout(accessToken(clientId, { interactive: false }), syncLimits.loginCheckMs, TOKEN_TIMEOUT_TEXT);
  } catch (error) {
    if (isLoginError(error)) {
      freshTokenNext = true;
      if (!state.loginNeeded) {
        const waiting = jobs().length;
        patch({ loginNeeded: true });
        if (waiting) toast(`Microsoft 로그인 필요 · ${waiting}건 대기 중`);
      }
      return false;
    }
    if (isTimeoutLike(error)) {
      // 토큰 응답이 없다: 기다리지 않고 대기로 둔다('보내는 중'이 붙잡히지 않게).
      const waiting = jobs().length;
      toast(waiting ? `${TOKEN_TIMEOUT_TEXT} (${waiting}건)` : TOKEN_TIMEOUT_TEXT);
      return false;
    }
    return true;
  }
  const account = currentAccount(clientId);
  if (state.loginNeeded || (account && account !== state.account)) patch({ loginNeeded: false, account: account ?? state.account });
  return true;
}

/** 로그인 끊김·시간 초과로 실패한 단계 표시를 지운다(그 파일은 대기 → 로그인·다시 시도 뒤 다시 보낸다). */
function deferSoftSteps(steps: FileSteps | undefined): FileSteps | undefined {
  if (!steps) return steps;
  const next: FileSteps = {};
  for (const file of SYNC_FILES) {
    const st = steps[file];
    if (st && !(st.status === "error" && (isLoginFailureText(st.message) || isStallFailureText(st.message)))) next[file] = st;
  }
  return next;
}

/**
 * 로그인이 끊겨(LoginRequiredError·새 토큰으로도 401) 또는 응답이 없어(시간 초과·중단) 실패한 단계는 실패가 아니라 대기다.
 * 다른 실패(409·412·423·400·네트워크 등)는 Graph 가 준 실제 이유 그대로 실패로 남긴다.
 */
function deferSoftFailures(outcome: GraphSyncOutcome): void {
  for (const rec of state.records) rec.steps = deferSoftSteps(rec.steps);
  for (const [id, steps] of outcome.steps) outcome.steps.set(id, deferSoftSteps(steps) ?? {});
  for (const file of SYNC_FILES) {
    const f = outcome.report.files[file];
    if (f.status === "error" && isLoginFailureText(f.message)) outcome.report.files[file] = { status: "idle", message: "로그인 대기" };
    else if (f.status === "error" && isStallFailureText(f.message)) outcome.report.files[file] = { status: "idle", message: "응답 없음 · 대기" };
  }
  if (outcome.report.lists) {
    for (const [book, f] of Object.entries(outcome.report.lists)) {
      if (f?.status === "error" && (isLoginFailureText(f.message) || isStallFailureText(f.message))) {
        outcome.report.lists = { ...outcome.report.lists, [book]: { status: "idle", message: isLoginFailureText(f.message) ? "로그인 대기" : "응답 없음 · 대기" } };
      }
    }
  }
}

/** 저장된 지난 보고서의 로그인·시간 초과 실패 줄을 대기로 바꾼다(다시 켰을 때 옛 '만료' 문구가 남지 않게). */
function sanitizeReport(report: SyncReport): void {
  deferSoftFailures({ steps: new Map(), excelRows: new Map(), report, discovered: [] });
}

function hasAnyFile(): boolean {
  const f = state.config.files;
  return !!(f.list?.id || listFileOf(state.config, "listGm")?.id || f.ilji?.id || f.statement?.id);
}

function applyOutcome(pending: InspectionRecord[], outcome: GraphSyncOutcome): void {
  const jobIds = new Set(pending.map((r) => r.id));
  if (outcome.discovered.length) {
    const known = new Set(state.records.map((r) => r.id));
    const fresh = outcome.discovered.filter((r) => !known.has(r.id)).map((r) => ({ ...r, sync: "synced" as const }));
    state.records = [...state.records, ...fresh];
  }
  for (const rec of state.records) {
    const changed = outcome.steps.get(rec.id);
    // 같은 라인 앱 행을 다시 정렬하면 이번 묶음이 아닌 기록의 행 번호도 바뀐다.
    const row = outcome.excelRows.get(rec.id);
    if (row && row > 0) rec.excelRow = row;
    if (!changed && !jobIds.has(rec.id)) continue;
    if (jobIds.has(rec.id) && rec.sync !== "syncing") continue; // 보내는 사이 다시 고친 기록: 다음 저장 때 처음부터.
    if (changed) rec.steps = { ...(rec.steps ?? {}), ...changed };
    const agg = aggregateSteps(rec.steps);
    if (jobIds.has(rec.id)) {
      rec.sync = agg.status;
      rec.syncError = agg.error;
      rec.syncNote = agg.note;
      if (rec.sync === "synced") delete rec.movedFrom;
    } else if (agg.status === "error") {
      // 백필 대상(이미 리스트에는 있음)이 작업일지·명세서에서 실패 → 대기열에 남겨 다시 켜도 재시도한다.
      rec.sync = "error";
      rec.syncError = agg.error;
    } else if (rec.sync === "error" && agg.status === "synced") {
      rec.sync = "synced";
      rec.syncError = undefined;
    } else if (agg.status === "pending" && SYNC_FILES.some((f) => rec.steps?.[f]?.status === "wait")) {
      // 백필 대상인데 거래명세서 그 달 시트가 아직 없다 → 대기열에 넣어 시트가 생기면 다음 동기화 때 채운다.
      rec.sync = "pending";
      rec.syncError = undefined;
    }
  }
  state.syncReport = outcome.report;
  patch({ syncReport: outcome.report });
  const failed = SYNC_FILES.filter((f) => outcome.report.files[f].status === "error").map((f) => SYNC_FILE_LABEL[f]);
  if (failed.length) toast(`${failed.join("·")} 반영 실패 · 위의 '다시 시도'를 누르세요`);
}

let initOnce: Promise<void> | null = null;

export function initSession(): Promise<void> {
  if (!initOnce) initOnce = doInit().catch((error) => {
    initOnce = null;
    throw error;
  });
  return initOnce;
}

async function doInit(): Promise<void> {
  try {
    const [savedRecords, savedConfig, savedMeta, savedReport, savedOwnership] = await Promise.all([
      kvGet<InspectionRecord[]>("records"),
      kvGet<AppConfig>("config"),
      kvGet<DayMeta>("meta"),
      kvGet<SyncReport>("syncReport"),
      kvGet<StatementOwnership>(OWNERSHIP_KEY),
    ]);
    restoreStatementOwnership(savedOwnership);
    // 보내는 중에 앱이 꺼졌으면 그 기록은 다시 보낸다(이미 끝난 파일은 steps 로 건너뜀).
    for (const rec of savedRecords ?? []) if (rec.sync === "syncing") rec.sync = "pending";
    // 예전 판이 '실패'로 남긴 로그인 만료(401)·시간 초과 단계는 대기로 되돌린다(로그인·다시 켜면 새 토큰으로 다시 보낸다).
    let hadLoginFailure = false;
    for (const rec of savedRecords ?? []) {
      if (!rec.steps || !SYNC_FILES.some((f) => rec.steps?.[f]?.status === "error")) continue;
      if (SYNC_FILES.some((f) => rec.steps?.[f]?.status === "error" && isLoginFailureText(rec.steps[f]?.message))) hadLoginFailure = true;
      rec.steps = deferSoftSteps(rec.steps);
      if (rec.sync === "error" && !SYNC_FILES.some((f) => rec.steps?.[f]?.status === "error")) {
        rec.sync = "pending";
        rec.syncError = undefined;
      }
    }
    const clientId = resolveClientId(savedConfig?.clientId ?? "");
    // 토큰 요청보다 먼저 Microsoft 응답 처리 + 캐시 계정 복원. 응답 처리 실패여도 캐시된 계정은 남는다.
    const redirected = clientId
      ? await withTimeout(completeRedirect(clientId), 25_000, TOKEN_TIMEOUT_TEXT).catch((error) => {
          console.warn("MSAL 시작 실패", error);
          return null;
        })
      : null;
    if (savedReport) sanitizeReport(savedReport);
    await loadBooks();
    let records = savedRecords;
    let inspectors: string[] = [];
    let lengths: Record<string, number> = {};
    if (!records) {
      const bytes = {
        list: await fetchBytes(SAMPLE.list),
        ilji: await fetchBytes(SAMPLE.ilji),
        statement: await fetchBytes(SAMPLE.statement),
      };
      const boot = bootstrapFromSamples(bytes);
      records = boot.records;
      inspectors = boot.inspectors;
      lengths = boot.lengths;
      await kvSet("records", records);
      for (const kind of ["list", "ilji", "statement"] as const) await kvSet(`book:${kind}`, bytes[kind]);
    }
    const config = mergeConfig(savedConfig, inspectors, lengths, sitesOf(records));
    const meta = { ...emptyMeta(), ...savedMeta };
    restoreUnsaved(records ?? [], meta);
    if (records && promoteStaleLocals(records, meta.workDate)) await kvSet("records", records);
    if (records && rerouteGmRecords(records)) await kvSet("records", records);
    const knownSites = sitesOf(records);
    if (!meta.lastSite || !knownSites.includes(meta.lastSite)) {
      meta.lastSite = knownSites.includes("드림 합정주공") ? "드림 합정주공" : knownSites[0] || "";
    }
    if (redirected?.fromRedirect && config.syncMode !== "graph") config.syncMode = "graph";
    // Microsoft 에서 막 돌아왔거나 지난번에 로그인 문제로 멈췄으면 첫 동기화는 새 토큰으로(캐시된 옛 토큰을 쓰지 않음).
    if (redirected?.fromRedirect || hadLoginFailure) freshTokenNext = true;
    if (!savedConfig || redirected?.fromRedirect) await kvSet("config", config);
    patch({
      ready: true,
      error: null,
      records,
      config,
      meta,
      account: redirected?.account?.username ?? null,
      // OneDrive 모드인데 캐시에 계정이 없으면 바로 '로그인 필요'(대기열은 그대로 둔다).
      // 캐시에 계정이 있어도 토큰으로 만료가 확인됐으면(initAuth) 로그인 필요.
      loginNeeded: config.syncMode === "graph" && (!redirected?.account || getAuthState() === "expired" || getAuthState() === "signedOut"),
      auth: getAuthState(),
      syncReport: savedReport ?? null,
    });
    if (redirected?.error && !redirected.account) toast(redirected.error);
    // 다시 켜면: 대기·실패를 보내고, OneDrive 모드면 이미 리스트에 들어간 앱 기록 중 작업일지·명세서에 빠진 것도 채운다.
    if (config.syncMode === "graph") requestSync({ backfill: true });
    else if (jobs().length) requestSync();
  } catch (error) {
    patch({ ready: true, error: error instanceof Error ? error.message : "기록을 불러오지 못했습니다." });
  }
}

function materialFor(site: string, line: string): string {
  const same = state.records.filter((r) => !r.deleted && r.site === site && r.line === line && r.material);
  return same.length ? same[same.length - 1].material : "";
}

let lastKey = "";
let lastAt = 0;
let mutationChain: Promise<void> = Promise.resolve();

/**
 * 기록을 바꾸는 작업은 모두 한 줄로 세운다. 앞 작업이 저장 중일 때 들어온 탭도
 * 끼어들지 않고 차례를 기다린다. (예전에는 "실행 중이면 바로 실행"이라
 * 빠르게 누른 S의 일부가 mirroring 중에 끼어들어 R을 건너뛰었다.)
 * 큐 안의 작업은 공개 함수(enqueue 래퍼)를 다시 부르면 안 된다 → *Now 함수를 직접 부른다.
 */
function enqueue<T>(task: () => Promise<T>): Promise<T> {
  const run = mutationChain.then(task);
  mutationChain = run.then(() => undefined, () => undefined);
  return run;
}

/** 큐에 쌓인 작업이 모두 끝날 때까지 기다린다. 테스트와 화면 전환에서 쓴다. */
export function settled(): Promise<void> {
  return mutationChain;
}

type Scope = { site: string; line: string; date: string };

function scopeNow(date = state.meta.workDate): Scope {
  return { site: state.meta.lastSite, line: state.meta.activeLine.trim(), date };
}

function scopeOf(rec: InspectionRecord): Scope {
  return { site: rec.site, line: rec.line.trim(), date: rec.date };
}

/** 같은 현장·라인·날짜·번호의 기록. 시간 순. S와 R은 이 목록의 순번으로 짝을 맞춘다. */
function samePoint(point: string, scope: Scope): InspectionRecord[] {
  const want = pointId(point);
  return state.records
    .filter((rec) => !rec.deleted && rec.site === scope.site && sameLine(rec.line, scope.line) && rec.date === scope.date && pointId(rec.point) === want)
    .sort((a, b) => a.ts - b.ts);
}

function ordinalOf(rec: InspectionRecord): number {
  return samePoint(rec.point, scopeOf(rec)).findIndex((item) => item.id === rec.id);
}

function pointId(point: string): string {
  return canonicalPoint(point) ?? point.trim().toUpperCase();
}

function sameLine(a: string, b: string): boolean {
  return a.trim().toUpperCase() === b.trim().toUpperCase();
}

const SR_KEY = "paut-sr-mirror";

/**
 * independent: 사용자가 R 결과를 직접 고친 번호(현장|라인|날짜|번호). 그 R 은 S 를 따라가지 않는다.
 * unsaved: 그중 마지막 저장 뒤에 생긴 표시. 초기화하면 R 과 함께 이 표시도 되돌린다.
 */
type SrMirrorStore = { enabled: boolean; independent: string[]; unsaved?: string[] };

function readSr(): SrMirrorStore {
  try {
    const raw = localStorage.getItem(SR_KEY);
    if (!raw) return { enabled: true, independent: [] };
    const parsed = JSON.parse(raw) as Partial<SrMirrorStore>;
    return { enabled: parsed.enabled !== false, independent: parsed.independent ?? [], unsaved: parsed.unsaved ?? [] };
  } catch {
    return { enabled: true, independent: [] };
  }
}

function sameSrKey(item: string, scope: Scope, rPoint?: string): boolean {
  const [storedSite, storedLine, storedDate, storedPoint] = item.split("|");
  if (storedSite !== scope.site || (storedLine ?? "").trim().toUpperCase() !== scope.line.trim().toUpperCase() || storedDate !== scope.date) return false;
  return rPoint === undefined || pointId(storedPoint ?? "") === pointId(rPoint);
}

/** 그 현장·라인·날짜·번호의 R 직접 편집 표시를 지운다. */
function clearDirectR(scope: Scope, rPoint: string): void {
  const cur = readSr();
  const hit = (item: string) => sameSrKey(item, scope, rPoint);
  if (!cur.independent.some(hit) && !(cur.unsaved ?? []).some(hit)) return;
  writeSr({ ...cur, independent: cur.independent.filter((item) => !hit(item)), unsaved: (cur.unsaved ?? []).filter((item) => !hit(item)) });
}

/** 초기화: 이 현장·날짜에서 마지막 저장 뒤에 생긴 R 직접 편집 표시를 지운다(되돌린 R 은 다시 S 를 따라간다). */
function resetUnsavedDirectR(site: string, date: string): void {
  const cur = readSr();
  const mine = (item: string) => {
    const [s, , d] = item.split("|");
    return s === site && d === date;
  };
  const drop = new Set((cur.unsaved ?? []).filter(mine));
  if (!drop.size) return;
  writeSr({ ...cur, independent: cur.independent.filter((item) => !drop.has(item)), unsaved: (cur.unsaved ?? []).filter((item) => !drop.has(item)) });
}

/** 저장: 이 현장·날짜의 R 직접 편집 표시는 이제 저장된 결정이다(초기화로 지우지 않음). */
function commitDirectR(site: string, date: string): void {
  const cur = readSr();
  const keep = (cur.unsaved ?? []).filter((item) => {
    const [s, , d] = item.split("|");
    return !(s === site && d === date);
  });
  if (keep.length !== (cur.unsaved ?? []).length) writeSr({ ...cur, unsaved: keep });
}

function writeSr(next: SrMirrorStore): void {
  try { localStorage.setItem(SR_KEY, JSON.stringify(next)); } catch { /* 저장 불가 */ }
}

export function srMirrorEnabled(): boolean {
  return readSr().enabled;
}

export function setSrMirrorEnabled(on: boolean): void {
  writeSr({ ...readSr(), enabled: on });
}

function srKey(scope: Scope, point: string): string {
  return `${scope.site}|${scope.line}|${scope.date}|${pointId(point)}`;
}

/**
 * 사용자가 R 카드를 직접 만졌다는 표시. 그 R 번호는 이후 S를 따라가지 않는다.
 * 자동 미러는 이 함수를 부르지 않으므로, 저장돼 있던 R 이라는 이유만으로는 표시되지 않는다.
 */
export function noteDirectR(point: string, date = state.meta.workDate, where?: { site: string; line: string }): void {
  const parsed = parsePoint(point);
  if (!parsed || parsed.prefix !== "R") return;
  const scope: Scope = where ? { site: where.site, line: where.line.trim(), date } : scopeNow(date);
  const cur = readSr();
  if (cur.independent.some((item) => sameSrKey(item, scope, point))) return;
  const key = srKey(scope, point);
  writeSr({ ...cur, independent: [...cur.independent, key], unsaved: [...(cur.unsaved ?? []), key] });
}

/**
 * 사용자가 결과를 직접 고친 R 이 지금도 있는가. 표시가 있어도 그 번호의 R 기록이 없으면(회색 카드뿐·지움·초기화)
 * 지킬 R 결과가 없으므로 S 를 따라간다. (예전: 'R 탭 +' 로 회색 R29~R31 을 만들기만 해도 표시가 남아
 * S29~S31 을 눌러도 R 이 영영 따라가지 않았다.)
 */
function directlyEdited(rPoint: string, scope: Scope): boolean {
  if (!samePoint(rPoint, scope).length) return false;
  return readSr().independent.some((item) => sameSrKey(item, scope, rPoint));
}

function mirrorBlocked(rPoint: string, scope: Scope): boolean {
  if (!readSr().enabled) return true;
  return directlyEdited(rPoint, scope);
}

/** S 목록의 index 번째와 짝인 R. index 가 없으면 가장 최근 R. */
function mirrorTarget(rPoint: string, scope: Scope, index?: number): InspectionRecord | undefined {
  const list = samePoint(rPoint, scope);
  if (index === undefined || index < 0) return list.at(-1);
  return list[index];
}

function draftKeyFor(prefix: string, scope: Scope = scopeNow()): string {
  return `${scope.site}||${scope.line}||${prefix}`;
}

function stripPlainDraft(point: string, scope: Scope): void {
  const parsed = parsePoint(point);
  if (!parsed) return;
  const key = draftKeyFor(parsed.prefix, scope);
  const list = state.meta.draftCards?.[key] ?? [];
  const next = list.filter((card) => !(pointId(card.point) === pointId(point) && !card.reinspectOf));
  if (next.length === list.length) return;
  const draftCards = { ...(state.meta.draftCards ?? {}) };
  if (next.length) draftCards[key] = next;
  else delete draftCards[key];
  state.meta = { ...state.meta, draftCards };
}

function addReinspectDraft(parentId: string, point: string, scope: Scope): void {
  const parsed = parsePoint(point);
  if (!parsed) return;
  const key = draftKeyFor(parsed.prefix, scope);
  const list = state.meta.draftCards?.[key] ?? [];
  if (list.some((card) => card.reinspectOf === parentId)) return;
  const draftCards = { ...(state.meta.draftCards ?? {}), [key]: [...list, { id: crypto.randomUUID(), point, reinspectOf: parentId }] };
  state.meta = { ...state.meta, draftCards };
}

function dropReinspectOf(parentId: string): void {
  const dropped = withoutReinspect(parentId);
  if (dropped.changed) state.meta = { ...state.meta, draftCards: dropped.draftCards };
}

async function persistMirror(): Promise<void> {
  patch({ records: [...state.records], meta: state.meta });
  await persistRecords();
  await persistMeta();
  mirrorUnsaved();
}

function blankRecord(point: string, result: ResultKind | "없음", scope: Scope, sync: InspectionRecord["sync"]): InspectionRecord {
  const sheetName = sheetNameFor(state.records, scope.site) ?? scope.site;
  // 같은 현장 기록이 있는 리스트 파일, 없으면 시트 이름으로(군포·마전 PA 시트 → 군포·마전 리스트).
  const sibling = state.records.find((r) => r.site === scope.site && r.listBook);
  const listBook = sibling?.listBook ?? (listBookOf({ sheetName, site: scope.site }) === "listGm" ? "listGm" : undefined);
  return {
    id: crypto.randomUUID(),
    site: scope.site,
    sheetName,
    ...(listBook ? { listBook } : {}),
    date: scope.date,
    line: scope.line,
    point,
    result,
    diameter: result === "없음" ? "" : diameterFromLine(scope.line) ?? "",
    material: result === "없음" ? "" : materialFor(scope.site, scope.line),
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync,
    ts: Date.now(),
  };
}

/**
 * S 기록 하나를 같은 번호의 R 에 옮긴다.
 * index = 그 S 가 같은 번호 S 목록에서 몇 번째인지(원본 0, 재검사 1 …). 같은 순번의 R 을 고치고, 없으면 만든다.
 */
async function mirrorResult(sPoint: string, result: ResultKind, scope: Scope, index: number): Promise<void> {
  const rPoint = mirrorPoint(sPoint);
  if (!rPoint || mirrorBlocked(rPoint, scope)) return;
  const existing = mirrorTarget(rPoint, scope, index);
  stripPlainDraft(rPoint, scope);
  if (existing) {
    // 'R2'(엑셀) 와 'R02' 는 같은 번호: 결과가 같으면 저장된 R 을 건드리지 않는다.
    if (existing.result !== result || pointId(existing.point) !== pointId(rPoint)) {
      existing.result = result;
      existing.point = rPoint;
      existing.sync = "local";
      existing.syncError = undefined;
    }
    if (result === FAIL) addReinspectDraft(existing.id, rPoint, scope);
    else dropReinspectOf(existing.id);
  } else {
    // 짝 R 기록이 없으면 지킬 R 결과도 없다: 남은 직접 편집 표시를 지우고 S 를 따라 만든다(이후 S 변경도 따라감).
    if (!samePoint(rPoint, scope).length) clearDirectR(scope, rPoint);
    // 재검사: 앞 순번의 R 불합격에 달린 회색 재검사 카드는 이 기록이 대신한다.
    const before = samePoint(rPoint, scope).at(-1);
    if (before) dropReinspectOf(before.id);
    const rec = blankRecord(rPoint, result, scope, "local");
    state.records = [...state.records, rec];
    if (result === FAIL) addReinspectDraft(rec.id, rPoint, scope);
  }
  await persistMirror();
}

async function mirrorRemove(sPoint: string, scope: Scope, index: number): Promise<void> {
  const rPoint = mirrorPoint(sPoint);
  if (!rPoint || mirrorBlocked(rPoint, scope)) return;
  const existing = mirrorTarget(rPoint, scope, index);
  if (existing) {
    dropReinspectOf(existing.id);
    if (existing.sync === "local" && !existing.excelRow) state.records = state.records.filter((rec) => rec.id !== existing.id);
    else {
      existing.deleted = true;
      existing.sync = existing.result === "없음" ? "synced" : "local";
    }
  }
  stripPlainDraft(rPoint, scope);
  await persistMirror();
}

export function mirrorRename(oldPoint: string, newPoint: string, date: string): Promise<void> {
  return enqueue(() => mirrorRenameNow(oldPoint, newPoint, scopeNow(date)));
}

async function mirrorRenameNow(oldPoint: string, newPoint: string, scope: Scope, index?: number): Promise<void> {
  const from = mirrorPoint(oldPoint);
  const to = mirrorPoint(newPoint);
  if (!from || !to || mirrorBlocked(to, scope) || mirrorBlocked(from, scope)) return;
  const existing = mirrorTarget(from, scope, index);
  if (existing) {
    existing.point = to;
    if (existing.result !== "없음") existing.sync = "local";
  }
  const parsed = parsePoint(from);
  if (parsed) {
    const key = draftKeyFor(parsed.prefix, scope);
    const list = (state.meta.draftCards?.[key] ?? []).map((card) => pointId(card.point) === pointId(from) && (!existing || !card.reinspectOf || card.reinspectOf === existing.id) ? { ...card, point: to } : card);
    const draftCards = { ...(state.meta.draftCards ?? {}), [key]: list };
    state.meta = { ...state.meta, draftCards };
  }
  await persistMirror();
}

export function mirrorGreyDraft(sPoint: string): Promise<void> {
  return enqueue(() => mirrorGreyDraftNow(sPoint));
}

async function mirrorGreyDraftNow(sPoint: string): Promise<void> {
  const scope = scopeNow();
  const rPoint = mirrorPoint(sPoint);
  if (!rPoint || mirrorBlocked(rPoint, scope)) return;
  if (mirrorTarget(rPoint, scope)) return;
  const key = draftKeyFor("R", scope);
  const list = state.meta.draftCards?.[key] ?? [];
  if (list.some((card) => pointId(card.point) === pointId(rPoint))) return;
  const draftCards = { ...(state.meta.draftCards ?? {}), [key]: [...list, { id: crypto.randomUUID(), point: rPoint }] };
  state.meta = { ...state.meta, draftCards };
  patch({ meta: state.meta });
  await persistMeta();
  mirrorUnsaved();
}

export function mirrorDropGrey(sPoint: string): Promise<void> {
  return enqueue(() => mirrorDropGreyNow(sPoint));
}

async function mirrorDropGreyNow(sPoint: string): Promise<void> {
  const scope = scopeNow();
  const rPoint = mirrorPoint(sPoint);
  if (!rPoint || mirrorBlocked(rPoint, scope)) return;
  if (mirrorTarget(rPoint, scope)) return;
  stripPlainDraft(rPoint, scope);
  patch({ meta: state.meta });
  await persistMeta();
  mirrorUnsaved();
}

/** S 가 없음이 되면 R 도 없음. removeIndex 가 있으면 그 순번의 R 기록을 지운다. */
async function mirrorAbsent(sPoint: string, scope: Scope, removeIndex: number | null): Promise<void> {
  const rPoint = mirrorPoint(sPoint);
  if (!rPoint || mirrorBlocked(rPoint, scope)) return;
  if (removeIndex !== null) {
    const existing = mirrorTarget(rPoint, scope, removeIndex);
    if (existing && existing.result !== "없음") {
      dropReinspectOf(existing.id);
      if (existing.sync === "local" && !existing.excelRow) state.records = state.records.filter((rec) => rec.id !== existing.id);
      else {
        existing.deleted = true;
        existing.sync = "local";
      }
    }
  }
  stripPlainDraft(rPoint, scope);
  const main = parsePoint(rPoint);
  const mainPoint = main ? fmtPoint(main.prefix, main.num) : rPoint;
  if (!samePoint(mainPoint, scope).some((rec) => rec.result === "없음")) {
    state.records = [...state.records, blankRecord(mainPoint, "없음", scope, "synced")];
  }
  await persistMirror();
}

/**
 * 칠하기가 이미 합격(또는 불합격)인 S 위를 지나갔을 때: S 는 그대로 두고, 짝 R 기록이 없으면 같은 결과로 만든다.
 * (예전: 이미 합격인 S29~S31 위를 칠하면 아무것도 하지 않아, 빠진 R29~R31 이 끝내 생기지 않았다.)
 * 짝 R 이 이미 있으면 건드리지 않는다(결과가 달라도 — 엑셀·사용자가 정한 값).
 */
export function ensurePairR(id: string): Promise<void> {
  return enqueue(async () => {
    const rec = state.records.find((r) => r.id === id && !r.deleted);
    if (!rec || parsePoint(rec.point)?.prefix !== "S" || (rec.result !== PASS && rec.result !== FAIL)) return;
    const rPoint = mirrorPoint(rec.point);
    const scope = scopeOf(rec);
    if (!rPoint || mirrorBlocked(rPoint, scope)) return;
    const index = ordinalOf(rec);
    if (mirrorTarget(rPoint, scope, index)) return;
    await mirrorResult(rec.point, rec.result as ResultKind, scope, index);
  });
}

export function savePoint(result: ResultKind, pointLabel: string): Promise<string | null> {
  return enqueue(() => savePointNow(result, pointLabel));
}

async function savePointNow(result: ResultKind, pointLabel: string): Promise<string | null> {
  const site = state.meta.lastSite;
  const line = state.meta.activeLine.trim();
  if (!site) {
    toast("현장을 선택하세요");
    return null;
  }
  if (!line) {
    toast("라인을 선택하세요");
    return null;
  }
  const parsed = parsePoint(pointLabel);
  if (!parsed) {
    toast("포인트 형식: S01");
    return null;
  }
  const point = parsed.suffix ? parsed.raw : fmtPoint(parsed.prefix, parsed.num);
  const key = `${site}|${line}|${point}|${result}|${state.meta.workDate}`;
  const now = Date.now();
  if (key === lastKey && now - lastAt < 700) return null;
  lastKey = key;
  lastAt = now;
  try { navigator.vibrate?.(12); } catch { /* 햅틱이 없는 기기 */ }
  const sheetName = sheetNameFor(state.records, site) ?? site;
  const rec: InspectionRecord = {
    id: crypto.randomUUID(),
    site,
    sheetName,
    date: state.meta.workDate,
    line,
    point,
    result,
    diameter: "",
    material: materialFor(site, line),
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "local",
    ts: now,
  };
  rec.diameter = diameterFromLine(line) ?? "";
  state.records = [...state.records, rec];
  patch({ records: state.records });
  await persistRecords();
  mirrorUnsaved();
  if (parsed.prefix === "R") noteDirectR(point, rec.date, rec);
  if (parsed.prefix === "S") await mirrorResult(point, result, scopeOf(rec), ordinalOf(rec));
  return rec.id;
}

/** 용접이 없는 번호. 제안에만 쓰이고 개소·엑셀에는 넣지 않는다. */
/** 선택한 날짜의 초안만 다른 접두어로 복사한다. 엑셀에는 저장 전까지 쓰지 않는다. */
export async function copyDayPrefix(from: string, to: string): Promise<number> {
  const site = state.meta.lastSite;
  const line = state.meta.activeLine.trim();
  const date = state.meta.workDate;
  if (!site || !line) {
    toast("라인을 선택하세요");
    return 0;
  }
  const sameDay = state.records.filter((rec) => !rec.deleted && rec.site === site && rec.line === line && rec.date === date);
  const destKey = `${site}||${line}||${to.toUpperCase()}`;
  const grey = (state.meta.draftCards?.[destKey] ?? []).map((card) => card.point);
  const planned = planPrefixCopy({
    from,
    to,
    drafts: sameDay.filter((rec) => rec.sync === "local"),
    existing: [...sameDay.map((rec) => rec.point), ...grey],
  });
  if (!planned.length) {
    toast("복사할 초안이 없습니다");
    return 0;
  }
  const now = Date.now();
  const sheetName = sheetNameFor(state.records, site) ?? site;
  const added: InspectionRecord[] = planned.map((row, index) => ({
    id: crypto.randomUUID(),
    site,
    sheetName,
    date,
    line,
    point: row.point,
    result: row.result,
    diameter: diameterFromLine(line) ?? "",
    material: materialFor(site, line),
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "local",
    ts: now + index,
  }));
  const draftCards = { ...(state.meta.draftCards ?? {}) };
  const list = [...(draftCards[destKey] ?? [])];
  for (const rec of added) {
    if (rec.result !== FAIL) continue;
    list.push({ id: crypto.randomUUID(), point: rec.point, reinspectOf: rec.id });
  }
  if (list.length) draftCards[destKey] = list;
  state.records = [...state.records, ...added];
  state.meta = { ...state.meta, draftCards, prefix: to.toUpperCase() };
  patch({ records: state.records, meta: state.meta });
  await persistRecords();
  await persistMeta();
  mirrorUnsaved();
  toast(`${from.toUpperCase()} → ${to.toUpperCase()} ${planned.length}개 복사`);
  return planned.length;
}

export function saveAbsent(pointLabel: string): Promise<void> {
  return enqueue(() => saveAbsentNow(pointLabel));
}

async function saveAbsentNow(pointLabel: string): Promise<void> {
  const site = state.meta.lastSite;
  const line = state.meta.activeLine.trim();
  const parsed = parsePoint(pointLabel);
  if (!site || !line || !parsed || parsed.suffix) {
    toast("본번호를 없음으로 표시할 수 있습니다");
    return;
  }
  const point = fmtPoint(parsed.prefix, parsed.num);
  if (state.records.some((r) => !r.deleted && r.site === site && r.line === line && pointId(r.point) === pointId(point) && r.result === "없음")) {
    toast(`${point} 은 이미 없음`);
    return;
  }
  state.records = [...state.records, {
    id: crypto.randomUUID(),
    site,
    sheetName: sheetNameFor(state.records, site) ?? site,
    date: state.meta.workDate,
    line,
    point,
    result: "없음",
    diameter: "",
    material: "",
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "synced",
    ts: Date.now(),
  }];
  patch({ records: state.records });
  await persistRecords();
  toast(`${point} 없음 · 엑셀에 쓰지 않음`);
  if (parsed.prefix === "R") noteDirectR(point, state.meta.workDate, { site, line });
  if (parsed.prefix === "S") await mirrorAbsent(point, { site, line, date: state.meta.workDate }, null);
}

export async function saveMany(rows: { point: string; result: ResultKind }[]): Promise<void> {
  const site = state.meta.lastSite;
  const line = state.meta.activeLine.trim();
  if (!site || !line) {
    toast("현장·라인이 필요합니다");
    return;
  }
  const sheetName = sheetNameFor(state.records, site) ?? site;
  const diameter = diameterFromLine(line) ?? "";
  const material = materialFor(site, line);
  const added: InspectionRecord[] = rows.map((row, i) => ({
    id: crypto.randomUUID(),
    site,
    sheetName,
    date: state.meta.workDate,
    line,
    point: row.point,
    result: row.result,
    diameter,
    material,
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "local",
    ts: Date.now() + i,
  }));
  state.records = [...state.records, ...added];
  patch({ records: state.records });
  await persistRecords();
  mirrorUnsaved();
}

export function setResult(id: string, result: ResultKind): Promise<void> {
  return enqueue(() => setResultNow(id, result));
}

async function setResultNow(id: string, result: ResultKind): Promise<void> {
  const rec = state.records.find((r) => r.id === id);
  if (!rec || rec.deleted) return;
  rec.result = result;
  rec.sync = "local";
  rec.syncError = undefined;
  patch({ records: [...state.records] });
  await persistRecords();
  mirrorUnsaved();
  const parsed = parsePoint(rec.point);
  if (parsed?.prefix === "R") noteDirectR(rec.point, rec.date, rec);
  if (parsed?.prefix === "S") await mirrorResult(rec.point, result, scopeOf(rec), ordinalOf(rec));
}

function withoutReinspect(parentId: string): { draftCards: Record<string, DraftCard[]>; changed: boolean } {
  const draftCards: Record<string, DraftCard[]> = {};
  let changed = false;
  for (const [key, list] of Object.entries(state.meta.draftCards ?? {})) {
    const next = list.filter((card) => card.reinspectOf !== parentId);
    draftCards[key] = next;
    if (next.length !== list.length) changed = true;
  }
  return { draftCards, changed };
}

export function removeRecord(id: string, quiet = false): Promise<void> {
  return enqueue(() => removeRecordNow(id, quiet));
}

async function removeRecordNow(id: string, quiet = false): Promise<void> {
  const rec = state.records.find((r) => r.id === id);
  if (!rec) return;
  const removedPoint = rec.point;
  const scope = scopeOf(rec);
  const index = ordinalOf(rec);
  const neverSent = rec.sync === "local" && !rec.excelRow;
  if (neverSent) state.records = state.records.filter((item) => item.id !== id);
  else {
    rec.deleted = true;
    rec.sync = rec.result === "없음" ? "synced" : "local";
  }
  const dropped = withoutReinspect(id);
  if (dropped.changed) state.meta = { ...state.meta, draftCards: dropped.draftCards };
  patch({ records: [...state.records], meta: state.meta });
  await persistRecords();
  if (dropped.changed) await persistMeta();
  mirrorUnsaved();
  if (!quiet) toast("삭제했습니다. 저장을 누르면 엑셀에도 반영됩니다");
  const parsed = parsePoint(removedPoint);
  if (parsed?.prefix === "R") noteDirectR(removedPoint, scope.date, scope);
  if (parsed?.prefix === "S") await mirrorRemove(removedPoint, scope, index);
}

export function clearGreyReinspect(parentId: string): Promise<void> {
  return enqueue(() => clearGreyReinspectNow(parentId));
}

async function clearGreyReinspectNow(parentId: string): Promise<void> {
  const dropped = withoutReinspect(parentId);
  if (!dropped.changed) return;
  state.meta = { ...state.meta, draftCards: dropped.draftCards };
  patch({ meta: state.meta });
  await persistMeta();
}

export function renamePoint(id: string, pointLabel: string): Promise<void> {
  return enqueue(() => renamePointNow(id, pointLabel));
}

async function renamePointNow(id: string, pointLabel: string): Promise<void> {
  const rec = state.records.find((r) => r.id === id && !r.deleted);
  const parsed = parsePoint(pointLabel);
  if (!rec || !parsed) {
    toast("포인트 형식: S01");
    return;
  }
  const previous = rec.point;
  const scope = scopeOf(rec);
  const index = ordinalOf(rec);
  const point = parsed.suffix ? parsed.raw : fmtPoint(parsed.prefix, parsed.num);
  rec.point = point;
  const draftCards = { ...(state.meta.draftCards ?? {}) };
  for (const [key, list] of Object.entries(draftCards)) {
    draftCards[key] = list.map((card) => (card.reinspectOf === id ? { ...card, point } : card));
  }
  state.meta = { ...state.meta, draftCards };
  if (rec.result !== "없음") rec.sync = "local";
  patch({ records: [...state.records], meta: state.meta });
  await persistRecords();
  await persistMeta();
  mirrorUnsaved();
  toast(`${point} 로 바꿈`);
  if (parsePoint(previous)?.prefix === "R") {
    noteDirectR(previous, scope.date, scope);
    noteDirectR(point, scope.date, scope);
  }
  // 큐 안이므로 공개 mirrorRename(큐에 다시 넣음)이 아니라 Now 를 직접 부른다.
  if (parsePoint(previous)?.prefix === "S") await mirrorRenameNow(previous, point, scope, index);
}

/**
 * 카드 길게 누르기 → 날짜 저장. 기록 날짜를 바꾸고 바로 엑셀에 다시 보낸다.
 * - 리스트: 같은 Y(앱 ID) 행을 고치고, 같은 라인의 앱 행을 날짜 → 라인 → 포인트 순으로 다시 놓는다(행 추가 없음).
 * - 작업일지·거래명세서: 옛 날짜·옛 달도 다시 계산해 옛 자리의 개수를 빼고 새 날짜·새 달에 쓴다(중복 없음).
 *   새 달 시트가 없으면 만들지 않고 대기(주황 안내)로 둔다.
 * S→R 자동으로 따라 만든 같은 번호 R 도 함께 옮긴다(직접 고친 R 은 그대로).
 */
export function changeRecordDate(id: string, date: string): Promise<void> {
  return enqueue(() => changeRecordDateNow(id, date));
}

async function changeRecordDateNow(id: string, date: string): Promise<void> {
  const rec = state.records.find((r) => r.id === id && !r.deleted);
  if (!rec || !/^\d{4}-\d{2}-\d{2}$/.test(date) || rec.date === date) return;
  const moving = [rec];
  const parsed = parsePoint(rec.point);
  if (parsed?.prefix === "S") {
    const rPoint = mirrorPoint(rec.point);
    const scope = scopeOf(rec);
    const twin = rPoint && !mirrorBlocked(rPoint, scope) ? mirrorTarget(rPoint, scope, ordinalOf(rec)) : undefined;
    if (twin && twin.date === rec.date) moving.push(twin);
  }
  for (const item of moving) {
    const from = item.date;
    item.date = date;
    if (item.result === "없음") continue;
    const sent = item.sync !== "local" || !!item.excelRow || !!item.steps;
    if (sent) item.movedFrom = [...new Set([...(item.movedFrom ?? []), from])].filter((d) => d !== date);
    item.sync = "pending";
    item.steps = undefined;
    item.syncError = undefined;
    item.syncNote = undefined;
  }
  patch({ records: [...state.records] });
  await persistRecords();
  mirrorUnsaved();
  toast(`${rec.point} 날짜 → ${date.slice(5).replace("-", "/")}`);
  requestSync();
}

/** 기존 합·불합격을 지우고, 그 본번호는 없음으로만 남긴다. 엑셀 행은 삭제된다. */
export function markNoWeld(id: string): Promise<void> {
  return enqueue(() => markNoWeldNow(id));
}

async function markNoWeldNow(id: string): Promise<void> {
  const rec = state.records.find((r) => r.id === id && !r.deleted);
  if (!rec) return;
  const original = rec.point;
  const scope = scopeOf(rec);
  const index = ordinalOf(rec);
  const parsed = parsePoint(rec.point);
  const main = parsed ? fmtPoint(parsed.prefix, parsed.num) : rec.point;
  rec.deleted = true;
  rec.sync = rec.result === "없음" || !rec.excelRow ? "local" : "local";
  const dropped = withoutReinspect(id);
  if (dropped.changed) state.meta = { ...state.meta, draftCards: dropped.draftCards };
  const hasSkip = state.records.some((r) => !r.deleted && r.site === rec.site && r.line === rec.line && pointId(r.point) === pointId(main) && r.result === "없음");
  if (!hasSkip) {
    state.records.push({
      ...rec,
      id: crypto.randomUUID(),
      point: main,
      result: "없음",
      deleted: false,
      sync: "synced",
      src: "user",
      ts: Date.now(),
    });
  }
  patch({ records: [...state.records], meta: state.meta });
  await persistRecords();
  if (dropped.changed) await persistMeta();
  toast(`${main} 없음`);
  mirrorUnsaved();
  if (parsePoint(original)?.prefix === "R") noteDirectR(original, scope.date, scope);
  if (parsePoint(original)?.prefix === "S") await mirrorAbsent(original, scope, index);
}

export function setLineDrafts(key: string, drafts: DraftCard[]): Promise<void> {
  return enqueue(() => setLineDraftsNow(key, drafts));
}

async function setLineDraftsNow(key: string, drafts: DraftCard[]): Promise<void> {
  const draftCards = { ...(state.meta.draftCards ?? {}) };
  if (drafts.length) draftCards[key] = drafts;
  else delete draftCards[key];
  state.meta = { ...state.meta, draftCards };
  patch({ meta: state.meta });
  await persistMeta();
  mirrorUnsaved();
}

export async function undoLast(): Promise<void> {
  const site = state.meta.lastSite;
  const date = state.meta.workDate;
  const last = [...state.records].reverse().find((r) => !r.deleted && r.src === "user" && r.site === site && r.date === date);
  if (!last) {
    toast("되돌릴 오늘 기록이 없습니다");
    return;
  }
  await removeRecord(last.id);
}

export async function updateMeta(partial: Partial<DayMeta>): Promise<void> {
  const byDate = { ...(state.meta.insp2ByDate ?? {}) };
  const next = { ...state.meta, ...partial, insp2ByDate: byDate };
  if (partial.workDate && partial.workDate !== state.meta.workDate) {
    byDate[state.meta.workDate] = state.meta.insp2;
    if (byDate[partial.workDate] !== undefined) next.insp2 = byDate[partial.workDate];
  }
  if (partial.insp2 !== undefined) byDate[next.workDate] = partial.insp2;
  state.meta = next;
  patch({ meta: state.meta });
  await persistMeta();
}

export async function selectSite(site: string): Promise<void> {
  const latest = latestLine(state.records.filter((rec) => !rec.deleted && rec.site === site));
  await updateMeta({ lastSite: site, activeLine: latest });
}

export async function addSite(name: string): Promise<void> {
  const site = name.trim();
  if (!site) return;
  if (isRemovedSite(site)) {
    toast(`${site} 현장은 앱에서 지웠습니다.`);
    return;
  }
  if (!state.config.mappings.some((m) => m.site === site)) {
    const mappings = [
      ...state.config.mappings,
      {
        site,
        iljiColumn: guessIljiColumn(site),
        statementPrefix: "",
        prefixConfirmed: false,
        unitPrice: null,
        priceConfirmed: false,
        priceNote: "직접 추가한 현장입니다. 작업일지 열·거래명세서 접두어·단가를 확인하세요.",
      },
    ];
    state.config = { ...state.config, mappings };
    patch({ config: state.config });
    await persistConfig();
  }
  await selectSite(site);
}

export async function addInspector(name: string): Promise<void> {
  const trimmed = name.trim();
  if (!trimmed) return;
  const added = state.config.addedInspectors ?? [];
  if (added.includes(trimmed)) return;
  state.config = { ...state.config, addedInspectors: [...added, trimmed] };
  patch({ config: state.config });
  await persistConfig();
}

export async function saveConfig(config: AppConfig): Promise<void> {
  state.config = config;
  patch({ config });
  await persistConfig();
  toast("설정을 저장했습니다");
}

/**
 * Microsoft 로그인(redirect). 평소에는 계정 선택 화면 없이 지난 계정으로 바로 간다.
 * switchAccount 는 설정의 '다른 계정' 버튼에서만 쓴다. 돌아오면 doInit 이 계정을 되살리고 대기열을 보낸다.
 */
export async function signIn(options: { switchAccount?: boolean } = {}): Promise<void> {
  const clientId = resolveClientId(state.config.clientId);
  if (!clientId) {
    toast("설정에서 Azure 클라이언트 ID를 입력하세요");
    return;
  }
  if (state.config.syncMode !== "graph") {
    state.config = { ...state.config, syncMode: "graph" };
    patch({ config: state.config });
  }
  await persistConfig();
  try {
    await loginRedirect(clientId, options);
  } catch (error) {
    toast(error instanceof Error ? `로그인을 시작하지 못했습니다: ${error.message}` : "로그인을 시작하지 못했습니다");
  }
}

export async function signOut(): Promise<void> {
  const clientId = resolveClientId(state.config.clientId);
  if (!clientId) return;
  await logoutRedirect(clientId);
}

export async function useOneDriveFile(kind: BookKind): Promise<void> {
  const id = state.config.files[kind]?.id;
  if (!id) throw new Error("파일을 먼저 고르세요.");
  const token = await accessToken(resolveClientId(state.config.clientId));
  const bytes = await downloadItem(token, id);
  rememberBaseline(kind, bytes);
  rememberWork(kind, bytes);
  await kvSet(`book:${kind}`, bytes);
  if (kind === "list" || kind === "listGm") {
    const parsed = parseListWorkbook(unzipXlsx(bytes), kind).flatMap((s) => s.records);
    state.records = replaceListRecords(kind, parsed, true);
    const config = mergeConfig(state.config, state.config.inspectors, {}, sitesOf(state.records));
    state.config = config;
    patch({ records: [...state.records], config });
    await persistRecords();
    await persistConfig();
    requestSync({ backfill: true });
  }
  toast("OneDrive 파일을 불러왔습니다");
}

export async function reimportLocal(kind: BookKind, bytes: Uint8Array): Promise<void> {
  rememberBaseline(kind, bytes);
  rememberWork(kind, bytes);
  await kvSet(`book:${kind}`, bytes);
  if (kind === "list" || kind === "listGm") {
    const parsed = parseListWorkbook(unzipXlsx(bytes), kind).flatMap((s) => s.records);
    const others = state.records.filter((r) => listBookOf(r) !== kind);
    const user = state.records.filter((r) => r.src === "user" && listBookOf(r) === kind);
    const merged = kind === "list" ? [...parsed, ...user, ...others] : [...others, ...parsed, ...user];
    const config = mergeConfig(state.config, state.config.inspectors, {}, sitesOf(merged));
    state.records = merged;
    patch({ records: [...state.records], config });
    await persistRecords();
    await persistConfig();
  }
  toast("파일을 반영했습니다");
}

export async function resetSamples(): Promise<void> {
  const bytes = {
    list: await fetchBytes(SAMPLE.list),
    ilji: await fetchBytes(SAMPLE.ilji),
    statement: await fetchBytes(SAMPLE.statement),
  };
  const boot = bootstrapFromSamples(bytes);
  const config = mergeConfig({ ...state.config, weldLengths: {} }, boot.inspectors, boot.lengths, sitesOf(boot.records));
  state.records = boot.records;
  patch({ records: boot.records, config, error: null });
  await persistRecords();
  await persistConfig();
  for (const kind of ["list", "ilji", "statement"] as const) await kvSet(`book:${kind}`, bytes[kind]);
  toast("샘플 통합 문서로 되돌렸습니다");
}

/** 이 기기에 그 리스트 파일 사본이 있는가(군포·마전은 OneDrive 에서 받은 뒤에만). */
export function hasListBook(book: ListBookKind): boolean {
  return hasBook(book);
}

export function downloadBook(kind: BookKind, filename: string): void {
  const bytes = exportBook(kind);
  const blob = new Blob([Uint8Array.from(bytes)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

export function requeueToday(): void {
  const { lastSite, workDate } = state.meta;
  let n = 0;
  for (const rec of state.records) {
    if (rec.deleted || rec.site !== lastSite || rec.date !== workDate || rec.result === "없음" || rec.sync === "local") continue;
    rec.sync = "pending";
    rec.steps = undefined;
    n += 1;
  }
  patch({ records: [...state.records] });
  void persistRecords();
  if (!n) toast("오늘 기록이 없습니다");
  else requestSync();
}

export function retryRecord(id: string): void {
  const rec = state.records.find((r) => r.id === id);
  if (!rec) return;
  rec.sync = "pending";
  rec.syncError = undefined;
  rec.steps = dropErrorSteps(rec.steps);
  patch({ records: [...state.records] });
  void persistRecords();
  requestSync();
}

export function mappingNote(site: string): string {
  return mappingFor(state.config.mappings, site).priceNote;
}

export function syncLabel(rec: InspectionRecord): string {
  if (rec.sync === "pending") return "대기";
  if (rec.sync === "syncing") return "전송";
  if (rec.sync === "error") {
    const failed = SYNC_FILES.filter((f) => rec.steps?.[f]?.status === "error").map((f) => SYNC_FILE_LABEL[f]);
    return failed.length ? `${failed.join("·")} 실패` : "실패";
  }
  if (state.config.syncMode === "graph") return "OneDrive";
  return "로컬";
}

export function bindOnline(): () => void {
  const on = () => {
    patch({ online: navigator.onLine });
    if (navigator.onLine) requestSync();
  };
  const off = () => patch({ online: false });
  window.addEventListener("online", on);
  window.addEventListener("offline", off);
  return () => {
    window.removeEventListener("online", on);
    window.removeEventListener("offline", off);
  };
}

export { PASS, FAIL, statementSkipReason };

/* ───────────── 세 번째 탭: OneDrive 검사리스트 바로 보기·고치기 ───────────── */

/**
 * 리스트 탭이 쓸 Graph 클라이언트. OneDrive 모드가 아니거나 로그인 전이면 null + 이유(화면 안내용).
 * 동기화와 같은 토큰을 화면 전환 없이(interactive:false) 받는다.
 */
export function liveListClient(): { client: GraphClient | null; reason: string } {
  if (state.config.syncMode !== "graph") return { client: null, reason: "OneDrive 모드가 아닙니다. 설정 → Microsoft 연결에서 OneDrive 를 고르고 로그인하세요." };
  const clientId = resolveClientId(state.config.clientId);
  // 로그인 중계 서버(paut-api) 모드: 토큰은 서버에서 받으므로 클라이언트 ID 가 없어도 읽는다(accessToken 이 서버 경로로 간다).
  if (!clientId && !backendEnabled()) return { client: null, reason: "Azure 앱 클라이언트 ID가 없습니다. 설정 → Microsoft 연결에서 입력하세요." };
  // loginNeeded·account 만 보고 막지 않는다: 동기화가 401·시간 초과로 '로그인 필요'를 켜 둔 채여도
  // 조용한 토큰이 되면 읽는다. 정말 로그인이 필요하면 토큰이 LoginRequiredError 를 던져 화면에 안내한다.
  return { client: new GraphClient({ getToken: (force) => accessToken(clientId, { forceRefresh: force, interactive: false }) }), reason: "" };
}

/**
 * 리스트 탭이 OneDrive 를 실제로 읽었다 = 로그인은 살아 있다. '로그인 필요'가 켜져 있었으면 끄고,
 * 대기 기록을 새 토큰으로 다시 보낸다.
 */
export function graphReadOk(): void {
  if (state.config.syncMode !== "graph" || !state.loginNeeded) return;
  freshTokenNext = true;
  const clientId = resolveClientId(state.config.clientId);
  patch({ loginNeeded: false, account: currentAccount(clientId) ?? state.account });
  if (jobs().length) requestSync();
}

/**
 * 리스트 탭에서 엑셀 칸을 고친 뒤, 같은 앱 기록의 그 필드만 맞춘다. 엑셀은 이미 그 값이므로 다시 보내지 않는다(대기 상태 그대로).
 */
export function applyListCellEdit(recordId: string, fields: Partial<InspectionRecord>): Promise<void> {
  return enqueue(async () => {
    const at = state.records.findIndex((r) => r.id === recordId && !r.deleted);
    if (at < 0) return;
    const next: Partial<InspectionRecord> = { ...fields };
    if (next.result !== undefined) next.result = normalizeResult(String(next.result));
    if (next.point !== undefined) next.point = canonicalPoint(next.point) ?? next.point.trim().toUpperCase();
    if (next.date !== undefined && !/^\d{4}-\d{2}-\d{2}$/.test(next.date)) delete next.date;
    // 새 객체로 바꿔 1번 탭 카드(입력·오늘)가 다시 그려지게 한다. 날짜가 바뀌면 같은 기록이 새 날짜로 옮겨 간다(복제 없음).
    const rec = { ...state.records[at], ...next };
    state.records[at] = rec;
    patch({ records: [...state.records] });
    await persistRecords();
    if (rec.sync === "local") mirrorUnsaved();
  });
}
