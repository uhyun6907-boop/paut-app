import type { FileSteps, InspectionRecord, ListBookKind, SyncFile, SyncReport } from "../domain/types";
import { LIST_BOOKS, LIST_BOOK_LABEL, SYNC_FILES, SYNC_FILE_LABEL } from "../domain/types";
import { listBookOf } from "../excel/listBook";
import { missingSheetsText, parseMissingSheetText } from "../excel/engine";
import { isLoginFailureText, type AuthState } from "../sync/authState";

/** 로그인 만료로 멈춘 단계·기록은 실패가 아니라 '로그인 대기'다(빨간 실패로 보이지 않게). */
function isLoginStep(message?: string): boolean {
  return isLoginFailureText(message);
}

export interface SyncFileLine {
  file: SyncFile;
  /** 검사리스트가 두 파일로 나뉘어 보일 때 그 파일. 화면 key 는 book ?? file. */
  book?: ListBookKind;
  label: string;
  state: "ok" | "error" | "waiting" | "skip" | "idle";
  /** 화면에 그대로 보이는 한 줄. 예: "작업일지 · 실패 — PC 엑셀에서 파일이 잠겨 있습니다 (7건)" */
  text: string;
}

export interface SyncPanelModel {
  visible: boolean;
  /** ask: 거래명세서 그 달 시트를 만들어 달라는 요청(실패 아님, 눈에 띄게) */
  tone: "bad" | "wait" | "ask" | "ok";
  title: string;
  lines: SyncFileLine[];
  /** 파일 단계 정보가 없는 실패(예: 데모 통합 문서 없음) */
  extra: string[];
  /** "11월 거래명세서 시트를 만들어주세요 (…)" 요청. 실패와 함께 있어도 따로 크게 보인다. */
  prompts: string[];
  waiting: number;
  canRetry: boolean;
  /** 로그인 만료로 대기 중. 버튼은 '다시 시도' 대신 '다시 로그인'. */
  login: boolean;
}

/**
 * 동기화 상태를 파일별로 정리한다. 실패는 어느 파일(검사리스트/작업일지/거래명세서)이 왜 실패했는지 적는다.
 * 기록의 steps 가 기준이고(앱을 다시 켜도 IndexedDB 에 남음), 마지막 보고서는 성공 개수·이유를 보탠다.
 */
export function syncPanelModel(
  records: InspectionRecord[],
  report: SyncReport | null,
  ctx: { online: boolean; syncing: boolean; mode: "demo" | "graph"; loginNeeded?: boolean },
): SyncPanelModel {
  const queue = records.filter((r) => r.result !== "없음" && (r.sync === "pending" || r.sync === "error" || r.sync === "syncing"));
  const lines: SyncFileLine[] = [];
  /** 실패가 아닌 기다림 이유(예: 11월 거래명세서 시트를 만들어주세요 (드림합정26.11)) */
  const calm: string[] = [];
  const prompts: string[] = [];
  let anyError = false;
  let loginWait = false;
  // 검사리스트: 군포·마전 리스트에 보낼 기록이나 결과가 있으면 파일마다 한 줄씩(평택 / 군포·마전).
  const splitLists = LIST_BOOKS.some((b) => b !== "list" && (queue.some((r) => listBookOf(r) === b) || !!report?.lists?.[b]));
  const units: { file: SyncFile; book?: SyncFileLine["book"]; label: string; recs: InspectionRecord[]; last: SyncReport["files"][SyncFile] | undefined }[] = [];
  for (const file of SYNC_FILES) {
    if (file === "list" && splitLists) {
      for (const book of LIST_BOOKS) {
        const label = LIST_BOOK_LABEL[book];
        const last = report?.lists?.[book] ?? (book === "list" && report?.files.list?.status !== "error" ? report?.files.list : undefined);
        units.push({ file, book, label, recs: queue.filter((r) => listBookOf(r) === book), last });
      }
    } else {
      units.push({ file, label: SYNC_FILE_LABEL[file], recs: queue, last: report?.files[file] });
    }
  }
  for (const { file, book, label, recs, last } of units) {
    const loginStuck = recs.filter((r) => r.steps?.[file]?.status === "error" && isLoginStep(r.steps?.[file]?.message));
    if (loginStuck.length) loginWait = true;
    const failed = recs.filter((r) => r.steps?.[file]?.status === "error" && !isLoginStep(r.steps?.[file]?.message));
    if (failed.length) {
      anyError = true;
      const reason = latestMessage(failed, file) ?? stripLabel(stripLabel(last?.message, label), SYNC_FILE_LABEL[file]) ?? "실패";
      lines.push({ file, book, label, state: "error", text: `${label} · 실패 — ${reason} (${failed.length}건)` });
      continue;
    }
    const waiting = recs.filter((r) => {
      const st = r.steps?.[file]?.status;
      return st !== "done" && st !== "skip";
    });
    if (waiting.length && ctx.loginNeeded && ctx.mode === "graph") {
      lines.push({ file, book, label, state: "waiting", text: `${label} · 로그인 대기 ${waiting.length}건` });
      continue;
    }
    if (waiting.length) {
      const why = waitReason(waiting, file);
      if (why) calm.push(why);
      if (why && file === "statement" && waiting.some((r) => parseMissingSheetText(r.steps?.statement?.message ?? ""))) prompts.push(why);
      lines.push({ file, book, label, state: "waiting", text: why ? `${label} · ${why} · 대기 ${waiting.length}건` : `${label} · 대기 ${waiting.length}건` });
      continue;
    }
    if (last?.status === "ok") lines.push({ file, book, label, state: "ok", text: `${label} · 반영됨${last.count ? ` (${last.count})` : ""}` });
    else if (last?.status === "skip") lines.push({ file, book, label, state: "skip", text: `${label} · 건너뜀 — ${last.message ?? ""}`.trim() });
    else if (last?.status === "error" && book && !isLoginStep(last.message)) lines.push({ file, book, label, state: "error", text: `${label} · 실패 — ${stripLabel(last.message, label) ?? "실패"}` });
    else lines.push({ file, book, label, state: "idle", text: `${label} · 보낼 것 없음` });
  }
  const extra = [...new Set(queue.filter((r) => r.sync === "error" && !r.steps && r.syncError && !isLoginStep(r.syncError)).map((r) => r.syncError!))];
  if (extra.length) anyError = true;
  const offline = ctx.mode === "graph" && !ctx.online;
  const waiting = queue.length;
  const login = ctx.mode === "graph" && (!!ctx.loginNeeded || loginWait);
  let title = "";
  let tone: SyncPanelModel["tone"] = "ok";
  if (offline && waiting) {
    title = `오프라인 · ${waiting}건 대기 — 연결되면 자동으로 엑셀에 보냅니다`;
    tone = anyError ? "bad" : "wait";
  } else if (anyError) {
    title = "엑셀 반영 실패 — 아래 파일만 다시 보냅니다";
    tone = "bad";
  } else if (login && waiting) {
    // 로그인 만료는 실패가 아니다: 기록은 이 기기에 대기로 남고, 다시 로그인하면 그대로 보낸다.
    title = `Microsoft 로그인 필요 · 대기 ${waiting}건 — 다시 로그인하면 바로 보냅니다`;
    tone = "wait";
  } else if (calm.length && !ctx.syncing) {
    // 실패가 아니다. 사용자가 시트를 만들면 다음 동기화 때 채운다. 시트 요청은 눈에 띄게(ask).
    title = calm.join(" · ");
    tone = prompts.length ? "ask" : "wait";
  } else if (ctx.syncing || waiting) {
    title = ctx.syncing ? "엑셀에 보내는 중…" : `엑셀 대기 ${waiting}건`;
    tone = "wait";
  }
  return {
    // 로그인만 필요할 때는 화면 위 '로그인 필요' 배너가 안내한다(패널을 겹쳐 띄우지 않음). 패널이 보일 때는 버튼이 '다시 로그인'.
    visible: anyError || (offline && waiting > 0) || calm.length > 0,
    tone,
    title,
    lines,
    extra,
    prompts,
    waiting,
    canRetry: !ctx.syncing && (!offline || ctx.mode === "demo") && waiting > 0,
    login: login && waiting > 0,
  };
}

/** 기다림(wait) 단계의 이유를 하나로 모은다. 시트 요청은 달별로: "11월 거래명세서 시트를 만들어주세요 (A, B)". */
function waitReason(recs: InspectionRecord[], file: SyncFile): string | undefined {
  const messages = [...new Set(recs.map((r) => r.steps?.[file]).filter((st) => st?.status === "wait" && st.message).map((st) => st!.message!))];
  if (!messages.length) return undefined;
  const asks = messages.map(parseMissingSheetText);
  if (asks.every(Boolean)) {
    const items = asks.flatMap((a) => a!.sheets.map((sheet) => ({ sheet, month: a!.month })));
    items.sort((a, b) => a.month - b.month);
    return missingSheetsText(items);
  }
  return messages.join(" · ");
}

function latestMessage(recs: InspectionRecord[], file: SyncFile): string | undefined {
  let best: { at: number; message?: string } | null = null;
  for (const r of recs) {
    const st = r.steps?.[file];
    if (!st?.message) continue;
    if (!best || (st.at ?? 0) >= best.at) best = { at: st.at ?? 0, message: st.message };
  }
  return best?.message;
}

function stripLabel(message: string | undefined, label: string): string | undefined {
  if (!message) return undefined;
  return message.startsWith(`${label}: `) ? message.slice(label.length + 2) : message;
}

/** 파일별 상태를 모아 기록 하나의 상태를 정한다. 한 파일이라도 실패면 실패(어느 파일인지 적는다). */
export function aggregateSteps(steps: FileSteps | undefined): { status: "synced" | "error" | "pending"; error?: string; note?: string } {
  const errors = SYNC_FILES.filter((f) => steps?.[f]?.status === "error");
  if (errors.length) {
    return { status: "error", error: errors.map((f) => `${SYNC_FILE_LABEL[f]}: ${steps?.[f]?.message ?? "실패"}`).join(" · ") };
  }
  const finished = SYNC_FILES.every((f) => steps?.[f]?.status === "done" || steps?.[f]?.status === "skip");
  if (!finished) return { status: "pending" };
  const note = SYNC_FILES.map((f) => (steps?.[f]?.status === "skip" ? steps?.[f]?.message : undefined)).find(Boolean);
  return { status: "synced", note };
}

export interface LoginBannerModel {
  visible: boolean;
  title: string;
  detail: string;
  button: string;
}

/**
 * '로그인 필요' 배너. OneDrive 모드에서 조용한 토큰 갱신이 안 될 때만 보인다(설정 화면으로 넘기지 않는다).
 * 누르면 Microsoft 로그인(redirect)으로 가고, 돌아오면 대기 기록을 바로 엑셀에 보낸다.
 */
export function loginBannerModel(ctx: { mode: "demo" | "graph"; loginNeeded: boolean; waiting: number; online: boolean }): LoginBannerModel {
  const visible = ctx.mode === "graph" && ctx.loginNeeded;
  const detail = ctx.waiting
    ? `대기 ${ctx.waiting}건은 이 기기에 안전하게 남아 있고, 로그인하면 바로 엑셀에 보냅니다.`
    : "로그인하면 OneDrive 엑셀 동기화를 이어 갑니다.";
  return {
    visible,
    title: "Microsoft 로그인 필요",
    detail: ctx.online ? detail : `${detail} (지금은 오프라인)`,
    button: "로그인",
  };
}

export type HeaderTone = "on" | "send" | "wait" | "login" | "bad" | "off";

export interface HeaderStatus {
  /** on 초록(동기화됨) · send 노랑(보내는 중) · wait 주황(대기) · login 파랑(로그인 필요) · bad 빨강(실패) · off 회색(오프라인) */
  tone: HeaderTone;
  text: string;
}

/**
 * 화면 위 한 줄 상태. '동기화 중'은 실제로 보내는 동안(syncing)만이다.
 * 로그인 대기·시간 초과 대기·시트 대기로 멈춘 줄은 '대기', 로그인이 필요하면 '로그인 필요', 진짜 실패만 빨강.
 * 초록은 보낼 것·실패가 하나도 없고 로그인돼 있을 때만.
 * queue: 실제로 엑셀에 보낼 기록(대기·실패) — 세션의 isJob 과 같은 기준(지운 현장·미저장·없음 제외).
 */
export function headerStatus(ctx: {
  mode: "demo" | "graph";
  online: boolean;
  syncing: boolean;
  loginNeeded: boolean;
  account: string | null;
  queue: InspectionRecord[];
  auth?: AuthState;
}): HeaderStatus {
  if (ctx.mode === "demo") return { tone: "on", text: "이 기기에 저장됨" };
  // 로그인 만료로 멈춘 기록은 실패(빨강)가 아니라 로그인 대기(파랑).
  const failed = ctx.queue.filter((r) => r.sync === "error" && !isLoginStep(r.syncError)).length;
  const waiting = ctx.queue.length;
  if (!ctx.online) return { tone: "off", text: waiting ? `오프라인 · 대기 ${waiting}건` : "오프라인" };
  if (ctx.syncing) return { tone: "send", text: "동기화 중…" };
  if (failed) return { tone: "bad", text: `엑셀 반영 실패 ${failed}건 · 다시 시도 필요` };
  if (ctx.loginNeeded || !ctx.account || ctx.auth === "expired" || ctx.auth === "signedOut") return { tone: "login", text: waiting ? `로그인 필요 · 대기 ${waiting}건` : "OneDrive 로그인 필요" };
  if (waiting) return { tone: "wait", text: `대기 ${waiting}건` };
  return { tone: "on", text: "OneDrive 동기화됨" };
}

export interface SettingsAuthModel {
  /** 계정 이름(보여도 로그인됐다는 뜻은 아니다) */
  email: string | null;
  /** '다시 로그인' 버튼(계정 이름이 보여도 토큰이 확인되지 않았으면 보인다) */
  showRelogin: boolean;
  /** 계정이 전혀 없을 때 'Microsoft 로그인' */
  showLogin: boolean;
  status: string;
}

/** 설정 화면의 로그인 칸. 이메일만으로는 '로그인됨'이 아니다: 토큰이 확인된(auth=ok) 때만 다시 로그인을 숨긴다. */
export function settingsAuthModel(ctx: { account: string | null; auth: AuthState; loginNeeded: boolean }): SettingsAuthModel {
  const email = ctx.account;
  if (!email) return { email: null, showRelogin: false, showLogin: true, status: "로그인 안 됨" };
  const ok = ctx.auth === "ok" && !ctx.loginNeeded;
  const status = ok ? "로그인됨" : ctx.auth === "expired" || ctx.loginNeeded ? "로그인 만료 · 다시 로그인 필요" : ctx.auth === "signedOut" ? "로그인 필요" : "로그인 확인 중 · 안 되면 다시 로그인";
  return { email, showRelogin: !ok, showLogin: false, status };
}
