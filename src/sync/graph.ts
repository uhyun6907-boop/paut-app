import type { ListWrite } from "../excel/listBook";
import { isoToExcelSerial } from "../domain/dates";
import { ABORTED_TEXT, TimeoutError, TOKEN_TIMEOUT_TEXT, withTimeout } from "./timeout";
import { AUTH_EXPIRED_TEXT, getAuthState, isLoginError, markExpired, trackRequest } from "./authState";

const GRAPH = "https://graph.microsoft.com/v1.0";

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  folder?: object;
  file?: object;
}

export function worksheetRangeUrl(itemId: string, worksheetId: string, address: string): string {
  return `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(worksheetId)}/range(address='${address}')`;
}

export function itemContentUrl(itemId: string): string {
  return `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/content`;
}

export interface ListRowOptions {
  /** 관경(L)을 값으로 쓸 때 그 값(군포·마전 리스트). 없으면 RIGHT(F,4) 수식(평택 리스트). */
  diameterValue?: string | null;
  /** 순번(A)이 비어 있는 새 행에 쓸 번호. 없으면 A 는 건드리지 않는다. */
  seq?: number;
}

/** A–Y. null 칸은 바꾸지 않는다. 관경(L)은 평택 리스트면 RIGHT 수식, 군포·마전 리스트면 값. */
export function listRowFormulas(row: number, input: ListWrite, options: ListRowOptions = {}): (string | number | null)[] {
  const cols: (string | number | null)[] = Array(25).fill(null);
  if (input.clear) {
    cols[4] = "";
    cols[5] = "";
    cols[6] = "";
    cols[10] = "";
    cols[11] = "";
    cols[15] = "";
    cols[24] = "";
    return cols;
  }
  if (options.seq !== undefined) cols[0] = options.seq;
  cols[4] = isoToExcelSerial(input.date);
  cols[5] = input.line;
  cols[6] = input.point;
  if (input.material) cols[10] = input.material;
  if (options.diameterValue !== undefined) cols[11] = options.diameterValue || null;
  else cols[11] = `=RIGHT(F${row},4)`;
  cols[15] = input.result;
  cols[24] = input.id;
  return cols;
}

/** 설정 화면(파일 고르기)용 단순 요청. 이것도 전역 inflight 에 등록해 만료 때 끊기고, 시간 제한이 실제로 abort 한다. */
async function graphFetch(url: string, token: string, init: RequestInit = {}, sessionId?: string, timeoutMs = graphTimeouts.transferMs): Promise<Response> {
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  if (sessionId) headers.set("workbook-session-id", sessionId);
  if (init.body && !(init.body instanceof ArrayBuffer) && !(init.body instanceof Blob) && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  const controller = new AbortController();
  const untrack = trackRequest(controller);
  const timer = setTimeout(() => controller.abort(new TimeoutError(GRAPH_TIMEOUT_TEXT)), timeoutMs);
  try {
    return await withTimeout(fetch(url, { ...init, headers, signal: controller.signal }), timeoutMs, GRAPH_TIMEOUT_TEXT);
  } catch (error) {
    if (controller.signal.aborted && isLoginError(controller.signal.reason)) throw new GraphError(401, GRAPH_LOGIN_TEXT, "expired");
    throw error;
  } finally {
    clearTimeout(timer);
    untrack();
  }
}

async function readError(res: Response): Promise<string> {
  const text = await res.text();
  try {
    const json = JSON.parse(text) as { error?: { message?: string } };
    return json.error?.message || text || `${res.status}`;
  } catch {
    return text || `${res.status}`;
  }
}

export async function listChildren(token: string, parentId: string | "root"): Promise<DriveItem[]> {
  const url =
    parentId === "root"
      ? `${GRAPH}/me/drive/root/children?$top=200&$select=id,name,folder,file,size`
      : `${GRAPH}/me/drive/items/${encodeURIComponent(parentId)}/children?$top=200&$select=id,name,folder,file,size`;
  const res = await graphFetch(url, token);
  if (!res.ok) throw new Error(await readError(res));
  const json = (await res.json()) as { value: DriveItem[] };
  return json.value || [];
}

export async function listXlsxFiles(token: string): Promise<DriveItem[]> {
  const queries = ["xlsx", "비파괴", "작업일지", "거래명세"];
  const batches = await Promise.all([
    listChildren(token, "root").catch(() => [] as DriveItem[]),
    ...queries.map((query) => searchDrive(token, query).catch(() => [] as DriveItem[])),
  ]);
  const map = new Map<string, DriveItem>();
  for (const item of batches.flat()) {
    if (item.folder) continue;
    if (!item.name.toLowerCase().endsWith(".xlsx")) continue;
    map.set(item.id, item);
  }
  return [...map.values()].sort((a, b) => a.name.localeCompare(b.name, "ko"));
}

export async function searchDrive(token: string, query: string): Promise<DriveItem[]> {
  const q = query.replace(/'/g, "''");
  const url = `${GRAPH}/me/drive/root/search(q='${encodeURIComponent(q)}')?$top=40&$select=id,name,folder,file,size`;
  const res = await graphFetch(url, token);
  if (!res.ok) throw new Error(await readError(res));
  const json = (await res.json()) as { value: DriveItem[] };
  return json.value || [];
}

export async function downloadItem(token: string, itemId: string): Promise<Uint8Array> {
  const res = await graphFetch(itemContentUrl(itemId), token);
  if (!res.ok) throw new Error(`파일을 받지 못했습니다. ${await readError(res)}`);
  return new Uint8Array(await res.arrayBuffer());
}

export interface GraphSheet {
  id: string;
  name: string;
}

/** Graph 실패. message 는 화면에 그대로 보여 줄 한국어. */
export class GraphError extends Error {
  status: number;
  detail: string;
  constructor(status: number, message: string, detail = "") {
    super(message);
    this.name = "GraphError";
    this.status = status;
    this.detail = detail;
  }
}

/** Graph 가 새로 받은 토큰으로도 401 을 돌려줄 때만 쓰는 문구. 다른 상태 코드에는 절대 쓰지 않는다. */
export const GRAPH_LOGIN_TEXT = `${AUTH_EXPIRED_TEXT}. 설정에서 다시 로그인하세요.`;
/** Graph 요청 하나가 시간 안에 끝나지 않음. */
export const GRAPH_TIMEOUT_TEXT = "OneDrive 응답이 없어(시간 초과) 멈췄습니다. 기록은 대기로 남습니다.";
export { TOKEN_TIMEOUT_TEXT };

/** 상태 코드 → 사용자에게 보일 이유. */
export function explainGraphStatus(status: number, detail = ""): string {
  if (status === 0) return "인터넷에 연결되지 않았습니다.";
  if (status === 401) return GRAPH_LOGIN_TEXT;
  if (status === 403) return "이 파일에 쓸 권한이 없습니다.";
  if (status === 404) return "OneDrive에서 파일이나 시트를 찾지 못했습니다(이동·삭제·이름 변경?).";
  if (status === 409 || status === 423) return "PC 엑셀 등에서 파일을 쓰는 중이라 잠겨 있습니다. 엑셀에서 저장하고 닫은 뒤 다시 시도하세요.";
  if (status === 412) return "올리는 사이에 다른 곳에서 파일이 바뀌었습니다. 다시 시도하면 새 내용으로 다시 계산합니다.";
  if (status === 429) return "OneDrive 요청이 많아 잠시 막혔습니다. 잠시 뒤 다시 시도하세요.";
  if (status >= 500) return `OneDrive 서버 오류(${status})입니다. 잠시 뒤 다시 시도하세요.`;
  return detail ? `OneDrive 오류(${status}): ${detail}` : `OneDrive 오류(${status})`;
}

/** 다시 보내면 되는 실패. 401 은 토큰을 새로 받아 다시 보낸다. */
export function isTransientStatus(status: number): boolean {
  return status === 0 || status === 409 || status === 423 || status === 429 || status >= 500;
}

/** 요청·토큰 시간 제한(ms). 아이폰 홈 화면 앱에서 응답 없는 요청이 '보내는 중…'을 영원히 붙잡지 않게 한다. */
export const graphTimeouts = {
  requestMs: 60_000,
  /** 파일 전체 받기·올리기 */
  transferMs: 180_000,
  tokenMs: 25_000,
};

export const graphRetry = {
  /** 한 요청을 최대 몇 번 보낼지(첫 시도 포함). */
  attempts: 4,
  /** 첫 대기(ms). 두 배씩 늘린다. Retry-After 가 있으면 그 값을 쓴다. */
  baseMs: 600,
  maxWaitMs: 8000,
};

export interface GraphDeps {
  /** force=true 면 캐시를 무시하고 새로 받는다. */
  getToken: (force: boolean) => Promise<string>;
  fetchImpl?: typeof fetch;
  sleep?: (ms: number) => Promise<void>;
  /** true 면 첫 토큰을 캐시 없이 새로 받는다(로그인 직후·401 뒤 첫 동기화). */
  freshToken?: boolean;
}

interface RequestOptions {
  sessionId?: string;
  /** 서버에서 실행됐는지 모를 실패(5xx·연결 끊김)에도 다시 보내도 되는가. 행 삽입은 false. */
  idempotent?: boolean;
  headers?: Record<string, string>;
  /** 이 요청의 시간 제한. 없으면 graphTimeouts.requestMs */
  timeoutMs?: number;
}

/**
 * 동기화용 Graph 클라이언트.
 * - 단계(파일)마다 refreshToken() 으로 토큰을 새로 받는다.
 * - 401 이면 토큰을 강제로 새로 받아 한 번 더 보낸다.
 * - 409/423(잠김)·429·5xx·연결 끊김은 잠시 기다렸다 다시 보낸다.
 */
export class GraphClient {
  private token = "";
  private readonly deps: GraphDeps;
  private forceNext: boolean;
  private aborted = false;
  private readonly controllers = new Set<AbortController>();
  private readonly abortWaiters = new Set<(error: Error) => void>();
  /** 테스트·진단용: 보낸 요청 수 */
  requests = 0;
  /** 새로 받은 토큰으로도 401 이었다(로그인이 실제로 끊김). */
  authFailed = false;
  /** 토큰·요청이 시간 안에 끝나지 않았다(응답 없음). 이번 묶음에서는 더 보내지 않는 것이 낫다. */
  stalled = false;

  constructor(deps: GraphDeps) {
    this.deps = deps;
    this.forceNext = !!deps.freshToken;
  }

  /** 진행 중인 요청·토큰 기다림을 모두 끊고, 이후 요청은 바로 실패시킨다(멈춘 동기화 정리). */
  abort(): void {
    if (this.aborted) return;
    this.aborted = true;
    for (const c of this.controllers) c.abort();
    this.controllers.clear();
    const error = new GraphError(0, ABORTED_TEXT, "aborted");
    for (const w of this.abortWaiters) w(error);
    this.abortWaiters.clear();
  }

  get isAborted(): boolean {
    return this.aborted;
  }

  private guard<T>(promise: Promise<T>): Promise<T> {
    if (this.aborted) {
      promise.catch(() => undefined); // 버린 약속의 늦은 실패가 처리 안 된 오류로 남지 않게
      return Promise.reject(new GraphError(0, ABORTED_TEXT, "aborted"));
    }
    return new Promise<T>((resolve, reject) => {
      const onAbort = (error: Error) => reject(error);
      this.abortWaiters.add(onAbort);
      promise.then(
        (v) => {
          this.abortWaiters.delete(onAbort);
          resolve(v);
        },
        (e) => {
          this.abortWaiters.delete(onAbort);
          reject(e);
        },
      );
    });
  }

  async refreshToken(force = false): Promise<void> {
    if (this.aborted) throw new GraphError(0, ABORTED_TEXT, "aborted");
    const fresh = force || this.forceNext;
    this.forceNext = false;
    try {
      this.token = await this.guard(withTimeout(this.deps.getToken(fresh), graphTimeouts.tokenMs, TOKEN_TIMEOUT_TEXT));
    } catch (error) {
      if (isLoginError(error)) this.authFailed = true;
      if (error instanceof TimeoutError || this.aborted) {
        this.stalled = true;
        // 토큰이 응답 없음: 같은 묶음의 다른 요청도 끊는다(로그인 상태는 바꾸지 않음).
        this.abort();
      }
      throw error;
    }
  }

  private sleep(ms: number): Promise<void> {
    if (this.deps.sleep) return this.guard(this.deps.sleep(ms));
    return this.guard(new Promise((resolve) => setTimeout(resolve, ms)));
  }

  /** fetch 하나: 시간 제한 + abort() 로 끊을 수 있게. */
  private async timedFetch(url: string, init: RequestInit, timeoutMs: number): Promise<Response> {
    if (this.aborted) throw new GraphError(0, ABORTED_TEXT, "aborted");
    const doFetch = this.deps.fetchImpl ?? ((input: RequestInfo | URL, i?: RequestInit) => fetch(input, i));
    const controller = new AbortController();
    this.controllers.add(controller);
    // 전역 등록: 토큰 만료(markExpired) 때 이 요청도 실제로 끊긴다.
    const untrack = trackRequest(controller);
    let timedOut = false;
    const timer = timeoutMs > 0 ? setTimeout(() => {
      timedOut = true;
      controller.abort(new TimeoutError(GRAPH_TIMEOUT_TEXT));
    }, timeoutMs) : null;
    // fetch 가 signal 을 무시해도 만료 abort 는 바로 끝나게 한다.
    const expired = new Promise<never>((_, reject) => {
      controller.signal.addEventListener("abort", () => {
        if (isLoginError(controller.signal.reason)) reject(controller.signal.reason);
      });
    });
    expired.catch(() => undefined);
    try {
      // fetch 가 signal 을 무시해도(일부 환경) 시간 제한·abort 는 지킨다.
      const res = await this.guard(withTimeout(Promise.race([doFetch(url, { ...init, signal: controller.signal }), expired]), timeoutMs, GRAPH_TIMEOUT_TEXT));
      return res;
    } catch (error) {
      if (isLoginError(error) || (controller.signal.aborted && isLoginError(controller.signal.reason))) {
        this.authFailed = true;
        throw new GraphError(401, GRAPH_LOGIN_TEXT, "expired");
      }
      if (this.aborted) throw new GraphError(0, ABORTED_TEXT, "aborted");
      if (timedOut || error instanceof TimeoutError) throw new TimeoutError(GRAPH_TIMEOUT_TEXT);
      throw error;
    } finally {
      if (timer) clearTimeout(timer);
      this.controllers.delete(controller);
      untrack();
    }
  }

  async request(url: string, init: RequestInit = {}, options: RequestOptions = {}): Promise<Response> {
    if (this.authFailed && getAuthState() === "expired") throw new GraphError(401, GRAPH_LOGIN_TEXT, "expired");
    if (!this.token || this.forceNext) await this.refreshToken(false);
    const idempotent = options.idempotent ?? true;
    const timeoutMs = options.timeoutMs ?? graphTimeouts.requestMs;
    let authRetried = false;
    let lastStatus = 0;
    let lastDetail = "";
    for (let attempt = 1; attempt <= graphRetry.attempts; attempt++) {
      const headers = new Headers(init.headers);
      headers.set("Authorization", `Bearer ${this.token}`);
      if (options.sessionId) headers.set("workbook-session-id", options.sessionId);
      for (const [k, v] of Object.entries(options.headers ?? {})) headers.set(k, v);
      if (typeof init.body === "string" && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");
      let res: Response | null = null;
      try {
        this.requests += 1;
        res = await this.timedFetch(url, { ...init, headers }, timeoutMs);
      } catch (error) {
        // 로그인 만료로 끊긴 요청: 다시 보내지 않는다.
        if (error instanceof GraphError && error.detail === "expired") throw error;
        // 시간 초과·중단은 다시 보내지 않는다(같은 요청이 또 멈추면 '보내는 중…'이 몇 분씩 붙잡힌다).
        if (error instanceof TimeoutError || this.aborted) {
          this.stalled = true;
          if (error instanceof GraphError) throw error;
          throw new GraphError(0, GRAPH_TIMEOUT_TEXT, "timeout");
        }
        lastStatus = 0;
        lastDetail = error instanceof Error ? error.message : String(error);
        if (!idempotent || attempt === graphRetry.attempts) break;
        await this.sleep(this.backoff(attempt, null));
        continue;
      }
      if (res.ok) return res;
      lastStatus = res.status;
      lastDetail = await readError(res);
      if (res.status === 401 && !authRetried) {
        authRetried = true;
        await this.refreshToken(true);
        attempt -= 1; // 토큰 갱신은 재시도 횟수에 넣지 않는다.
        continue;
      }
      if (res.status === 401) {
        this.authFailed = true;
        console.warn("Graph 401(새 토큰으로도)", lastDetail);
        // 로그인이 실제로 끊김: 동기화가 끝나기를 기다리지 않고 바로 전역 만료 → 진행 중 요청 전부 끊김, 화면은 '로그인 필요'.
        markExpired();
        throw new GraphError(401, GRAPH_LOGIN_TEXT, lastDetail);
      }
      const retryable = isTransientStatus(res.status) && (idempotent || res.status === 409 || res.status === 423 || res.status === 429);
      if (!retryable || attempt === graphRetry.attempts) break;
      await this.sleep(this.backoff(attempt, res.headers.get("Retry-After")));
    }
    throw new GraphError(lastStatus, explainGraphStatus(lastStatus, lastDetail), lastDetail);
  }

  private backoff(attempt: number, retryAfter: string | null): number {
    const sec = retryAfter && /^\d+$/.test(retryAfter) ? Number(retryAfter) * 1000 : 0;
    const exp = graphRetry.baseMs * 2 ** (attempt - 1);
    return Math.min(graphRetry.maxWaitMs, Math.max(sec, exp));
  }

  async json<T>(url: string, init: RequestInit = {}, options: RequestOptions = {}): Promise<T> {
    const res = await this.request(url, init, options);
    const text = await this.guard(withTimeout(res.text(), options.timeoutMs ?? graphTimeouts.requestMs, GRAPH_TIMEOUT_TEXT));
    return (text ? JSON.parse(text) : {}) as T;
  }

  async itemMeta(itemId: string): Promise<{ id: string; name: string; eTag?: string }> {
    return this.json(`${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}?$select=id,name,eTag`);
  }

  /** 드라이브 루트에서 이름이 정확히 같은 파일. 저장된 항목 ID 로 못 찾을 때(이동·다시 올림) 쓴다. */
  async findRootFile(name: string): Promise<DriveItem | null> {
    let url: string | undefined = `${GRAPH}/me/drive/root/children?$top=200&$select=id,name,file,folder`;
    for (let page = 0; url && page < 10; page++) {
      const json: { value?: DriveItem[]; "@odata.nextLink"?: string } = await this.json(url);
      const hit = (json.value ?? []).find((item) => !item.folder && item.name === name);
      if (hit) return hit;
      url = json["@odata.nextLink"];
    }
    return null;
  }

  async download(itemId: string): Promise<Uint8Array> {
    const res = await this.request(itemContentUrl(itemId), {}, { timeoutMs: graphTimeouts.transferMs });
    return new Uint8Array(await this.guard(withTimeout(res.arrayBuffer(), graphTimeouts.transferMs, GRAPH_TIMEOUT_TEXT)));
  }

  /**
   * 파일 전체 업로드. 새 달 거래명세서 시트를 만들 때만 쓴다.
   * ifMatch 를 주면 받은 뒤 다른 곳(PC 엑셀)에서 바뀐 경우 412 로 막혀 덮어쓰지 않는다.
   */
  async upload(itemId: string, bytes: Uint8Array, ifMatch?: string, contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"): Promise<void> {
    const body = new Blob([Uint8Array.from(bytes)], { type: contentType });
    const guard: Record<string, string> = ifMatch ? { "If-Match": ifMatch } : {};
    if (bytes.byteLength < 3_500_000) {
      await this.request(itemContentUrl(itemId), { method: "PUT", body }, { headers: guard, timeoutMs: graphTimeouts.transferMs });
      return;
    }
    // 큰 파일: 업로드 세션. If-Match 는 세션을 만들 때 건다.
    const session = await this.json<{ uploadUrl: string }>(
      `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/createUploadSession`,
      { method: "POST", body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } }) },
      { headers: guard },
    );
    const put = await this.timedFetch(session.uploadUrl, {
      method: "PUT",
      headers: { "Content-Length": String(bytes.byteLength), "Content-Range": `bytes 0-${bytes.byteLength - 1}/${bytes.byteLength}` },
      body,
    }, graphTimeouts.transferMs);
    if (!put.ok) {
      const detail = await readError(put);
      throw new GraphError(put.status, explainGraphStatus(put.status, detail), detail);
    }
  }

  /** 통합 문서의 시트 목록(세션 없이 읽기). */
  async worksheetsOf(itemId: string): Promise<GraphSheet[]> {
    const json = await this.json<{ value?: GraphSheet[] }>(`${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets?$select=id,name`);
    return json.value ?? [];
  }

  /** 시트에서 쓰인 범위(주소·값·수식). 시트를 만들거나 바꾸지 않는다. */
  async usedRange(itemId: string, sheetId: string): Promise<{ address: string; values: GraphCell[][]; formulas: GraphCell[][] }> {
    const url = `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(sheetId)}/usedRange?$select=address,values,formulas`;
    const json = await this.json<{ address?: string; values?: GraphCell[][]; formulas?: GraphCell[][] }>(url);
    return { address: json.address ?? "", values: json.values ?? [], formulas: json.formulas ?? json.values ?? [] };
  }

  /**
   * 값이 있는 칸의 범위 주소만(값·수식은 받지 않음). 서식만 있는 칸은 넣지 않는다(valuesOnly).
   * 실제 리스트 시트는 서식이 65536행까지 있어 usedRange 전체(값+수식)는 수백만 칸이라 아이폰에서 끝나지 않는다.
   */
  async valuesExtent(itemId: string, sheetId: string): Promise<string> {
    const url = `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(sheetId)}/usedRange(valuesOnly=true)?$select=address`;
    const json = await this.json<{ address?: string }>(url);
    return json.address ?? "";
  }

  /** 범위의 값·수식(세션 없이 읽기). */
  async readRangeCells(itemId: string, sheetId: string, address: string): Promise<{ values: GraphCell[][]; formulas: GraphCell[][] }> {
    const json = await this.json<{ values?: GraphCell[][]; formulas?: GraphCell[][] }>(worksheetRangeUrl(itemId, sheetId, address) + "?$select=values,formulas");
    return { values: json.values ?? [], formulas: json.formulas ?? json.values ?? [] };
  }

  async readRange(itemId: string, sheetId: string, address: string): Promise<GraphCell[][]> {
    const json = await this.json<{ values?: GraphCell[][] }>(worksheetRangeUrl(itemId, sheetId, address) + "?$select=values");
    return json.values ?? [];
  }

  /** 칸 하나(또는 범위)만 값으로 쓴다. ifMatch 가 있으면 그 사이 파일이 바뀐 경우 412 로 막힌다. */
  async patchValues(itemId: string, sheetId: string, address: string, values: GraphCell[][], ifMatch?: string): Promise<void> {
    await this.request(
      worksheetRangeUrl(itemId, sheetId, address),
      { method: "PATCH", body: JSON.stringify({ values }) },
      { headers: ifMatch ? { "If-Match": ifMatch } : {} },
    );
  }

  async openWorkbook(itemId: string): Promise<WorkbookSession> {
    const base = `${GRAPH}/me/drive/items/${encodeURIComponent(itemId)}/workbook`;
    const session = await this.json<{ id: string }>(`${base}/createSession`, {
      method: "POST",
      body: JSON.stringify({ persistChanges: true }),
    });
    return new WorkbookSession(this, itemId, session.id);
  }
}

export type GraphCell = string | number | null;

/** 통합 문서 세션. 셀 단위로 읽고 쓴다(파일 전체를 바꾸지 않는다). */
export class WorkbookSession {
  readonly client: GraphClient;
  readonly itemId: string;
  readonly id: string;
  private sheets: GraphSheet[] | null = null;

  constructor(client: GraphClient, itemId: string, id: string) {
    this.client = client;
    this.itemId = itemId;
    this.id = id;
  }

  private get base(): string {
    return `${GRAPH}/me/drive/items/${encodeURIComponent(this.itemId)}/workbook`;
  }

  async worksheets(refresh = false): Promise<GraphSheet[]> {
    if (this.sheets && !refresh) return this.sheets;
    const json = await this.client.json<{ value: GraphSheet[] }>(`${this.base}/worksheets?$select=id,name`, {}, { sessionId: this.id });
    this.sheets = json.value || [];
    return this.sheets;
  }

  /** 이름이 같은 시트. 앞뒤 공백(예: ' 드림 합정주공', ' PAUT FIELD')은 무시하고도 찾는다. */
  async sheet(name: string): Promise<GraphSheet | null> {
    const list = await this.worksheets();
    return list.find((s) => s.name === name) ?? list.find((s) => s.name.trim() === name.trim()) ?? null;
  }

  async readValues(sheetId: string, address: string): Promise<GraphCell[][]> {
    const json = await this.client.json<{ values?: GraphCell[][] }>(
      worksheetRangeUrl(this.itemId, sheetId, address) + "?$select=values",
      {},
      { sessionId: this.id },
    );
    return json.values || [];
  }

  /** formulas 의 null 칸은 바꾸지 않는다. '=' 로 시작하면 수식. */
  async writeFormulas(sheetId: string, address: string, formulas: GraphCell[][]): Promise<void> {
    await this.client.request(
      worksheetRangeUrl(this.itemId, sheetId, address),
      { method: "PATCH", body: JSON.stringify({ formulas }) },
      { sessionId: this.id },
    );
  }

  /** 행 삽입(아래로 밀기). 다시 보내면 두 번 들어갈 수 있어 잠김·429·401 일 때만 다시 보낸다. */
  async insertRows(sheetId: string, fromRow: number, count: number): Promise<void> {
    await this.client.request(
      worksheetRangeUrl(this.itemId, sheetId, `${fromRow}:${fromRow + count - 1}`) + "/insert",
      { method: "POST", body: JSON.stringify({ shift: "Down" }) },
      { sessionId: this.id, idempotent: false },
    );
  }

  async close(): Promise<void> {
    await this.client
      .request(`${this.base}/closeSession`, { method: "POST" }, { sessionId: this.id, idempotent: false, timeoutMs: 15_000 })
      .catch(() => undefined);
  }
}
