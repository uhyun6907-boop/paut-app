import { useEffect, useMemo, useRef, useState } from "react";
import { isRemovedSite } from "../domain/mapping";
import { LIST_BOOK_LABEL } from "../domain/types";
import { cellText, isEditable, liveTarget, loadLiveList, matchRecord, recordPatch, writeLiveCell, LiveConflictError, type LiveCell, type LiveRow, type LiveTable } from "../excel/liveList";
import { fileOf } from "../excel/engine";
import { CACHE_STALE_TEXT, REFRESHING_TEXT, openList, saveListCache, savedAtText, type CacheSlot, type ListView } from "./listLoad";
import { FilterPanel, FilterBar } from "./ListFilterPanel";
import { applyFilters, type ListFilters } from "./listFilter";
import { DATE_PAGE, MODE_LABEL, dotDate, shortLine, summarizeByDate, toggleDate, toggleEditing, toggleRow, verdictOf, visibleColumns, type ColumnMode, type DateFocus, type DateGroup } from "./listLayout";

const MODE_KEY = "paut-list-columns";
function readMode(): ColumnMode {
  try {
    return localStorage.getItem(MODE_KEY) === "all" ? "all" : "key";
  } catch {
    return "key";
  }
}
import {
  applyListCellEdit,
  deleteSavedRecord,
  getState,
  graphReadOk,
  liveListClient,
  pullFromList,
  rememberFile,
  repairSavedRecord,
  settled,
  signIn,
  syncIdle,
  useSession,
} from "../store/session";
import { FAIL, type InspectionRecord } from "../domain/types";
import { GRAPH_TIMEOUT_TEXT, GraphError, type GraphClient } from "../sync/graph";
import { isLoginError, isLoginFailureText } from "../sync/authState";
import { ABORTED_TEXT, isTimeoutLike, TOKEN_TIMEOUT_TEXT, withTimeout } from "../sync/timeout";

/** 리스트 한 번 읽기 전체 시간 제한(ms). 넘으면 요청을 끊고 안내한다('읽는 중…'에 붙잡히지 않게). */
export const LIST_LOAD_MS = 90_000;
/** 로그인이 만료됐거나 필요할 때. 이 문구일 때 화면에 '다시 로그인' 버튼이 보인다. */
export const LIST_LOGIN_TEXT = "Microsoft 로그인 필요 — 로그인이 만료되어 OneDrive 검사리스트를 읽지 못했습니다. 아래 '다시 로그인'을 누르세요.";
export const LIST_TIMEOUT_TEXT = "OneDrive 응답이 없어(시간 초과) 리스트를 읽지 못했습니다. 인터넷 연결을 확인하고 '새로 불러오기'를 누르세요.";

/** 리스트 탭 오류 → 화면 문구(한국어). 로그인·시간 초과·Graph 실제 이유를 구분한다. */
export function listErrorText(error: unknown): string {
  if (isLoginError(error)) return LIST_LOGIN_TEXT;
  if (error instanceof GraphError && (error.status === 401 || error.detail === "expired")) return LIST_LOGIN_TEXT;
  if (error instanceof Error && isLoginFailureText(error.message)) return LIST_LOGIN_TEXT;
  if (isTimeoutLike(error)) return LIST_TIMEOUT_TEXT;
  if (error instanceof GraphError) {
    if (error.detail === "timeout" || error.detail === "aborted" || error.message === GRAPH_TIMEOUT_TEXT || error.message === ABORTED_TEXT) return LIST_TIMEOUT_TEXT;
    return `검사리스트를 읽지 못했습니다 — ${error.message}`;
  }
  if (error instanceof Error && error.message) {
    if (error.message === TOKEN_TIMEOUT_TEXT) return LIST_TIMEOUT_TEXT;
    return `검사리스트를 읽지 못했습니다 — ${error.message}`;
  }
  return "검사리스트를 읽지 못했습니다.";
}

/** 클라이언트를 만들고 시간 제한 안에 읽는다. 실패는 문구로 돌려준다(던지지 않음). */
export async function readLiveList(
  site: string,
  opts: { timeoutMs?: number; client?: GraphClient; onClient?: (client: GraphClient) => void } = {},
): Promise<{ ok: true; table: LiveTable } | { ok: false; message: string }> {
  const timeoutMs = opts.timeoutMs ?? LIST_LOAD_MS;
  if (!site) return { ok: false, message: "현장이 없습니다." };
  const { client, reason } = opts.client ? { client: opts.client, reason: "" } : liveListClient();
  if (!client) return { ok: false, message: reason };
  opts.onClient?.(client);
  try {
    const live = getState();
    const res = await withTimeout(loadLiveList(client, live.config, site, live.records), timeoutMs, LIST_TIMEOUT_TEXT);
    if (res.ok) {
      graphReadOk();
      await rememberFile(res.table.book, { id: res.table.itemId, name: res.table.fileName });
      await pullFromList(res.table, site);
      return res;
    }
    return { ok: false, message: isLoginFailureText(res.message) ? LIST_LOGIN_TEXT : res.message || "검사리스트를 읽지 못했습니다." };
  } catch (error) {
    client.abort();
    return { ok: false, message: listErrorText(error) };
  }
}

/** 확인 창에 쓰는 행 이름. 예: S38 (26.06.18) */
export function rowLabel(row: LiveRow): string {
  const text = (col: string) => {
    const cell = row.cells.find((c) => c.col === col);
    return cell ? cellText(cell).trim() : "";
  };
  const date = dotDate(text("E"));
  const point = text("G") || `${row.row}행`;
  return date ? `${point} (${date})` : point;
}

export const deleteConfirmText = (row: LiveRow) => `${rowLabel(row)} 기록을 지울까요? 엑셀 행 내용이 지워집니다.`;
export const repairConfirmText = (row: LiveRow) => `${rowLabel(row)} 리페어 처리할까요? 판정이 불합격이 되고 재검사 카드가 생깁니다.`;

export type RowAction = "repair" | "delete";

const ACTION_LABEL: Record<RowAction, string> = { repair: "리페어 처리", delete: "삭제" };

/**
 * 「수정」 중인 날짜 묶음의 행 하나에 리페어 처리·삭제. 확인을 받은 뒤에만 앱 기록을 바꾸고 엑셀 대기열에 넣는다.
 * done 이면 recordId 의 동기화가 끝난 뒤 sentNote 로 결과를 보인다.
 */
export async function applyRowAction(
  action: RowAction,
  table: Pick<LiveTable, "book" | "sheetName">,
  row: LiveRow,
  confirm: (text: string) => boolean,
): Promise<{ done: false; note: string } | { done: true; note: string; recordId: string }> {
  const rec = matchRecord(getState().records, table, row);
  if (!rec) return { done: false, note: `${rowLabel(row)}: 이 행과 맞는 앱 기록이 없어 바꾸지 않았습니다. '새로 불러오기' 뒤 다시 하세요.` };
  if (!confirm(action === "delete" ? deleteConfirmText(row) : repairConfirmText(row))) return { done: false, note: "" };
  await (action === "delete" ? deleteSavedRecord(rec.id) : repairSavedRecord(rec.id));
  return { done: true, note: `${rowLabel(row)} ${ACTION_LABEL[action]} · 엑셀에 보내는 중`, recordId: rec.id };
}

/** 동기화가 끝난 뒤 문구: 검사리스트 행까지 썼는지. 지움을 다 보낸 기록은 앱에서 치워져 rec 이 없다. */
export function sentNote(action: RowAction, row: LiveRow, rec: InspectionRecord | undefined): string {
  const head = `${rowLabel(row)} ${ACTION_LABEL[action]}`;
  return !rec || rec.steps?.list?.status === "done" ? `${head} · 검사리스트에 반영됨` : `${head} · 엑셀에 아직 못 보냄(위 동기화 상태를 확인하세요)`;
}

/** 「수정」 중 연 행의 칸들(보이는 열)과 리페어 처리·삭제. 고칠 수 있는 칸만 눌린다. */
function RowEditor(props: {
  row: LiveRow;
  cols: { col: string; title: string }[];
  record: InspectionRecord | null;
  editing: string | null;
  onCell: (cell: LiveCell) => void;
  onAction: (action: RowAction, row: LiveRow) => void;
}) {
  const { row, cols, record, editing, onCell, onAction } = props;
  return (
    <div className="dr-edit">
      <div className="dr-cells">
        {cols.map(({ col, title }) => {
          const cell = row.cells.find((c) => c.col === col);
          if (!cell) return null;
          const editable = isEditable(cell);
          return (
            <button
              key={col}
              type="button"
              className={`lc-f${editable ? "" : " ro"}${editing === cell.address ? " on" : ""}`}
              aria-label={`${title} ${cell.address}`}
              disabled={!editable}
              onClick={() => onCell(cell)}
            >
              <span className="lc-k">{title}</span>
              <span className="lc-v">{cellText(cell) || "\u00a0"}</span>
            </button>
          );
        })}
      </div>
      <div className="lc-actions">
        {record ? (
          <>
            <button className="btn small" type="button" disabled={record.result === FAIL} onClick={() => onAction("repair", row)}>리페어 처리</button>
            <button className="btn small danger" type="button" onClick={() => onAction("delete", row)}>삭제</button>
          </>
        ) : (
          <span className="lc-norec">앱 기록 없음 · 칸만 고칠 수 있음</span>
        )}
      </div>
    </div>
  );
}

/** 펼친 날짜의 행 한 칸: 포인트 · 합부 · 라인 번호 · 비고. */
function RowLine({ row }: { row: LiveRow }) {
  const text = (col: string) => {
    const cell = row.cells.find((c) => c.col === col);
    return cell ? cellText(cell).trim() : "";
  };
  const verdict = verdictOf(row);
  return (
    <>
      <span className="dr-point">{text("G") || `${row.row}행`}</span>
      <span className={`dr-verdict ${verdict}`}>{text("P") || "—"}</span>
      <span className="dr-line">{[shortLine(text("F")), text("S")].filter(Boolean).join(" · ")}</span>
    </>
  );
}

/**
 * 검사 날짜 하나: 접힌 요약(날짜·행 수·합격/불합격·라인별 포인트 범위). 누르면 그 날짜의 행이 한 줄씩 펼쳐진다(읽기만).
 * 펼친 날짜의 「수정」을 눌러야 행을 눌러 칸을 고치고 리페어 처리·삭제를 할 수 있다.
 */
export function DateSummaryCard(props: {
  group: DateGroup;
  focus: DateFocus;
  cols: { col: string; title: string }[];
  busy: boolean;
  editing: LiveCell | null;
  recordOf: (row: LiveRow) => InspectionRecord | null;
  onDate: (date: string) => void;
  onEditing: () => void;
  onRow: (row: number) => void;
  onCell: (cell: LiveCell) => void;
  onAction: (action: RowAction, row: LiveRow) => void;
}) {
  const { group, focus, cols, busy, editing, recordOf, onDate, onEditing, onRow, onCell, onAction } = props;
  const open = focus?.date === group.date;
  const unlocked = open && focus.editing;
  return (
    <section className={`ds${open ? " open" : ""}${unlocked ? " unlocked" : ""}`}>
      <button type="button" className="ds-head" aria-expanded={open} onClick={() => onDate(group.date)}>
        <span className="ds-top">
          <strong>{dotDate(group.date) || "날짜 없음"}</strong>
          <span>{group.rows.length}행</span>
          <span className="ds-ok">합격 {group.pass}</span>
          <span className={`ds-bad${group.fail ? " some" : ""}`}>불합격 {group.fail}</span>
          <span className="ds-caret" aria-hidden="true">{open ? "▴" : "▾"}</span>
        </span>
        <span className="ds-lines">{group.lines.map((l) => `${l.short} · ${l.ranges}`).join("  /  ")}</span>
      </button>
      {open && (
        <div className="ds-body">
          <div className="ds-bar">
            <span>{unlocked ? "행을 눌러 고치기" : "읽기 전용"}</span>
            <button className={`btn small${unlocked ? " primary" : ""}`} type="button" disabled={busy} onClick={onEditing}>
              {unlocked ? "완료" : "수정"}
            </button>
          </div>
          <div className="ds-rows">
            {group.rows.map((row) => {
              const rowOpen = unlocked && focus.row === row.row;
              return (
                <div key={row.row} className={`dr${verdictOf(row) === "bad" ? " fail" : ""}${rowOpen ? " on" : ""}`}>
                  {unlocked ? (
                    <button type="button" className="dr-row" aria-expanded={rowOpen} onClick={() => onRow(row.row)}>
                      <RowLine row={row} />
                    </button>
                  ) : (
                    <div className="dr-row">
                      <RowLine row={row} />
                    </div>
                  )}
                  {rowOpen && (
                    <RowEditor
                      row={row}
                      cols={cols}
                      record={recordOf(row)}
                      editing={editing && editing.row === row.row ? editing.address : null}
                      onCell={onCell}
                      onAction={onAction}
                    />
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </section>
  );
}

/** 세 번째 탭: 현장별 OneDrive 검사리스트 시트를 그대로 보여 주고, 칸 하나를 고치면 그 칸만 엑셀에 쓴다. */
export function ListScreen() {
  const session = useSession();
  const sites = useMemo(() => {
    const names = session.config.mappings.map((m) => m.site).filter((s) => !isRemovedSite(s));
    for (const rec of session.records) if (rec.site && !names.includes(rec.site) && !isRemovedSite(rec.site)) names.push(rec.site);
    return names;
  }, [session.config.mappings, session.records]);
  const [site, setSite] = useState(sites.includes(session.meta.lastSite) ? session.meta.lastSite : sites[0] ?? "");
  const [view, setView] = useState<ListView>({ status: "loading" });
  const [edit, setEdit] = useState<{ cell: LiveCell; text: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState("");
  const [filters, setFilters] = useState<ListFilters>({});
  const [filterCol, setFilterCol] = useState<string | null>(null);
  const [mode, setMode] = useState<ColumnMode>(readMode);
  const [limit, setLimit] = useState(DATE_PAGE);
  const [focus, setFocus] = useState<DateFocus>(null);
  function chooseMode(next: ColumnMode): void {
    setMode(next);
    try { localStorage.setItem(MODE_KEY, next); } catch { /* 저장 불가 */ }
  }
  const loadNo = useRef(0);
  const clientRef = useRef<GraphClient | null>(null);
  const viewRef = useRef(view);
  viewRef.current = view;
  const busyRef = useRef(busy);
  busyRef.current = busy;
  const expired = session.config.syncMode === "graph" && (session.auth === "expired" || session.auth === "signedOut");
  const target = site ? liveTarget(site, session.records) : null;
  const slot: CacheSlot | null = target
    ? { book: target.book, fileId: fileOf(session.config, target.book)?.id ?? "", sheetName: target.sheetName }
    : null;

  /** keep: 같은 현장 새로 불러오기 — 지금 표를 그대로 둔 채 뒤에서 읽는다. */
  async function load(keep = false): Promise<void> {
    const no = ++loadNo.current;
    const current = keep && viewRef.current.status === "ok" ? viewRef.current.table : null;
    if (!keep) setEdit(null);
    await openList({
      slot,
      current,
      alive: () => no === loadNo.current,
      show: (v) => {
        if (no === loadNo.current) setView(v);
      },
      read: async () => {
        const res = await readLiveList(site, { onClient: (c) => (clientRef.current = c) });
        if (no === loadNo.current) clientRef.current = null;
        return res;
      },
    });
  }

  // 현장·로그인 상태가 바뀔 때만 다시 읽는다(기록이 바뀔 때마다 OneDrive 를 부르지 않는다).
  // 현장이 바뀌면 열 필터도 비운다(열 값이 다르다).
  const lastSite = useRef(site);
  useEffect(() => {
    setNote("");
    const siteChanged = lastSite.current !== site;
    lastSite.current = site;
    if (siteChanged) {
      setFocus(null);
      setLimit(DATE_PAGE);
      setFilters({});
      setFilterCol(null);
    }
    void load(!siteChanged && viewRef.current.status === "ok");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [site, session.account, session.config.syncMode]);

  // 읽는 도중 로그인이 만료되면 '읽는 중…'·'업데이트 중'에 남지 않는다: 요청을 끊고 바로 로그인 안내(보이던 표는 그대로).
  useEffect(() => {
    if (!expired) return;
    setView((v) => {
      if (v.status === "loading") {
        loadNo.current += 1;
        clientRef.current?.abort();
        clientRef.current = null;
        return { status: "error", message: LIST_LOGIN_TEXT };
      }
      if (v.status === "ok" && v.refreshing) {
        loadNo.current += 1;
        clientRef.current?.abort();
        clientRef.current = null;
        return { ...v, refreshing: false, staleError: LIST_LOGIN_TEXT };
      }
      return v;
    });
  }, [expired]);

  const onCell = useMemo(
    () => (cell: LiveCell) => {
      if (busyRef.current) return;
      setNote("");
      setEdit({ cell, text: cellText(cell) });
    },
    [],
  );

  function refocus(next: (focus: DateFocus) => DateFocus): void {
    setEdit(null);
    setNote("");
    setFocus(next);
  }

  async function act(action: RowAction, row: LiveRow): Promise<void> {
    if (view.status !== "ok" || busy) return;
    setEdit(null);
    setBusy(true);
    try {
      const res = await applyRowAction(action, view.table, row, (text) => window.confirm(text));
      setNote(res.note);
      if (!res.done) return;
      await settled();
      await syncIdle();
      setNote(sentNote(action, row, getState().records.find((r) => r.id === res.recordId)));
      await load(true);
    } finally {
      setBusy(false);
    }
  }

  async function save(): Promise<void> {
    if (!edit || view.status !== "ok" || busy) return;
    const { client, reason } = liveListClient();
    if (!client) {
      setNote(reason);
      return;
    }
    const table = view.table;
    const row = table.rows.find((r) => r.row === edit.cell.row);
    const cell = row?.cells.find((c) => c.col === edit.cell.col);
    if (!row || !cell) return;
    if (edit.text.trim() === cellText(cell)) {
      setEdit(null);
      return;
    }
    setBusy(true);
    try {
      // 앱 기록은 쓰기 전 행 모습(라인·포인트·Y)으로 찾는다.
      const rec = matchRecord(getState().records, table, structuredClone(row));
      const written = await writeLiveCell(client, table, cell, edit.text);
      const fields = recordPatch(cell.col, written);
      if (rec && fields) await applyListCellEdit(rec.id, fields);
      setNote(`${cell.address} 저장됨${rec && fields ? ` · 입력 탭 카드도 고침${fields.date ? ` (${fields.date})` : ""}` : ""}`);
      setEdit(null);
      // 고친 행만 새 객체로(그 행만 다시 그린다). 캐시도 고친 값으로.
      const next: LiveTable = { ...table, rows: table.rows.map((r) => (r === row ? { ...row, cells: [...row.cells] } : r)) };
      const wasRefreshing = view.refreshing;
      if (wasRefreshing) {
        // 뒤에서 읽던 결과는 고치기 전 값일 수 있다: 버리고 고친 표 위에서 다시 읽는다.
        loadNo.current += 1;
        clientRef.current?.abort();
        clientRef.current = null;
      }
      setView({ ...view, table: next, refreshing: false });
      viewRef.current = { ...view, table: next, refreshing: false };
      if (slot) void saveListCache(slot, next);
      if (wasRefreshing) void load(true);
    } catch (error) {
      setNote(error instanceof LiveConflictError ? error.message : listErrorText(error));
      if (error instanceof LiveConflictError) void load(true);
    } finally {
      setBusy(false);
    }
  }

  const table = view.status === "ok" ? view.table : null;
  const shownRows = useMemo(() => (table ? applyFilters(table.rows, filters) : []), [table, filters]);
  useEffect(() => setLimit(DATE_PAGE), [filters]);
  const cols = useMemo(() => (table ? visibleColumns(table.header, mode) : []), [table, mode]);
  const groups = useMemo(() => summarizeByDate(shownRows), [shownRows]);
  const recordOf = (row: LiveRow) => (table ? matchRecord(session.records, table, row) : null);
  const actRef = useRef(act);
  actRef.current = act;
  const onAction = useMemo(() => (action: RowAction, row: LiveRow) => void actRef.current(action, row), []);
  const loginStale = view.status === "ok" && view.staleError === LIST_LOGIN_TEXT;

  return (
    <div className="list-screen">
      <div className="card">
        <div className="k" style={{ marginBottom: 6 }}>현장</div>
        <select className="site-select" value={site} onChange={(e) => setSite(e.target.value)}>
          {sites.map((name) => (
            <option key={name} value={name}>{name}</option>
          ))}
        </select>
        {target && (
          <p className="nav-hint">
            {LIST_BOOK_LABEL[target.book]} · 시트 「{target.sheetName}」 · OneDrive 에서 바로 읽음. 날짜를 눌러 펼치고 「수정」을 눌러야 고칠 수 있고, 고친 칸만 엑셀에 씁니다(시트·행은 만들지 않음).
          </p>
        )}
        <button className="btn block" type="button" disabled={view.status === "loading" || busy || (view.status === "ok" && view.refreshing)} onClick={() => void load(view.status === "ok")}>
          {view.status === "loading" ? "읽는 중…" : view.status === "ok" && view.refreshing ? REFRESHING_TEXT : "새로 불러오기"}
        </button>
      </div>

      {view.status === "ok" && (view.refreshing || view.staleError) && (
        <div className={`list-refresh${view.staleError ? " stale" : ""}`} role="status">
          {view.refreshing ? (
            <>
              <span className="spin" aria-hidden="true" />
              {REFRESHING_TEXT}
              {view.fromCache && view.savedAt ? ` · 캐시 ${savedAtText(view.savedAt)}` : ""}
            </>
          ) : (
            <>
              {view.fromCache ? CACHE_STALE_TEXT : "갱신 실패 · 이전 내용 표시 중"}
              {view.savedAt ? ` (${savedAtText(view.savedAt)})` : ""}
              {view.staleError && view.staleError !== LIST_LOGIN_TEXT ? <small>{view.staleError}</small> : null}
            </>
          )}
        </div>
      )}
      {loginStale && (
        <button className="btn primary block" type="button" disabled={!session.online} onClick={() => void signIn()}>다시 로그인</button>
      )}

      {note && <div className="nav-hint" role="status">{note}</div>}

      {edit && (
        <div className="card live-edit">
          <div className="title">{edit.cell.address} 고치기</div>
          <input
            className="field-input"
            autoFocus
            value={edit.text}
            aria-label={`${edit.cell.address} 값`}
            onChange={(e) => setEdit({ ...edit, text: e.target.value })}
            onKeyDown={(e) => {
              if (e.key === "Enter") void save();
              if (e.key === "Escape") setEdit(null);
            }}
          />
          <div className="row">
            <button className="btn primary" type="button" disabled={busy} onClick={() => void save()}>{busy ? "저장 중…" : "엑셀에 저장"}</button>
            <button className="btn" type="button" disabled={busy} onClick={() => setEdit(null)}>취소</button>
          </div>
        </div>
      )}

      {view.status === "loading" && <div className="empty-note" role="status">OneDrive 에서 {target ? `「${target.sheetName.trim()}」 시트를` : "검사리스트를"} 읽는 중…</div>}
      {view.status === "error" && <div className="empty-note list-error" role="alert">{view.message || "검사리스트를 읽지 못했습니다."}</div>}
      {view.status === "error" && view.message === LIST_LOGIN_TEXT && (
        <button className="btn primary block" type="button" disabled={!session.online} onClick={() => void signIn()}>다시 로그인</button>
      )}
      {table && (
        table.rows.length === 0 ? (
          <div className="empty-note">「{table.sheetName}」 시트에 행이 없습니다.</div>
        ) : (
          <>
            <div className="lc-mode" role="group" aria-label="보이는 열">
              {(["key", "all"] as const).map((m) => (
                <button key={m} type="button" className={`chip${mode === m ? " on" : ""}`} aria-pressed={mode === m} onClick={() => chooseMode(m)}>
                  {MODE_LABEL[m]}
                </button>
              ))}
            </div>
            <FilterBar
              header={table.header}
              filters={filters}
              shown={shownRows.length}
              total={table.rows.length}
              onOpen={(col) => setFilterCol(col)}
              onClearAll={() => {
                setFilters({});
                setFilterCol(null);
              }}
            />
            {filterCol && (
              <FilterPanel
                col={filterCol}
                title={table.header.find((h) => h.col === filterCol)?.title || filterCol}
                rows={table.rows}
                filters={filters}
                onChange={setFilters}
                onClose={() => setFilterCol(null)}
              />
            )}
            {shownRows.length === 0 ? (
              <div className="empty-note">필터에 맞는 행이 없습니다.</div>
            ) : (
              <>
                <div className="lc-list">
                  {groups.slice(0, limit).map((group) => (
                    <DateSummaryCard
                      key={group.date}
                      group={group}
                      focus={focus}
                      cols={cols}
                      busy={busy}
                      editing={edit?.cell ?? null}
                      recordOf={recordOf}
                      onDate={(date) => refocus((f) => toggleDate(f, date))}
                      onEditing={() => refocus(toggleEditing)}
                      onRow={(row) => refocus((f) => toggleRow(f, row))}
                      onCell={onCell}
                      onAction={onAction}
                    />
                  ))}
                </div>
                {groups.length > limit && (
                  <button type="button" className="btn block lc-more" onClick={() => setLimit(limit + DATE_PAGE)}>
                    더 보기 ({(groups.length - limit).toLocaleString("ko-KR")}일 남음)
                  </button>
                )}
                <div className="lc-end" />
              </>
            )}
          </>
        )
      )}
    </div>
  );
}
