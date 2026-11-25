import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { isRemovedSite } from "../domain/mapping";
import { LIST_BOOK_LABEL } from "../domain/types";
import { cellText, duplicateAppRows, duplicateRowsText, isDateCol, isEditable, liveTarget, loadLiveList, matchRecord, recordPatch, writeLiveCell, writeLiveCells, LiveConflictError, type LiveCell, type LiveCellsResult, type LiveRow, type LiveTable } from "../excel/liveList";
import { LiveMoveRefused, moveLiveRow, moveOrder, restoreLiveRows, type MoveSnapshot } from "../excel/liveMove";
import { fileOf } from "../excel/engine";
import { CACHE_STALE_TEXT, REFRESHING_TEXT, openList, saveListCache, savedAtText, type CacheSlot, type ListView } from "./listLoad";
import { FilterPanel, FilterBar } from "./ListFilterPanel";
import { applyFilters, type ListFilters } from "./listFilter";
import { DATE_PAGE, MODE_LABEL, adjacentRows, dotDate, gridText, summarizeByDate, toggleDate, toggleEditing, toggleRow, verdictOf, visibleColumns, type ColumnMode, type DateFocus, type DateGroup } from "./listLayout";

const MODE_KEY = "paut-list-columns";
function readMode(): ColumnMode {
  try {
    return localStorage.getItem(MODE_KEY) === "all" ? "all" : "key";
  } catch {
    return "key";
  }
}
import {
  applyListCellEdits,
  applyListRowMove,
  dayStaffOn,
  deleteSavedRecord,
  getState,
  graphReadOk,
  liveListClient,
  pullFromList,
  rememberFile,
  repairSavedRecord,
  setDayStaff,
  settled,
  signIn,
  syncIdle,
  useSession,
  type StaffEdit,
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

export const moveConfirmText = (row: LiveRow, neighbor: LiveRow) =>
  `행 ${row.row} ↔ ${neighbor.row} 순서를 바꿀까요? ${rowLabel(row)} ↔ ${rowLabel(neighbor)} · 엑셀 B–Y 값·서식을 맞바꾸고 순번(A)은 그대로 둡니다.`;

/**
 * 「수정」 중 행 하나를 바로 위·아래 행과 맞바꾼다(엑셀 두 행의 B–Y 를 다시 씀). 제외 현장·아직 엑셀에 다 보내지 않은 카드가 있으면 하지 않는다.
 * 확인을 받은 뒤에만 쓴다. 쓴 뒤 확인이 맞으면 카드의 excelRow 를 새 행으로, 다르면 다른 칸과 되돌리기용 snapshot 을 돌려준다.
 * 다른 곳에서 바뀐 범위면 LiveConflictError, 옮길 수 없는 범위(병합 칸·메모 등)면 LiveMoveRefused.
 */
export async function applyRowMove(
  client: GraphClient,
  table: LiveTable,
  site: string,
  row: LiveRow,
  neighbor: LiveRow,
  confirm: (text: string) => boolean,
): Promise<{ done: false; note: string } | { done: true; note: string; rows: LiveRow[]; snapshot: MoveSnapshot | null }> {
  if (isRemovedSite(site)) return { done: false, note: `${site}: 제외된 현장이라 순서를 바꾸지 않습니다.` };
  const records = getState().records;
  const recs = [row, neighbor].map((r) => matchRecord(records, table, r));
  const unsent = recs.filter((r) => r && r.sync !== "synced");
  if (unsent.length) {
    return { done: false, note: `${unsent.map((r) => r!.point).join(", ")}: 엑셀에 아직 다 보내지 않은 카드가 있어 순서를 바꾸지 않습니다. 동기화가 끝난 뒤 다시 하세요.` };
  }
  if (!confirm(moveConfirmText(row, neighbor))) return { done: false, note: "" };
  const res = await moveLiveRow(client, table, row.row, neighbor.row);
  if (res.mismatches.length) {
    return { done: true, note: `행 ${row.row} ↔ ${neighbor.row} 쓰기 뒤 확인이 맞지 않습니다: ${res.mismatches.join(", ")}. 아래 「되돌리기」로 옮기기 전 모습을 다시 쓸 수 있습니다.`, rows: res.rows, snapshot: res.snapshot };
  }
  const { lo, sources } = moveOrder(row.row, neighbor.row);
  const moved = [row, neighbor].flatMap((r, i) => (recs[i] ? [{ id: recs[i]!.id, excelRow: res.moved.get(r.row)! }] : []));
  await applyListRowMove(table, lo, sources, moved);
  return { done: true, note: `행 ${row.row} ↔ ${neighbor.row} 순서 바꿈 · 엑셀 다시 읽어 확인함${moved.length ? ` · 카드 ${moved.length}개 행 번호 맞춤` : ""}`, rows: res.rows, snapshot: null };
}

export const moveDateConfirmText = (group: DateGroup, date: string) => `${dotDate(group.date)} 의 ${group.rows.length}행을 ${dotDate(date)} 로 바꿀까요?`;

function failedRowsText(failed: LiveCellsResult["failed"]): string {
  const byReason = new Map<string, number[]>();
  for (const { cell, error } of failed) {
    const reason = error instanceof LiveConflictError ? "다른 곳에서 바뀜" : listErrorText(error);
    byReason.set(reason, [...(byReason.get(reason) ?? []), cell.row]);
  }
  return [...byReason].map(([reason, rows]) => `${rows.map((r) => `${r}행`).join(", ")} (${reason})`).join(" · ");
}

/**
 * 「수정」 중인 날짜 묶음의 모든 행 E(검사 날짜)를 새 날짜로. 확인을 받은 뒤에만 쓰고, 이어진 행은 PATCH 한 번으로 묶는다.
 * 쓴 행과 맞는 앱 기록(카드)의 날짜도 바꾼다. 일부만 실패하면 실패한 행과 이유를 note 에 적는다.
 */
export async function moveDateGroup(
  client: GraphClient,
  table: LiveTable,
  group: DateGroup,
  date: string,
  confirm: (text: string) => boolean,
  onProgress?: (done: number, total: number) => void,
): Promise<{ done: false; note: string } | { done: true; note: string; rows: Set<number> }> {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || date === group.date) return { done: false, note: "" };
  if (!confirm(moveDateConfirmText(group, date))) return { done: false, note: "" };
  const records = getState().records;
  const recs = new Map(group.rows.map((row) => [row.row, matchRecord(records, table, row)]));
  const cells = group.rows.map((row) => row.cells.find((c) => c.col === "E")!);
  const locked = cells.filter((c) => !isEditable(c));
  const res = await writeLiveCells(client, table, cells.filter(isEditable), date, onProgress);
  res.failed.push(...locked.map((cell) => ({ cell, error: new Error("수식 칸은 고치지 않습니다.") })));
  const fields = recordPatch("E", res.value);
  const edits = fields ? res.written.flatMap((c) => (recs.get(c.row) ? [{ id: recs.get(c.row)!.id, fields }] : [])) : [];
  await applyListCellEdits(edits);
  const head = `${dotDate(group.date)} → ${dotDate(date)} · ${res.written.length}행 바꿈 · 카드 ${edits.length}개 고침`;
  return { done: true, note: res.failed.length ? `${head} · 실패 ${res.failed.length}행: ${failedRowsText(res.failed)}` : head, rows: new Set(res.written.map((c) => c.row)) };
}

export function staffText(staff: StaffEdit): string {
  const names = [staff.insp1, staff.insp2].filter(Boolean).join(" · ") || "검사자 없음";
  const span = staff.timeStart || staff.timeEnd ? `${staff.timeStart || "-"}–${staff.timeEnd || "-"}` : "시간 모름";
  return `${names} · ${span}`;
}

export const staffConfirmText = (date: string, staff: StaffEdit) =>
  `${dotDate(date)} 검사자·검사시간을 ${staffText(staff)} 로 바꿀까요? 작업일지(PAUT FIELD) 그날 행에 보냅니다.`;

/**
 * 「수정」 중인 날짜의 검사자·검사시간. 검사리스트에는 이 칸이 없어 작업일지 그날 행에 쓴다(저장과 같은 동기화).
 * 그날 행과 맞는 앱 기록이 실어 간다 — 맞는 기록이 없으면 보내지 않는다.
 */
export async function applyDayStaff(
  table: Pick<LiveTable, "book" | "sheetName">,
  group: DateGroup,
  staff: StaffEdit,
  confirm: (text: string) => boolean,
): Promise<{ done: false; note: string } | { done: true; note: string; ids: string[] }> {
  const records = getState().records;
  const ids = [...new Set(group.rows.flatMap((row) => matchRecord(records, table, row)?.id ?? []))];
  const none = `${dotDate(group.date)}: 이 날짜 행과 맞는 앱 기록이 없어 작업일지에 보낼 수 없습니다. '새로 불러오기' 뒤 다시 하세요.`;
  if (!ids.length) return { done: false, note: none };
  if (!confirm(staffConfirmText(group.date, staff))) return { done: false, note: "" };
  if (!(await setDayStaff(group.date, staff, ids))) return { done: false, note: none };
  return { done: true, note: `${dotDate(group.date)} 검사자·검사시간 · 작업일지에 보내는 중`, ids };
}

export function staffSentNote(date: string, recs: InspectionRecord[]): string {
  const head = `${dotDate(date)} 검사자·검사시간`;
  return recs.length && recs.every((r) => r.steps?.ilji?.status === "done") ? `${head} · 작업일지에 반영됨` : `${head} · 작업일지에 아직 못 보냄(위 동기화 상태를 확인하세요)`;
}

function DayPanel(props: {
  group: DateGroup;
  staff: StaffEdit;
  busy: boolean;
  onMoveDate: (group: DateGroup, date: string) => void;
  onStaff: (group: DateGroup, staff: StaffEdit) => void;
}) {
  const { group, busy, onMoveDate, onStaff } = props;
  const [date, setDate] = useState(group.date);
  const [staff, setStaff] = useState(props.staff);
  const field = (key: keyof StaffEdit, label: string, type: "text" | "time") => (
    <label className="dp-f">
      <span>{label}</span>
      <input className="field-input" type={type} value={staff[key]} aria-label={`${label} (그날)`} onChange={(e) => setStaff({ ...staff, [key]: e.target.value })} />
    </label>
  );
  return (
    <details className="dp">
      <summary>
        그날 한꺼번에 <span className="dp-caret">▸</span>
        <span className="dp-staff">{staffText(props.staff)}</span>
      </summary>
      <div className="dp-row">
        <label className="dp-f wide">
          <span>검사 날짜(E) · 그날 {group.rows.length}행</span>
          <input className="field-input" type="date" value={date} aria-label="검사 날짜 (그날 모든 행)" onChange={(e) => setDate(e.target.value)} />
        </label>
        <button className="btn small primary" type="button" disabled={busy || !date || date === group.date} onClick={() => onMoveDate(group, date)}>날짜 바꾸기</button>
      </div>
      <div className="dp-grid">
        {field("insp1", "검사자 1", "text")}
        {field("insp2", "검사자 2", "text")}
        {field("timeStart", "시작", "time")}
        {field("timeEnd", "종료", "time")}
      </div>
      <div className="dp-row">
        <span className="dp-hint">검사자·검사시간은 리스트에 칸이 없어 작업일지에 씁니다(동기화)</span>
        <button className="btn small primary" type="button" disabled={busy} onClick={() => onStaff(group, staff)}>작업일지에 저장</button>
      </div>
    </details>
  );
}

function keepInsideGrid(el: HTMLDivElement | null): void {
  const wrap = el?.closest(".xg-wrap")?.getBoundingClientRect();
  const anchor = el?.parentElement?.getBoundingClientRect();
  if (!el || !wrap || !anchor) return;
  el.classList.toggle("end", anchor.left + el.offsetWidth > wrap.right && anchor.right - el.offsetWidth >= wrap.left);
  el.classList.toggle("up", anchor.bottom + el.offsetHeight > wrap.bottom && anchor.top - el.offsetHeight >= wrap.top);
  el.querySelector("input")?.focus({ preventScroll: true });
}

export interface CellEdit {
  cell: LiveCell;
  text: string;
}

function CellEditor(props: { edit: CellEdit; busy: boolean; onText: (text: string) => void; onSave: () => void; onCancel: () => void }) {
  const { edit, busy, onText, onSave, onCancel } = props;
  const place = useCallback(keepInsideGrid, []);
  return (
    <div className="xg-pop" ref={place} role="dialog" aria-label={`${edit.cell.address} 고치기`}>
      <input
        className="field-input"
        type={isDateCol(edit.cell.col) ? "date" : "text"}
        value={edit.text}
        aria-label={`${edit.cell.address} 값`}
        onChange={(e) => onText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter") onSave();
          if (e.key === "Escape") onCancel();
        }}
      />
      <div className="xg-pop-row">
        <button className="btn small primary" type="button" disabled={busy} onClick={onSave}>{busy ? "저장 중…" : "저장"}</button>
        <button className="btn small" type="button" disabled={busy} onClick={onCancel}>취소</button>
      </div>
    </div>
  );
}

function RowMenu(props: {
  row: LiveRow;
  up: LiveRow | null;
  down: LiveRow | null;
  record: InspectionRecord | null;
  busy: boolean;
  onMove: (row: LiveRow, neighbor: LiveRow) => void;
  onAction: (action: RowAction, row: LiveRow) => void;
  onClose: () => void;
}) {
  const { row, up, down, record, busy, onMove, onAction, onClose } = props;
  const place = useCallback(keepInsideGrid, []);
  return (
    <div className="xg-pop xg-menu" ref={place} role="menu" aria-label={`${row.row}행`}>
      <button className="btn small" type="button" role="menuitem" disabled={busy || !up} onClick={() => up && onMove(row, up)}>위로</button>
      <button className="btn small" type="button" role="menuitem" disabled={busy || !down} onClick={() => down && onMove(row, down)}>아래로</button>
      {record ? (
        <>
          <button className="btn small" type="button" role="menuitem" disabled={busy || record.result === FAIL} onClick={() => onAction("repair", row)}>리페어 처리</button>
          <button className="btn small danger" type="button" role="menuitem" disabled={busy} onClick={() => onAction("delete", row)}>삭제</button>
        </>
      ) : (
        <span className="xg-norec">앱 기록 없음</span>
      )}
      <button className="btn small" type="button" role="menuitem" onClick={onClose}>취소</button>
    </div>
  );
}

function DateGrid(props: {
  group: DateGroup;
  cols: { col: string; title: string }[];
  unlocked: boolean;
  menuRow: number | null;
  busy: boolean;
  edit: CellEdit | null;
  recordOf: (row: LiveRow) => InspectionRecord | null;
  onRow: (row: number) => void;
  onCell: (cell: LiveCell) => void;
  onEditText: (text: string) => void;
  onSave: () => void;
  onCancelEdit: () => void;
  onMove: (row: LiveRow, neighbor: LiveRow) => void;
  onAction: (action: RowAction, row: LiveRow) => void;
}) {
  const { group, cols, unlocked, menuRow, busy, edit, recordOf, onRow, onCell, onEditText, onSave, onCancelEdit, onMove, onAction } = props;
  return (
    <div className="xg-wrap">
      <table className={`xg${unlocked ? " unlocked" : ""}`}>
        <thead>
          <tr>
            <th className="xg-rn" scope="col" aria-label="행" />
            {cols.map(({ col, title }) => (
              <th key={col} scope="col" className={`xg-h c-${col}`}>{title}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {group.rows.map((row) => {
            const menu = unlocked && menuRow === row.row;
            return (
              <tr key={row.row}>
                <th scope="row" className={`xg-rn${menu ? " open" : ""}`}>
                  {unlocked ? (
                    <button type="button" className="xg-rnb" aria-expanded={menu} aria-label={`${row.row}행 메뉴`} onClick={() => onRow(row.row)}>{row.row}</button>
                  ) : (
                    row.row
                  )}
                  {menu && <RowMenu row={row} {...adjacentRows(group.rows, row)} record={recordOf(row)} busy={busy} onMove={onMove} onAction={onAction} onClose={() => onRow(row.row)} />}
                </th>
                {cols.map(({ col, title }) => {
                  const cell = row.cells.find((c) => c.col === col)!;
                  const text = gridText(cell);
                  const verdict = col === "P" ? verdictOf(row) : "";
                  const cls = [`c-${col}`, typeof cell.value === "number" && !isDateCol(col) ? "num" : "", verdict ? `v-${verdict}` : ""].filter(Boolean);
                  if (!unlocked) return <td key={col} className={cls.join(" ")}>{text}</td>;
                  if (!isEditable(cell)) return <td key={col} className={[...cls, "ro"].join(" ")}>{text}</td>;
                  const editing = edit?.cell.address === cell.address;
                  return (
                    <td key={col} className={[...cls, "ed", editing ? "on" : ""].filter(Boolean).join(" ")}>
                      <button type="button" className="xg-cb" aria-label={`${title} ${cell.address}`} disabled={busy} onClick={() => onCell(cell)}>{text}</button>
                      {editing && <CellEditor edit={edit} busy={busy} onText={onEditText} onSave={onSave} onCancel={onCancelEdit} />}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/**
 * 검사 날짜 하나: 접힌 요약(날짜·행 수·합격/불합격·라인별 포인트 범위). 누르면 그 날짜의 행이 엑셀 같은 표로 펼쳐진다(읽기만).
 * 펼친 날짜의 「수정」을 눌러야 칸을 고치고, 행 번호로 순서 바꾸기·리페어 처리·삭제를 할 수 있다.
 */
export function DateSummaryCard(props: {
  group: DateGroup;
  focus: DateFocus;
  cols: { col: string; title: string }[];
  busy: boolean;
  edit: CellEdit | null;
  recordOf: (row: LiveRow) => InspectionRecord | null;
  staffOf: (date: string) => StaffEdit;
  onDate: (date: string) => void;
  onEditing: () => void;
  onRow: (row: number) => void;
  onCell: (cell: LiveCell) => void;
  onEditText: (text: string) => void;
  onSave: () => void;
  onCancelEdit: () => void;
  onMove: (row: LiveRow, neighbor: LiveRow) => void;
  onAction: (action: RowAction, row: LiveRow) => void;
  onMoveDate: (group: DateGroup, date: string) => void;
  onStaff: (group: DateGroup, staff: StaffEdit) => void;
}) {
  const { group, focus, busy, staffOf, onDate, onEditing, onMoveDate, onStaff } = props;
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
            <span>{unlocked ? "칸을 눌러 고치기 · 행 번호: 순서·리페어·삭제" : "읽기 전용"}</span>
            <button className={`btn small${unlocked ? " primary" : ""}`} type="button" disabled={busy} onClick={onEditing}>
              {unlocked ? "완료" : "수정"}
            </button>
          </div>
          {unlocked && group.date && <DayPanel key={group.date} group={group} staff={staffOf(group.date)} busy={busy} onMoveDate={onMoveDate} onStaff={onStaff} />}
          <DateGrid {...props} unlocked={unlocked} menuRow={unlocked ? focus.row : null} />
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
  /** 같은 앱 ID(Y)가 두 줄 이상(엑셀에서 앱 행을 Y 째로 복사함) — 앱과 엑셀이 어긋나는 원인이라 위에 알린다. */
  const duplicates = useMemo(() => (view.status === "ok" ? duplicateAppRows(view.table.rows) : []), [view]);
  const [edit, setEdit] = useState<CellEdit | null>(null);
  const [undo, setUndo] = useState<MoveSnapshot | null>(null);
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
      setUndo(null);
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
    setFocus((f) => f && { ...f, row: null });
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

  function showEdited(table: LiveTable, rows: Set<number>): void {
    showTable({ ...table, rows: table.rows.map((r) => (rows.has(r.row) ? { ...r, cells: [...r.cells] } : r)) });
  }

  function showRows(table: LiveTable, changed: LiveRow[]): void {
    const byRow = new Map(changed.map((r) => [r.row, r]));
    showTable({ ...table, rows: table.rows.map((r) => byRow.get(r.row) ?? r) });
  }

  function showTable(next: LiveTable): void {
    const current = viewRef.current;
    if (current.status !== "ok") return;
    const wasRefreshing = current.refreshing;
    if (wasRefreshing) {
      loadNo.current += 1;
      clientRef.current?.abort();
      clientRef.current = null;
    }
    const shown = { ...current, table: next, refreshing: false };
    setView(shown);
    viewRef.current = shown;
    if (slot) void saveListCache(slot, next);
    if (wasRefreshing) void load(true);
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
      if (rec && fields) await applyListCellEdits([{ id: rec.id, fields }]);
      setNote(`${cell.address} 저장됨${rec && fields ? ` · 입력 탭 카드도 고침${fields.date ? ` (${fields.date})` : ""}` : ""}`);
      setEdit(null);
      showEdited(table, new Set([row.row]));
      if (cell.col === "E" && cellText(cell)) setFocus({ date: cellText(cell), editing: true, row: null });
    } catch (error) {
      setNote(error instanceof LiveConflictError ? error.message : listErrorText(error));
      if (error instanceof LiveConflictError) void load(true);
    } finally {
      setBusy(false);
    }
  }

  async function move(row: LiveRow, neighbor: LiveRow): Promise<void> {
    if (view.status !== "ok" || busy) return;
    const { client, reason } = liveListClient();
    if (!client) {
      setNote(reason);
      return;
    }
    setEdit(null);
    setUndo(null);
    setBusy(true);
    const table = view.table;
    try {
      const res = await applyRowMove(client, table, site, row, neighbor, (text) => window.confirm(text));
      setNote(res.note);
      if (!res.done) return;
      showRows(table, res.rows);
      setUndo(res.snapshot);
      setFocus((f) => f && { ...f, row: res.snapshot ? null : neighbor.row });
    } catch (error) {
      setNote(error instanceof LiveConflictError || error instanceof LiveMoveRefused ? error.message : listErrorText(error));
      if (error instanceof LiveConflictError) void load(true);
    } finally {
      setBusy(false);
    }
  }

  async function restore(): Promise<void> {
    if (view.status !== "ok" || busy || !undo) return;
    const { client, reason } = liveListClient();
    if (!client) {
      setNote(reason);
      return;
    }
    const snapshot = undo;
    const hi = snapshot.lo + snapshot.rows.length - 1;
    if (!window.confirm(`${snapshot.lo}–${hi}행을 옮기기 전 모습으로 다시 쓸까요?`)) return;
    setUndo(null);
    setBusy(true);
    const table = view.table;
    try {
      const res = await restoreLiveRows(client, table, snapshot);
      showRows(table, res.rows);
      setNote(res.mismatches.length ? `되돌리기 뒤에도 다른 칸: ${res.mismatches.join(", ")}. PC 엑셀에서 ${snapshot.lo}–${hi}행을 확인하세요.` : `${snapshot.lo}–${hi}행 되돌림 · 엑셀 다시 읽어 확인함`);
    } catch (error) {
      setNote(listErrorText(error));
    } finally {
      setBusy(false);
    }
  }

  async function moveDate(group: DateGroup, date: string): Promise<void> {
    if (view.status !== "ok" || busy) return;
    const { client, reason } = liveListClient();
    if (!client) {
      setNote(reason);
      return;
    }
    setEdit(null);
    setBusy(true);
    try {
      const table = view.table;
      const res = await moveDateGroup(client, table, group, date, (text) => window.confirm(text), (done, total) =>
        setNote(`${dotDate(group.date)} → ${dotDate(date)} 바꾸는 중… ${done}/${total}행`),
      );
      setNote(res.note);
      if (!res.done) return;
      showEdited(table, res.rows);
      if (res.rows.size) setFocus({ date, editing: true, row: null });
    } catch (error) {
      setNote(listErrorText(error));
    } finally {
      setBusy(false);
    }
  }

  async function saveStaff(group: DateGroup, staff: StaffEdit): Promise<void> {
    if (view.status !== "ok" || busy) return;
    setEdit(null);
    setBusy(true);
    try {
      const res = await applyDayStaff(view.table, group, staff, (text) => window.confirm(text));
      setNote(res.note);
      if (!res.done) return;
      await settled();
      await syncIdle();
      setNote(staffSentNote(group.date, getState().records.filter((r) => res.ids.includes(r.id))));
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
  const rowMoveRef = useRef(move);
  rowMoveRef.current = move;
  const onMove = useMemo(() => (row: LiveRow, neighbor: LiveRow) => void rowMoveRef.current(row, neighbor), []);
  const moveRef = useRef(moveDate);
  moveRef.current = moveDate;
  const onMoveDate = useMemo(() => (group: DateGroup, date: string) => void moveRef.current(group, date), []);
  const staffRef = useRef(saveStaff);
  staffRef.current = saveStaff;
  const onStaff = useMemo(() => (group: DateGroup, staff: StaffEdit) => void staffRef.current(group, staff), []);
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
        {view.status === "ok" && duplicates.length > 0 && <div className="warn">{duplicateRowsText(view.table.sheetName, duplicates)}</div>}
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
      {undo && (
        <button className="btn block" type="button" disabled={busy} onClick={() => void restore()}>되돌리기 ({undo.lo}–{undo.lo + undo.rows.length - 1}행)</button>
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
                      edit={edit}
                      recordOf={recordOf}
                      staffOf={dayStaffOn}
                      onDate={(date) => refocus((f) => toggleDate(f, date))}
                      onEditing={() => refocus(toggleEditing)}
                      onRow={(row) => refocus((f) => toggleRow(f, row))}
                      onCell={onCell}
                      onEditText={(text) => setEdit((e) => e && { ...e, text })}
                      onSave={() => void save()}
                      onCancelEdit={() => setEdit(null)}
                      onMove={onMove}
                      onAction={onAction}
                      onMoveDate={onMoveDate}
                      onStaff={onStaff}
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
