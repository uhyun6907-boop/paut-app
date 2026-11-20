import { useCallback, useEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { iljiInput, iljiText, loadIlji, planIljiEdit, rowStaff, writeIljiCell, LOCK_TEXT, type IljiCell, type IljiEditPlan, type IljiTable } from "../excel/liveIlji";
import { LiveConflictError } from "../excel/liveList";
import { columnIndex } from "../excel/ooxml";
import { adoptIljiDayStaff, getState, graphReadOk, liveListClient, rememberFile, signIn, useSession } from "../store/session";
import { GraphError, type GraphClient } from "../sync/graph";
import { isLoginError } from "../sync/authState";
import { isTimeoutLike } from "../sync/timeout";

export const ILJI_LOGIN_TEXT = "Microsoft 로그인 필요 — 로그인이 만료되어 OneDrive 작업일지를 읽지 못했습니다. 아래 '다시 로그인'을 누르세요.";

function iljiErrorText(error: unknown): string {
  if (isLoginError(error) || (error instanceof GraphError && error.status === 401)) return ILJI_LOGIN_TEXT;
  if (isTimeoutLike(error)) return "OneDrive 응답이 없어(시간 초과) 작업일지를 읽거나 쓰지 못했습니다. 인터넷 연결을 확인하고 '새로 불러오기'를 누르세요.";
  return `작업일지 — ${error instanceof Error ? error.message : String(error)}`;
}

type IljiView = { status: "loading" } | { status: "error"; message: string } | { status: "ok"; table: IljiTable };

export interface IljiEdit {
  cell: IljiCell;
  text: string;
}

export function monthLabel(month: string): string {
  const [y, m] = month.split("-");
  return `${y.slice(2)}년 ${Number(m)}월`;
}

export function defaultMonth(table: IljiTable, workDate: string): string {
  const want = workDate.slice(0, 7);
  return table.months.includes(want) ? want : table.months.at(-1) ?? "";
}

export function columnTitle(table: IljiTable, col: string): string {
  if (col === table.staff.start) return "근무 시작";
  if (col === table.staff.end) return "근무 종료";
  const at = columnIndex(col);
  const top = table.head[0].find((h) => columnIndex(h.col) <= at && at < columnIndex(h.col) + h.colSpan)?.text ?? "";
  const sub = table.head[1].find((h) => h.col === col)?.text ?? "";
  return [top, sub].filter(Boolean).join(" ") || col;
}

export const confirmText = (table: IljiTable, plan: IljiEditPlan) =>
  `${plan.cell.address} (${columnTitle(table, plan.cell.col)})\n${plan.before} → ${plan.after}\n작업일지(${table.file.name})에 바로 씁니다. 바꿀까요?`;

export type SaveResult = { done: false; note: string; conflict?: boolean } | { done: true; note: string; table: IljiTable };

/**
 * 칸 하나 저장: 확인(전 → 후)을 받은 뒤 쓰고 다시 읽어 확인한다. 날짜 행의 구분·검사자·근무시간이면 앱의 그날 값도 엑셀에 맞춘다.
 * 돌려주는 table 은 그 칸만 새 값으로 바꾼 사본.
 */
export async function saveIljiEdit(client: GraphClient, table: IljiTable, edit: IljiEdit, confirm: (text: string) => boolean): Promise<SaveResult> {
  let plan: IljiEditPlan;
  try {
    plan = planIljiEdit(table, edit.cell, edit.text);
  } catch (error) {
    return { done: false, note: error instanceof Error ? error.message : String(error) };
  }
  if (!confirm(confirmText(table, plan))) return { done: false, note: "" };
  let written: Awaited<ReturnType<typeof writeIljiCell>>;
  try {
    written = await writeIljiCell(client, table, plan);
  } catch (error) {
    if (error instanceof LiveConflictError) return { done: false, note: error.message, conflict: true };
    return { done: false, note: iljiErrorText(error) };
  }
  if (written.file.id !== table.file.id) await rememberFile("ilji", written.file);
  const rows = table.rows.map((row) =>
    row.row !== plan.cell.row ? row : { ...row, cells: row.cells.map((c) => (c.address === plan.cell.address ? { ...c, value: written.value } : c)) },
  );
  const next: IljiTable = { ...table, file: written.file, rows };
  const row = rows.find((r) => r.row === plan.cell.row)!;
  const staffCols = ["C", table.staff.insp1, table.staff.insp2, table.staff.start, table.staff.end];
  const synced = !!row.date && staffCols.includes(plan.cell.col);
  if (synced) await adoptIljiDayStaff(row.date, rowStaff(next, row));
  return { done: true, note: `${plan.cell.address} 저장됨 (${plan.before} → ${plan.after}) · 엑셀 다시 읽어 확인함${synced ? " · 앱의 그날 검사자·시간도 맞춤" : ""}`, table: next };
}

function keepInside(el: HTMLDivElement | null): void {
  const wrap = el?.closest(".xg-wrap")?.getBoundingClientRect();
  const anchor = el?.parentElement?.getBoundingClientRect();
  if (!el || !wrap || !anchor) return;
  el.classList.toggle("end", anchor.left + el.offsetWidth > wrap.right && anchor.right - el.offsetWidth >= wrap.left);
  el.classList.toggle("up", anchor.bottom + el.offsetHeight > wrap.bottom && anchor.top - el.offsetHeight >= wrap.top);
  el.querySelector<HTMLElement>("input, select")?.focus({ preventScroll: true });
}

function IljiEditor(props: { title: string; edit: IljiEdit; busy: boolean; onText: (text: string) => void; onSave: () => void; onCancel: () => void }) {
  const { title, edit, busy, onText, onSave, onCancel } = props;
  const place = useCallback(keepInside, []);
  const { kind, address } = edit.cell;
  const label = `${address} ${title} 값`;
  const keys = (e: KeyboardEvent) => {
    if (e.key === "Enter") onSave();
    if (e.key === "Escape") onCancel();
  };
  return (
    <div className="xg-pop" ref={place} role="dialog" aria-label={`${address} 고치기`}>
      <div className="xg-pop-title">{address} · {title}</div>
      {kind === "dayType" ? (
        <select className="field-input" value={edit.text} aria-label={label} onChange={(e) => onText(e.target.value)}>
          {["평일", "공휴일", ""].map((v) => (
            <option key={v} value={v}>{v || "(비움)"}</option>
          ))}
        </select>
      ) : (
        <input
          className="field-input"
          type={kind === "date" ? "date" : kind === "time" ? "time" : "text"}
          inputMode={kind === "count" ? "numeric" : undefined}
          value={edit.text}
          aria-label={label}
          onChange={(e) => onText(e.target.value)}
          onKeyDown={keys}
        />
      )}
      <div className="xg-pop-row">
        <button className="btn small primary" type="button" disabled={busy} onClick={onSave}>{busy ? "저장 중…" : "저장"}</button>
        <button className="btn small" type="button" disabled={busy} onClick={onCancel}>취소</button>
      </div>
    </div>
  );
}

export function IljiGrid(props: {
  table: IljiTable;
  month: string;
  unlocked: boolean;
  busy: boolean;
  edit: IljiEdit | null;
  onCell: (cell: IljiCell) => void;
  onEditText: (text: string) => void;
  onSave: () => void;
  onCancelEdit: () => void;
}) {
  const { table, month, unlocked, busy, edit, onCell, onEditText, onSave, onCancelEdit } = props;
  const rows = table.rows.filter((r) => r.month === month);
  const shown = new Set(rows.map((r) => r.row));
  const visibleSpan = (cell: IljiCell) => {
    let n = 0;
    while (n < (cell.span?.rows ?? 1) && shown.has(cell.row + n)) n++;
    return n;
  };
  return (
    <div className={`xg-wrap${unlocked ? " ilji-edit" : ""}`}>
      <table className={`xg ilji${unlocked ? " unlocked" : ""}`}>
        <thead>
          {table.head.map((cells, i) => (
            <tr key={i}>
              {i === 0 && <th className="xg-rn" scope="col" rowSpan={2} aria-label="행" />}
              {cells.map((h) => (
                <th key={h.col} scope="col" className={`xg-h c-${h.col}`} rowSpan={h.rowSpan} colSpan={h.colSpan}>{h.text}</th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.row} className={row.date ? "day" : "cont"}>
              <th scope="row" className="xg-rn">{row.row}</th>
              {row.cells.map((cell) => {
                if (cell.coveredBy && shown.has(Number(cell.coveredBy.replace(/^[A-Z]+/, "")))) return null;
                const text = iljiText(cell);
                const span = cell.span ? { rowSpan: visibleSpan(cell), colSpan: cell.span.cols } : {};
                const cls = [`c-${cell.col}`, `k-${cell.kind}`, typeof cell.value === "number" && cell.kind !== "date" && cell.kind !== "time" ? "num" : "", cell.lock === "excluded" ? "ex" : ""].filter(Boolean);
                if (!unlocked) return <td key={cell.col} className={cls.join(" ")} {...span}>{text}</td>;
                if (cell.lock) return <td key={cell.col} className={[...cls, "ro"].join(" ")} title={LOCK_TEXT[cell.lock]} {...span}>{text}</td>;
                const editing = edit?.cell.address === cell.address;
                const title = columnTitle(table, cell.col);
                return (
                  <td key={cell.col} className={[...cls, "ed", editing ? "on" : ""].filter(Boolean).join(" ")} {...span}>
                    <button type="button" className="xg-cb" aria-label={`${title} ${cell.address}`} disabled={busy} onClick={() => onCell(cell)}>{text}</button>
                    {editing && <IljiEditor title={title} edit={edit} busy={busy} onText={onEditText} onSave={onSave} onCancel={onCancelEdit} />}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** 두 번째 탭 「근무일지」: OneDrive 작업일지를 달별로 엑셀처럼 보여 주고, 「수정」 중 칸 하나를 고치면 그 칸만 쓴다. */
export function IljiScreen() {
  const session = useSession();
  const [view, setView] = useState<IljiView>({ status: "loading" });
  const [month, setMonth] = useState("");
  const [unlocked, setUnlocked] = useState(false);
  const [edit, setEdit] = useState<IljiEdit | null>(null);
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState("");
  const loadNo = useRef(0);

  async function load(keepMonth: boolean): Promise<void> {
    const no = ++loadNo.current;
    const { client, reason } = liveListClient();
    if (!client) {
      setView({ status: "error", message: reason });
      return;
    }
    if (!keepMonth) setView({ status: "loading" });
    try {
      const res = await loadIlji(client, getState().config);
      if (no !== loadNo.current) return;
      if (!res.ok) {
        setView({ status: "error", message: res.message });
        return;
      }
      graphReadOk();
      if (res.table.file.id !== getState().config.files.ilji?.id) await rememberFile("ilji", res.table.file);
      setView({ status: "ok", table: res.table });
      setMonth((m) => (keepMonth && res.table.months.includes(m) ? m : defaultMonth(res.table, getState().meta.workDate)));
    } catch (error) {
      if (no === loadNo.current) setView({ status: "error", message: iljiErrorText(error) });
    }
  }

  useEffect(() => {
    void load(false);
  }, []);

  async function save(): Promise<void> {
    if (!edit || view.status !== "ok" || busy) return;
    const { client, reason } = liveListClient();
    if (!client) {
      setNote(reason);
      return;
    }
    setBusy(true);
    try {
      const res = await saveIljiEdit(client, view.table, edit, (text) => window.confirm(text));
      if (res.note) setNote(res.note);
      if (res.done) {
        setView({ status: "ok", table: res.table });
        setEdit(null);
      } else if (res.conflict) {
        setEdit(null);
        await load(true);
      }
    } finally {
      setBusy(false);
    }
  }

  const table = view.status === "ok" ? view.table : null;
  const at = table ? table.months.indexOf(month) : -1;
  const dayCount = useMemo(() => (table ? table.rows.filter((r) => r.month === month && r.date).length : 0), [table, month]);

  return (
    <div className="ilji-screen">
      <div className="ilji-bar">
        <button className="btn small" type="button" aria-label="이전 달" disabled={!table || at <= 0} onClick={() => table && setMonth(table.months[at - 1])}>‹</button>
        <select className="chip chip-select ilji-month" aria-label="달" value={month} disabled={!table} onChange={(e) => setMonth(e.target.value)}>
          {(table?.months ?? []).map((m) => (
            <option key={m} value={m}>{monthLabel(m)}</option>
          ))}
        </select>
        <button className="btn small" type="button" aria-label="다음 달" disabled={!table || at < 0 || at >= table.months.length - 1} onClick={() => table && setMonth(table.months[at + 1])}>›</button>
        <button className="btn small" type="button" disabled={view.status === "loading" || busy} onClick={() => void load(true)}>새로 불러오기</button>
        <button
          className={`btn small${unlocked ? " primary" : ""}`}
          type="button"
          disabled={!table || busy}
          onClick={() => {
            setUnlocked(!unlocked);
            setEdit(null);
          }}
        >
          {unlocked ? "완료" : "수정"}
        </button>
      </div>
      {table && (
        <p className="ilji-hint">
          {table.file.name} · 「{table.sheetName.trim()}」 · {dayCount}일 · {unlocked ? "칸을 눌러 고치기 (회색은 고치지 않음)" : "읽기 전용 · 「수정」을 눌러 고치기"}
        </p>
      )}
      {note && <div className="nav-hint" role="status">{note}</div>}
      {view.status === "loading" && <div className="empty-note" role="status">OneDrive 에서 작업일지를 읽는 중…</div>}
      {view.status === "error" && <div className="empty-note list-error" role="alert">{view.message}</div>}
      {view.status === "error" && view.message === ILJI_LOGIN_TEXT && (
        <button className="btn primary block" type="button" disabled={!session.online} onClick={() => void signIn()}>다시 로그인</button>
      )}
      {table && (
        <IljiGrid
          table={table}
          month={month}
          unlocked={unlocked}
          busy={busy}
          edit={edit}
          onCell={(cell) => setEdit({ cell, text: iljiInput(cell) })}
          onEditText={(text) => setEdit((e) => e && { ...e, text })}
          onSave={() => void save()}
          onCancelEdit={() => setEdit(null)}
        />
      )}
      <div className="lc-end" />
    </div>
  );
}
