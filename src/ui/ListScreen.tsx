import { useEffect, useMemo, useRef, useState } from "react";
import { isRemovedSite } from "../domain/mapping";
import { LIST_BOOK_LABEL } from "../domain/types";
import { cellText, isEditable, liveTarget, loadLiveList, matchRecord, recordPatch, writeLiveCell, LiveConflictError, type LiveCell, type LiveTable } from "../excel/liveList";
import { applyListCellEdit, getState, graphReadOk, liveListClient, signIn, useSession } from "../store/session";
import { GRAPH_TIMEOUT_TEXT, GraphError, type GraphClient } from "../sync/graph";
import { isLoginError, isLoginFailureText } from "../sync/authState";
import { ABORTED_TEXT, isTimeoutLike, TOKEN_TIMEOUT_TEXT, withTimeout } from "../sync/timeout";

type View = { status: "loading" } | { status: "error"; message: string } | { status: "ok"; table: LiveTable };

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
    if (res.ok) graphReadOk();
    if (res.ok) return res;
    return { ok: false, message: isLoginFailureText(res.message) ? LIST_LOGIN_TEXT : res.message || "검사리스트를 읽지 못했습니다." };
  } catch (error) {
    client.abort();
    return { ok: false, message: listErrorText(error) };
  }
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
  const [view, setView] = useState<View>({ status: "loading" });
  const [edit, setEdit] = useState<{ cell: LiveCell; text: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [note, setNote] = useState("");
  const loadNo = useRef(0);
  const clientRef = useRef<GraphClient | null>(null);
  const expired = session.config.syncMode === "graph" && (session.auth === "expired" || session.auth === "signedOut");
  const target = site ? liveTarget(site, session.records) : null;

  async function load(): Promise<void> {
    const no = ++loadNo.current;
    setEdit(null);
    setView({ status: "loading" });
    const res = await readLiveList(site, { onClient: (c) => (clientRef.current = c) });
    if (no !== loadNo.current) return;
    clientRef.current = null;
    setView(res.ok ? { status: "ok", table: res.table } : { status: "error", message: res.message });
  }

  // 현장·로그인 상태가 바뀔 때만 다시 읽는다(기록이 바뀔 때마다 OneDrive 를 부르지 않는다).
  useEffect(() => {
    setNote("");
    void load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [site, session.account, session.config.syncMode]);

  // 읽는 도중 로그인이 만료되면 '읽는 중…'에 남지 않는다: 요청을 끊고 바로 로그인 안내.
  useEffect(() => {
    if (!expired) return;
    setView((v) => {
      if (v.status !== "loading") return v;
      loadNo.current += 1;
      clientRef.current?.abort();
      clientRef.current = null;
      return { status: "error", message: LIST_LOGIN_TEXT };
    });
  }, [expired]);

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
      setView({ status: "ok", table: { ...table } });
    } catch (error) {
      setNote(error instanceof LiveConflictError ? error.message : listErrorText(error));
      if (error instanceof LiveConflictError) void load();
    } finally {
      setBusy(false);
    }
  }

  return (
    <div>
      <div className="card">
        <div className="k" style={{ marginBottom: 6 }}>현장</div>
        <select className="site-select" value={site} onChange={(e) => setSite(e.target.value)}>
          {sites.map((name) => (
            <option key={name} value={name}>{name}</option>
          ))}
        </select>
        {target && (
          <p className="nav-hint">
            {LIST_BOOK_LABEL[target.book]} · 시트 「{target.sheetName}」 · OneDrive 에서 바로 읽음. 칸을 눌러 고치면 그 칸만 엑셀에 씁니다(시트·행은 만들지 않음).
          </p>
        )}
        <button className="btn block" type="button" disabled={view.status === "loading" || busy} onClick={() => void load()}>
          {view.status === "loading" ? "읽는 중…" : "새로 불러오기"}
        </button>
      </div>

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
      {view.status === "ok" && (
        view.table.rows.length === 0 ? (
          <div className="empty-note">「{view.table.sheetName}」 시트에 행이 없습니다.</div>
        ) : (
          <div className="live-wrap">
            <table className="live-table">
              <thead>
                <tr>
                  <th>행</th>
                  {view.table.header.map((h) => (
                    <th key={h.col} title={h.col}>{h.title || h.col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {view.table.rows.map((row) => (
                  <tr key={row.row}>
                    <th>{row.row}</th>
                    {row.cells.map((cell) => {
                      const editable = isEditable(cell);
                      const on = edit?.cell.address === cell.address;
                      return (
                        <td
                          key={cell.address}
                          className={`${editable ? "" : "ro"}${on ? " on" : ""}`}
                          onClick={() => {
                            if (!editable || busy) return;
                            setNote("");
                            setEdit({ cell, text: cellText(cell) });
                          }}
                        >
                          {cellText(cell)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )
      )}
    </div>
  );
}
