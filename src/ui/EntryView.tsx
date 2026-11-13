import { useEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from "react";
import { buildBill, recordsForMonth } from "../domain/billing";
import { DayHistory } from "./DayHistory";
import { defectStats, latestLine, lineSummary, linesByLatestDate, shortDate, sortLineRows } from "../domain/history";
import {
  canonicalPoint,
  cycleCard,
  fmtFullPoint,
  directionFromTrail,
  nextDraftPoint,
  orderPointCards,
  parsePoint,
  reinspectLabel,
  type CardPhase,
  type WalkDirection,
} from "../domain/point";
import { isRemovedSite, mappingFor } from "../domain/mapping";
import { FAIL, PASS, type DraftCard, type InspectionRecord } from "../domain/types";
import {
  addInspector,
  changeRecordDate,
  clearGreyReinspect,
  commitDay,
  ensurePairR,
  mirrorDropGrey,
  mirrorGreyDraft,
  mirrorRename,
  noteDirectR,
  setSrMirrorEnabled,
  srMirrorEnabled,
  getState,
  inspectorsOnDate,
  markNoWeld,
  removeRecord,
  renamePoint,
  resetUnsaved,
  resettableCount,
  saveAbsent,
  savePoint,
  addSite,
  selectSite,
  setLineDrafts,
  setResult,
  updateMeta,
  useSession,
} from "../store/session";
import { useDrill } from "./drill";
import { LongPress, PaintGesture, paintAction, type PaintCardInfo, type PaintHit } from "./paint";
import { LinePointList } from "./TodayScreen";

const TAP_WAIT = 280;
const HISTORY_SORT_KEY = "paut-history-sort";
/** 초기화 버튼 확인 문구 */
export const RESET_CONFIRM = "저장 안 한 표시를 되돌릴까요?";

function readHistorySort(): "recent" | "point" {
  try {
    return localStorage.getItem(HISTORY_SORT_KEY) === "point" ? "point" : "recent";
  } catch {
    return "recent";
  }
}
const INSPECTORS = ["조우현", "심명보", "이주호", "방현민", "이연행"];
const TIME_STEPS = Array.from({ length: 33 }, (_, index) => {
  const minutes = 6 * 60 + index * 30;
  return `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
});
type DaySheet = "day" | "insp1" | "insp2" | "start" | "end";

export function EntryView() {
  const session = useSession();
  const { meta, records, config } = session;
  const site = meta.lastSite;
  const [prefix, setPrefix] = useState(meta.prefix || "S");
  const [query, setQuery] = useState("");
  const [addingLine, setAddingLine] = useState(false);
  const [lineDraft, setLineDraft] = useState("");
  const [siteOpen, setSiteOpen] = useState(false);
  const [newSite, setNewSite] = useState("");
  const [lineOpen, setLineOpen] = useState(false);
  const [daySheet, setDaySheet] = useState<DaySheet | null>(null);
  const [customName, setCustomName] = useState("");
  const [customTime, setCustomTime] = useState("");
  const [edit, setEdit] = useState<{ kind: "record" | "draft"; id: string } | null>(null);
  const [hint, setHint] = useState<EditHint | null>(null);
  /** 기록 카드를 길게 눌렀을 때 여는 날짜 고치기 창 */
  const [dateEdit, setDateEdit] = useState<{ id: string; point: string; date: string } | null>(null);
  const [dateDraft, setDateDraft] = useState("");
  const [rename, setRename] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [listOpen, setListOpen] = useState(false);
  useDrill(siteOpen, () => setSiteOpen(false));
  useDrill(lineOpen, () => setLineOpen(false));
  useDrill(daySheet != null, () => setDaySheet(null));
  useDrill(edit != null, () => setEdit(null));
  useDrill(listOpen, () => setListOpen(false));
  const [historySort, setHistorySort] = useState<"recent" | "point">(readHistorySort);
  const [srMirror, setSrMirror] = useState(srMirrorEnabled);
  const queue = useRef(Promise.resolve());
  const pending = useRef(new Map<string, { steps: number; timer: number }>());
  const paint = useRef(new PaintGesture());
  const swallowClick = useRef(false);
  const paintToken = useRef(0);
  /** 드래그 중 미리보기(카드 key → 시작 카드인지). 저장된 것이 아니다. */
  const [preview, setPreview] = useState<Map<string, boolean>>(() => new Map());
  /** 손을 뗀 뒤 저장이 끝날 때까지 칠한 모양 유지. */
  const [committing, setCommitting] = useState<Map<string, boolean>>(() => new Map());
  const grid = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const node = grid.current;
    if (!node) return;
    // iOS Safari: 칠하는 동안에는 페이지가 움직이지 않게. React 의 onTouchMove 는 passive 라 직접 붙인다.
    const block = (event: TouchEvent) => {
      if (paint.current.painting && event.cancelable) event.preventDefault();
    };
    node.addEventListener("touchmove", block, { passive: false });
    return () => node.removeEventListener("touchmove", block);
  }, [meta.activeLine]);

  useEffect(() => {
    if (!hint) return;
    const dismiss = (event: Event) => {
      const target = event.target as HTMLElement | null;
      if (target?.closest(".edit-pill")) return;
      setHint(null);
    };
    const onScroll = () => setHint(null);
    document.addEventListener("pointerdown", dismiss, true);
    window.addEventListener("scroll", onScroll, true);
    return () => {
      document.removeEventListener("pointerdown", dismiss, true);
      window.removeEventListener("scroll", onScroll, true);
    };
  }, [hint]);

  useEffect(() => {
    setHint(null);
  }, [prefix, meta.activeLine, site]);

  const sites = useMemo(() => {
    const names = config.mappings.map((item) => item.site);
    for (const rec of records) if (rec.site && !names.includes(rec.site) && !isRemovedSite(rec.site)) names.push(rec.site);
    return names;
  }, [config.mappings, records]);
  const siteRecords = useMemo(() => records.filter((rec) => rec.site === site && !rec.deleted), [records, site]);
  const lineRecords = useMemo(
    () => siteRecords.filter((rec) => rec.line.trim().toUpperCase() === meta.activeLine.trim().toUpperCase()),
    [siteRecords, meta.activeLine],
  );
  const draftKey = `${site}||${meta.activeLine}||${prefix}`;
  const drafts = meta.draftCards?.[draftKey] ?? [];
  const prefixRecords = useMemo(
    () => lineRecords.filter((rec) => parsePoint(rec.point)?.prefix === prefix).sort(byPointThenTime),
    [lineRecords, prefix],
  );
  const trail = useMemo(() => walkTrail(prefixRecords), [prefixRecords]);
  const direction: WalkDirection = directionFromTrail(trail) ?? "up";
  const cards = useMemo(() => layoutCards(prefixRecords, drafts, "up"), [prefixRecords, drafts]);
  const unsaved = useMemo(() => {
    const locals = records.filter((rec) => rec.sync === "local" && rec.site === site && rec.date === meta.workDate).length;
    const grey = Object.entries(meta.draftCards ?? {}).filter(([key]) => key.startsWith(`${site}||`)).reduce((sum, [, list]) => sum + list.length, 0);
    return locals + grey;
  }, [records, site, meta.workDate, meta.draftCards]);
  // 초기화: 마지막 저장 뒤 바꾼 카드 표시가 있을 때만 켠다.
  const resettable = useMemo(() => resettableCount(records, meta), [records, meta]);
  const historyRows = useMemo(() => historyOfLine(lineRecords, historySort, "up"), [lineRecords, historySort]);
  const defects = useMemo(() => defectStats(siteRecords), [siteRecords]);
  const recent = useMemo(() => linesByLatestDate(siteRecords), [siteRecords]);
  const needle = query.trim().toUpperCase();
  const lineChoices = useMemo(() => {
    const filtered = needle ? recent.filter((row) => row.line.toUpperCase().includes(needle)) : recent;
    return sortLineRows(filtered, "name");
  }, [recent, needle]);

  function chooseHistorySort(mode: "recent" | "point") {
    setHistorySort(mode);
    try { localStorage.setItem(HISTORY_SORT_KEY, mode); } catch { /* 저장 불가 */ }
  }

  const summary = meta.activeLine ? lineSummary(siteRecords, meta.activeLine) : "";
  const today = siteRecords.filter((rec) => rec.date === meta.workDate && rec.result !== "없음");
  const pass = today.filter((rec) => rec.result === PASS).length;
  const fail = today.filter((rec) => rec.result === FAIL).length;
  const mapping = mappingFor(config.mappings, site);
  const monthBill = mapping.unitPrice != null ? buildBill(recordsForMonth(siteRecords, meta.workDate.slice(0, 7)), mapping.unitPrice, config.weldLengths) : null;

  useEffect(() => {
    if (!(edit || lineOpen || daySheet || siteOpen || listOpen || dateEdit)) return;
    const active = document.activeElement;
    if (active instanceof HTMLElement) active.blur();
  }, [edit, lineOpen, daySheet, siteOpen, listOpen, dateEdit]);

  useEffect(() => {
    const onLeave = (event: BeforeUnloadEvent) => {
      if (unsaved > 0) {
        event.preventDefault();
        event.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", onLeave);
    return () => window.removeEventListener("beforeunload", onLeave);
  }, [unsaved]);

  useEffect(() => {
    if (!site || meta.activeLine) return;
    const latest = latestLine(siteRecords);
    if (latest) void updateMeta({ activeLine: latest });
  }, [site, meta.activeLine, siteRecords]);

  function run(task: () => Promise<void>): Promise<void> {
    queue.current = queue.current.then(task, task);
    return queue.current;
  }

  function tapCard(key: string): void {
    try { navigator.vibrate?.(12); } catch { /* 햅틱이 없는 기기 */ }
    const slot = pending.current.get(key) ?? { steps: 0, timer: 0 };
    slot.steps += 1;
    window.clearTimeout(slot.timer);
    slot.timer = window.setTimeout(() => {
      const steps = slot.steps % 3;
      slot.steps = 0;
      pending.current.delete(key);
      if (!steps) return;
      run(() => applyTap(key, steps));
    }, TAP_WAIT);
    pending.current.set(key, slot);
  }

  /** 미리보기 갱신. 저장은 하지 않는다(손을 뗄 때만 저장). */
  function showPreview(): void {
    setPreview(new Map(paint.current.path.map((hit) => [hit.key, hit.fromStart])));
  }

  /** 손을 뗀 순간 남아 있는 길만 실제로 저장한다. 되돌린 카드는 저장도, 엑셀 대기도 생기지 않는다. */
  function commitPaint(hits: PaintHit[]): void {
    const token = ++paintToken.current;
    const saving = new Map<string, boolean>();
    let last: Promise<void> | null = null;
    for (const hit of hits) {
      const slot = pending.current.get(hit.key);
      if (slot) {
        // 칠한 카드에 남아 있던 탭 대기는 버린다. 아니면 칠한 뒤에 탭이 다시 결과를 돌린다.
        window.clearTimeout(slot.timer);
        pending.current.delete(hit.key);
      }
      if (paintAction(paintInfo(hit.key), hit.fromStart) === "skip") {
        // 이미 합격인 S 위를 지나감: S 는 그대로, 빠진 짝 R 만 만든다(applyPaint 안에서).
        if (hit.key.startsWith("rec:")) last = run(() => applyPaint(hit.key, hit.fromStart));
        continue;
      }
      saving.set(hit.key, hit.fromStart);
      last = run(() => applyPaint(hit.key, hit.fromStart));
    }
    // 저장이 끝날 때까지 칠한 모양을 유지해서 깜빡이지 않게 한다.
    setCommitting(saving);
    setPreview(new Map());
    const clear = () => { if (paintToken.current === token) setCommitting(new Map()); };
    if (last) void last.then(clear, clear);
    else clear();
  }

  function onGridPointerDown(event: ReactPointerEvent<HTMLDivElement>): void {
    swallowClick.current = false;
    if (event.pointerType === "mouse" && event.button !== 0) return;
    const target = event.target as HTMLElement | null;
    if (target?.closest(".pcard.add")) {
      cancelPaint();
      return;
    }
    paint.current.down(event.clientX, event.clientY, paintKeyAt(event.clientX, event.clientY));
  }

  function onGridPointerMove(event: ReactPointerEvent<HTMLDivElement>): void {
    if (paint.current.phase === "idle") return;
    if (event.pointerType === "mouse" && event.buttons === 0) {
      onGridPointerEnd();
      return;
    }
    if (paint.current.move(event.clientX, event.clientY, paintKeyAt(event.clientX, event.clientY))) {
      try { navigator.vibrate?.(8); } catch { /* 햅틱이 없는 기기 */ }
      showPreview();
    }
  }

  function onGridPointerEnd(): void {
    if (paint.current.phase === "idle") return;
    const release = paint.current.up();
    if (!release.swallowClick) return;
    swallowClick.current = true;
    commitPaint(release.commit);
  }

  /** 스크롤로 넘어감·길게 누르기: 미리보기만 지우고 아무것도 저장하지 않는다. */
  function cancelPaint(): void {
    paint.current.cancel();
    setPreview((prev) => (prev.size ? new Map() : prev));
  }

  function paintedLook(key: string, info: PaintCardInfo): boolean {
    const fromStart = preview.get(key) ?? committing.get(key);
    return fromStart !== undefined && paintAction(info, fromStart) === "pass";
  }

  function toggleSrMirror() {
    const next = !srMirror;
    setSrMirror(next);
    setSrMirrorEnabled(next);
  }

  function pointNow() {
    return editingRecord?.point ?? editingDraft?.point ?? rename;
  }

  async function applyNumber(target: { kind: "record" | "draft"; id: string }, next: string | null) {
    if (!next) return;
    setRename(next);
    await renameEdit(target, next);
  }

  function addUpcoming(): void {
    try { navigator.vibrate?.(12); } catch { /* 햅틱이 없는 기기 */ }
    run(async () => {
      const live = getState();
      const key = `${live.meta.lastSite}||${live.meta.activeLine}||${prefix}`;
      const list = live.meta.draftCards?.[key] ?? [];
      const plain = list.filter((card) => !card.reinspectOf);
      const occupied = occupiedMains(prefixRecordsOf(live.records, live.meta, prefix), list);
      const tail = plain.length ? parsePoint(plain[plain.length - 1].point)?.num ?? null : trail.at(-1) ?? null;
      const label = nextDraftPoint({ direction, prefix, occupied, lastEntered: tail });
      await setLineDrafts(key, [...list, { id: crypto.randomUUID(), point: label }]);
      // 회색 R 카드를 더하는 것은 R 결과를 고친 것이 아니다 → S→R 자동을 막지 않는다(예전: R 탭 + 로 만든 R29~R31 이 S 를 영영 안 따라감).
      if (prefix === "S") await mirrorGreyDraft(label);
    });
  }

  const editingRecord = edit?.kind === "record" ? lineRecords.find((rec) => rec.id === edit.id) ?? null : null;
  const editingDraft = edit?.kind === "draft" ? drafts.find((card) => card.id === edit.id) ?? findDraft(meta.draftCards, edit?.id ?? "") : null;

  const crew = `${meta.insp1 || "검사자"}${meta.insp2 ? ` / ${meta.insp2}` : " · 혼자"} · ${timeSpan(meta.timeStart, meta.timeEnd)}`;

  return (
    <div className="entry">
      <div className="meta-stack">
      <details className="meta-row">
        <summary draggable={false}>
          <span className="meta-copy">
            <span className="k">날짜</span>
            <strong>{shortWeek(meta.workDate)} · {meta.dayType}</strong>
            <small>{crew}</small>
          </span>
          <span className="chev" aria-hidden="true">›</span>
        </summary>
        <div className="day-compact">
          <label className="mini">
            <span>날짜</span>
            <input className="field-input" type="date" value={meta.workDate} onChange={(e) => void updateMeta({ workDate: e.target.value })} />
          </label>
          <button className="mini" type="button" onClick={() => setDaySheet("day")}>
            <span>구분</span>
            <strong>{meta.dayType}</strong>
          </button>
          <button className="mini" type="button" onClick={() => { setCustomName(""); setDaySheet("insp1"); }}>
            <span>검사자 1</span>
            <strong>{meta.insp1 || "선택"}</strong>
          </button>
          <button className="mini" type="button" onClick={() => { setCustomName(""); setDaySheet("insp2"); }}>
            <span>검사자 2</span>
            <strong>{meta.insp2 || "없음(혼자)"}</strong>
          </button>
          <button className="mini" type="button" onClick={() => { setCustomTime(meta.timeStart); setDaySheet("start"); }}>
            <span>시작</span>
            <strong>{meta.timeStart || "— (미정)"}</strong>
          </button>
          <button className="mini" type="button" onClick={() => { setCustomTime(meta.timeEnd); setDaySheet("end"); }}>
            <span>종료</span>
            <strong>{meta.timeEnd || "— (미정)"}</strong>
          </button>
          <button className="mini wide" type="button" onClick={() => void updateMeta({ timeStart: "", timeEnd: "" })}>시간 모름</button>
        </div>
      </details>
      <button className="meta-row" type="button" draggable={false} onClick={() => setSiteOpen(true)}>
        <span className="meta-copy">
          <span className="k">현장</span>
          <strong>{site || "현장 선택"}</strong>
        </span>
        <span className="chev" aria-hidden="true">›</span>
      </button>
      <button className="meta-row line" type="button" draggable={false} onClick={() => { setQuery(""); setAddingLine(false); setLineDraft(""); setLineOpen(true); }}>
        <span className="meta-copy">
          <span className="k">라인</span>
          <strong>{meta.activeLine || "라인 선택"}</strong>
          {summary && <small><SummaryText text={summary} /></small>}
        </span>
        <span className="chev" aria-hidden="true">›</span>
      </button>
      </div>
      {meta.activeLine && (
        <>
          <div className="tool-row">
            <div className="segment">
              {["S", "R", "P"].map((item) => (
                <button key={item} type="button" className={prefix === item ? "on" : ""} onClick={() => { setPrefix(item); void updateMeta({ prefix: item }); }}>{item}</button>
              ))}
            </div>
            <button className={`mirror-toggle${srMirror ? " on" : ""}`} type="button" onClick={toggleSrMirror}>S→R 자동</button>
          </div>
          <div
            className="point-grid"
            ref={grid}
            onPointerDown={onGridPointerDown}
            onPointerMove={onGridPointerMove}
            onPointerUp={onGridPointerEnd}
            onPointerCancel={cancelPaint}
            onClickCapture={(event) => {
              if (!swallowClick.current) return;
              swallowClick.current = false;
              event.stopPropagation();
              event.preventDefault();
            }}
          >
            {cards.map((card) => card.kind === "record" ? (
              <PointCard
                key={card.key}
                label={card.rec.point}
                paintKey={card.key}
                date={card.rec.date ? shortDate(card.rec.date) : ""}
                state={paintedLook(card.key, { kind: "record", result: card.rec.result }) ? "pass" : card.rec.result === FAIL ? "fail" : card.rec.result === "없음" ? "absent" : "pass"}
                caption={paintedLook(card.key, { kind: "record", result: card.rec.result }) ? PASS : card.rec.result === "없음" ? "없음" : card.rec.result}
                re={card.re}
                armed={hint?.kind === "record" && hint.id === card.rec.id}
                onTap={() => { if (card.rec.result !== "없음") tapCard(card.key); }}
                onLong={() => { cancelPaint(); setHint(null); setDateDraft(card.rec.date); setDateEdit({ id: card.rec.id, point: card.rec.point, date: card.rec.date }); }}
              />
            ) : (
              <PointCard
                key={card.key}
                label={card.draft.point}
                paintKey={card.key}
                state={paintedLook(card.key, { kind: "draft", reinspect: !!card.draft.reinspectOf }) ? "pass" : "draft"}
                caption={paintedLook(card.key, { kind: "draft", reinspect: !!card.draft.reinspectOf }) ? PASS : card.draft.reinspectOf ? "재검사" : "미입력"}
                re={!!card.draft.reinspectOf}
                armed={hint?.kind === "draft" && hint.id === card.draft.id}
                onTap={() => tapCard(card.key)}
                onLong={(box) => { cancelPaint(); setHint({ kind: "draft", id: card.draft.id, point: card.draft.point, ...box }); }}
              />
            ))}
            <button className="pcard add" type="button" aria-label="다음 번호" onClick={addUpcoming}>+</button>
          </div>
          <p className="defect-line">
            <span className="scope">현장 전체</span>
            총 포인트 {defects.pass} · 불합격 {defects.fail > 0 ? <b>{defects.fail}</b> : defects.fail} · 불량률 {defects.fail > 0 && defects.pass > 0 ? <b>{defects.rate}</b> : defects.rate}
          </p>
          <div className="history">
            <div className="history-head">
              <div className="history-title">
                <span>검사 기록</span>
                <strong>{meta.activeLine}</strong>
              </div>
              <div className="segment line-sort" role="group" aria-label="기록 정렬">
                <button type="button" className={historySort === "recent" ? "on" : ""} onClick={() => chooseHistorySort("recent")}>최신순</button>
                <button type="button" className={historySort === "point" ? "on" : ""} onClick={() => chooseHistorySort("point")}>포인트순</button>
              </div>
            </div>
            {historyRows.length === 0 && <div className="empty-note">이 라인의 검사 기록이 없습니다.</div>}
            {historyRows.map((row) => (
              <div key={row.id} className={`history-row${row.result === FAIL ? " fail" : ""}`}>
                <b>{row.point}{row.re ? <em>재</em> : null}</b>
                <span>{shortDate(row.date)}</span>
                <span className={row.result === FAIL ? "bad" : "ok"}>{row.result}</span>
                <span className="who">{inspectorLabel(row.date, meta)}</span>
              </div>
            ))}
          </div>
        </>
      )}

      {/* 날짜별 기록: 위 '검사 기록'(라인별)은 그대로 두고 그 아래에 더한다. 읽기만 한다. */}
      {site && <DayHistory records={siteRecords} site={site} />}

      <div className="bottom-dock">
        <button
          className="reset-main"
          type="button"
          disabled={resettable === 0}
          onClick={() => {
            if (!window.confirm(RESET_CONFIRM)) return;
            cancelPaint();
            void resetUnsaved();
          }}
        >
          초기화
        </button>
        <button className="save-main" type="button" disabled={unsaved === 0} onClick={() => void commitDay()}>
          {unsaved > 0 ? `저장 · 미저장 ${unsaved}` : "저장됨"}
        </button>
      </div>

      <div className="stat-strip">
        오늘 {today.length}개소 · 합격 {pass} · 불합격 {fail}
        {monthBill ? ` · 이번달 ₩${Math.round(monthBill.grand).toLocaleString("ko-KR")}` : ""}
      </div>

      {hint && (
        <button
          className="edit-pill"
          type="button"
          draggable={false}
          style={editPillStyle(hint)}
          onContextMenu={(event) => event.preventDefault()}
          onClick={() => {
            openEdit(hint.kind, hint.id, hint.point);
            setHint(null);
          }}
        >
          편집
        </button>
      )}

      {dateEdit && (
        <div className="modal-backdrop show" onClick={() => setDateEdit(null)}>
          <div className="modal date-modal" onClick={(e) => e.stopPropagation()}>
            <h3>{dateEdit.point} 날짜</h3>
            <input
              className="field-input"
              type="date"
              aria-label="검사 날짜"
              value={dateDraft}
              onChange={(e) => setDateDraft(e.target.value)}
            />
            <div className="date-actions">
              <button className="btn" type="button" onClick={() => setDateEdit(null)}>취소</button>
              <button
                className="btn primary"
                type="button"
                disabled={!/^\d{4}-\d{2}-\d{2}$/.test(dateDraft)}
                onClick={() => {
                  const target = dateEdit;
                  setDateEdit(null);
                  if (dateDraft !== target.date) void run(() => changeRecordDate(target.id, dateDraft));
                }}
              >
                저장
              </button>
            </div>
            <button className="btn small block" type="button" onClick={() => { const target = dateEdit; setDateEdit(null); openEdit("record", target.id, target.point); }}>번호·결과 편집</button>
          </div>
        </div>
      )}

      {edit && (
        <div className="modal-backdrop show" onClick={() => setEdit(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h3>{editingRecord?.point ?? editingDraft?.point ?? "포인트"}</h3>
            <div className="edit-grid">
              <button className="pass" type="button" onClick={() => void run(() => setEditResult(edit, PASS)).then(() => setEdit(null))}>합격</button>
              <button className="fail" type="button" onClick={() => void run(() => setEditResult(edit, FAIL)).then(() => setEdit(null))}>불합격</button>
              <button type="button" onClick={() => void run(() => markEditGrey(edit)).then(() => setEdit(null))}>미입력</button>
              <button type="button" onClick={() => void run(() => markEditAbsent(edit)).then(() => setEdit(null))}>없음</button>
              <button className="wide" type="button" onClick={() => void applyNumber(edit, addSub(pointNow()))}>-1</button>
              <button type="button" onClick={() => void applyNumber(edit, stepMain(pointNow(), -1))}>번호 -1</button>
              <button type="button" onClick={() => void applyNumber(edit, stepMain(pointNow(), 1))}>번호 +1</button>
            </div>
            <input className="field-input" value={rename} aria-label="포인트 번호" placeholder="번호 직접 입력" onChange={(e) => setRename(e.target.value.toUpperCase())} onKeyDown={(e) => {
              if (e.key === "Enter") void run(() => renameEdit(edit, rename)).then(() => setEdit(null));
            }} />
            <button className="btn block" type="button" onClick={() => void run(() => renameEdit(edit, rename)).then(() => setEdit(null))}>이 번호로</button>
            {confirmDelete ? (
              <button className="btn danger block" type="button" onClick={() => void run(() => deleteEdit(edit)).then(() => { setConfirmDelete(false); setEdit(null); })}>정말 삭제</button>
            ) : (
              <button className="btn danger block" type="button" onClick={() => setConfirmDelete(true)}>삭제</button>
            )}
            <button className="btn block" type="button" onClick={() => setEdit(null)}>닫기</button>
          </div>
        </div>
      )}
      {daySheet && (
        <div className="modal-backdrop show" onClick={() => setDaySheet(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h3>{daySheet === "day" ? "구분" : daySheet === "insp1" ? "검사자 1" : daySheet === "insp2" ? "검사자 2" : daySheet === "start" ? "시작" : "종료"}</h3>
            {daySheet === "day" && (
              <div className="sheet-list">
                {(["평일", "공휴일"] as const).map((item) => (
                  <button key={item} className={meta.dayType === item ? "choice on" : "choice"} type="button" onClick={() => { void updateMeta({ dayType: item }); setDaySheet(null); }}>{item}</button>
                ))}
              </div>
            )}
            {(daySheet === "insp1" || daySheet === "insp2") && (
              <>
                <div className="sheet-list">
                  {daySheet === "insp2" && (
                    <button className={!meta.insp2 ? "choice on" : "choice"} type="button" onClick={() => { void updateMeta({ insp2: "" }); setDaySheet(null); }}>없음(혼자)</button>
                  )}
                  {inspectorNames(config.addedInspectors, daySheet === "insp1" ? meta.insp1 : meta.insp2).map((name) => (
                    <button key={name} className={(daySheet === "insp1" ? meta.insp1 : meta.insp2) === name ? "choice on" : "choice"} type="button" onClick={() => { void updateMeta(daySheet === "insp1" ? { insp1: name } : { insp2: name }); setDaySheet(null); }}>{name}</button>
                  ))}
                </div>
                <input className="field-input" value={customName} placeholder="새 이름" aria-label="새 검사자" onChange={(e) => setCustomName(e.target.value)} />
                <button className="btn primary block" type="button" onClick={() => {
                  const name = customName.trim();
                  if (!name) return;
                  void addInspector(name);
                  void updateMeta(daySheet === "insp1" ? { insp1: name } : { insp2: name });
                  setCustomName("");
                  setDaySheet(null);
                }}>이름 추가</button>
              </>
            )}
            {(daySheet === "start" || daySheet === "end") && (
              <>
                <button className="choice" type="button" onClick={() => { void updateMeta({ timeStart: "", timeEnd: "" }); setDaySheet(null); }}>시간 모름</button>
                <button className={(daySheet === "start" ? meta.timeStart : meta.timeEnd) ? "choice" : "choice on"} type="button" onClick={() => { void updateMeta(daySheet === "start" ? { timeStart: "" } : { timeEnd: "" }); setDaySheet(null); }}>— (미정)</button>
                <div className="time-grid">
                  {TIME_STEPS.map((time) => (
                    <button key={time} className={(daySheet === "start" ? meta.timeStart : meta.timeEnd) === time ? "on" : ""} type="button" onClick={() => { void updateMeta(daySheet === "start" ? { timeStart: time } : { timeEnd: time }); setDaySheet(null); }}>{time}</button>
                  ))}
                </div>
                <label className="mini">직접 입력
                  <input className="field-input" type="time" value={customTime} aria-label="직접 입력 시간" onChange={(e) => setCustomTime(e.target.value)} />
                </label>
                <button className="btn primary block" type="button" onClick={() => {
                  if (!customTime) return;
                  void updateMeta(daySheet === "start" ? { timeStart: customTime } : { timeEnd: customTime });
                  setDaySheet(null);
                }}>이 시간으로</button>
              </>
            )}
            <button className="btn block" type="button" onClick={() => setDaySheet(null)}>닫기</button>
          </div>
        </div>
      )}
      {siteOpen && (
        <div className="modal-backdrop show" onClick={() => setSiteOpen(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h3>현장</h3>
            <div className="sheet-list">
              {sites.map((name) => (
                <button key={name} className={name === site ? "choice on" : "choice"} type="button" onClick={() => { void selectSite(name); setSiteOpen(false); }}>{name}</button>
              ))}
            </div>
            <input className="field-input" value={newSite} placeholder="새 현장" aria-label="새 현장" onChange={(e) => setNewSite(e.target.value)} />
            <button className="btn primary block" type="button" onClick={() => { if (!newSite.trim()) return; void addSite(newSite.trim()); setNewSite(""); setSiteOpen(false); }}>현장 추가</button>
            <button className="btn block" type="button" onClick={() => setSiteOpen(false)}>닫기</button>
          </div>
        </div>
      )}
      {lineOpen && (
        <div className="modal-backdrop show" onClick={() => setLineOpen(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h3>라인</h3>
            <div className="sheet-list lines">
              {lineChoices.map((row) => {
                const detail = lineSummary(siteRecords, row.line);
                return (
                  <button key={row.line} className={row.line.toUpperCase() === meta.activeLine.trim().toUpperCase() ? "line-choice on" : "line-choice"} type="button" onClick={() => { void updateMeta({ activeLine: row.line }); setLineOpen(false); }}>
                    <span className="name">{row.line}</span>
                    <span className="meta">{detail ? <SummaryText text={detail} /> : shortDate(row.date)}</span>
                  </button>
                );
              })}
              {lineChoices.length === 0 && <div className="empty-note">이 현장 라인이 없습니다.</div>}
            </div>
            <input className="field-input" value={query} placeholder="라인 검색" aria-label="라인 검색" onChange={(e) => setQuery(e.target.value)} />
            {addingLine ? (
              <>
                <input className="field-input" value={lineDraft} placeholder="새 라인 번호" aria-label="새 라인 번호" onChange={(e) => setLineDraft(e.target.value)} onKeyDown={(e) => {
                  if (e.key === "Enter" && lineDraft.trim()) {
                    void updateMeta({ activeLine: lineDraft.trim() });
                    setLineDraft("");
                    setAddingLine(false);
                    setLineOpen(false);
                  }
                }} />
                <button className="btn primary block" type="button" onClick={() => {
                  if (!lineDraft.trim()) return;
                  void updateMeta({ activeLine: lineDraft.trim() });
                  setLineDraft("");
                  setAddingLine(false);
                  setLineOpen(false);
                }}>추가</button>
              </>
            ) : (
              <button className="btn primary block" type="button" onClick={() => setAddingLine(true)}>+ 라인 추가</button>
            )}
            <button className="btn block" type="button" onClick={() => setLineOpen(false)}>닫기</button>
          </div>
        </div>
      )}
      {listOpen && (
        <div className="modal-backdrop show" onClick={() => setListOpen(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h3>{meta.activeLine}</h3>
            <LinePointList records={siteRecords} line={meta.activeLine} />
            <button className="btn block" type="button" onClick={() => setListOpen(false)}>닫기</button>
          </div>
        </div>
      )}
    </div>
  );

  function openEdit(kind: "record" | "draft", id: string, point: string) {
    setRename(point);
    setConfirmDelete(false);
    setEdit({ kind, id });
  }
}

function timeSpan(start: string, end: string): string {
  if (!start && !end) return "-";
  return `${start || "-"}–${end || "-"}`;
}

function shortWeek(iso: string): string {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  if (!match) return iso;
  const week = ["일", "월", "화", "수", "목", "금", "토"][new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]))).getUTCDay()];
  return `${match[2]}/${match[3]} (${week})`;
}

function inspectorLabel(date: string, meta: { workDate: string; insp1: string; insp2: string }): string {
  if (date === meta.workDate) return [meta.insp1, meta.insp2].filter(Boolean).join(" · ") || "혼자";
  const staff = inspectorsOnDate(date);
  if (!staff) return "";
  return [staff.insp1, staff.insp2].filter(Boolean).join(" · ");
}

function historyOfLine(records: InspectionRecord[], mode: "recent" | "point", direction: WalkDirection) {
  const rows = records
    .filter((rec) => rec.result === PASS || rec.result === FAIL)
    .map((rec) => ({
      id: rec.id,
      point: parsePoint(rec.point)?.suffix ? parsePoint(rec.point)!.raw : rec.point,
      date: rec.date,
      result: rec.result,
      re: !!parentFail(records, rec),
    }));
  if (mode === "point") return orderPointCards(rows, direction);
  return [...rows].sort((a, b) => b.date.localeCompare(a.date) || a.point.localeCompare(b.point, "ko", { numeric: true }));
}

function inspectorNames(added: string[] | undefined, current: string): string[] {
  return [...new Set([...INSPECTORS, ...(added ?? []), current].filter(Boolean))];
}

function SummaryText({ text }: { text: string }) {
  const match = text.match(/^(.*?)( · 불합격 [^·]+)(.*)$/);
  if (!match) return <>{text}</>;
  return <>{match[1]}<span className="fail-bit">{match[2]}</span>{match[3]}</>;
}

type EditHint = {
  kind: "record" | "draft";
  id: string;
  point: string;
  top: number;
  left: number;
  width: number;
  height: number;
};

function editPillStyle(hint: EditHint): { left: number; top: number; transform: string } {
  const center = hint.left + hint.width / 2;
  const left = Math.min(window.innerWidth - 36, Math.max(36, center));
  const above = hint.top > 64;
  return {
    left,
    top: above ? hint.top - 8 : hint.top + hint.height + 8,
    transform: above ? "translate(-50%, -100%)" : "translate(-50%, 0)",
  };
}

function PointCard({ label, paintKey, date, state, caption, re, armed, onTap, onLong }: {
  label: string;
  paintKey: string;
  date?: string;
  state: "pass" | "fail" | "draft" | "absent";
  caption: string;
  re: boolean;
  armed: boolean;
  onTap: () => void;
  onLong: (box: { top: number; left: number; width: number; height: number }) => void;
}) {
  const button = useRef<HTMLButtonElement>(null);
  const longFn = useRef(onLong);
  longFn.current = onLong;
  const press = useRef<LongPress | null>(null);
  press.current ??= new LongPress(() => {
    const node = button.current;
    if (!node) return;
    try { navigator.vibrate?.(12); } catch { /* 햅틱이 없는 기기 */ }
    const box = node.getBoundingClientRect();
    longFn.current({ top: box.top, left: box.left, width: box.width, height: box.height });
  });
  useEffect(() => () => press.current?.cancel(), []);
  return (
    <button
      ref={button}
      type="button"
      className={`pcard ${state}${armed ? " hold" : ""}`}
      data-point={label}
      data-paint-key={paintKey}
      data-state={state}
      aria-label={`${label} ${caption}${date ? ` ${date}` : ""}`}
      draggable={false}
      onDragStart={(e) => e.preventDefault()}
      onPointerDown={(e) => press.current?.down(e.clientX, e.clientY)}
      onPointerMove={(e) => press.current?.move(e.clientX, e.clientY)}
      onPointerUp={() => press.current?.up()}
      onPointerCancel={() => press.current?.cancel()}
      onClick={() => {
        if (press.current?.click() === false) return;
        onTap();
      }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {re && <em className="re">재</em>}
      <b>{label}</b>
      {date ? <small className="when">{date}</small> : null}
    </button>
  );
}

async function applyTap(key: string, steps: number): Promise<void> {
  const live = getState();
  if (key.startsWith("draft:")) {
    const id = key.slice(6);
    const found = findDraftEntry(live.meta.draftCards, id);
    if (!found) return;
    let phase: CardPhase = "draft";
    for (let i = 0; i < steps; i++) phase = cycleCard(phase);
    if (phase === "draft") return;
    const saved = await savePoint(phase === "pass" ? PASS : FAIL, found.card.point);
    if (!saved) return;
    await dropDraft(id);
    if (phase === "fail") await ensureReinspect(saved, found.card.point);
    return;
  }
  const id = key.slice(4);
  const rec = getState().records.find((item) => item.id === id && !item.deleted);
  if (!rec || rec.result === "없음") return;
  let phase: CardPhase = rec.result === FAIL ? "fail" : "pass";
  for (let i = 0; i < steps; i++) phase = cycleCard(phase);
  if (phase === "fail") {
    if (rec.result !== FAIL) await setResult(id, FAIL);
    await ensureReinspect(id, rec.point);
    return;
  }
  if (phase === "pass") {
    await clearGreyReinspect(id);
    if (rec.result !== PASS) await setResult(id, PASS);
    return;
  }
  await clearGreyReinspect(id);
  const parent = parentFail(getState().records, rec);
  const point = canonicalPoint(rec.point);
  await removeRecord(id, true);
  if (parent && point) await ensureReinspect(parent.id, point);
  else if (point) await addPlainDraft(point);
}

/** 손가락 아래의 카드. 터치는 처음 카드에 잡혀 있으므로 target 대신 좌표로 찾는다. */
function paintKeyAt(x: number, y: number): string | null {
  const node = document.elementFromPoint(x, y) as HTMLElement | null;
  return node?.closest<HTMLElement>("[data-paint-key]")?.dataset.paintKey ?? null;
}

function paintInfo(key: string): { kind: "record" | "draft"; result?: string; reinspect?: boolean } {
  const live = getState();
  if (key.startsWith("draft:")) {
    const found = findDraftEntry(live.meta.draftCards, key.slice(6));
    return { kind: "draft", reinspect: !!found?.card.reinspectOf };
  }
  const rec = live.records.find((item) => item.id === key.slice(4) && !item.deleted);
  return { kind: "record", result: rec?.result ?? "없음" };
}

/** 칠하기 한 칸. 실제로 저장하기 직전에 최신 상태로 규칙을 다시 확인한다. S면 R 이 따라간다(savePoint/setResult). */
export async function applyPaint(key: string, fromStart: boolean): Promise<void> {
  if (paintAction(paintInfo(key), fromStart) === "skip") {
    // 이미 합격(또는 지나가기만 한 불합격)인 S: S 는 그대로 두고 빠진 짝 R 만 만든다(S→R 자동).
    if (key.startsWith("rec:")) await ensurePairR(key.slice(4));
    return;
  }
  if (key.startsWith("draft:")) {
    const id = key.slice(6);
    const found = findDraftEntry(getState().meta.draftCards, id);
    if (!found) return;
    const saved = await savePoint(PASS, found.card.point);
    if (saved) await dropDraft(id);
    return;
  }
  const id = key.slice(4);
  await clearGreyReinspect(id);
  await setResult(id, PASS);
}

async function markEditGrey(target: { kind: "record" | "draft"; id: string }) {
  if (target.kind === "draft") return;
  const rec = getState().records.find((item) => item.id === target.id && !item.deleted);
  if (!rec) return;
  await clearGreyReinspect(target.id);
  const parent = parentFail(getState().records, rec);
  const point = canonicalPoint(rec.point);
  await removeRecord(target.id, true);
  if (parent && point) await ensureReinspect(parent.id, point);
  else if (point) await addPlainDraft(point);
}

function stepMain(point: string, delta: number): string | null {
  const parsed = parsePoint(point);
  if (!parsed) return null;
  const num = parsed.num + delta;
  if (num < 1) return null;
  return fmtFullPoint(parsed.prefix, num, parsed.suffix);
}

function addSub(point: string): string | null {
  const parsed = parsePoint(point);
  if (!parsed) return null;
  if (!parsed.suffix) return fmtFullPoint(parsed.prefix, parsed.num, "1");
  if (/^\d+$/.test(parsed.suffix)) return fmtFullPoint(parsed.prefix, parsed.num, String(Number(parsed.suffix) + 1));
  return fmtFullPoint(parsed.prefix, parsed.num, "1");
}

async function renameEdit(edit: { kind: "record" | "draft"; id: string }, raw: string): Promise<void> {
  const label = canonicalPoint(raw);
  if (!label) return;
  if (edit.kind === "record") {
    await renamePoint(edit.id, label);
    return;
  }
  const found = findDraftEntry(getState().meta.draftCards, edit.id);
  if (!found) return;
  const previous = found.card.point;
  await setLineDrafts(found.key, found.list.map((card) => card.id === edit.id ? { ...card, point: label } : card));
  if (parsePoint(previous)?.prefix === "R") {
    noteDirectR(previous);
    noteDirectR(label);
  }
  if (parsePoint(previous)?.prefix === "S") await mirrorRename(previous, label, getState().meta.workDate);
}

async function setEditResult(edit: { kind: "record" | "draft"; id: string }, result: typeof PASS | typeof FAIL): Promise<void> {
  if (edit.kind === "draft") {
    const found = findDraftEntry(getState().meta.draftCards, edit.id);
    if (!found) return;
    const saved = await savePoint(result, found.card.point);
    if (!saved) return;
    await dropDraft(edit.id);
    if (result === FAIL) await ensureReinspect(saved, found.card.point);
    return;
  }
  const rec = getState().records.find((item) => item.id === edit.id && !item.deleted);
  if (!rec) return;
  if (result === PASS) await clearGreyReinspect(edit.id);
  if (rec.result !== result) await setResult(edit.id, result);
  if (result === FAIL) await ensureReinspect(edit.id, rec.point);
}

async function markEditAbsent(edit: { kind: "record" | "draft"; id: string }): Promise<void> {
  if (edit.kind === "draft") {
    const found = findDraftEntry(getState().meta.draftCards, edit.id);
    if (!found) return;
    await saveAbsent(found.card.point);
    await dropDraft(edit.id);
    return;
  }
  await markNoWeld(edit.id);
}

async function deleteEdit(edit: { kind: "record" | "draft"; id: string }): Promise<void> {
  if (edit.kind === "draft") {
    await dropDraft(edit.id);
    return;
  }
  await removeRecord(edit.id);
}

async function ensureReinspect(parentId: string, point: string): Promise<void> {
  const label = reinspectLabel(point);
  const parsed = label ? parsePoint(label) : null;
  if (!label || !parsed) return;
  const meta = getState().meta;
  const key = `${meta.lastSite}||${meta.activeLine}||${parsed.prefix}`;
  const list = meta.draftCards?.[key] ?? [];
  if (list.some((card) => card.reinspectOf === parentId)) return;
  await setLineDrafts(key, [...list, { id: crypto.randomUUID(), point: label, reinspectOf: parentId }]);
}

async function addPlainDraft(point: string): Promise<void> {
  const parsed = parsePoint(point);
  if (!parsed) return;
  const meta = getState().meta;
  const key = `${meta.lastSite}||${meta.activeLine}||${parsed.prefix}`;
  const list = meta.draftCards?.[key] ?? [];
  if (list.some((card) => card.point === point && !card.reinspectOf)) return;
  await setLineDrafts(key, [...list, { id: crypto.randomUUID(), point }]);
  if (parsed.prefix === "R") noteDirectR(point);
  if (parsed.prefix === "S") await mirrorGreyDraft(point);
}

async function dropDraft(id: string): Promise<void> {
  const found = findDraftEntry(getState().meta.draftCards, id);
  if (!found) return;
  const point = found.card.point;
  const parsed = parsePoint(point);
  await setLineDrafts(found.key, found.list.filter((card) => card.id !== id));
  // R 회색 카드를 치우는 것만으로는 R 결과를 고친 것이 아니다(저장한 R 은 savePoint 가 표시한다).
  if (parsed?.prefix === "S" && !found.card.reinspectOf) await mirrorDropGrey(point);
}

function findDraft(all: Record<string, DraftCard[]> | undefined, id: string): DraftCard | null {
  return findDraftEntry(all, id)?.card ?? null;
}

function findDraftEntry(all: Record<string, DraftCard[]> | undefined, id: string): { key: string; list: DraftCard[]; card: DraftCard } | null {
  for (const [key, list] of Object.entries(all ?? {})) {
    const card = list.find((item) => item.id === id);
    if (card) return { key, list, card };
  }
  return null;
}

function parentFail(records: InspectionRecord[], rec: InspectionRecord): InspectionRecord | null {
  return records
    .filter((other) => !other.deleted && other.id !== rec.id && other.site === rec.site && other.line === rec.line && (canonicalPoint(other.point) ?? other.point) === (canonicalPoint(rec.point) ?? rec.point) && other.result === FAIL && other.ts <= rec.ts)
    .sort((a, b) => b.ts - a.ts)[0] ?? null;
}

type Shown =
  | { key: string; kind: "record"; rec: InspectionRecord; re: boolean }
  | { key: string; kind: "draft"; draft: DraftCard };

function layoutCards(records: InspectionRecord[], drafts: DraftCard[], direction: WalkDirection): Shown[] {
  const out: Shown[] = [
    ...records.map((rec) => ({ key: `rec:${rec.id}`, kind: "record" as const, rec, re: !!parentFail(records, rec) })),
    ...drafts.map((draft) => ({ key: `draft:${draft.id}`, kind: "draft" as const, draft, re: !!draft.reinspectOf })),
  ];
  return orderPointCards(out.map((card) => ({ ...card, point: card.kind === "record" ? card.rec.point : card.draft.point })), direction);
}

function walkTrail(records: InspectionRecord[]): number[] {
  const failed = new Set<string>();
  const nums: number[] = [];
  for (const rec of [...records].sort((a, b) => a.ts - b.ts)) {
    const parsed = parsePoint(rec.point);
    if (!parsed || parsed.suffix || rec.result === "없음") continue;
    if (failed.has(rec.point)) continue;
    if (rec.result === FAIL) failed.add(rec.point);
    nums.push(parsed.num);
  }
  return nums;
}

function occupiedMains(records: InspectionRecord[], drafts: DraftCard[]): number[] {
  const nums: number[] = [];
  for (const rec of records) {
    const parsed = parsePoint(rec.point);
    if (parsed && !parsed.suffix) nums.push(parsed.num);
  }
  for (const draft of drafts) {
    if (draft.reinspectOf) continue;
    const parsed = parsePoint(draft.point);
    if (parsed && !parsed.suffix) nums.push(parsed.num);
  }
  return nums;
}

function prefixRecordsOf(records: InspectionRecord[], meta: { lastSite: string; activeLine: string }, prefix: string): InspectionRecord[] {
  return records.filter((rec) => !rec.deleted && rec.site === meta.lastSite && rec.line.trim().toUpperCase() === meta.activeLine.trim().toUpperCase() && parsePoint(rec.point)?.prefix === prefix);
}

function byPointThenTime(a: InspectionRecord, b: InspectionRecord): number {
  const pa = parsePoint(a.point);
  const pb = parsePoint(b.point);
  if (!pa || !pb) return a.point.localeCompare(b.point);
  if (pa.num !== pb.num) return pa.num - pb.num;
  const as = pa.suffix ?? "";
  const bs = pb.suffix ?? "";
  if (as !== bs) {
    if (!as) return -1;
    if (!bs) return 1;
    return as.localeCompare(bs, undefined, { numeric: true });
  }
  return a.ts - b.ts;
}
