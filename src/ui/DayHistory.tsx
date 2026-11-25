/**
 * 1번 탭: 카드·'검사 기록' 아래에 더하는 '날짜별 기록'. 날짜를 누르면 그날 검사 내역(도면·개수·판정)과
 * 작업일지의 근무자·근무시간을 보여 준다. 그날을 펼치고 「수정」을 누르면 날짜·검사자·시간·현장·라인·포인트·판정을 고치거나 지울 수 있다
 * (확인을 받은 뒤 저장 버튼 없이 바로 엑셀에 보냄).
 * 일지 탭·리스트 탭에서 고친 값이 바로 보이게: 기록은 앱 기록 그대로, 근무자·시간은 앱이 그날 값으로 기억한 것(dayStaff —
 * 일지 탭·리스트 탭 「수정」이 맞춰 둠)을 기기의 작업일지 사본보다 먼저 쓴다.
 * <details> 를 써서 화면 위에 띄우지 않으므로 두 손가락 확대 상태에서도 그대로 펼치고 읽을 수 있다.
 */
import { useMemo, useState } from "react";
import { dayHistory, NO_WORKLOG_TEXT, workersText, workTimeText, type WorkLogInfo } from "../domain/dayHistory";
import { FAIL, PASS, type DayStaff, type InspectionRecord } from "../domain/types";
import { workLogByDate } from "../excel/engine";
import { dayStaffOn, editSavedRecords, getState, setDayStaff, useSession, type SavedEdit, type StaffEdit } from "../store/session";

/** 기기의 작업일지 사본에 앱이 기억한 그날 검사자·시간(dayStaff)을 덮어쓴다. 입력은 바꾸지 않는다. */
export function mergeDayStaff(log: Record<string, WorkLogInfo>, dayStaff: Record<string, DayStaff> | undefined): Record<string, WorkLogInfo> {
  if (!dayStaff) return log;
  const out: Record<string, WorkLogInfo> = { ...log };
  for (const [date, st] of Object.entries(dayStaff)) {
    out[date] = {
      dayType: st.dayType || log[date]?.dayType || "",
      workers: [st.insp1, st.insp2].filter(Boolean),
      timeStart: st.timeStart,
      timeEnd: st.timeEnd,
    };
  }
  return out;
}

const PAGE = 30;

function dayLabel(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return iso;
  const week = ["일", "월", "화", "수", "목", "금", "토"][new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))).getUTCDay()];
  return `${m[1].slice(2)}.${m[2]}.${m[3]} (${week})`;
}

interface PointDraft {
  id: string;
  point: string;
  result: string;
  remove: boolean;
  /** 수정 화면에서 새로 더한 포인트(저장하면 새 기록). */
  isNew?: boolean;
}

interface LineDraft {
  line: string;
  newLine: string;
  site: string;
  points: PointDraft[];
}

/** S → R → P → 그 밖, 같은 접두어는 번호순. */
function byPoint(a: InspectionRecord, b: InspectionRecord): number {
  const rank = (p: string) => ({ S: 0, R: 1, P: 2 })[p.charAt(0).toUpperCase() as "S" | "R" | "P"] ?? 3;
  return rank(a.point) - rank(b.point) || a.point.localeCompare(b.point, "en", { numeric: true });
}

/** 그날 한 현장의 기록을 고치는 칸. 저장을 누르면 확인을 받고 바뀐 것만 보낸다. */
function DayEditor({ date, site, recs, sites, onDone }: { date: string; site: string; recs: InspectionRecord[]; sites: string[]; onDone: (note: string) => void }) {
  const staff0 = useMemo(() => dayStaffOn(date), [date]);
  const [newDate, setNewDate] = useState(date);
  const [staff, setStaff] = useState<StaffEdit>(staff0);
  const [busy, setBusy] = useState(false);
  const [lines, setLines] = useState<LineDraft[]>(() => {
    const map = new Map<string, InspectionRecord[]>();
    for (const r of recs) map.set(r.line, [...(map.get(r.line) ?? []), r]);
    return [...map].map(([line, list]) => ({
      line,
      newLine: line,
      site,
      points: [...list].sort(byPoint).map((r) => ({ id: r.id, point: r.point, result: r.result, remove: false })),
    }));
  });
  const original = useMemo(() => new Map(recs.map((r) => [r.id, r])), [recs]);

  function setLine(i: number, partial: Partial<LineDraft>) {
    setLines(lines.map((l, k) => (k === i ? { ...l, ...partial } : l)));
  }
  function setPoint(i: number, j: number, partial: Partial<PointDraft>) {
    setLines(lines.map((l, k) => (k === i ? { ...l, points: l.points.map((p, m) => (m === j ? { ...p, ...partial } : p)) } : l)));
  }

  async function save() {
    const edits: SavedEdit[] = [];
    const notes: string[] = [];
    if (newDate !== date) notes.push(`날짜 ${date.slice(5)} → ${newDate.slice(5)}`);
    for (const l of lines) {
      const lineChanged = l.newLine.trim() && l.newLine.trim() !== l.line;
      if (lineChanged) notes.push(`라인 ${l.line} → ${l.newLine.trim()}`);
      if (l.site !== site) notes.push(`${l.line} 현장 → ${l.site}`);
      for (const p of l.points) {
        if (p.isNew) {
          const point = p.point.trim().toUpperCase();
          if (!point) continue;
          edits.push({ id: p.id, add: { site: l.site, line: l.newLine.trim() || l.line, date: newDate, point, result: p.result } });
          notes.push(`${point} ${p.result} 추가`);
          continue;
        }
        const rec = original.get(p.id);
        if (!rec) continue;
        if (p.remove) {
          edits.push({ id: p.id, delete: true });
          notes.push(`${rec.point} 삭제`);
          continue;
        }
        const e: SavedEdit = { id: p.id };
        if (newDate !== date) e.date = newDate;
        if (lineChanged) e.line = l.newLine.trim();
        if (l.site !== site) e.site = l.site;
        if (p.point.trim().toUpperCase() !== rec.point) {
          e.point = p.point;
          notes.push(`${rec.point} → ${p.point.trim().toUpperCase()}`);
        }
        if (p.result !== rec.result) {
          e.result = p.result;
          notes.push(`${p.point} ${p.result}`);
        }
        if (Object.keys(e).length > 1) edits.push(e);
      }
    }
    const staffChanged = (["insp1", "insp2", "timeStart", "timeEnd"] as const).some((k) => staff[k] !== staff0[k]);
    if (staffChanged) notes.push(`검사자·시간 → ${[staff.insp1, staff.insp2].filter(Boolean).join(" · ") || "없음"} ${staff.timeStart || "-"}–${staff.timeEnd || "-"}`);
    if (!edits.length && !staffChanged) return onDone("바뀐 것이 없습니다.");
    if (!window.confirm(`이렇게 바꿀까요? 바로 엑셀에 보냅니다.\n\n${notes.join("\n")}`)) return;
    setBusy(true);
    try {
      let n = 0;
      let staffNote = "";
      if (edits.length) {
        // 고친 기록이 새 날짜의 검사자·시간을 작업일지까지 실어 간다.
        n = await editSavedRecords(edits, staffChanged ? { date: newDate, staff } : undefined);
        if (staffChanged) staffNote = " · 검사자·시간 함께 보냄";
      } else if (staffChanged) {
        const ids = getState().records.filter((r) => r.date === date && !r.deleted).map((r) => r.id);
        staffNote = (await setDayStaff(date, staff, ids)) ? "검사자·시간 작업일지에 보내는 중" : "검사자·시간은 그날 기록이 없어 보내지 못함";
      }
      onDone(`${n ? `기록 ${n}개 고침 · 엑셀에 보내는 중` : "기록은 그대로"}${staffNote}`);
    } finally {
      setBusy(false);
    }
  }

  const field = (key: keyof StaffEdit, label: string, type: "text" | "time") => (
    <label className="de-f">
      <span>{label}</span>
      <input className="field-input" type={type} value={staff[key]} onChange={(e) => setStaff({ ...staff, [key]: e.target.value })} />
    </label>
  );

  return (
    <div className="day-editor">
      <label className="de-f wide">
        <span>날짜</span>
        <input className="field-input" type="date" value={newDate} onChange={(e) => setNewDate(e.target.value || date)} />
      </label>
      <div className="de-grid">
        {field("insp1", "검사자 1", "text")}
        {field("insp2", "검사자 2", "text")}
        {field("timeStart", "시작", "time")}
        {field("timeEnd", "종료", "time")}
      </div>
      {lines.map((l, i) => (
        <div key={l.line} className="de-line">
          <label className="de-f wide">
            <span>라인</span>
            <input className="field-input" value={l.newLine} onChange={(e) => setLine(i, { newLine: e.target.value })} />
          </label>
          <label className="de-f wide">
            <span>현장</span>
            <select className="field-input" value={l.site} onChange={(e) => setLine(i, { site: e.target.value })}>
              {sites.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </label>
          {l.points.map((p, j) => (
            <div key={p.id} className={`de-pt${p.remove ? " removed" : ""}`}>
              <input className="field-input" value={p.point} aria-label="포인트" placeholder={p.isNew ? "예: S05" : undefined} disabled={p.remove} onChange={(e) => setPoint(i, j, { point: e.target.value })} />
              <select className="field-input" value={p.result} aria-label="판정" disabled={p.remove} onChange={(e) => setPoint(i, j, { result: e.target.value })}>
                <option value={PASS}>합격</option>
                <option value={FAIL}>불합격</option>
                {p.result !== PASS && p.result !== FAIL && <option value={p.result}>{p.result}</option>}
              </select>
              {p.isNew ? (
                <button type="button" className="btn small" onClick={() => setLine(i, { points: l.points.filter((_, m) => m !== j) })}>빼기</button>
              ) : (
                <button type="button" className={`btn small${p.remove ? " danger" : ""}`} onClick={() => setPoint(i, j, { remove: !p.remove })}>
                  {p.remove ? "삭제 취소" : "삭제"}
                </button>
              )}
            </div>
          ))}
          <button
            type="button"
            className="btn small de-add"
            onClick={() => setLine(i, { points: [...l.points, { id: `new-${Date.now()}-${l.points.length}`, point: "", result: FAIL, remove: false, isNew: true }] })}
          >
            + 포인트 추가 (불합격)
          </button>
        </div>
      ))}
      <div className="de-actions">
        <button type="button" className="btn" disabled={busy} onClick={() => onDone("")}>취소</button>
        <button type="button" className="btn primary" disabled={busy} onClick={() => void save()}>{busy ? "보내는 중…" : "저장"}</button>
      </div>
    </div>
  );
}

export function DayHistory({ records, site, sites }: { records: InspectionRecord[]; site: string; sites: string[] }) {
  const dayStaff = useSession().meta.dayStaff;
  const copy = workLogByDate();
  const workLog = useMemo(() => mergeDayStaff(copy, dayStaff), [copy, dayStaff]);
  const days = useMemo(() => dayHistory(records, workLog), [records, workLog]);
  const [limit, setLimit] = useState(PAGE);
  const [editing, setEditing] = useState<string | null>(null);
  const [note, setNote] = useState<{ date: string; text: string } | null>(null);
  return (
    <section className="day-history" aria-label="날짜별 기록">
      <div className="day-history-head">
        <span>날짜별 기록</span>
        <strong>{site}</strong>
      </div>
      {days.length === 0 && <div className="empty-note">날짜별 기록이 없습니다.</div>}
      {days.slice(0, limit).map((d) => {
        const time = workTimeText(d.work);
        const who = workersText(d.work);
        return (
          <details key={d.date} className="day-entry">
            <summary>
              <b>{dayLabel(d.date)}</b>
              <span className="cnt">{`${d.total}개소`}{d.fail ? <em> 불 {d.fail}</em> : null}</span>
              <span className="who">{who || time ? [who, time].filter(Boolean).join(" ") : "근무자·시간 없음"}</span>
            </summary>
            <div className="day-work">
              {d.work ? (
                <>
                  <span>근무자 <b>{who || "-"}</b></span>
                  <span>근무시간 <b>{time || "-"}</b>{d.work.dayType ? ` · ${d.work.dayType}` : ""}</span>
                </>
              ) : (
                <span className="none">{NO_WORKLOG_TEXT}</span>
              )}
            </div>
            <div className="day-lines">
              {d.drawings.map((g) => (
                <div key={g.line} className={`day-line${g.fail ? " fail" : ""}`}>
                  <div className="day-line-head">
                    <span className="dwg">{g.line}{g.drawingNo ? <small> {g.drawingNo}</small> : null}</span>
                    <span className="sum">
                      {g.count}개{g.pass ? <i className="ok"> · 합 {g.pass}</i> : null}
                      {g.fail ? <i className="bad"> · 불 {g.fail}</i> : null}
                      {g.other ? <i> · 기타 {g.other}</i> : null}
                    </span>
                  </div>
                  <div className="day-points" aria-label={`${g.line} 검사 포인트`}>
                    {g.points.map((p, i) => (
                      <span key={`${p.point}-${i}`} className={`day-pt${p.fail ? " bad" : p.result === "합격" ? "" : " other"}`} title={p.result}>
                        {p.point}{p.fail ? " 불" : ""}
                      </span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
            {editing === d.date ? (
              <DayEditor
                key={`${d.date}|${site}`}
                date={d.date}
                site={site}
                recs={records.filter((r) => r.date === d.date && !r.deleted)}
                sites={sites}
                onDone={(text) => {
                  setEditing(null);
                  setNote(text ? { date: d.date, text } : null);
                }}
              />
            ) : (
              <button type="button" className="btn block day-edit-btn" onClick={() => { setEditing(d.date); setNote(null); }}>수정</button>
            )}
            {note?.date === d.date && <div className="empty-note">{note.text}</div>}
          </details>
        );
      })}
      {days.length > limit && (
        <button type="button" className="btn block" onClick={() => setLimit(limit + PAGE)}>이전 날짜 더 보기 ({days.length - limit})</button>
      )}
    </section>
  );
}
