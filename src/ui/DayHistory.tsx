/**
 * 1번 탭: 카드·'검사 기록' 아래에 더하는 '날짜별 기록'. 날짜를 누르면 그날 검사 내역(도면·개수·판정)과
 * 작업일지의 근무자·근무시간을 보여 준다. 읽기만 한다(엑셀·기록에 쓰지 않음).
 * 일지 탭·리스트 탭에서 고친 값이 바로 보이게: 기록은 앱 기록 그대로, 근무자·시간은 앱이 그날 값으로 기억한 것(dayStaff —
 * 일지 탭 칸 수정·리스트 탭 「수정」이 맞춰 둠)을 기기의 작업일지 사본보다 먼저 쓴다.
 * <details> 를 써서 화면 위에 띄우지 않으므로 두 손가락 확대 상태에서도 그대로 펼치고 읽을 수 있다.
 */
import { useMemo, useState } from "react";
import { dayHistory, NO_WORKLOG_TEXT, workersText, workTimeText, type WorkLogInfo } from "../domain/dayHistory";
import type { DayStaff, InspectionRecord } from "../domain/types";
import { workLogByDate } from "../excel/engine";
import { useSession } from "../store/session";

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

export function DayHistory({ records, site }: { records: InspectionRecord[]; site: string }) {
  const dayStaff = useSession().meta.dayStaff;
  const copy = workLogByDate();
  const workLog = useMemo(() => mergeDayStaff(copy, dayStaff), [copy, dayStaff]);
  const days = useMemo(() => dayHistory(records, workLog), [records, workLog]);
  const [limit, setLimit] = useState(PAGE);
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
          </details>
        );
      })}
      {days.length > limit && (
        <button type="button" className="btn block" onClick={() => setLimit(limit + PAGE)}>이전 날짜 더 보기 ({days.length - limit})</button>
      )}
    </section>
  );
}
