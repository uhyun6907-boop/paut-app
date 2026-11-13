/**
 * 1번 탭: 카드·'검사 기록' 아래에 더하는 '날짜별 기록'. 날짜를 누르면 그날 검사 내역(도면·개수·판정)과
 * 작업일지의 근무자·근무시간을 보여 준다. 읽기만 한다(엑셀·기록에 쓰지 않음).
 * <details> 를 써서 화면 위에 띄우지 않으므로 두 손가락 확대 상태에서도 그대로 펼치고 읽을 수 있다.
 */
import { useMemo, useState } from "react";
import { dayHistory, NO_WORKLOG_TEXT, workersText, workTimeText } from "../domain/dayHistory";
import type { InspectionRecord } from "../domain/types";
import { workLogByDate } from "../excel/engine";

const PAGE = 30;

function dayLabel(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return iso;
  const week = ["일", "월", "화", "수", "목", "금", "토"][new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]))).getUTCDay()];
  return `${m[1].slice(2)}.${m[2]}.${m[3]} (${week})`;
}

export function DayHistory({ records, site }: { records: InspectionRecord[]; site: string }) {
  const workLog = workLogByDate();
  const days = useMemo(() => dayHistory(records, workLog, { workOnlyDays: true }), [records, workLog]);
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
              <span className="cnt">{d.total ? `${d.total}개소` : "검사 없음"}{d.fail ? <em> 불 {d.fail}</em> : null}</span>
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
            {d.drawings.length === 0 ? (
              <div className="day-none">이 현장 검사 기록 없음</div>
            ) : (
              <div className="day-lines">
                <div className="day-line head"><span>도면</span><span>개수</span><span>판정</span></div>
                {d.drawings.map((g) => (
                  <div key={g.line} className={`day-line${g.fail ? " fail" : ""}`}>
                    <span className="dwg">{g.line}{g.drawingNo ? <small> {g.drawingNo}</small> : null}</span>
                    <span>{g.count}</span>
                    <span>
                      {g.pass ? <i className="ok">합 {g.pass}</i> : null}
                      {g.fail ? <i className="bad"> 불 {g.fail}{g.failPoints.length ? ` (${g.failPoints.join(", ")})` : ""}</i> : null}
                      {g.other ? <i> 기타 {g.other}</i> : null}
                    </span>
                  </div>
                ))}
              </div>
            )}
          </details>
        );
      })}
      {days.length > limit && (
        <button type="button" className="btn block" onClick={() => setLimit(limit + PAGE)}>이전 날짜 더 보기 ({days.length - limit})</button>
      )}
    </section>
  );
}
