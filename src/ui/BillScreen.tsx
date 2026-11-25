/**
 * 명세서 탭: 달을 고르면(11월·12월…) 그 달에 한 작업을 현장별로 모아 얼마치인지 보여 준다.
 * 앱 기록만 읽는다(엑셀에 쓰지 않음). 현장을 누르면 날짜·관경별 줄이 펼쳐진다.
 */
import { useMemo, useState } from "react";
import { monthsOf, monthSummary } from "../domain/monthly";
import { useSession } from "../store/session";

const wonText = (n: number) => `${Math.round(n).toLocaleString("ko-KR")}원`;

function monthLabel(ym: string): string {
  const [y, m] = ym.split("-");
  return `${y.slice(2)}년 ${Number(m)}월`;
}

function dayLabel(iso: string): string {
  const [, m, d] = iso.split("-");
  const week = ["일", "월", "화", "수", "목", "금", "토"][new Date(`${iso}T00:00:00Z`).getUTCDay()];
  return `${Number(m)}/${Number(d)} (${week})`;
}

export function BillScreen() {
  const session = useSession();
  const months = useMemo(() => monthsOf(session.records, session.meta.workDate), [session.records, session.meta.workDate]);
  const [month, setMonth] = useState(session.meta.workDate.slice(0, 7));
  const sum = useMemo(() => monthSummary(session.records, session.config, month), [session.records, session.config, month]);

  return (
    <div className="bill">
      <div className="bill-months" role="tablist" aria-label="달 고르기">
        {months.map((ym) => (
          <button key={ym} type="button" role="tab" aria-selected={ym === month} className={ym === month ? "on" : ""} onClick={() => setMonth(ym)}>
            {monthLabel(ym)}
          </button>
        ))}
      </div>

      <div className="card bill-total">
        <div className="k">{monthLabel(month)} 합계 (부가세 포함, 천 원 미만 버림)</div>
        <div className="bill-grand">{wonText(sum.grand)}</div>
        <div className="sub">
          공급가액 {wonText(sum.supply)} · 합격 {sum.count}개소{sum.fail ? ` · 불합격 ${sum.fail}` : ""} · {sum.sites.length}개 현장
        </div>
      </div>

      {sum.sites.length === 0 && <div className="empty-note">{monthLabel(month)}에 한 작업이 없습니다.</div>}

      {sum.sites.map((s) => (
        <details key={s.site} className="card bill-site">
          <summary>
            <span className="name">{s.site}</span>
            <span className="amt">{s.unpriced ? "단가 없음" : wonText(s.grand)}</span>
            <span className="meta">
              {s.count}개소{s.fail ? ` · 불 ${s.fail}` : ""}
              {s.totalM ? ` · ${s.totalM.toFixed(1)}m` : ""} · {s.priceText}
            </span>
          </summary>
          <div className="bill-lines">
            {s.lines.map((l) => (
              <div key={`${l.date}|${l.diameter}`} className="bill-line">
                <span className="d">{dayLabel(l.date)}</span>
                <span className="dia">{l.diameter}</span>
                <span className="n">{l.count}개소{l.qtyM != null ? ` · ${l.qtyM}m` : ""}</span>
                <span className="a">{l.amount != null ? wonText(l.amount) : "—"}</span>
              </div>
            ))}
            {!s.unpriced && (
              <div className="bill-line foot">
                <span className="d">공급가액</span>
                <span className="a">{wonText(s.supply)}</span>
                <span className="d">부가세</span>
                <span className="a">{wonText(s.vat)}</span>
              </div>
            )}
            {s.missing.length > 0 && <div className="warn">금액을 셀 수 없는 관경: {s.missing.join(", ")} (설정에서 길이·단가 확인)</div>}
          </div>
        </details>
      ))}
      <p className="nav-hint">앱에 입력한 기록으로 센 금액입니다. 엑셀 명세서에 사람이 따로 적은 줄은 들어 있지 않을 수 있어요.</p>
    </div>
  );
}
