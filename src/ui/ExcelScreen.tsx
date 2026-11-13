import { useMemo, useState } from "react";
import { buildBill, countByDiameter, iljiDiameterLabel, passingRecords, recordsForDay, recordsForMonth, won } from "../domain/billing";
import { isRemovedSite, mappingFor } from "../domain/mapping";
import { PASS } from "../domain/types";
import { downloadBook, hasListBook, reimportLocal, requeueToday, useSession } from "../store/session";
import { listBookOf } from "../excel/listBook";
import { LIST_BOOK_LABEL } from "../domain/types";

export function ExcelScreen({ part = "all" }: { part?: "all" | "log" | "bill" }) {
  const session = useSession();
  const [site, setSite] = useState(session.meta.lastSite);
  const [date, setDate] = useState(session.meta.workDate);
  const [month, setMonth] = useState(session.meta.workDate.slice(0, 7));
  const activeSite = site || session.meta.lastSite;
  const names = useMemo(() => {
    const set = new Set(session.config.mappings.map((m) => m.site));
    for (const rec of session.records) if (!isRemovedSite(rec.site)) set.add(rec.site);
    return [...set];
  }, [session.config.mappings, session.records]);

  const siteRecords = session.records.filter((r) => r.site === activeSite && !r.deleted);
  const mapping = mappingFor(session.config.mappings, activeSite);
  const dayCounts = countByDiameter(recordsForDay(siteRecords, date));
  const monthRecs = recordsForMonth(siteRecords, month);
  const bill = mapping.unitPrice != null ? buildBill(monthRecs, mapping.unitPrice, session.config.weldLengths) : null;
  const preview = siteRecords.slice().reverse().slice(0, 12);
  const pending = session.records.filter((r) => !r.deleted && (r.sync === "pending" || r.sync === "error")).length;
  // 현장이 들어가는 리스트 파일(군포·마전 PA 시트 → 군포·마전 리스트).
  const siteBook = listBookOf({ sheetName: siteRecords[0]?.sheetName ?? activeSite, site: activeSite, listBook: siteRecords.find((r) => r.listBook)?.listBook });
  const siteBookReady = hasListBook(siteBook);

  return (
    <div>
      {part === "all" && <p className="nav-hint">불합격은 검사리스트에만 포함. 작업일지·거래명세서 개소는 합격만 (포인트 1 = 개소 1). 하위 번호도 합격이면 1개소입니다.</p>}
      {part === "all" && (<>
      <div className="card">
        <div className="title">{session.config.syncMode === "graph" ? "OneDrive 동기화" : "데모 · 이 브라우저"}</div>
        <div className="v">{session.syncing ? "반영하는 중…" : pending ? `대기 ${pending}건` : "저장 대기 없음"}</div>
        <p className="nav-hint">
          {session.config.syncMode === "graph"
            ? "네 파일(검사리스트 평택·군포마전, 작업일지, 거래명세서) 모두 Graph 통합 문서 API로 바뀐 칸만 씁니다(파일 전체를 덮어쓰지 않음). 파일마다 따로 성공·실패하며, 실패한 파일만 다시 보냅니다. 거래명세서 새 달 시트는 만들지 않습니다. 시트가 없으면 위에 '11월 거래명세서 시트를 만들어주세요 (지성 마전 26.11)'처럼 만들 시트 이름을 알려 주고 기록은 대기합니다. PC 엑셀에서 그 이름으로 시트를 만들면 다음 동기화 때 자동으로 채웁니다."
            : "로그인 없이 첨부된 통합 문서를 이 브라우저에서 고칩니다. 아래 버튼으로 xlsx를 받을 수 있습니다."}
        </p>
        <button className="btn block" type="button" onClick={() => requeueToday()}>
          오늘 기록을 엑셀에 다시 쓰기
        </button>
      </div>

      <div className="card">
        <div className="title">미리보기 · 내보내기</div>
        <div className="k" style={{ marginBottom: 8 }}>대상 현장</div>
        <select className="site-select" value={activeSite} onChange={(e) => setSite(e.target.value)} style={{ marginBottom: 10 }}>
          {names.map((name) => (
            <option key={name} value={name}>{name}</option>
          ))}
        </select>
        <div className="meta-grid">
          <div className="meta-card">
            <div className="k">작업일지 일자</div>
            <input className="field-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ marginTop: 4, padding: 8, minHeight: 40 }} />
          </div>
          <div className="meta-card">
            <div className="k">거래명세서 월</div>
            <input className="field-input" type="month" value={month} onChange={(e) => setMonth(e.target.value)} style={{ marginTop: 4, padding: 8, minHeight: 40 }} />
          </div>
        </div>
        <button className="btn primary block" type="button" disabled={!siteBookReady} onClick={() => downloadBook(siteBook, `비파괴검사_리스트_${activeSite}.xlsx`)}>
          {LIST_BOOK_LABEL[siteBook]} .xlsx
        </button>
        <button className="btn block" type="button" onClick={() => downloadBook("ilji", "작업일지.xlsx")}>
          작업일지 .xlsx
        </button>
        <button className="btn block" type="button" onClick={() => downloadBook("statement", "거래명세서.xlsx")}>
          거래명세서 .xlsx
        </button>
        {(["list", "listGm"] as const).map((book) => (
          <label key={book} className="btn block" style={{ display: "block", textAlign: "center" }}>
            {LIST_BOOK_LABEL[book]} 다시 불러오기
            <input
              className="hidden"
              type="file"
              accept=".xlsx"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                void file.arrayBuffer().then((buf) => reimportLocal(book, new Uint8Array(buf)));
                e.target.value = "";
              }}
            />
          </label>
        ))}
      </div>
      </>)}

      {part === "all" && <div className="section-title">미리보기 · 검사리스트 (전체 기록)</div>}
      {part === "all" && preview.length === 0 && <div className="empty-note">기록 없음</div>}
      {part === "all" && preview.map((rec) => (
        <div className="rec-item" key={rec.id}>
          <div>
            <div className="main">{rec.line} · {rec.point}</div>
            <div className="sub">{rec.date} · {rec.diameter}</div>
          </div>
          <span className={`badge ${rec.result === PASS ? "ok" : "ng"}`}>{rec.result || "—"}</span>
        </div>
      ))}
      {part === "all" && siteRecords.length > 12 && <div className="nav-hint">…외 {siteRecords.length - 12}건 (xlsx에 전부)</div>}

      {part !== "bill" && <div className="section-title">작업일지 · 합격만 개소</div>}
      {part !== "bill" && (Object.keys(dayCounts).length === 0 ? (
        <div className="empty-note">{date} 합격 개소 없음 (불합격만 있으면 0)</div>
      ) : (
        <div className="card">
          <div className="title">{date} · {session.meta.dayType} · {session.meta.insp1}/{session.meta.insp2}</div>
          <div className="main" style={{ fontWeight: 800 }}>{activeSite} → 작업일지 열 「{mapping.iljiColumn || "그 외"}」</div>
          <div className="sub" style={{ marginTop: 6, color: "var(--muted)" }}>
            {Object.entries(dayCounts).map(([d, n]) => `${iljiDiameterLabel(d, mapping.iljiColumn, activeSite)} × ${n}`).join(", ")}
            {" · "}근무 {session.meta.timeStart}~{session.meta.timeEnd}
          </div>
        </div>
      ))}

      {part !== "log" && <div className="section-title">거래명세서 · 합격만</div>}
      {part !== "log" && (!mapping.prefixConfirmed || !mapping.statementPrefix ? (
        <div className="empty-note">접두어가 확정되지 않았습니다. 설정에서 거래명세서 시트 접두어를 입력하고 확인을 켜세요. 그때까지 이 현장은 명세서에 쓰지 않습니다.</div>
      ) : mapping.unitPrice == null ? (
        <div className="empty-note">단가가 없습니다. 설정에서 원/m 를 입력하기 전에는 금액을 쓰지 않습니다. {mapping.priceNote}</div>
      ) : passingRecords(monthRecs).length === 0 || !bill ? (
        <div className="empty-note">{month} 합격 개소 없음</div>
      ) : (
        <div className="card">
          <div className="title">{activeSite} · {month} · 단가 {mapping.unitPrice.toLocaleString("ko-KR")}원/M · 시트 {mapping.statementPrefix}</div>
          {bill.lines.map((line) => (
            <div className="rec-item" key={`${line.date}-${line.diameter}`}>
              <div>
                <div className="main">{line.diameter} ({line.lengthM != null ? `${line.lengthM}M` : "길이 미정"})</div>
                <div className="sub">{line.date} · 개소 {line.count} · {line.qtyM != null ? `${line.qtyM} M` : "—"}</div>
              </div>
              <span className="badge ok">{line.amount != null ? won(line.amount) : "—"}</span>
            </div>
          ))}
          <div className="meta-grid" style={{ marginTop: 8 }}>
            <div className="meta-card"><div className="k">공급가액</div><div className="v">{Math.round(bill.supply).toLocaleString("ko-KR")}</div></div>
            <div className="meta-card"><div className="k">부가세</div><div className="v">{Math.round(bill.vat).toLocaleString("ko-KR")}</div></div>
            <div className="meta-card wide"><div className="k">합계 (천 원 절사)</div><div className="v">{Math.round(bill.grand).toLocaleString("ko-KR")} · {bill.totalM.toFixed(2)}M</div></div>
          </div>
          {bill.missingLengths.length > 0 && <div className="warn">길이표에 없는 관경: {bill.missingLengths.join(", ")}. 설정에서 m를 입력하세요.</div>}
        </div>
      ))}
    </div>
  );
}
