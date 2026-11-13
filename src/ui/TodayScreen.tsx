import { useEffect, useMemo, useState } from "react";
import { linePointRows, lineSummary, pointsOnLine, recentLineProgress, shortDate, siteLastDay, type PrefixRange } from "../domain/history";
import type { InspectionRecord } from "../domain/types";
import { expandRange, fmtPoint, nextNumber, parsePoint, splitFailList } from "../domain/point";
import { sizeFromLine } from "../domain/size";
import { PAUT_OD } from "../domain/reference";
import { isRemovedSite } from "../domain/mapping";
import type { ResultKind } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import {
  addSite,
  removeRecord,
  retryRecord,
  saveMany,
  savePoint,
  selectSite,
  setResult,
  syncLabel,
  undoLast,
  updateMeta,
  useSession,
} from "../store/session";

const PARTNER_NAMES = ["심명보", "이주호", "방현민", "이연행", "김정환"];

function partnerNames(fromFile: string[], current: string): string[] {
  const names = [...fromFile, ...PARTNER_NAMES].map((name) => name.trim()).filter((name) => name && name !== "조우현");
  if (current && !names.includes(current)) names.unshift(current);
  return [...new Set(names)];
}

function RangeText({ ranges }: { ranges: PrefixRange[] }) {
  const fails = ranges.flatMap((range) => range.fails);
  return (
    <span className="range-text">
      {ranges.map((range, index) => (
        <span key={range.prefix}>
          {index > 0 ? " · " : ""}
          {range.text}
        </span>
      ))}
      {fails.length > 0 && <span className="fail-bit"> 불합격 {fails.join(", ")}</span>}
    </span>
  );
}

export function LinePointList({ records, line }: { records: InspectionRecord[]; line: string }) {
  const rows = linePointRows(records, line);
  if (!rows.length) return <div className="empty-note">이 라인 기록 없음</div>;
  return (
    <div className="point-log">
      <div className="point-log-sum">{lineSummary(records, line)}</div>
      <div className="point-log-list">
        {rows.map((row) => (
          <div key={`${row.id}-${row.point}-${row.date}`} className={`point-log-row${row.fail ? " fail" : ""}`}>
            <span className="pt">{row.point}</span>
            <span className="dt">{row.dateLabel}</span>
            <span className={`badge ${row.fail ? "ng" : "ok"}`}>{row.result}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function ScanMini({ line }: { line: string }) {
  const info = sizeFromLine(line, PAUT_OD);
  if (!info) {
    return (
      <div className="scan-mini">
        관경·스캔길이는 라인명 끝 <strong>-350A</strong> 형태로 자동 추출
      </div>
    );
  }
  return (
    <div className="scan-mini">
      관경 <strong>{info.label}</strong>
      {info.scan != null ? (
        <>
          {" "}
          · 스캔 <strong>{info.scan} mm</strong>
        </>
      ) : (
        " · 스캔 데이터 없음"
      )}
    </div>
  );
}

export function TodayScreen({ view = "all" }: { view?: "all" | "lines" }) {
  const session = useSession();
  const { meta, records, config } = session;
  const site = meta.lastSite;
  const siteRecords = useMemo(() => records.filter((r) => r.site === site && !r.deleted), [records, site]);
  const names = useMemo(() => {
    const set = new Set<string>(config.mappings.map((m) => m.site));
    for (const rec of records) if (!isRemovedSite(rec.site)) set.add(rec.site);
    return [...set];
  }, [config.mappings, records]);
  const recents = recentLineProgress(siteRecords);
  const allLines = recentLineProgress(siteRecords, 200);
  const lastDay = siteLastDay(siteRecords);
  const todayRows = siteRecords.filter((r) => r.date === meta.workDate).slice().reverse();
  const linePoints = pointsOnLine(siteRecords, meta.activeLine);

  const [lineDraft, setLineDraft] = useState(meta.activeLine);
  const [openLine, setOpenLine] = useState<string | null>(null);
  const typedMatch = recents.find((row) => row.line.trim().toUpperCase() === lineDraft.trim().toUpperCase());
  const [prefix, setPrefix] = useState(meta.prefix || "S");
  const [num, setNum] = useState(1);
  const [draft, setDraft] = useState("S01");
  const [custom, setCustom] = useState("");
  const [showLine, setShowLine] = useState(!meta.activeLine);
  const [recentOpen, setRecentOpen] = useState(false);
  const [siteOpen, setSiteOpen] = useState(false);
  const [siteName, setSiteName] = useState("");
  const [rangeFrom, setRangeFrom] = useState("");
  const [rangeTo, setRangeTo] = useState("");
  const [rangeNg, setRangeNg] = useState("");
  const [editId, setEditId] = useState<string | null>(null);
  const [editResult, setEditResult] = useState<ResultKind>(PASS);
  const editing = siteRecords.find((r) => r.id === editId) ?? null;

  const signature = linePoints.join(",");
  useEffect(() => {
    if (!meta.activeLine) return;
    const n = nextNumber(linePoints, prefix);
    setNum(n);
    setDraft(fmtPoint(prefix, n));
    // 라인·접두어·저장된 포인트가 바뀔 때만 다음 번호로 맞춘다.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.activeLine, prefix, signature]);

  useEffect(() => {
    setLineDraft(meta.activeLine);
    setShowLine(!meta.activeLine);
  }, [meta.activeLine, site]);

  function activate(line: string) {
    const name = line.trim();
    if (!site) return;
    if (!name) return;
    void updateMeta({ activeLine: name });
    setShowLine(false);
    setRecentOpen(false);
  }

  function choosePrefix(next: string) {
    setPrefix(next);
    setCustom("");
    void updateMeta({ prefix: next });
  }

  function setPointNumber(n: number) {
    const next = Math.max(1, n);
    setNum(next);
    setDraft(fmtPoint(prefix, next));
  }

  async function commit(result: ResultKind) {
    await savePoint(result, draft);
  }

  async function applyRange() {
    const rows = expandRange(rangeFrom, rangeTo, splitFailList(rangeNg));
    if (!rows) {
      window.alert("범위 확인: 같은 접두어, 시작이 끝보다 작거나 같아야 합니다.");
      return;
    }
    await saveMany(rows);
    const parsed = parsePoint(rangeTo);
    if (parsed) {
      setPrefix(parsed.prefix);
      void updateMeta({ prefix: parsed.prefix });
    }
    setRangeFrom("");
    setRangeTo("");
    setRangeNg("");
  }

  if (view === "lines") {
    return (
      <div>
        <div className="section-title">라인</div>
        {!site && <div className="empty-note">현장을 선택하면 검사리스트의 라인이 나옵니다.</div>}
        {site && allLines.length === 0 && <div className="empty-note">불러온 라인 없음</div>}
        {allLines.map((line) => (
          <div key={line.line}>
            <button className="line-row" type="button" onClick={() => { setOpenLine(openLine === line.line ? null : line.line); void updateMeta({ activeLine: line.line }); }}>
              <span className="main">{line.line}</span>
              <span className="sub">{lineSummary(siteRecords, line.line) || "기록 없음"}</span>
            </button>
            {openLine === line.line && <LinePointList records={siteRecords} line={line.line} />}
          </div>
        ))}
      </div>
    );
  }

  return (
    <div>
      <div className="card">
        <div className="title">1. 현장 · 날짜 · 근무</div>
        <div className="row">
          <select className="site-select" value={site} onChange={(e) => void selectSite(e.target.value)}>
            <option value="">현장 선택…</option>
            {names.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
          <button className="btn" type="button" onClick={() => setSiteOpen(true)}>
            + 현장
          </button>
        </div>
        <div className="day-bar">
          <div>
            <div className="k">검사일자</div>
            <input className="field-input" type="date" value={meta.workDate} onChange={(e) => void updateMeta({ workDate: e.target.value })} style={{ marginTop: 4 }} />
          </div>
          <div>
            <div className="k">구분</div>
            <select className="site-select" value={meta.dayType} onChange={(e) => void updateMeta({ dayType: e.target.value as "평일" | "공휴일" })} style={{ marginTop: 4 }}>
              <option value="평일">평일</option>
              <option value="공휴일">공휴일</option>
            </select>
          </div>
        </div>
        <div className="meta-grid">
          <div className="meta-card">
            <div className="k">검사자 1</div>
            <input className="field-input" value={meta.insp1} onChange={(e) => void updateMeta({ insp1: e.target.value })} style={{ marginTop: 4, padding: 8, minHeight: 40, borderWidth: 1 }} />
          </div>
          <div className="meta-card">
            <div className="k">검사자 2</div>
            <select className="site-select" value={meta.insp2} onChange={(e) => void updateMeta({ insp2: e.target.value })} style={{ marginTop: 4 }}>
              <option value="">없음(혼자)</option>
              {partnerNames(config.inspectors, meta.insp2).map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
          <div className="meta-card">
            <div className="k">근무 시작</div>
            <input className="field-input" type="time" value={meta.timeStart} onChange={(e) => void updateMeta({ timeStart: e.target.value })} style={{ marginTop: 4, padding: 8, minHeight: 40, borderWidth: 1 }} />
          </div>
          <div className="meta-card">
            <div className="k">근무 종료</div>
            <input className="field-input" type="time" value={meta.timeEnd} onChange={(e) => void updateMeta({ timeEnd: e.target.value })} style={{ marginTop: 4, padding: 8, minHeight: 40, borderWidth: 1 }} />
          </div>
        </div>
        {lastDay && (
          <div className="site-last">
            <div className="k">지난 작업 {shortDate(lastDay.date)}</div>
            {lastDay.lines.slice(0, 3).map((line) => (
              <div key={line.line} className="site-last-line">
                <span>{line.line}</span> <RangeText ranges={line.ranges} />
              </div>
            ))}
            {lastDay.lines.length > 3 && <div className="k">외 {lastDay.lines.length - 3}라인</div>}
          </div>
        )}
      </div>

      {showLine && (
        <div className="card">
          <div className="title">2. 배관라인 (line No.)</div>
          <input
            className="field-input line-big"
            value={lineDraft}
            placeholder="예: HJ-M-002-350A"
            autoComplete="off"
            onChange={(e) => setLineDraft(e.target.value)}
          />
          <ScanMini line={lineDraft} />
          {typedMatch && <LinePointList records={siteRecords} line={typedMatch.line} />}
          <div className="row" style={{ marginTop: 10, marginBottom: 0 }}>
            <button className="btn primary" type="button" onClick={() => activate(lineDraft)}>
              이 라인으로
            </button>
            <button className="btn" type="button" onClick={() => setRecentOpen(true)}>
              최근 라인
            </button>
          </div>
          <div style={{ marginTop: 10 }}>
            {recents.length === 0 && <div className="nav-hint">최근 라인 없음</div>}
            {recents.slice(0, 8).map((line) => (
              <button key={line.line} className="line-row" type="button" onClick={() => activate(line.line)}>
                <span className="main">{line.line}</span>
                <span className="sub">
                  <RangeText ranges={line.ranges} />
                </span>
              </button>
            ))}
          </div>
        </div>
      )}

      {!showLine && meta.activeLine && (
        <div className="card hero">
          <div className="title">3. 다음 포인트 · 누르면 즉시 저장</div>
          <div className="row" style={{ marginBottom: 8 }}>
            <button className="btn" type="button" style={{ flex: 1, textAlign: "left", fontWeight: 900 }} onClick={() => setShowLine(true)}>
              {meta.activeLine}
            </button>
            <button className="btn" type="button" onClick={() => setShowLine(true)}>
              변경
            </button>
          </div>
          <LinePointList records={siteRecords} line={meta.activeLine} />
          <ScanMini line={meta.activeLine} />
          <div className="prefix-row" style={{ marginTop: 10 }}>
            {["S", "R", "P"].map((p) => (
              <button key={p} className={`prefix-btn ${prefix === p && !custom ? "active" : ""}`} type="button" onClick={() => choosePrefix(p)}>
                {p}
              </button>
            ))}
            <input
              className="field-input prefix-custom"
              maxLength={3}
              placeholder="직접"
              value={custom}
              onChange={(e) => {
                const v = e.target.value.toUpperCase().replace(/[^A-Z]/g, "").slice(0, 3);
                setCustom(v);
                if (v) {
                  setPrefix(v);
                  void updateMeta({ prefix: v });
                }
              }}
            />
          </div>
          <div className="point-display">
            <button className="adj" type="button" onClick={() => setPointNumber(num - 1)} aria-label="이전 번호">
              −
            </button>
            <input className="pt" value={draft} aria-label="포인트" onChange={(e) => setDraft(e.target.value.toUpperCase())} />
            <button className="adj" type="button" onClick={() => setPointNumber(num + 1)} aria-label="다음 번호">
              +
            </button>
          </div>
          <div className="nav-hint" style={{ textAlign: "center" }}>
            마지막 기준 번호 다음. R05-1 을 직접 입력하면 기준은 5, 저장 후 R06.
          </div>
          <div className="pass-row">
            <button className="pass-btn ok" type="button" onClick={() => void commit(PASS)}>
              합격
            </button>
            <button className="pass-btn ng" type="button" onClick={() => void commit(FAIL)}>
              불합격
            </button>
          </div>
          <details className="range-box">
            <summary>범위 일괄 입력 (예: S05~S10)</summary>
            <div className="range-row">
              <input className="field-input" value={rangeFrom} placeholder="S05" onChange={(e) => setRangeFrom(e.target.value)} />
              <span>~</span>
              <input className="field-input" value={rangeTo} placeholder="S10" onChange={(e) => setRangeTo(e.target.value)} />
            </div>
            <div className="nav-hint">그중 불합격만 (쉼표)</div>
            <input className="field-input" value={rangeNg} placeholder="S07, S09" onChange={(e) => setRangeNg(e.target.value)} />
            <button className="btn primary block" type="button" onClick={() => void applyRange()}>
              범위 적용 (기본 합격)
            </button>
          </details>
        </div>
      )}

      <div className="section-title">
        오늘 기록 {todayRows.length}건{site ? ` · ${site}` : ""}
      </div>
      {todayRows.length === 0 && <div className="empty-note">오늘 기록 없음. 라인 선택 후 합격을 누르세요.</div>}
      {todayRows.map((rec, index) => (
        <div className="rec-item" key={rec.id}>
          <div>
            <div className="main">
              {rec.line} · {rec.point}
            </div>
            <div className="sub">
              {rec.date}
              {rec.diameter ? ` · ${rec.diameter}` : ""} · {syncLabel(rec)}
              {rec.syncNote ? ` · ${rec.syncNote}` : ""}
              {rec.syncError ? ` · ${rec.syncError}` : ""}
            </div>
            <div className="actions">
              {index === 0 && rec.src === "user" && (
                <button className="btn small" type="button" onClick={() => void undoLast()}>
                  실행취소
                </button>
              )}
              <button
                className="btn small"
                type="button"
                onClick={() => {
                  setEditId(rec.id);
                  setEditResult(rec.result === FAIL ? FAIL : PASS);
                }}
              >
                수정
              </button>
              <button className="btn small danger" type="button" onClick={() => void removeRecord(rec.id)}>
                삭제
              </button>
              {rec.sync === "error" && (
                <button className="btn small" type="button" onClick={() => retryRecord(rec.id)}>
                  재시도
                </button>
              )}
            </div>
          </div>
          <span className={`badge ${rec.result === PASS ? "ok" : rec.result === FAIL ? "ng" : "wait"} ${rec.sync === "error" ? "bad" : rec.sync === "synced" ? "" : "wait"}`}>
            {rec.result || "미판정"}
          </span>
        </div>
      ))}

      <div className="section-title">라인별 기록</div>
      {!site && <div className="empty-note">현장을 선택하면 검사리스트의 라인이 나옵니다.</div>}
      {site && recents.length === 0 && <div className="empty-note">불러온 라인 없음</div>}
      {site && recents.length > 0 && (
        <>
          <div className="nav-hint">총 {siteRecords.length}건 · 라인 {recents.length}개. 라인을 누르면 포인트와 날짜가 펼쳐집니다.</div>
          {allLines.map((line) => (
            <div key={line.line}>
              <button className="line-row" type="button" onClick={() => setOpenLine(openLine === line.line ? null : line.line)}>
                <span className="main">{line.line}</span>
                <span className="sub">{lineSummary(siteRecords, line.line)}</span>
              </button>
              {openLine === line.line && <LinePointList records={siteRecords} line={line.line} />}
            </div>
          ))}
        </>
      )}

      <div className={`modal-backdrop ${recentOpen ? "show" : ""}`} onClick={() => setRecentOpen(false)}>
        <div className="modal" onClick={(e) => e.stopPropagation()}>
          <h3>최근 라인</h3>
          {recents.length === 0 && <div className="empty-note">없음</div>}
          {recents.map((line) => (
            <button key={line.line} className="choice" type="button" onClick={() => activate(line.line)}>
              <div>{line.line}</div>
              <div className="sub">
                <RangeText ranges={line.ranges} />
              </div>
            </button>
          ))}
          <button className="btn block" type="button" onClick={() => setRecentOpen(false)}>
            닫기
          </button>
        </div>
      </div>

      <div className={`modal-backdrop ${siteOpen ? "show" : ""}`} onClick={() => setSiteOpen(false)}>
        <div className="modal" onClick={(e) => e.stopPropagation()}>
          <h3>현장 추가</h3>
          <input className="field-input" value={siteName} placeholder="현장 이름" onChange={(e) => setSiteName(e.target.value)} />
          <p className="nav-hint">작업일지 열과 거래명세서 접두어는 설정에서 확인하세요. 비어 있으면 그 외, 명세서는 만들지 않습니다.</p>
          <button
            className="btn primary block"
            type="button"
            onClick={() => {
              void addSite(siteName);
              setSiteName("");
              setSiteOpen(false);
            }}
          >
            추가
          </button>
        </div>
      </div>

      <div className={`modal-backdrop ${editing ? "show" : ""}`} onClick={() => setEditId(null)}>
        <div className="modal" onClick={(e) => e.stopPropagation()}>
          <h3>기록 수정</h3>
          {editing && (
            <>
              <div className="main" style={{ fontWeight: 800, marginBottom: 8 }}>
                {editing.line} · {editing.point}
              </div>
              <div className="pass-row">
                <button className={`pass-btn ok ${editResult === PASS ? "" : ""}`} type="button" onClick={() => setEditResult(PASS)} style={{ opacity: editResult === PASS ? 1 : 0.45 }}>
                  합격
                </button>
                <button className="pass-btn ng" type="button" onClick={() => setEditResult(FAIL)} style={{ opacity: editResult === FAIL ? 1 : 0.45 }}>
                  불합격
                </button>
              </div>
              <button
                className="btn primary block"
                type="button"
                onClick={() => {
                  void setResult(editing.id, editResult);
                  setEditId(null);
                }}
              >
                저장
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
