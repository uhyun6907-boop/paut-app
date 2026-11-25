/**
 * 저장 전 확인: 「저장」을 누르면 바로 보내지 않고, 이번에 저장될 것(저장 안 한 카드 — 두 현장이면 둘 다)을
 * 현장 → 라인 → 포인트로 한 번에 보여 준다. 여기서 판정을 바꾸거나 지우고, 날짜·검사자·시간을 확인한 뒤 「저장」.
 * 바꾸는 것은 아직 저장 안 한 카드라 엑셀에는 「저장」을 눌러야 간다.
 */
import { useMemo } from "react";
import { FAIL, PASS, type InspectionRecord } from "../domain/types";
import { commitDay, removeRecord, setResult, updateMeta, useSession } from "../store/session";
import { useDrill } from "./drill";

function byPoint(a: InspectionRecord, b: InspectionRecord): number {
  const rank = (p: string) => ({ S: 0, R: 1, P: 2 })[p.charAt(0).toUpperCase() as "S" | "R" | "P"] ?? 3;
  return rank(a.point) - rank(b.point) || a.point.localeCompare(b.point, "en", { numeric: true });
}

function dayText(iso: string): string {
  const [, m, d] = iso.split("-");
  const week = ["일", "월", "화", "수", "목", "금", "토"][new Date(`${iso}T00:00:00Z`).getUTCDay()];
  return `${Number(m)}/${Number(d)} (${week})`;
}

export function SaveReview({ open, onClose }: { open: boolean; onClose: () => void }) {
  const { records, meta } = useSession();
  useDrill(open, onClose);
  const groups = useMemo(() => {
    const local = records.filter((r) => r.sync === "local" && r.result !== "없음");
    const map = new Map<string, Map<string, InspectionRecord[]>>();
    for (const r of local) {
      const site = map.get(r.site) ?? new Map<string, InspectionRecord[]>();
      site.set(r.line, [...(site.get(r.line) ?? []), r]);
      map.set(r.site, site);
    }
    return [...map].map(([site, lines]) => ({ site, lines: [...lines].map(([line, recs]) => ({ line, recs: [...recs].sort(byPoint) })) }));
  }, [records]);
  const live = groups.flatMap((g) => g.lines.flatMap((l) => l.recs)).filter((r) => !r.deleted);
  const pass = live.filter((r) => r.result === PASS).length;
  const fail = live.filter((r) => r.result === FAIL).length;

  if (!open) return null;
  return (
    <div className="modal-backdrop show" onClick={onClose}>
      <div className="modal save-review" onClick={(e) => e.stopPropagation()}>
        <h3>이대로 저장할까요?</h3>
        <div className="sr-staff">
          <label>
            <span>날짜</span>
            <input className="field-input" type="date" value={meta.workDate} onChange={(e) => e.target.value && void updateMeta({ workDate: e.target.value })} />
          </label>
          <label>
            <span>검사자 1</span>
            <input className="field-input" value={meta.insp1} onChange={(e) => void updateMeta({ insp1: e.target.value })} />
          </label>
          <label>
            <span>검사자 2</span>
            <input className="field-input" value={meta.insp2} onChange={(e) => void updateMeta({ insp2: e.target.value })} />
          </label>
          <label>
            <span>시작</span>
            <input className="field-input" type="time" value={meta.timeStart} onChange={(e) => void updateMeta({ timeStart: e.target.value })} />
          </label>
          <label>
            <span>종료</span>
            <input className="field-input" type="time" value={meta.timeEnd} onChange={(e) => void updateMeta({ timeEnd: e.target.value })} />
          </label>
        </div>
        <div className="sr-sum">{dayText(meta.workDate)} · 합격 {pass} · 불합격 {fail}{groups.length > 1 ? ` · ${groups.length}개 현장` : ""}</div>
        {groups.length === 0 && <div className="empty-note">저장 안 한 카드가 없습니다(검사자·시간만 저장).</div>}
        {groups.map((g) => (
          <div key={g.site} className="sr-site">
            <div className="sr-site-name">{g.site}</div>
            {g.lines.map((l) => (
              <div key={l.line} className="sr-line">
                <div className="sr-line-name">{l.line}</div>
                {l.recs.map((r) => (
                  <div key={r.id} className={`sr-pt${r.deleted ? " removed" : ""}`}>
                    <span className="p">{r.point}{r.date !== meta.workDate ? <small> {dayText(r.date)}</small> : null}</span>
                    {r.deleted ? (
                      <span className="gone">지움</span>
                    ) : (
                      <>
                        <span className="seg">
                          <button type="button" className={r.result === PASS ? "on ok" : ""} onClick={() => void setResult(r.id, PASS)}>합격</button>
                          <button type="button" className={r.result === FAIL ? "on bad" : ""} onClick={() => void setResult(r.id, FAIL)}>불합격</button>
                        </span>
                        <button type="button" className="btn small" onClick={() => void removeRecord(r.id, true)}>지우기</button>
                      </>
                    )}
                  </div>
                ))}
              </div>
            ))}
          </div>
        ))}
        <div className="sr-actions">
          <button type="button" className="btn" onClick={onClose}>계속 입력</button>
          <button
            type="button"
            className="btn primary"
            onClick={() => {
              onClose();
              void commitDay((text) => window.confirm(text));
            }}
          >
            저장
          </button>
        </div>
      </div>
    </div>
  );
}
