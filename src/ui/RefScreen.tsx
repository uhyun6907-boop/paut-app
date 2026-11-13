import { useEffect, useState } from "react";
import { JACKET_ROWS, jacketButtonLabel, jacketPaut, jacketScan, jacketSizeLabel, type JacketRow } from "../domain/jacket";
import { PAUT_OD, SCH_ORDER, SCHEDULE, inchKey, nearestThickness, sizeChoices } from "../domain/reference";
import { inchLabel } from "../domain/size";
import { useDrill } from "./drill";

const SIZES = sizeChoices();

export function RefScreen() {
  const [tab, setTab] = useState<"od" | "jacket">("od");
  const [inch, setInch] = useState<number | null>(null);
  const [jacketA, setJacketA] = useState<number | null>(null);
  const [mat, setMat] = useState<"cs" | "ss">("cs");
  const [openSch, setOpenSch] = useState<number | null>(null);
  const paut = PAUT_OD.find((p) => p.inch === inch);
  const sch = inch == null ? undefined : SCHEDULE[inchKey(inch)];
  const jacket = JACKET_ROWS.find((row) => row.a === jacketA) ?? null;
  const odOpen = tab === "od" && inch != null;
  const jacketOpen = tab === "jacket" && jacket != null;

  useDrill(odOpen, () => { setInch(null); setOpenSch(null); });
  useDrill(jacketOpen, () => setJacketA(null));

  useEffect(() => {
    if (odOpen || jacketOpen) window.scrollTo(0, 0);
  }, [odOpen, jacketOpen]);

  return (
    <div className="ref">
      <div className="encoder">
        <span>Encoder · ROTIX 일반 <strong>18.900</strong></span>
        <span>수동 <strong>19.050</strong></span>
        <span>이중보온관 <strong>15.800</strong></span>
      </div>
      <div className="segment">
        <button className={tab === "od" ? "on" : ""} type="button" onClick={() => setTab("od")}>외경/스캔</button>
        <button className={tab === "jacket" ? "on" : ""} type="button" onClick={() => setTab("jacket")}>이중보온관</button>
      </div>
      {tab === "od" && inch == null && (
        <div className="size-grid">
          {SIZES.map((s) => (
            <button key={s.inch} className="size-btn" type="button" onClick={() => { setInch(s.inch); setOpenSch(null); }}>
              {inchLabel(s.inch)} / {s.a}A
              <small>{s.hasScan ? "스캔" : "스케줄만"}</small>
            </button>
          ))}
        </div>
      )}
      {odOpen && (
        <div>
          <button className="ref-back" type="button" onClick={() => { setInch(null); setOpenSch(null); }}>‹ 관경 목록</button>
          <div className="scan-hero">
            <div className="label">스캔길이</div>
            <div className="value">{paut?.scan ?? "—"}</div>
            <div className="who">{inchLabel(inch!)} / {(paut || sch)?.a}A</div>
          </div>
          <div className="meta-grid">
            {paut && <div className="meta-card"><div className="k">외경 JIS</div><div className="v">{paut.od} mm</div></div>}
            {sch && <div className="meta-card"><div className="k">외경 ANSI</div><div className="v">{sch.ansi} mm</div></div>}
            {paut?.circ ? <div className="meta-card"><div className="k">원둘레</div><div className="v">{paut.circ} mm</div></div> : null}
            {paut && (
              <div className="meta-card">
                <div className="k">ROTIX</div>
                <div className="v">{sl(paut.rs, paut.rl) ?? "—"}</div>
              </div>
            )}
            {paut && (paut.ds != null || paut.dl != null) && (
              <div className="meta-card">
                <div className="k">이중보온관</div>
                <div className="v">{sl(paut.ds, paut.dl)}</div>
              </div>
            )}
          </div>
          {sch && (
            <>
              <div className="section-title">SCH별 두께</div>
              <div className="segment ref-mat">
                <button className={mat === "cs" ? "on" : ""} type="button" onClick={() => { setMat("cs"); setOpenSch(null); }}>탄소강</button>
                <button className={mat === "ss" ? "on" : ""} type="button" onClick={() => { setMat("ss"); setOpenSch(null); }}>스테인리스</button>
              </div>
              <div className="sch-grid">
                {SCH_ORDER.filter((n) => (mat === "cs" ? sch.cs : sch.ss)[n] != null).map((n) => {
                  const thk = (mat === "cs" ? sch.cs : sch.ss)[n];
                  const row = nearestThickness(thk);
                  return (
                    <div key={n} style={{ display: "contents" }}>
                      <button className={`sch-btn${openSch === n ? " on" : ""}`} type="button" onClick={() => setOpenSch(openSch === n ? null : n)}>
                        <div style={{ fontWeight: 800 }}>SCH {n}</div>
                        <div className="thk">{thk} <span style={{ fontSize: 12, color: "var(--muted)" }}>mm</span></div>
                      </button>
                      {openSch === n && (
                        <div className="setting-box">
                          <div className="applied">적용 두께 T={row.t} 기준 (실제 {thk}mm)</div>
                          <div className="set-grid">
                            <div className="set-item full"><div className="k">Probe</div><div className="v">{row.probe}</div></div>
                            <div className="set-item"><div className="k">EL.Q</div><div className="v">{row.elq}</div></div>
                            <div className="set-item"><div className="k">비드 폭</div><div className="v">{row.bead}</div></div>
                            <div className="set-item"><div className="k">Focus 1CH</div><div className="v">{row.f1}</div></div>
                            <div className="set-item"><div className="k">Focus 2CH</div><div className="v">{row.f2}</div></div>
                            <div className="set-item full"><div className="k">Index</div><div className="v">{row.idx}</div></div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </>
          )}
          <button className="btn block" type="button" onClick={() => { setInch(null); setOpenSch(null); }}>← 관경 다시</button>
        </div>
      )}
      {tab === "jacket" && !jacket && (
        <div className="size-grid">
          {JACKET_ROWS.map((row) => (
            <button key={row.a} className="size-btn" type="button" onClick={() => setJacketA(row.a)}>
              {jacketButtonLabel(row)}
              <small>{row.mat}</small>
            </button>
          ))}
        </div>
      )}
      {jacketOpen && jacket && <JacketDetail row={jacket} onBack={() => setJacketA(null)} />}
    </div>
  );
}

function JacketDetail({ row, onBack }: { row: JacketRow; onBack: () => void }) {
  const paut = jacketPaut(row.a);
  const length = jacketScan(row.a);
  const scan = paut && (paut.ds != null || paut.dl != null) ? sl(paut.ds, paut.dl) : null;
  return (
    <div>
      <button className="ref-back" type="button" onClick={onBack}>‹ 관경 목록</button>
      <div className="scan-hero">
        <div className="label">스캔길이</div>
        <div className="value">{length ?? "—"}{length != null ? <span className="unit"> mm</span> : null}</div>
        <div className="who">{jacketSizeLabel(row)}</div>
      </div>
      <div className="meta-grid jacket-ods">
        <div className="meta-card"><div className="k">내관 외경 D</div><div className="v">{row.innerD} mm</div></div>
        <div className="meta-card"><div className="k">외관 외경 D</div><div className="v">{row.outerD} mm</div></div>
      </div>
      <div className="meta-grid">
        <div className="meta-card"><div className="k">내관 재질</div><div className="v">{row.mat}</div></div>
        <div className="meta-card"><div className="k">내관 두께 T</div><div className="v">{row.innerT} mm</div></div>
        <div className="meta-card"><div className="k">외관 두께</div><div className="v">{row.outerT} mm</div></div>
        <div className="meta-card"><div className="k">보온재 두께</div><div className="v">{row.insul} mm</div></div>
        <div className="meta-card"><div className="k">단위중량</div><div className="v">{row.kg} kg/m</div></div>
        <div className="meta-card"><div className="k">길이</div><div className="v">{row.len} mm</div></div>
        {scan && (
          <div className="meta-card wide">
            <div className="k">이중보온관 S·L · Encoder 15.800</div>
            <div className="v">{scan}</div>
          </div>
        )}
      </div>
      <p className="jacket-note">내관 강관 / 외관 HDPE(매설용), AL(노출용) / 보온재 경질 폴리우레탄 폼</p>
      <button className="btn block" type="button" onClick={onBack}>← 관경 다시</button>
    </div>
  );
}

function sl(s?: number, l?: number): string | null {
  const parts = [s != null ? `S ${s}` : null, l != null ? `L ${l}` : null].filter(Boolean);
  return parts.length ? parts.join(" · ") : null;
}
