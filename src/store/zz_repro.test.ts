import { beforeEach, describe, expect, it, vi } from "vitest";
const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
vi.mock("./db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => { mem.set(key, structuredClone(value)); },
  kvDel: async (key: string) => { mem.delete(key); },
}));
vi.stubGlobal("localStorage", {
  getItem: (key: string) => bag.get(key) ?? null,
  setItem: (key: string, value: string) => { bag.set(key, String(value)); },
  removeItem: (key: string) => { bag.delete(key); },
});
import { FAIL, PASS, type InspectionRecord } from "../domain/types";
import { getState, mirrorGreyDraft, setLineDrafts, setResult, settled, updateMeta } from "./session";
import { applyPaint } from "../ui/EntryView";
const DATE = "2026-11-12"; const SITE = "현장"; let LINE = "L"; let n = 0;
const today = (p: string) => getState().records.filter((r) => !r.deleted && r.date === DATE && r.line.trim() === LINE && r.point.toUpperCase() === p);
function xl(point: string, result: string, date = DATE): InspectionRecord {
  return { id: crypto.randomUUID(), site: SITE, sheetName: SITE, listBook: undefined as any, date, line: LINE, point, result: result as any, diameter: "", material: "", sch: "", reportNo: "", reportDate: "", defectNo: "", drawingNo: "", weldNo: "", rtSheets: "", flawType: "", flawLen: "", note: "", issued: "", src: "xlsx" as any, sync: "synced", ts: Date.now(), excelRow: 10 } as InspectionRecord;
}
describe("repro", () => {
  beforeEach(async () => {
    await settled(); mem.clear(); bag.clear();
    getState().records.splice(0);
    n++; LINE = `L-${n}`;
    await updateMeta({ lastSite: SITE, activeLine: LINE, workDate: DATE, prefix: "S", draftCards: {} });
  });
  it("drafts S29-31 painted", async () => {
    for (let i = 1; i <= 28; i++) getState().records.push(xl(`S${String(i).padStart(2,"0")}`, PASS), xl(`R${i}`, PASS));
    const key = `${SITE}||${LINE}||S`;
    await setLineDrafts(key, ["S29","S30","S31"].map((p) => ({ id: `d-${p}`, point: p })));
    for (const p of ["S29","S30","S31"]) await mirrorGreyDraft(p);
    for (const p of ["S29","S30","S31"]) await applyPaint(`draft:d-${p}`, false);
    await settled();
    for (const p of ["R29","R30","R31"]) expect(today(p).map((r) => r.result), p).toEqual([PASS]);
  });
  it("S already pass from excel, R missing", async () => {
    for (const p of ["S29","S30","S31"]) getState().records.push(xl(p, PASS));
    const s = today("S29")[0];
    await setResult(s.id, FAIL); await setResult(s.id, PASS);
    expect(today("R29").map((r) => r.result)).toEqual([PASS]);
  });
  it("S on older date", async () => {
    getState().records.push(xl("S29", FAIL, "2026-11-10"));
    const s = getState().records.find((r) => r.point === "S29")!;
    await setResult(s.id, PASS);
    console.log(getState().records.filter(r=>r.point==="R29").map(r=>[r.date,r.result]));
  });
});
