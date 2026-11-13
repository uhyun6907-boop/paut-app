/** 날짜 고치기 — 로컬(데모) 경로: 이 기기 사본에만 쓴다. 리스트 행이 늘지 않고 정렬되며, 옛 달·옛 날에서 빠진다. */
import { beforeEach, describe, expect, it } from "vitest";
import { PASS } from "../domain/types";
import { fixture, graphConfig, meta, rec } from "../test/realBooks";
import { userCreatesMonthSheet } from "../test/userSheet";
import { bootstrapFromSamples, exportBook, rememberWork, restoreStatementOwnership, syncPendingDemo } from "./engine";
import { parseListWorkbook, planLineReorder } from "./listBook";
import { unzipXlsx, workbookSheets, zipXlsx } from "./ooxml";
import { readPautRows } from "./statementBook";

beforeEach(() => restoreStatementOwnership({}));

const LINE = "HJ-T-900-350A";

describe("날짜 고치기 — 로컬 사본", () => {
  it("다른 달 기존 시트로 옮기면 옛 달 줄은 지우고, 리스트는 같은 Y 행을 고쳐 날짜 순으로 둔다. 새 시트는 만들지 않는다", () => {
    const boot = bootstrapFromSamples({ list: fixture("real-list.xlsx"), ilji: fixture("real-worklog.xlsx"), statement: fixture("real-statement.xlsx") });
    // 사용자가 만든 11월 시트
    const stmt = unzipXlsx(exportBook("statement"));
    userCreatesMonthSheet(stmt, "드림합정", 2026, 11);
    rememberWork("statement", zipXlsx(stmt));
    const recs = [rec("S01", PASS, { line: LINE }), rec("S02", PASS, { line: LINE }), rec("S03", PASS, { line: LINE })];
    const all = [...boot.records, ...recs];
    syncPendingDemo(recs, all, graphConfig(), meta());
    const sheetsBefore = workbookSheets(unzipXlsx(exportBook("statement"))).map((s) => s.name);
    expect(readPautRows(unzipXlsx(exportBook("statement")), "드림합정26.11").map((r) => `${r.date}|${r.count}`)).toContain("2026-11-12|3");

    const s02 = recs[1];
    s02.date = "2026-10-30";
    s02.movedFrom = ["2026-11-12"];
    s02.steps = undefined;
    const res = syncPendingDemo([s02], all, graphConfig(), meta());
    expect(res.get(s02.id)?.status).toBe("synced");

    const out = unzipXlsx(exportBook("statement"));
    expect(workbookSheets(out).map((s) => s.name)).toEqual(sheetsBefore);
    expect(readPautRows(out, "드림합정26.11").map((r) => `${r.date}|${r.count}`)).toContain("2026-11-12|2");
    expect(readPautRows(out, "드림합정26.10").map((r) => `${r.date}|${r.count}`)).toContain("2026-10-30|1");
    const rows = parseListWorkbook(unzipXlsx(exportBook("list")))
      .flatMap((s) => s.records)
      .filter((r) => r.line === LINE)
      .sort((a, b) => (a.excelRow ?? 0) - (b.excelRow ?? 0))
      .map((r) => `${r.date}|${r.point}`);
    expect(rows).toEqual(["2026-10-30|S02", "2026-11-12|S01", "2026-11-12|S03"]);
  });

  it("planLineReorder: 같은 라인 앱 행 자리 안에서만 바꾸고, 날짜를 바꾼 기록이 없으면 아무것도 옮기지 않는다", () => {
    const a = rec("S01", PASS, { line: LINE, date: "2026-11-10" });
    const b = rec("S02", PASS, { line: LINE, date: "2026-11-11" });
    const other = rec("S01", PASS, { line: "OTHER-200A", date: "2026-11-01" });
    const rows = new Map([[a.id, 20], [other.id, 21], [b.id, 30]]);
    expect(planLineReorder(rows, [a, b, other]).size).toBe(0);
    b.date = "2026-11-01";
    b.movedFrom = ["2026-11-11"];
    const moves = planLineReorder(rows, [a, b, other]);
    expect([...moves.entries()].map(([id, m]) => [id, m.row])).toEqual([[b.id, 20], [a.id, 30]]);
  });
});
