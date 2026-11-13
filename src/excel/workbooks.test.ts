import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { buildBill, countByDiameter } from "../domain/billing";
import { nextPointLabel } from "../domain/point";
import { defaultWeldLengths } from "../domain/mapping";
import type { InspectionRecord } from "../domain/types";
import { applyIljiDay, detectPairs, detectStaffColumns, findIljiSheet, readIljiCell } from "./iljiBook";
import { parseListWorkbook, writeListRow } from "./listBook";
import { readSheet, unzipXlsx, zipXlsx, type ZipFiles } from "./ooxml";
import { readLengths, readStatement, writePautLines } from "./statementBook";
import { userCreatesMonthSheet } from "../test/userSheet";

function load(name: string): ZipFiles {
  return unzipXlsx(readFileSync(new URL(`../../public/samples/${name}`, import.meta.url)));
}

function clone(files: ZipFiles): ZipFiles {
  return { ...files };
}

describe("실제 통합 문서", () => {
  const list = load("ndt-list.xlsx");
  const ilji = load("work-log.xlsx");
  const statement = load("statement.xlsx");
  const sites = parseListWorkbook(list);

  it("현장 시트 이름(앞 공백 포함)을 읽고 RT 시트는 버린다", () => {
    const names = sites.map((s) => s.sheetName);
    expect(names).toContain(" 드림 합정주공");
    expect(names).toContain(" 지성 브레인");
    expect(names).toContain("택지 드림");
    expect(names).toContain("드림 화양");
    expect(names).toContain(" 브레인대지 인입");
    expect(names.some((n) => n.includes("RT"))).toBe(false);
    const hap = sites.find((s) => s.site === "드림 합정주공")!;
    expect(hap.records.length).toBe(427);
    expect(hap.records.filter((r) => r.result === "불합격")).toHaveLength(6);
    expect(hap.records.filter((r) => r.result === "합격")).toHaveLength(421);
    const row = hap.records.find((r) => r.line === "HJ-M-002-350A" && r.point === "S13")!;
    expect(row.date).toBe("2026-03-25");
    expect(row.diameter).toBe("350A");
    expect(row.material).toBe("SPPS380");
  });

  it("불합격 재검사는 접미어 없이 같은 포인트 번호의 다음 행이다", () => {
    for (const site of sites) {
      for (const fail of site.records.filter((rec) => rec.result === "불합격")) {
        expect(fail.point).not.toMatch(/재|-R$|RE\d/i);
        expect(fail.note).not.toMatch(/재검/);
        const laterPass = site.records.filter(
          (rec) => rec.line === fail.line && rec.point === fail.point && rec.result === "합격" && (rec.excelRow ?? 0) > (fail.excelRow ?? 0),
        );
        expect(laterPass.length).toBeGreaterThan(0);
      }
    }
  });

  it("HJ-M-005-350A 의 다음 R 은 R22", () => {
    const hap = sites.find((s) => s.site === "드림 합정주공")!;
    const pts = hap.records.filter((r) => r.line === "HJ-M-005-350A").map((r) => r.point);
    expect(nextPointLabel(pts, "R")).toBe("R22");
    expect(nextPointLabel(pts, "S")).toBe("S22");
  });

  it("행을 추가하면 관경은 RIGHT 수식이고, 다시 읽으면 350A", () => {
    const files = clone(list);
    const row = writeListRow(files, " 드림 합정주공", {
      id: "test-point-1",
      date: "2026-11-09",
      line: "HJ-M-099-350A",
      point: "S01",
      result: "합격",
      material: "SPPS380",
    });
    expect(row).toBe(433);
    const again = parseListWorkbook(files).find((s) => s.site === "드림 합정주공")!;
    const added = again.records.find((r) => r.id === "test-point-1")!;
    expect(added.point).toBe("S01");
    expect(added.diameter).toBe("350A");
    expect(added.date).toBe("2026-11-09");
    expect(added.result).toBe("합격");
    const zipped = zipXlsx(files);
    const reopened = parseListWorkbook(unzipXlsx(zipped)).find((s) => s.site === "드림 합정주공")!;
    expect(reopened.records.some((r) => r.id === "test-point-1")).toBe(true);
  });

  it("불합격은 작업일지 개소에서 빠진다", () => {
    const files = clone(ilji);
    const passes = [
      { result: "합격", diameter: "350A", point: "S01" },
      { result: "합격", diameter: "350A", point: "R01" },
      { result: "불합격", diameter: "350A", point: "S02" },
    ];
    const counted = countByDiameter(
      passes.map((p, i) => ({
        id: String(i),
        site: "드림 합정주공",
        sheetName: " 드림 합정주공",
        date: "2026-11-09",
        line: "HJ-M-099-350A",
        point: p.point,
        result: p.result,
        diameter: p.diameter,
        material: "",
        sch: "",
        reportNo: "",
        reportDate: "",
        defectNo: "",
        drawingNo: "",
        weldNo: "",
        rtSheets: "",
        flawType: "",
        flawLen: "",
        note: "",
        issued: "",
        src: "user" as const,
        sync: "pending" as const,
        ts: i,
      })),
    );
    expect(counted).toEqual({ "350A": 2 });
    const row = applyIljiDay(files, {
      date: "2026-11-09",
      column: "드림 합정주공",
      groups: Object.entries(counted).map(([diameter, count]) => ({ diameter, count })),
      dayType: "평일",
      insp1: "조우현",
      insp2: "심명보",
      timeStart: "09:00",
      timeEnd: "18:00",
    });
    const pairs = detectPairs(readSheet(files, findIljiSheet(files)).cells);
    expect(pairs.find((p) => p.column === "드림 합정주공")?.diaCol).toBe("J");
    expect(readIljiCell(files, row, "J")).toBe("350A");
    expect(readIljiCell(files, row, "K")).toBe("2");
    expect(readIljiCell(files, row, "B")).toBe("2026-11-09");
    expect(readIljiCell(files, row, "P")).toBe("조우현");
    expect(readIljiCell(files, row, "Q")).toBe("심명보");
    expect(readIljiCell(files, row, "P")).not.toContain("/");
    expect(readIljiCell(files, row, "S")).toBe("09:00");
    expect(readIljiCell(files, row, "T")).toBe("18:00");
  });

  it("검사자는 P·Q에 한 명씩, J·K는 합정 관경이다", () => {
    const sheet = findIljiSheet(ilji);
    const cells = readSheet(ilji, sheet).cells;
    const staff = detectStaffColumns(cells);
    expect(staff).toEqual({ insp1: "P", insp2: "Q", start: "S", end: "T" });
    expect(detectPairs(cells).find((p) => p.column === "드림 합정주공")).toMatchObject({ diaCol: "J", countCol: "K" });
    expect(readIljiCell(ilji, 6, "P")).toBe("조우현");
    expect(readIljiCell(ilji, 6, "Q")).toBe("심명보");
    expect(readIljiCell(ilji, 6, "J")).toBe("");

    const files = clone(ilji);
    const row = applyIljiDay(files, {
      date: "2026-11-09",
      column: "드림 합정주공",
      groups: [{ diameter: "350A", count: 2 }],
      dayType: "평일",
      insp1: "조우현/심명보",
      insp2: "방현민",
      timeStart: "09:00",
      timeEnd: "18:00",
    });
    expect(readIljiCell(files, row, "P")).toBe("조우현");
    expect(readIljiCell(files, row, "Q")).toBe("방현민");
    expect(readIljiCell(files, row, "J")).toBe("350A");
    expect(readIljiCell(files, row, "K")).toBe("2");
    expect(readIljiCell(files, row, "L")).toBe("");
    expect(readIljiCell(files, row, "M")).toBe("");
    const alone = applyIljiDay(files, {
      date: "2026-11-09",
      column: "드림 합정주공",
      groups: [{ diameter: "350A", count: 2 }],
      dayType: "평일",
      insp1: "조우현",
      insp2: "",
      timeStart: "09:00",
      timeEnd: "18:00",
    });
    expect(readIljiCell(files, alone, "P")).toBe("조우현");
    expect(readIljiCell(files, alone, "Q")).toBe("");
  });

  it("근무시간을 비우면 작업일지 시각 칸은 빈칸이다", () => {
    const files = clone(ilji);
    const base = {
      date: "2026-11-09",
      column: "드림 합정주공",
      groups: [{ diameter: "350A", count: 2 }],
      dayType: "평일",
      insp1: "조우현",
      insp2: "심명보",
    };
    const row = applyIljiDay(files, { ...base, timeStart: "09:00", timeEnd: "18:00" });
    expect(readIljiCell(files, row, "S")).toBe("09:00");
    applyIljiDay(files, { ...base, timeStart: "", timeEnd: "" });
    expect(readIljiCell(files, row, "S")).toBe("");
    expect(readIljiCell(files, row, "T")).toBe("");
    const xml = readSheet(files, findIljiSheet(files)).xml;
    for (const col of ["S", "T"]) {
      const cell = xml.match(new RegExp(`<c r="${col}${row}"(?:[^>]*/>|[^>]*>[\\s\\S]*?</c>)`))?.[0] ?? "";
      expect(cell).not.toContain("<v>");
    }
    applyIljiDay(files, { ...base, timeStart: "09:00", timeEnd: "" });
    expect(readIljiCell(files, row, "S")).toBe("09:00");
    expect(readIljiCell(files, row, "T")).toBe("");
  });

  it("드림합정 26.10 청구액은 리스트의 합격 개소와 같다", () => {
    const view = readStatement(statement, "드림합정26.10");
    expect(view.unitPrice).toBe(75000);
    expect(view.lengths["350"]).toBe(1.2);
    expect(view.lengths["200"]).toBe(0.7);
    const allLengths = readLengths(statement);
    expect(allLengths["750"]).toBe(2.4);
    expect(allLengths["600"]).toBe(1.9);
    expect(view.paut.map((p) => [p.date, p.count, p.unitPrice])).toEqual([
      ["2026-10-15", 8, 75000],
      ["2026-10-21", 10, 75000],
      ["2026-10-28", 10, 75000],
    ]);
    const hap = sites.find((s) => s.site === "드림 합정주공")!;
    const october = hap.records.filter((r) => r.date.startsWith("2026-10"));
    const bill = buildBill(october, view.unitPrice!, { ...defaultWeldLengths(), ...view.lengths });
    expect(bill.supply).toBe(2520000);
    expect(bill.vat).toBe(252000);
    expect(bill.grand).toBe(2772000);
    expect(bill.lines.map((l) => l.count)).toEqual([8, 10, 10]);
  });

  it("사용자가 만든 11월 시트에 합격만 적고, 비고(I열)에는 아무것도 쓰지 않는다", () => {
    const files = clone(statement);
    const ensured = { name: userCreatesMonthSheet(files, "드림합정", 2026, 11) };
    expect(ensured.name).toBe("드림합정26.11");
    const hap = sites.find((s) => s.site === "드림 합정주공")!.records.filter((r) => r.date === "2026-11-09");
    expect(hap.length).toBe(0);
    const fake: InspectionRecord[] = [
      base("a", "합격", "S01"),
      base("b", "합격", "R01"),
      base("c", "불합격", "S02"),
    ];
    const bill = buildBill(fake, 75000, defaultWeldLengths());
    writePautLines(
      files,
      ensured.name,
      bill.lines.map((l) => ({
        date: l.date,
        diameter: l.diameter,
        lengthM: l.lengthM!,
        count: l.count,
        unitPrice: 75000,
      })),
    );
    const view = readStatement(files, ensured.name);
    expect(view.paut).toEqual([{ date: "2026-11-09", item: "350A(1.2M)", count: 2, unitPrice: 75000 }]);
    expect(readSheet(files, ensured.name).cells.get("I12")?.v ?? "").toBe("");
    const zipped = zipXlsx(files);
    expect(zipped.byteLength).toBeGreaterThan(1000);
    const names = [...new TextDecoder().decode(unzipXlsx(zipped)["xl/workbook.xml"])];
    expect(names.join("")).toContain("드림합정26.11");
  });
});

function base(id: string, result: string, point: string): InspectionRecord {
  return {
    id,
    site: "드림 합정주공",
    sheetName: " 드림 합정주공",
    date: "2026-11-09",
    line: "HJ-M-099-350A",
    point,
    result,
    diameter: "350A",
    material: "",
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "pending",
    ts: 1,
  };
}
