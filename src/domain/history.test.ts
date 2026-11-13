import { describe, expect, it } from "vitest";
import { defectStats, formatInspectionDate, latestLine, linePointRows, linesByLatestDate, lineSummary, rangesFromPoints, recentLineProgress, sortLineRows } from "./history";
import type { InspectionRecord } from "./types";

function row(partial: Pick<InspectionRecord, "line" | "point" | "result" | "date">): InspectionRecord {
  return {
    id: partial.point + partial.date,
    site: "드림 합정주공",
    sheetName: " 드림 합정주공",
    date: partial.date,
    line: partial.line,
    point: partial.point,
    result: partial.result,
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
    src: "xlsx",
    sync: "synced",
    ts: 1,
  };
}

describe("라인 전체 구간", () => {
  it("날짜를 모아 S01~S05 · R01~R21 로 붙인다", () => {
    const records = [
      ...[1, 2, 3, 4, 5].map((n) => row({ line: "HJ-M-005-350A", point: `S0${n}`, result: "합격", date: "2026-10-28" })),
      ...[1, 2, 10, 21].map((n) => row({ line: "HJ-M-005-350A", point: `R${String(n).padStart(2, "0")}`, result: "합격", date: "2026-10-21" })),
      ...[3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20].map((n) =>
        row({ line: "HJ-M-005-350A", point: `R${String(n).padStart(2, "0")}`, result: "합격", date: "2026-10-28" }),
      ),
    ];
    const [line] = recentLineProgress(records);
    expect(line.lastDate).toBe("2026-10-28");
    expect(line.ranges.map((r) => r.text)).toEqual(["S01~S05", "R01~R21"]);
  });

  it("빠진 번호는 구간을 끊고, 불합격은 따로 모은다", () => {
    const records = [1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21].map((n) =>
      row({
        line: "L",
        point: `R${String(n).padStart(2, "0")}`,
        result: n === 10 ? "불합격" : "합격",
        date: "2026-10-28",
      }),
    );
    const ranges = rangesFromPoints(records);
    expect(ranges[0].text).toBe("R01~R08, R10~R21");
    expect(ranges[0].fails).toEqual(["R10"]);
  });

  it("같은 번호가 나중에 합격이면 그 포인트는 끝난 것으로 본다", () => {
    const records = [
      row({ line: "L", point: "S14", result: "합격", date: "2026-10-28" }),
      row({ line: "L", point: "S15", result: "불합격", date: "2026-10-28" }),
      row({ line: "L", point: "S15", result: "합격", date: "2026-10-28" }),
      row({ line: "L", point: "S16", result: "합격", date: "2026-10-28" }),
    ];
    records[1].ts = 1;
    records[2].ts = 2;
    const ranges = rangesFromPoints(records);
    expect(ranges[0].text).toBe("S14~S16");
    expect(ranges[0].fails).toEqual([]);
  });

  it("한 번호만 있으면 물결을 붙이지 않는다", () => {
    expect(rangesFromPoints([row({ line: "L", point: "S05", result: "합격", date: "2026-10-28" })])[0].text).toBe("S05");
  });

  it("라인의 포인트를 S·R 순으로 나열하고 날짜와 요일을 붙인다", () => {
    const records = [
      row({ line: "HJ-M-005-350A", point: "R02", result: "불합격", date: "2026-11-08" }),
      row({ line: "HJ-M-005-350A", point: "S02", result: "합격", date: "2026-10-28" }),
      row({ line: "HJ-M-005-350A", point: "S01", result: "합격", date: "2026-10-21" }),
      row({ line: "다른라인", point: "S01", result: "합격", date: "2026-11-08" }),
      row({ line: "hj-m-005-350a", point: "R01", result: "합격", date: "2026-11-09" }),
    ];
    const listed = linePointRows(records, "HJ-M-005-350A");
    expect(listed.map((item) => item.point)).toEqual(["S01", "S02", "R01", "R02"]);
    expect(listed[0].dateLabel).toBe(formatInspectionDate("2026-10-21"));
    expect(listed[2].dateLabel).toBe(formatInspectionDate("2026-11-09"));
    expect(listed[3].fail).toBe(true);
    expect(formatInspectionDate("2026-11-08")).toBe("2026-11-08 (일)");
    expect(lineSummary(records, "HJ-M-005-350A")).toBe("S01~S02, R01~R02 · 불합격 R02 · 마지막 11/09");
  });

  it("라인을 최근 검사일 순으로 고른다", () => {
    const records = [
      row({ line: "OLD", point: "S01", result: "합격", date: "2026-01-02" }),
      row({ line: "NEW", point: "S01", result: "합격", date: "2026-11-08" }),
      row({ line: "OLD", point: "S02", result: "합격", date: "2026-03-01" }),
      row({ line: "MID", point: "R01", result: "합격", date: "2026-06-01" }),
    ];
    expect(linesByLatestDate(records).map((item) => [item.line, item.date])).toEqual([
      ["NEW", "2026-11-08"],
      ["MID", "2026-06-01"],
      ["OLD", "2026-03-01"],
    ]);
    expect(latestLine(records)).toBe("NEW");
  });

  it("라인순은 이름에 들어 있는 숫자 순이다", () => {
    const rows = linesByLatestDate([
      row({ line: "HJ-M-10-200A", point: "S01", result: "합격", date: "2026-11-01" }),
      row({ line: "HJ-M-2-200A", point: "S01", result: "합격", date: "2026-01-01" }),
      row({ line: "HJ-M-9-200A", point: "S01", result: "합격", date: "2026-06-01" }),
    ]);
    expect(sortLineRows(rows, "recent").map((item) => item.line)).toEqual(["HJ-M-10-200A", "HJ-M-9-200A", "HJ-M-2-200A"]);
    expect(sortLineRows(rows, "name").map((item) => item.line)).toEqual(["HJ-M-2-200A", "HJ-M-9-200A", "HJ-M-10-200A"]);
  });
});

describe("불량률", () => {
  function rec(partial: Pick<InspectionRecord, "point" | "result"> & Partial<InspectionRecord>): InspectionRecord {
    return {
      ...row({ line: "L", point: partial.point, result: partial.result, date: partial.date ?? "2026-07-15" }),
      id: partial.id ?? `${partial.line ?? "L"}-${partial.point}-${partial.result}-${partial.date ?? "2026-07-15"}-${partial.ts ?? 1}`,
      line: partial.line ?? "L",
      ts: partial.ts ?? 1,
      deleted: partial.deleted,
    };
  }

  it("재검사 합격은 한 번만 세고 불합격 행은 남긴다. 1/10 이지 1/11 이 아니다", () => {
    const records = [
      ...Array.from({ length: 9 }, (_, index) => rec({ point: `S${String(index + 1).padStart(2, "0")}`, result: "합격" })),
      rec({ id: "fail", point: "S10", result: "불합격", date: "2026-07-10", ts: 1 }),
      rec({ id: "again", point: "S10", result: "합격", date: "2026-07-15", ts: 2 }),
    ];
    expect(defectStats(records)).toEqual({ pass: 10, fail: 1, rate: "10%" });
  });

  it("하위 번호와 S·R 은 각자 1개소다", () => {
    const records = [
      rec({ point: "S01", result: "합격" }),
      rec({ point: "S01-1", result: "합격" }),
      rec({ point: "R01", result: "합격" }),
      rec({ point: "s1", result: "불합격", date: "2026-07-16", ts: 3 }),
    ];
    expect(defectStats(records)).toEqual({ pass: 2, fail: 1, rate: "50%" });
  });

  it("소수 한 자리만 남기고 .0 은 뺀다. 합격이 없으면 대시", () => {
    expect(defectStats([
      rec({ point: "S01", result: "합격" }),
      rec({ point: "S02", result: "합격" }),
      rec({ point: "S03", result: "합격" }),
      rec({ point: "S04", result: "불합격" }),
    ]).rate).toBe("33.3%");
    expect(defectStats([
      ...Array.from({ length: 8 }, (_, index) => rec({ point: `S${index + 1}`, result: "합격" })),
      rec({ point: "S09", result: "불합격" }),
    ]).rate).toBe("12.5%");
    expect(defectStats([rec({ point: "S01", result: "불합격" })])).toEqual({ pass: 0, fail: 1, rate: "–" });
    expect(defectStats([rec({ point: "S01", result: "합격", deleted: true })])).toEqual({ pass: 0, fail: 0, rate: "–" });
  });

  it("다른 라인의 같은 번호는 현장 전체에서 따로 센다", () => {
    const records = [
      rec({ line: "A", point: "S01", result: "합격" }),
      rec({ line: "B", point: "S01", result: "불합격", date: "2026-07-10", ts: 1 }),
      rec({ id: "b-again", line: "B", point: "S01", result: "합격", date: "2026-07-15", ts: 2 }),
    ];
    expect(defectStats(records)).toEqual({ pass: 2, fail: 1, rate: "50%" });
  });
});
