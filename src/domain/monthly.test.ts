import { describe, expect, it } from "vitest";
import { graphConfig, rec } from "../test/realBooks";
import { FAIL, PASS } from "./types";
import { monthSummary, monthsOf } from "./monthly";

describe("명세서 탭 · 달별 작업·금액", () => {
  const recs = [
    rec("S01", PASS, { date: "2026-11-12" }),
    rec("R01", PASS, { date: "2026-11-12" }),
    rec("S02", FAIL, { date: "2026-11-12" }),
    rec("S03", PASS, { date: "2026-12-01" }),
    rec("S10", PASS, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: "AJ-750A", diameter: "750A", date: "2026-11-20" }),
    rec("S11", PASS, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: "AJ-750A", diameter: "750A", date: "2026-11-20" }),
    rec("S12", PASS, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: "AJ-750A", diameter: "750A", date: "2026-11-21" }),
    rec("S13", PASS, { site: "지성 브레인", date: "2026-11-12" }),
  ];

  it("11월: 현장별 개소·금액(합정주공 원/m, 아전 군포 개소당 183,600), 불합격은 따로, 지운 현장은 빼고 합계는 현장별 합계를 더한 것", () => {
    const sum = monthSummary(recs, graphConfig(), "2026-11");
    const hj = sum.sites.find((s) => s.site === "드림 합정주공")!;
    // 350A 1.2m × 2개소 × 75,000 = 180,000 → +부가세 18,000 = 198,000
    expect(hj).toMatchObject({ count: 2, fail: 1, supply: 180000, grand: 198000 });
    expect(hj.lines).toEqual([{ date: "2026-11-12", diameter: "350A", count: 2, qtyM: 2.4, amount: 180000 }]);
    const aj = sum.sites.find((s) => s.site === "아전 군포 PA")!;
    // 3개소 × 183,600 = 550,800 → +55,080 = 605,880 → 천 원 미만 버림 605,000
    expect(aj).toMatchObject({ count: 3, supply: 550800, grand: 605000 });
    expect(aj.lines.map((l) => [l.date, l.count, l.amount])).toEqual([
      ["2026-11-20", 2, 367200],
      ["2026-11-21", 1, 183600],
    ]);
    expect(sum.sites.map((s) => s.site)).not.toContain("지성 브레인");
    expect(sum).toMatchObject({ count: 5, fail: 1, supply: 730800, grand: 803000 });
  });

  it("12월은 12월 기록만, 달 목록은 기록이 있는 달 + 이번 달(최근 먼저)", () => {
    const dec = monthSummary(recs, graphConfig(), "2026-12");
    expect(dec.sites.map((s) => [s.site, s.count])).toEqual([["드림 합정주공", 1]]);
    expect(monthsOf(recs, "2027-01-05")).toEqual(["2027-01", "2026-12", "2026-11"]);
  });
});
