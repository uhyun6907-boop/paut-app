import { describe, expect, it } from "vitest";
import { findMonthSheet, latestSheet, matchStatementSheet, suggestSheetName, defaultMappings, isRemovedSite } from "./mapping";

const names = [
  "지성군포 26.10",
  "지성군포 2611",
  "지성브레인26.6",
  "지성브레인26.7",
  "지성브레인2610",
  "지성브레인2611",
  "지성 마전 26.8",
  "지성 마전 26.5 AM",
  "드림합정26.10",
  "드림화양 26.10",
];

describe("거래명세서 시트 이름", () => {
  it("점 있음, 점 없음, 앞 공백, AM을 구분한다", () => {
    expect(matchStatementSheet("드림합정26.10", "드림합정")).toMatchObject({ year: 2026, month: 10 });
    expect(matchStatementSheet("지성브레인2611", "지성브레인")).toMatchObject({ year: 2026, month: 11, style: { dot: "", monthWidth: 2 } });
    expect(matchStatementSheet("지성군포 26.4", "지성군포")?.style.gap).toBe(" ");
    expect(matchStatementSheet("지성 마전 26.5 AM", "지성 마전")?.suffix).toBe(" AM");
    expect(matchStatementSheet("드림합정26.10", "지성브레인")).toBeNull();
  });

  it("최근 시트 표기를 따라 다음 달을 제안하고 AM 시트는 건너뛴다", () => {
    const brain = latestSheet(names, "지성브레인");
    expect(brain?.name).toBe("지성브레인2611");
    expect(suggestSheetName("지성브레인", brain, 2026, 12)).toBe("지성브레인2612");
    const gp = latestSheet(names, "지성군포");
    expect(gp?.name).toBe("지성군포 2611");
    expect(suggestSheetName("지성군포", gp, 2026, 12)).toBe("지성군포 2612");
    const maj = latestSheet(names, "지성 마전");
    expect(maj?.name).toBe("지성 마전 26.8");
    expect(findMonthSheet(names, "드림합정", 2026, 10)?.name).toBe("드림합정26.10");
    expect(findMonthSheet(names, "지성 마전", 2026, 5)).toBeNull();
  });
});

describe("앱에서 지운 현장", () => {
  it("지성 브레인·택지 드림(드림 택지)·드림 화양·브레인대지 인입은 기본 현장 목록에 없고, 다른 현장은 남는다", () => {
    const sites = defaultMappings().map((m) => m.site);
    expect(sites).not.toContain("지성 브레인");
    expect(sites).not.toContain("택지 드림");
    expect(sites).not.toContain("드림 화양");
    expect(sites).not.toContain("브레인대지 인입");
    expect(sites).toEqual(["드림 합정주공", "지성 군포 PA", "지성 마전 PA", "아전 군포 PA"]);
    for (const name of ["지성 브레인", " 지성 브레인", "지성브레인", "택지 드림", "드림 택지", "택지드림", "드림 화양", "드림화양", "브레인대지 인입", " 브레인대지 인입"]) expect(isRemovedSite(name)).toBe(true);
    for (const name of ["드림 합정주공", "지성 군포 PA", "지성 마전 PA", "아전 군포 PA", "", undefined]) expect(isRemovedSite(name)).toBe(false);
  });
});
