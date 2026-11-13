/**
 * 1번 탭 '날짜별 기록'(검사 기록 아래에 더함): 날짜마다 그날 검사 내역(도면·개수·판정) + 작업일지의 근무자·근무시간.
 * 실제 작업일지·리스트의 로컬 사본만 쓴다. 네트워크 없음, 쓰기 없음. 기존 '검사 기록' 블록은 그대로인지 확인한다.
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { fixture, realListRecords } from "../test/realBooks";
import { unzipXlsx } from "../excel/ooxml";
import { readWorkLog } from "../excel/iljiBook";
import { rememberWork, workLogByDate, inspectorsOnDate } from "../excel/engine";
import { dayHistory, NO_WORKLOG_TEXT, workersText, workTimeText } from "./dayHistory";
import { FAIL, PASS } from "./types";

const SITE = "드림 합정주공";

describe("날짜별 기록 · 검사 내역 + 근무일지 근무자·시간", () => {
  it("날짜를 열면 그날 이 현장 검사 내역(도면·개수·판정)과 근무일지의 근무자·근무시간", () => {
    const files = unzipXlsx(fixture("real-worklog.xlsx"));
    const log = readWorkLog(files);
    const recs = realListRecords().filter((r) => r.site === SITE);
    const days = dayHistory(recs, log);
    const d = days.find((x) => x.date === "2026-10-21")!;
    const expected = recs.filter((r) => r.date === "2026-10-21" && !r.deleted);
    expect(d.total).toBe(expected.length);
    expect(d.pass).toBe(expected.filter((r) => r.result === PASS).length);
    expect(d.fail).toBe(expected.filter((r) => r.result === FAIL).length);
    // 도면(라인)별 개수 합 = 그날 개수, 라인마다 판정 개수
    expect(d.drawings.reduce((n, g) => n + g.count, 0)).toBe(d.total);
    for (const g of d.drawings) {
      const mine = expected.filter((r) => r.line.trim().toUpperCase() === g.line.toUpperCase());
      expect(g.count).toBe(mine.length);
      expect(g.pass + g.fail + g.other).toBe(g.count);
    }
    expect(d.work).not.toBeNull();
    expect(workersText(d.work)).toBe("조우현 · 심명보");
    expect(workTimeText(d.work)).toBe("09:00–20:00");
    // 최근 날짜가 먼저
    expect(days.map((x) => x.date)).toEqual([...days.map((x) => x.date)].sort().reverse());
  });

  it("근무일지에 그날이 없어도 검사 내역은 보이고, 근무자·시간 없음 안내", () => {
    const log = readWorkLog(unzipXlsx(fixture("real-worklog.xlsx")));
    expect(log["2026-11-11"]).toBeUndefined();
    const recs = realListRecords().filter((r) => r.site === SITE);
    const d = dayHistory(recs, log).find((x) => x.date === "2026-11-11")!;
    expect(d.total).toBe(recs.filter((r) => r.date === "2026-11-11").length);
    expect(d.total).toBeGreaterThan(0);
    expect(d.work).toBeNull();
    expect(NO_WORKLOG_TEXT).toMatch(/근무자.*근무시간이 없습니다/);
  });

  it("이 현장에 검사가 없는 날(작업일지에만 있는 날·다른 현장 근무)은 보이지 않는다. 지운 기록은 세지 않는다", () => {
    const log = readWorkLog(unzipXlsx(fixture("real-worklog.xlsx")));
    const recs = realListRecords().filter((r) => r.site === SITE);
    const days = dayHistory(recs, log);
    const inspected = new Set(recs.filter((r) => !r.deleted && /^\d{4}-\d{2}-\d{2}$/.test(r.date)).map((r) => r.date));
    expect(days.every((d) => d.total > 0)).toBe(true);
    expect(new Set(days.map((d) => d.date))).toEqual(inspected);
    // 작업일지엔 있지만 합정주공 검사는 없는 날(예: 10/22 지성 브레인 근무)
    const workOnly = Object.keys(log).filter((d) => !inspected.has(d));
    expect(workOnly).toContain("2026-10-22");
    for (const d of workOnly) expect(days.some((x) => x.date === d)).toBe(false);
    // 다른 현장만 검사한 날도 없다
    const otherSite = realListRecords().filter((r) => r.site !== SITE).map((r) => r.date).filter((d) => !inspected.has(d));
    for (const d of otherSite) expect(days.some((x) => x.date === d)).toBe(false);
    const one = recs.find((r) => r.date === "2026-10-21")!;
    const before = days.find((d) => d.date === "2026-10-21")!.total;
    const after = dayHistory(recs.map((r) => (r === one ? { ...r, deleted: true } : r)), log).find((d) => d.date === "2026-10-21")!.total;
    expect(after).toBe(before - 1);
  });

  it("날짜를 열면 그날 검사한 포인트 번호가 도면(라인)별로, 번호순, 불합격은 표시(fail)", () => {
    const log = readWorkLog(unzipXlsx(fixture("real-worklog.xlsx")));
    const recs = realListRecords().filter((r) => r.site === SITE);
    for (const d of dayHistory(recs, log)) {
      for (const g of d.drawings) {
        const mine = recs.filter((r) => r.date === d.date && !r.deleted && r.line.trim().toUpperCase() === g.line.toUpperCase());
        expect(g.points.map((p) => p.point).sort()).toEqual(mine.map((r) => r.point.trim()).sort());
        expect(g.points.map((p) => p.point)).toEqual([...g.points.map((p) => p.point)].sort((a, b) => a.localeCompare(b, "ko", { numeric: true })));
        expect(g.points.filter((p) => p.fail).map((p) => p.point).sort()).toEqual([...g.failPoints].sort());
        expect(g.points.filter((p) => p.fail).length).toBe(g.fail);
      }
    }
    // 불합격이 있는 날이 적어도 하나 있고, 그 포인트는 fail 로 표시된다.
    const failDay = dayHistory(recs, log).find((d) => d.fail > 0)!;
    expect(failDay).toBeTruthy();
    const failed = recs.filter((r) => r.date === failDay.date && r.result === FAIL).map((r) => r.point.trim());
    expect(failDay.drawings.flatMap((g) => g.points.filter((p) => p.fail).map((p) => p.point)).sort()).toEqual(failed.sort());
    // 근무자·시간도 그대로(작업일지)
    const d1021 = dayHistory(recs, log).find((d) => d.date === "2026-10-21")!;
    expect(d1021.drawings.flatMap((g) => g.points).length).toBe(d1021.total);
    expect(workersText(d1021.work)).toBe("조우현 · 심명보");
  });

  it("읽기만 한다: 기록·작업일지 사본을 바꾸지 않는다(엑셀에 쓰지 않음)", () => {
    const files = unzipXlsx(fixture("real-worklog.xlsx"));
    const snapshot = Object.fromEntries(Object.entries(files).map(([k, v]) => [k, Buffer.from(v).toString("base64")]));
    const recs = realListRecords();
    const recSnap = structuredClone(recs);
    rememberWork("ilji", fixture("real-worklog.xlsx"));
    const log = workLogByDate();
    dayHistory(recs, log);
    readWorkLog(files);
    expect(recs).toEqual(recSnap);
    expect(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, Buffer.from(v).toString("base64")]))).toEqual(snapshot);
    // 기기 작업일지 사본에서 읽은 값 = 기존 검사자 경로와 같은 사람
    const staff = inspectorsOnDate("2026-10-21")!;
    expect(log["2026-10-21"].workers).toEqual([staff.insp1, staff.insp2].filter(Boolean));
  });
});

describe("기존 '검사 기록' 블록은 그대로 두고 날짜별 기록을 그 아래에 더한다", () => {
  const src = readFileSync(new URL("../ui/EntryView.tsx", import.meta.url), "utf8");

  it("검사 기록(라인별·최신순/포인트순) 블록과 데이터 경로가 남아 있다", () => {
    expect(src).toContain('<span>검사 기록</span>');
    expect(src).toContain('className="history"');
    expect(src).toContain("historyOfLine(lineRecords, historySort");
    expect(src).toContain("이 라인의 검사 기록이 없습니다.");
    expect(src).toContain('onClick={() => chooseHistorySort("recent")}>최신순');
    expect(src).toContain('onClick={() => chooseHistorySort("point")}>포인트순');
    expect(src).toContain("inspectorLabel(row.date, meta)");
  });

  it("날짜별 기록은 검사 기록 뒤에 따로 들어간다(대신하지 않음)", () => {
    const history = src.indexOf('<span>검사 기록</span>');
    const day = src.indexOf("<DayHistory ");
    expect(day).toBeGreaterThan(history);
    expect(src.match(/<DayHistory /g)).toHaveLength(1);
  });
});
