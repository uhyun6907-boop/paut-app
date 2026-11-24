/**
 * 실제 사고(2026-11-20 지성 마전): 11/17 에 쓰던 작업 날짜가 그대로 남아 11/20 에 넣은 기록이 모두 11/17 로 엑셀에 들어갔다.
 * 앱을 새 날에 켜면 작업 날짜를 오늘로 맞추고, 지난 날짜에 저장 안 한 기록이 있으면 묻는다.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
vi.mock("./db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => void mem.set(key, structuredClone(value)),
  kvDel: async (key: string) => void mem.delete(key),
}));
vi.stubGlobal("localStorage", {
  getItem: (k: string) => bag.get(k) ?? null,
  setItem: (k: string, v: string) => void bag.set(k, String(v)),
  removeItem: (k: string) => void bag.delete(k),
});

import { PASS, type DayMeta, type DayStaff, type InspectionRecord } from "../domain/types";
import { fixture } from "../test/realBooks";

type Session = typeof import("./session");
const OLD = "2026-11-17";
const TODAY = "2026-11-20";
const SITE = "지성 마전";
const LINE = "MJ-LINE-1";

function onDay(iso: string, hour = 8): void {
  const [y, m, d] = iso.split("-").map(Number);
  vi.setSystemTime(new Date(y, m - 1, d, hour));
}

async function boot(): Promise<Session> {
  vi.resetModules();
  const s = await import("./session");
  await s.initSession();
  await s.settled();
  return s;
}

function rec(point: string, extra: Partial<InspectionRecord> = {}): InspectionRecord {
  return {
    id: crypto.randomUUID(), site: SITE, sheetName: SITE, date: OLD, line: LINE, point, result: PASS,
    diameter: "", material: "", sch: "", reportNo: "", reportDate: "", defectNo: "", drawingNo: "", weldNo: "",
    rtSheets: "", flawType: "", flawLen: "", note: "", issued: "", src: "user", sync: "synced", ts: 1, ...extra,
  };
}

function seed(records: InspectionRecord[], meta: Partial<DayMeta> = {}): void {
  mem.set("records", records);
  mem.set("meta", { lastSite: SITE, activeLine: LINE, workDate: OLD, dayType: "평일", prefix: "S", insp1: "조우현", insp2: "심명보", timeStart: "09:00", timeEnd: "18:00", draftCards: {}, ...meta });
  mem.set("book:list", fixture("real-list.xlsx"));
  mem.set("book:ilji", fixture("real-worklog.xlsx"));
  mem.set("book:statement", fixture("real-statement.xlsx"));
}

function byPoint(s: Session, point: string): InspectionRecord {
  return s.getState().records.find((r) => r.point === point && !r.deleted)!;
}

beforeEach(() => {
  mem.clear();
  bag.clear();
  vi.useFakeTimers({ toFake: ["Date"] });
});
afterEach(() => vi.useRealTimers());

describe("작업 날짜가 오늘로 넘어간다", () => {
  it("지난 날짜에 저장 안 한 기록이 없으면 켜자마자 오늘로, 그날 남긴 검사자·시간이 있으면 그 값", async () => {
    const todayStaff: DayStaff = { dayType: "평일", insp1: "이연행", insp2: "", timeStart: "07:30", timeEnd: "16:00" };
    seed([rec("S01")], { dayStaff: { [TODAY]: todayStaff } });
    onDay(TODAY);
    const s = await boot();
    const { meta, staleWorkDate } = s.getState();
    expect(meta.workDate).toBe(TODAY);
    expect(meta).toMatchObject({ insp1: "이연행", insp2: "", timeStart: "07:30", timeEnd: "16:00" });
    expect(staleWorkDate).toBeNull();
    expect(byPoint(s, "S01").date).toBe(OLD);
    expect((mem.get("meta") as DayMeta).workDate).toBe(TODAY);
  });

  it("그날 값이 없으면 검사자는 조우현 한 명(검사자 2 비움), 시간은 그대로, 구분은 오늘 요일로(토요일 → 공휴일)", async () => {
    seed([rec("S01")], { dayType: "평일", insp1: "김정환", insp2: "심명보" });
    onDay("2026-11-21");
    const s = await boot();
    expect(s.getState().meta).toMatchObject({ workDate: "2026-11-21", dayType: "공휴일", insp1: "조우현", insp2: "", timeStart: "09:00" });
  });

  it("지난 날짜에 저장 안 한 기록이 있으면 바꾸지 않고 묻는다. 「오늘로 바꾸기」는 그 미저장 기록만 오늘로 옮긴다", async () => {
    seed([
      rec("S01"),
      rec("S02", { sync: "local" }),
      rec("R02", { sync: "local" }),
      rec("S03", { sync: "local", date: "2026-11-16" }),
    ]);
    onDay(TODAY);
    const s = await boot();
    expect(s.getState().meta.workDate).toBe(OLD);
    expect(s.getState().staleWorkDate).toEqual({ workDate: OLD, today: TODAY });

    await s.checkWorkDate();
    expect(s.getState().meta.workDate).toBe(OLD);

    await s.moveStaleToToday();
    const st = s.getState();
    expect(st.meta.workDate).toBe(TODAY);
    expect(st.staleWorkDate).toBeNull();
    expect(["S01", "S02", "R02", "S03"].map((p) => [p, byPoint(s, p).date, byPoint(s, p).sync])).toEqual([
      ["S01", OLD, "synced"],
      ["S02", TODAY, "local"],
      ["R02", TODAY, "local"],
      ["S03", "2026-11-16", "local"],
    ]);
    expect((mem.get("records") as InspectionRecord[]).find((r) => r.point === "S02")?.date).toBe(TODAY);
  });

  it("「11/17 그대로」는 오늘 하루 그 날짜를 쓴다. 다시 켜도 묻지 않고, 다음 날 켜면 다시 묻는다", async () => {
    seed([rec("S02", { sync: "local" })]);
    onDay(TODAY);
    let s = await boot();
    await s.keepStaleWorkDate();
    expect(s.getState().meta.workDate).toBe(OLD);
    expect(s.getState().staleWorkDate).toBeNull();
    expect(byPoint(s, "S02").date).toBe(OLD);

    onDay(TODAY, 15);
    s = await boot();
    await s.checkWorkDate();
    expect(s.getState().meta.workDate).toBe(OLD);
    expect(s.getState().staleWorkDate).toBeNull();

    onDay("2026-11-21");
    await s.checkWorkDate();
    expect(s.getState().staleWorkDate).toEqual({ workDate: OLD, today: "2026-11-21" });
  });

  it("사용자가 고른 지난 날짜는 같은 날 화면으로 돌아와도 그대로, 다음 날 처음 돌아오면 오늘로", async () => {
    seed([rec("S01")]);
    onDay(TODAY);
    const s = await boot();
    expect(s.getState().meta.workDate).toBe(TODAY);

    await s.updateMeta({ workDate: "2026-11-18" });
    onDay(TODAY, 20);
    await s.checkWorkDate();
    expect(s.getState().meta.workDate).toBe("2026-11-18");

    onDay("2026-11-21");
    await s.checkWorkDate();
    expect(s.getState().meta.workDate).toBe("2026-11-21");
  });
});

describe("오늘이 아닌 날짜로 저장", () => {
  it("한 번 묻고, 취소하면 아무것도 보내지 않는다. 확인하면 저장한다", async () => {
    seed([rec("S02", { sync: "local" })]);
    onDay(TODAY);
    const s = await boot();
    await s.keepStaleWorkDate();

    const asked: string[] = [];
    await s.commitDay((text) => {
      asked.push(text);
      return false;
    });
    await s.settled();
    expect(asked).toEqual(["11/17 날짜로 저장할까요? 오늘은 11/20 이에요."]);
    expect(byPoint(s, "S02").sync).toBe("local");
    expect(s.syncQueue(s.getState().records)).toHaveLength(0);

    await s.commitDay((text) => {
      asked.push(text);
      return true;
    });
    await s.settled();
    expect(asked).toHaveLength(2);
    expect(byPoint(s, "S02")).toMatchObject({ date: OLD });
    expect(byPoint(s, "S02").sync).not.toBe("local");
  });

  it("작업 날짜가 오늘이면 묻지 않는다", async () => {
    seed([rec("S02", { sync: "local", date: TODAY })], { workDate: TODAY });
    onDay(TODAY);
    const s = await boot();
    const asked: string[] = [];
    await s.commitDay((text) => {
      asked.push(text);
      return false;
    });
    await s.settled();
    expect(asked).toEqual([]);
    expect(byPoint(s, "S02").sync).not.toBe("local");
  });
});
