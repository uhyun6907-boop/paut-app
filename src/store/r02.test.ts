/**
 * "r02 불꺼졌잖아": 저장(엑셀에 들어간) R 표시는 초기화·다시 켜기·S→R 자동 뒤에도 꺼지면 안 된다.
 * - 초기화 배포 전에 미저장(local)으로 남은 기록은 저장 때 모습(메모리 사본)이 없다 → 새 카드로 잘못 보고 초기화가 지웠다.
 * - 엑셀에서 읽은 'R2'(0 없음)와 앱의 'R02'(두 자리)는 같은 포인트다.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

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

import { FAIL, PASS, type InspectionRecord } from "../domain/types";
import { fixture } from "../test/realBooks";

type Session = typeof import("./session");
const DATE = "2026-11-12";
const SITE = "드림 합정주공";
const LINE = "R02-LINE-1";

async function boot(): Promise<Session> {
  vi.resetModules();
  const s = await import("./session");
  await s.initSession();
  await s.settled();
  return s;
}

function base(point: string, result: string, extra: Partial<InspectionRecord> = {}): InspectionRecord {
  return {
    id: crypto.randomUUID(), site: SITE, sheetName: SITE, date: DATE, line: LINE, point, result,
    diameter: "", material: "", sch: "", reportNo: "", reportDate: "", defectNo: "", drawingNo: "", weldNo: "",
    rtSheets: "", flawType: "", flawLen: "", note: "", issued: "", src: "user", sync: "synced", ts: 1, ...extra,
  };
}

const sent = { list: { status: "done" as const, at: 1 }, ilji: { status: "done" as const, at: 1 }, statement: { status: "done" as const, at: 1 } };

/** 화면의 R 카드(현장·라인, 지운 것 제외) → '번호:결과'. */
function rCards(s: Session): string[] {
  return s.getState().records
    .filter((r) => r.site === SITE && r.line === LINE && !r.deleted && r.point.startsWith("R"))
    .map((r) => `${r.point}:${r.result}`)
    .sort();
}

function seed(records: InspectionRecord[], unsaved: InspectionRecord[] = []): void {
  mem.set("records", records);
  mem.set("meta", { lastSite: SITE, activeLine: LINE, workDate: DATE, prefix: "S", insp1: "조우현", insp2: "심명보", draftCards: {} });
  mem.set("book:list", fixture("real-list.xlsx"));
  mem.set("book:ilji", fixture("real-worklog.xlsx"));
  mem.set("book:statement", fixture("real-statement.xlsx"));
  bag.set("paut-sr-mirror", JSON.stringify({ enabled: true, independent: [] }));
  if (unsaved.length) bag.set("paut-unsaved-v1", JSON.stringify({ records: unsaved, draftCards: {}, stamps: {} }));
}

describe("R02 저장 표시 유지", () => {
  beforeEach(() => {
    mem.clear();
    bag.clear();
  });

  it("초기화 배포 전부터 미저장(local)으로 남은, 엑셀에 이미 들어간 R02 는 다시 켠 뒤 초기화해도 지워지지 않는다", async () => {
    const s02 = base("S02", PASS, { sync: "local", excelRow: 440, steps: sent, ts: 10 });
    const r02 = base("R02", FAIL, { sync: "local", excelRow: 441, steps: sent, ts: 11 });
    seed([s02, r02], [s02, r02]);
    // 초기화 기억(paut-unsaved-originals-v1)이 없는 상태로 켠다(배포 직후).
    let s = await boot();
    expect(rCards(s)).toEqual(["R02:불합격"]);
    await s.resetUnsaved();
    await s.settled();
    expect(rCards(s)).toEqual(["R02:불합격"]);
    s = await boot();
    expect(rCards(s)).toEqual(["R02:불합격"]);
    await s.resetUnsaved();
    await s.settled();
    expect(rCards(s)).toEqual(["R02:불합격"]);
  });

  it("엑셀의 'R2'(0 없음)와 앱의 'R02'는 같은 포인트: S→R 자동·초기화·다시 켜기 뒤에도 R 하나가 저장 결과로 켜져 있다", async () => {
    const s2 = base("S2", PASS, { id: "xlsx:드림 합정주공:300", src: "xlsx", excelRow: 300, ts: 300 });
    const r2 = base("R2", PASS, { id: "xlsx:드림 합정주공:301", src: "xlsx", excelRow: 301, ts: 301 });
    seed([s2, r2]);
    let s = await boot();
    expect(rCards(s)).toEqual(["R2:합격"]);

    // S2 를 불합격으로 잘못 누름 → R2 가 따라감(같은 기록, 새 R 카드 없음)
    await s.setResult(s2.id, FAIL);
    await s.settled();
    expect(rCards(s)).toEqual(["R02:불합격"]);
    expect(s.getState().records.find((r) => r.id === r2.id)?.result).toBe(FAIL);

    // 다시 켜도 미저장 표시는 남고, 초기화하면 저장 때 모습(R2 합격)으로
    s = await boot();
    expect(rCards(s)).toEqual(["R02:불합격"]);
    await s.resetUnsaved();
    await s.settled();
    expect(rCards(s)).toEqual(["R2:합격"]);
    expect(s.getState().records.find((r) => r.id === r2.id)?.sync).toBe("synced");

    // 다시 켜고 초기화를 또 눌러도 꺼지지 않는다
    s = await boot();
    await s.resetUnsaved();
    await s.settled();
    expect(rCards(s)).toEqual(["R2:합격"]);

    // S→R 자동: S2 를 다시 합격(같은 결과) → R2 는 그대로 저장 상태
    await s.setResult(s2.id, PASS);
    await s.settled();
    expect(rCards(s)).toEqual(["R2:합격"]);
  });

  it("저장한 R02 는 초기화·다시 켜기에 그대로이고, 저장 안 한 탭만 지운다", async () => {
    seed([]);
    let s = await boot();
    await s.savePoint(PASS, "S2");
    await s.commitDay();
    await s.settled();
    expect(rCards(s)).toEqual(["R02:합격"]);
    await s.savePoint(FAIL, "S3"); // 저장 안 한 탭
    await s.settled();
    expect(rCards(s)).toEqual(["R02:합격", "R03:불합격"]);
    s = await boot();
    await s.resetUnsaved();
    await s.settled();
    expect(rCards(s)).toEqual(["R02:합격"]);
    s = await boot();
    expect(rCards(s)).toEqual(["R02:합격"]);
  });
});
