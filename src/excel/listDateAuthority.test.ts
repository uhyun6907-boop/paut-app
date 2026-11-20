/**
 * 리스트 E(날짜)가 카드 날짜보다 우선(2026-11-21).
 * 11/20 지성 마전 카드가 11/17 로 저장됐고 엑셀을 손으로 11/20 으로 고쳤는데, 앱을 켜자 백필이 카드의 옛 날짜(11/17)로
 * 근무일지 11/17 행을 다시 채웠다. 이제 앱을 켜면 리스트 Y 행의 날짜를 카드에 맞추고, 그 기록으로는 아무것도 쓰지 않는다.
 * 실제 파일 복사본 + 가짜 Graph 로만 시험한다.
 */
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
vi.mock("../store/db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => void mem.set(key, structuredClone(value)),
  kvDel: async (key: string) => void mem.delete(key),
}));
vi.stubGlobal("window", { setTimeout, clearTimeout });
vi.stubGlobal("localStorage", { getItem: () => null, setItem: () => undefined, removeItem: () => undefined });

import { isoToExcelSerial } from "../domain/dates";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import type { FakeDrive } from "../test/fakeGraph";
import { ILJI_ID, LIST_GM_ID, LIST_ID, STMT_ID, graphConfig, meta, realDrive, realGmRecords, realListRecords, rec } from "../test/realBooks";
import { userCreatesMonthSheet } from "../test/userSheet";
import { getState, pullFromList } from "../store/session";
import { restoreStatementOwnership, syncGraph } from "./engine";
import { readIljiDay } from "./iljiBook";
import { loadLiveList } from "./liveList";

const MJ = "MJDH-D-D-P-58-003-200A";
const at = 1;
let drive: FakeDrive;
const client = () => new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
const writesTo = (item: string, from: number) => drive.log.slice(from).filter((e) => e.item === item && e.method !== "GET" && !/\/(create|close)Session$/.test(e.path));
const iljiDay = (date: string) => readIljiDay(drive.files(ILJI_ID), date, "지성 마전");
const dayMeta = () =>
  meta({
    workDate: "2026-11-20",
    lastSite: "지성 마전 PA",
    dayStaff: {
      "2026-11-17": { dayType: "평일", insp1: "조우현", insp2: "심명보", timeStart: "08:00", timeEnd: "17:00" },
      "2026-11-20": { dayType: "평일", insp1: "조우현", insp2: "김정환", timeStart: "09:00", timeEnd: "15:30" },
    },
  });

function cards(date: string): InspectionRecord[] {
  return ["S17", "S18", "S19", "S20", "R17", "R18", "R19", "R20"].map((point) =>
    rec(point, PASS, { site: "지성 마전 PA", sheetName: "지성 마전 PA", listBook: "listGm", line: MJ, diameter: "200A", material: "", date }),
  );
}

async function sync(jobs: InspectionRecord[], all: InspectionRecord[], backfill = false) {
  return syncGraph({ client: client(), jobs, all, config: graphConfig(), meta: dayMeta(), backfill, sleep: async () => undefined });
}

/** IndexedDB 에 남은 옛 카드: 리스트는 보냈고 작업일지·명세서는 백필 확인 대상. */
function staleSynced(recs: InspectionRecord[], date: string): void {
  for (const r of recs) Object.assign(r, { date, sync: "synced", steps: { list: { status: "done", at } } });
}

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});
beforeEach(() => {
  restoreStatementOwnership({});
  drive = realDrive();
  userCreatesMonthSheet(drive.files(STMT_ID), "지성 마전", 2026, 11);
});

describe("리스트 E(날짜)가 카드 날짜보다 우선", () => {
  it("엑셀 E=11/20, 카드 11/17(작업일지 11/17 마전 칸 비어 있음): 앱을 켜도 작업일지·명세서·리스트에 아무것도 쓰지 않고 카드 날짜를 11/20 으로 맞춘다", async () => {
    const recs = cards("2026-11-20");
    const first = await sync(recs, [...realGmRecords(), ...recs]);
    expect(first.report.files.ilji.status).toBe("ok");
    expect(first.report.files.statement.status).toBe("ok");
    expect(iljiDay("2026-11-20")).toEqual([{ diameter: "200A", count: 8 }]);
    expect(iljiDay("2026-11-17")).toBeNull();

    staleSynced(recs, "2026-11-17");
    const logAt = drive.log.length;
    const out = await sync([], [...realGmRecords(), ...recs], true);

    expect(writesTo(ILJI_ID, logAt)).toEqual([]);
    expect(writesTo(STMT_ID, logAt)).toEqual([]);
    expect(writesTo(LIST_GM_ID, logAt)).toEqual([]);
    expect(iljiDay("2026-11-17")).toBeNull();
    expect(iljiDay("2026-11-20")).toEqual([{ diameter: "200A", count: 8 }]);
    expect(out.pulled.map((p) => [p.id, p.patch.date])).toEqual(recs.map((r) => [r.id, "2026-11-20"]));
    for (const r of recs) {
      expect(r.date).toBe("2026-11-20");
      expect(r.movedFrom).toBeUndefined();
    }
  });

  it("엑셀 E 가 카드 날짜와 같고 작업일지 그날 칸이 비어 있으면 예전처럼 백필로 채운다", async () => {
    const recs = cards("2026-11-17");
    drive.fail((m, p) => m === "PATCH" && p.includes(ILJI_ID), 423, Infinity);
    await sync(recs, [...realGmRecords(), ...recs]);
    drive.clearFaults();
    expect(iljiDay("2026-11-17")).toBeNull();

    staleSynced(recs, "2026-11-17");
    for (const r of recs) r.steps = { ...r.steps, statement: { status: "done", at } };
    const out = await sync([], [...realGmRecords(), ...recs], true);

    expect(out.pulled.filter((p) => p.patch.date)).toEqual([]);
    expect(out.report.files.ilji.status).toBe("ok");
    expect(iljiDay("2026-11-17")).toEqual([{ diameter: "200A", count: 8 }]);
    expect(recs.every((r) => r.steps?.ilji?.status === "done")).toBe(true);
  });
});

describe("리스트 탭 pullFromList", () => {
  it("대기 카드도 Y 가 같은 행의 날짜로 맞추고 movedFrom·단계·상태는 그대로", async () => {
    const S22 = "62980e41-3f0b-436e-9a5d-0be156928211";
    const records = [...realListRecords(), ...realGmRecords()];
    const card = records.find((r) => r.id === S22)!;
    const steps = { list: { status: "done" as const, at } };
    Object.assign(card, { sync: "pending", steps });
    getState().records.splice(0, getState().records.length, ...records);
    const c = client();
    const sheetId = (await c.worksheetsOf(LIST_ID)).find((w) => w.name.trim() === "드림 합정주공")!.id;
    await c.patchValues(LIST_ID, sheetId, "E433", [[isoToExcelSerial("2026-11-09")]]);
    const res = await loadLiveList(c, graphConfig(), "드림 합정주공", getState().records);
    if (!res.ok) throw new Error(res.message);

    await pullFromList(res.table, "드림 합정주공");

    const after = getState().records.find((r) => r.id === S22)!;
    expect(after).toMatchObject({ date: "2026-11-09", sync: "pending", steps });
    expect(after.movedFrom).toBeUndefined();
  });
});
