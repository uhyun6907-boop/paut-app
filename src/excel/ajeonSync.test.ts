/** 아전 군포 거래명세서는 01(아전·승재) 파일의 아전 시트에 쓴다. 접두어·단가는 그 파일에서 읽는다. 가짜 OneDrive 로만 확인한다. */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import { STMT_AJ_ID, STMT_ID, graphConfig, meta, realDrive, rec } from "../test/realBooks";
import { detectAjStatement, restoreStatementOwnership, statementBookOf, syncGraph } from "./engine";
import { readPautRows } from "./statementBook";
import { userCreatesMonthSheet } from "../test/userSheet";
import { unzipXlsx } from "./ooxml";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});
beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
}

const aj = (point: string, partial: Partial<InspectionRecord> = {}) =>
  rec(point, PASS, { site: "아전 군포 PA", sheetName: "아전 군포 PA", line: "AJ-750A", diameter: "750A", date: "2026-11-03", steps: { list: { status: "done" }, ilji: { status: "done" } }, ...partial });

describe("아전 군포 거래명세서", () => {
  it("아전 군포는 01 파일, 다른 현장은 02 파일(예전 설정에 statementBook 이 없어도)", () => {
    const old = graphConfig();
    old.mappings = old.mappings.map((m) => ({ ...m, statementBook: undefined }));
    expect(statementBookOf(old, "아전 군포 PA")).toBe("statementAj");
    expect(statementBookOf(old, "아전 군포")).toBe("statementAj");
    expect(statementBookOf(old, "지성 군포 PA")).toBe("statement");
  });

  it("01 파일의 최근 아전 시트에서 시트 이름 앞부분과 단가를 읽는다", () => {
    const d = realDrive();
    expect(detectAjStatement(d.files(STMT_AJ_ID))).toEqual({ prefix: "아전군포", unitPrice: 75000 });
  });

  it("아전 군포 기록은 01 파일(이름으로 찾음)의 그 달 아전 시트에 쓰고, 02 파일에는 쓰지 않는다", async () => {
    const d = realDrive();
    // 사용자가 PC 엑셀에서 12월 아전 시트를 만들어 둔 상태.
    const dec = userCreatesMonthSheet(d.files(STMT_AJ_ID), "아전군포", 2026, 12);
    const jobs = [aj("S1", { date: "2026-12-03" }), aj("S2", { date: "2026-12-03" }), aj("S3", { date: "2026-12-05" })];
    const out = await syncGraph({ client: client(d), jobs, all: jobs, config: graphConfig(), meta: meta(), sleep: async () => undefined });
    for (const r of jobs) expect(out.steps.get(r.id)?.statement).toMatchObject({ status: "done" });
    const rows = readPautRows(d.files(STMT_AJ_ID), dec).map((r) => `${r.date}|${r.diameter}|${r.count}`);
    expect(dec).toBe("아전군포 2612");
    expect(rows).toEqual(["2026-12-03|750A|2", "2026-12-05|750A|1"]);
    expect(d.requests((e) => e.item === STMT_ID && e.method !== "GET")).toHaveLength(0);
    expect(out.report.files.statement.status).toBe("ok");
  });

  it("그 달 아전 시트가 없으면 만들지 않고 대기, 01 파일에 아전 시트가 아예 없으면 안내하고 대기", async () => {
    const d = realDrive();
    const jobs = [aj("S1", { date: "2026-12-03" })];
    const out = await syncGraph({ client: client(d), jobs, all: jobs, config: graphConfig(), meta: meta(), sleep: async () => undefined });
    expect(out.steps.get(jobs[0].id)?.statement).toMatchObject({ status: "wait", message: "12월 거래명세서 시트를 만들어주세요 (아전군포 2612)" });

    const none = realDrive({ ajeon: false });
    const plain = unzipXlsx(none.bytes(STMT_ID));
    none.put(STMT_AJ_ID, "01_PAUT 거래명세서 (아전, 승재).xlsx", none.bytes(STMT_ID));
    expect(detectAjStatement(plain)).toBeNull();
    const out2 = await syncGraph({ client: client(none), jobs: [aj("S9")], all: [aj("S9")], config: graphConfig(), meta: meta(), sleep: async () => undefined });
    expect([...out2.steps.values()][0]?.statement).toMatchObject({ status: "wait" });
    expect([...out2.steps.values()][0]?.statement?.message).toContain("아전 시트");
    expect(none.requests((e) => e.method !== "GET" && (e.item === STMT_AJ_ID || e.item === STMT_ID))).toHaveLength(0);
  });

  it("동기화 상태 화면의 거래명세서는 한 줄이다", () => {
    const r = aj("S1");
    r.sync = "pending";
    const model = syncPanelModel([r], null, { online: true, syncing: false, mode: "graph" });
    expect(model.lines.filter((l) => l.file === "statement")).toHaveLength(1);
  });
});

describe("예전 판이 건너뛴 아전 군포 명세서", () => {
  it("앱을 켜면 아전 군포 기록의 명세서 '건너뜀' 표시만 지운다(다른 현장·다른 단계는 그대로)", async () => {
    const { reopenAjeonStatement } = await import("../store/session");
    const a = aj("S1", { steps: { list: { status: "done" }, ilji: { status: "done" }, statement: { status: "skip", message: "접두어" } } });
    const other = rec("S2", PASS, { steps: { statement: { status: "skip" } } });
    expect(reopenAjeonStatement([a, other])).toBe(true);
    expect(a.steps).toMatchObject({ list: { status: "done" }, ilji: { status: "done" } });
    expect(a.steps?.statement).toBeUndefined();
    expect(other.steps?.statement?.status).toBe("skip");
    expect(reopenAjeonStatement([a, other])).toBe(false);
  });
});
