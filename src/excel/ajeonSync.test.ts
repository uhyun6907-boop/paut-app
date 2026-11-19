/** 아전 군포는 거래명세서에 시트가 없어(접두어 미확정) 명세서 단계를 건너뛴다. 가짜 OneDrive 로만 확인한다. */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import { STMT_ID, graphConfig, meta, realDrive, rec } from "../test/realBooks";
import { restoreStatementOwnership, statementSkipReason, syncGraph } from "./engine";

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
  it("아전 군포 기록은 명세서를 건너뛰고 거래명세서 파일에 쓰지 않는다", async () => {
    const d = realDrive();
    const jobs = [aj("S1"), aj("S2", { date: "2026-11-05" })];
    const out = await syncGraph({ client: client(d), jobs, all: jobs, config: graphConfig(), meta: meta(), sleep: async () => undefined, backfill: true });
    const reason = statementSkipReason(graphConfig(), "아전 군포 PA");
    expect(reason).toBeTruthy();
    for (const r of jobs) expect(out.steps.get(r.id)?.statement).toMatchObject({ status: "skip", message: reason });
    expect(d.requests((e) => e.item === STMT_ID && e.method !== "GET")).toHaveLength(0);
    expect(d.requests((e) => e.method === "PUT")).toHaveLength(0);
  });

  it("동기화 상태 화면의 거래명세서는 한 줄이다", () => {
    const r = aj("S1");
    r.sync = "pending";
    const model = syncPanelModel([r], null, { online: true, syncing: false, mode: "graph" });
    expect(model.lines.filter((l) => l.file === "statement")).toHaveLength(1);
  });
});
