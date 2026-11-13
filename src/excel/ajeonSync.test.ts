/**
 * 아전 거래명세서(.xls)는 이번 배포의 동기화 경로에 없다(실제 OneDrive 사본을 받기 전까지).
 * 가짜 OneDrive 로만 확인한다: 아전 군포 기록은 명세서를 건너뛰고(접두어 미확정), .xls 항목은 받지도 올리지도 않는다.
 * .xls 쓰기 코드(src/excel/xls/)는 저장소에 남아 있고 그 자체 시험은 xls/ajeonStatement.test.ts 에 있다.
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import { AJEON_ID, graphConfig, meta, realDrive, rec } from "../test/realBooks";
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

describe("아전 거래명세서(.xls) — 이번 배포에서는 연결 안 함", () => {
  it("아전 군포 기록은 명세서를 건너뛰고, .xls 항목에는 요청조차 하지 않는다", async () => {
    const d = realDrive();
    const before = d.raw(AJEON_ID).slice();
    const jobs = [aj("S1"), aj("S2", { date: "2026-11-05" })];
    const out = await syncGraph({ client: client(d), jobs, all: jobs, config: graphConfig(), meta: meta(), sleep: async () => undefined, backfill: true });
    const reason = statementSkipReason(graphConfig(), "아전 군포 PA");
    expect(reason).toBeTruthy();
    for (const r of jobs) expect(out.steps.get(r.id)?.statement).toMatchObject({ status: "skip", message: reason });
    expect(out.report.ajeon).toBeUndefined();
    expect(d.requests((e) => e.item === AJEON_ID)).toHaveLength(0);
    expect(d.requests((e) => e.method === "PUT")).toHaveLength(0);
    expect(Buffer.from(d.raw(AJEON_ID)).equals(Buffer.from(before))).toBe(true);
  });

  it("동기화 상태 화면에 아전 .xls 줄이 없다(거래명세서는 한 줄)", () => {
    const r = aj("S1");
    r.sync = "pending";
    const model = syncPanelModel([r], null, { online: true, syncing: false, mode: "graph" });
    expect(model.lines.filter((l) => l.file === "statement")).toHaveLength(1);
    expect(model.lines.some((l) => l.text.includes(".xls"))).toBe(false);
  });
});
