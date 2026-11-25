/**
 * 아전 군포 거래명세서. 실제 OneDrive 확인(2026-11-25): '01_PAUT 거래명세서 (아전, 승재).xls' — .xls(옛 형식)라 Graph 로 칸을 쓸 수 없고,
 * 아전 군포 시트 양식도 02 와 다르다(일자·품명·관경·수량(개소)·단가(개소당 183,600)·금액). 그래서 01 전용 입력을 만들기 전까지
 * 01 파일에는 아무것도 쓰지 않고 '대기'로 둔다. 다른 현장은 그대로 02 파일에 쓴다. 가짜 OneDrive 로만 확인한다.
 */
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import type { InspectionRecord } from "../domain/types";
import { PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { syncPanelModel } from "../store/syncStatus";
import { FakeDrive } from "../test/fakeGraph";
import { STMT_ID, graphConfig, meta, realDrive, rec } from "../test/realBooks";
import { AJEON_WAIT_TEXT, restoreStatementOwnership, statementBookOf, syncGraph } from "./engine";
import { readPautRows } from "./statementBook";
import { workbookSheets } from "./ooxml";

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
  it("아전 군포는 01 파일 몫, 다른 현장은 02 파일(예전 설정에 statementBook 이 없어도)", () => {
    const old = graphConfig();
    old.mappings = old.mappings.map((m) => ({ ...m, statementBook: undefined }));
    expect(statementBookOf(old, "아전 군포 PA")).toBe("statementAj");
    expect(statementBookOf(old, "아전 군포")).toBe("statementAj");
    expect(statementBookOf(old, "지성 군포 PA")).toBe("statement");
  });

  it("아전 군포 기록은 어느 명세서 파일에도 쓰지 않고 '대기'(실패 아님). 같은 동기화의 합정주공은 02 파일에 쓴다", async () => {
    const d = realDrive();
    const hj = rec("S28", PASS);
    const jobs = [aj("S1"), aj("S2", { date: "2026-11-05" }), hj];
    const snapshot = () =>
      Object.fromEntries(workbookSheets(d.files(STMT_ID)).map(({ name }) => [name, JSON.stringify(readPautRows(d.files(STMT_ID), name))]));
    const before = snapshot();
    const out = await syncGraph({ client: client(d), jobs, all: jobs, config: graphConfig(), meta: meta(), sleep: async () => undefined });
    for (const r of jobs.slice(0, 2)) expect(out.steps.get(r.id)?.statement).toEqual(expect.objectContaining({ status: "wait", message: AJEON_WAIT_TEXT }));
    expect(out.steps.get(hj.id)?.statement?.status).toBe("done");
    expect(readPautRows(d.files(STMT_ID), "드림합정26.11").map((r) => `${r.date}|${r.diameter}|${r.count}`)).toContain("2026-11-12|350A|1");
    // 02 파일은 합정주공 11월 시트만 바뀌었고(아전 줄 없음), 01 파일은 찾지도 열지도 않는다.
    const after = snapshot();
    expect(Object.keys(after).filter((name) => after[name] !== before[name])).toEqual(["드림합정26.11"]);
    expect(readPautRows(d.files(STMT_ID), "드림합정26.11").some((r) => r.date === "2026-11-03" || r.date === "2026-11-05")).toBe(false);
    expect(d.requests((e) => e.path.includes("/search(") || e.path.includes("root:/"))).toHaveLength(0);
    expect(out.report.files.statement.status).not.toBe("error");
  });

  it("동기화 창: 거래명세서는 한 줄, 빨간 실패가 아니라 주황 대기로 이유를 보인다", () => {
    const r = aj("S1", { sync: "pending", steps: { list: { status: "done" }, ilji: { status: "done" }, statement: { status: "wait", message: AJEON_WAIT_TEXT } } });
    const model = syncPanelModel([r], null, { online: true, syncing: false, mode: "graph" });
    const lines = model.lines.filter((l) => l.file === "statement");
    expect(lines).toHaveLength(1);
    expect(lines[0].state).toBe("waiting");
    expect(lines[0].text).toContain(AJEON_WAIT_TEXT);
    expect(model.tone).not.toBe("bad");
  });
});

describe("예전 판이 남긴 아전 군포 명세서 표시", () => {
  it("앱을 켜면 '건너뜀'·'실패(01 파일 못 찾음)'를 지우고, 명세서 실패만 있던 기록은 실패에서 대기로", async () => {
    const { reopenAjeonStatement } = await import("../store/session");
    const skipped = aj("S1", { sync: "synced", steps: { list: { status: "done" }, ilji: { status: "done" }, statement: { status: "skip", message: "접두어" } } });
    const failed = aj("S2", {
      sync: "error",
      syncError: "거래명세서: OneDrive 에서 「01_PAUT 거래명세서 (아전, 승재).xlsx」 파일을 찾지 못했어요.",
      steps: { list: { status: "done" }, ilji: { status: "done" }, statement: { status: "error", message: "못 찾음" } },
    });
    const other = rec("S3", PASS, { steps: { statement: { status: "skip" } } });
    expect(reopenAjeonStatement([skipped, failed, other])).toBe(true);
    expect(skipped.steps?.statement).toBeUndefined();
    expect(failed).toMatchObject({ sync: "pending", syncError: undefined });
    expect(failed.steps?.statement).toBeUndefined();
    expect(failed.steps?.list?.status).toBe("done");
    expect(other.steps?.statement?.status).toBe("skip");
    expect(reopenAjeonStatement([skipped, failed, other])).toBe(false);
  });
});
