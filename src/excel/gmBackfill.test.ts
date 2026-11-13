/**
 * 마전·아전 기록이 군포·마전 리스트에 하나도 없던 문제(2026-11-13).
 * 원인: 군포·마전 리스트가 연결되기 전에 '새 현장'으로 직접 추가한 현장('지성 마전'·'아전 군포', PA 없음)의 기록은
 * 평택 리스트로 보내졌다(그 파일에는 시트가 없어 실패 또는 리스트 단계가 끝난 것으로 남음). 군포·마전 리스트는
 * 이런 기록을 다시 확인하지 않아 한 번도 쓰이지 않았다.
 * 고친 뒤: 앱을 열면(OneDrive 로그인 상태) 그 기록을 군포·마전 PA 시트로 옮겨, Y(앱 ID)·날짜+라인+포인트로 확인하고
 * 빠진 것만 날짜 → 라인 → 포인트(S 다음 R) 순으로 마지막 행 아래에 쓴다. 실제 파일 복사본 + 가짜 Graph 로만 시험한다.
 */
import { writeFileSync } from "node:fs";
import { beforeAll, beforeEach, describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../domain/dates";
import type { InspectionRecord } from "../domain/types";
import { FAIL, PASS } from "../domain/types";
import { GraphClient, graphRetry } from "../sync/graph";
import { FakeDrive } from "../test/fakeGraph";
import { LIST_GM_ID, LIST_ID, graphConfig, meta, realDrive, realGmRecords, realListRecords, rec } from "../test/realBooks";
import { restoreStatementOwnership, syncGraph } from "./engine";
import { gmSheetOf, listBookOf, parseListWorkbook, rerouteGmRecords } from "./listBook";
import { readSheet } from "./ooxml";

beforeAll(() => {
  graphRetry.baseMs = 0;
  graphRetry.maxWaitMs = 0;
});
beforeEach(() => restoreStatementOwnership({}));

function client(drive: FakeDrive): GraphClient {
  return new GraphClient({ fetchImpl: drive.fetch, sleep: async () => undefined, getToken: async () => drive.validToken });
}

const MJ = "MJDH-D-D-P-58-005-200A";
const AJ = "GP-001-750A";
const GP = "군포산본 이원화 750A-5";
const at = 1;

/** 휴대폰에 남아 있던 기록 모양(앱 ID = UUID 형태). */
function phone(): InspectionRecord[] {
  return [
    // '새 현장'으로 추가한 '지성 마전' — 평택 리스트로 가려다 리스트 단계가 끝난 것으로 남음(동기화됨)
    rec("S01", PASS, { site: "지성 마전", sheetName: "지성 마전", line: MJ, date: "2026-11-10", sync: "synced", steps: { list: { status: "done", at }, ilji: { status: "done", at }, statement: { status: "skip", at } } }),
    // 평택 리스트에 시트가 없어 실패
    rec("R01", FAIL, { site: "지성 마전", sheetName: "지성 마전", line: MJ, date: "2026-11-10", sync: "error", steps: { list: { status: "error", message: "검사리스트(평택)에 '지성 마전' 시트가 없습니다.", at } } }),
    // 재검사: 같은 번호, 새 행
    rec("R01", PASS, { site: "지성 마전", sheetName: "지성 마전", line: MJ, date: "2026-11-11", sync: "error" }),
    rec("S02", PASS, { site: "지성 마전", sheetName: "지성 마전", line: MJ, date: "2026-11-10", sync: "synced", steps: { list: { status: "skip", at } } }),
    // '아전 군포'(PA 없음)
    rec("R17", PASS, { site: "아전 군포", sheetName: "아전 군포", line: AJ, date: "2026-11-05", sync: "synced" }),
    rec("S17", PASS, { site: "아전 군포", sheetName: "아전 군포", line: AJ, date: "2026-11-05", sync: "pending", steps: { list: { status: "skip", at } } }),
    // 이미 손으로 적힌 행(10/27 R16, Y 없음)과 같은 검사 → 다시 쓰지 않는다
    rec("R16", PASS, { site: "아전 군포", sheetName: "아전 군포", line: AJ, date: "2026-10-27", sync: "synced" }),
    // 정식 PA 현장인데 보낸 것으로만 표시되고 파일에 없음(62일 지난 것도 확인)
    rec("S07", PASS, { site: "지성 군포 PA", sheetName: "지성 군포 PA", line: GP, date: "2026-11-06", sync: "synced", listBook: "listGm", steps: { list: { status: "done", at }, ilji: { status: "done", at }, statement: { status: "done", at } } }),
  ];
}

const isJob = (r: InspectionRecord) => r.sync === "pending" || r.sync === "error";

function rowsOf(drive: FakeDrive, sheet: string) {
  return parseListWorkbook(drive.files(LIST_GM_ID), "listGm")
    .find((s) => s.sheetName === sheet)!
    .records.map((r) => ({ row: r.excelRow!, date: r.date, line: r.line, point: r.point, result: r.result, note: r.note, id: r.id }));
}

describe("군포·마전 리스트 백필(마전·아전 기록이 없던 문제)", () => {
  it("이름 맞추기: '지성 마전'·'지성마전'·'아전 군포 PA' → PA 시트, RT 시트·평택 현장은 아님", () => {
    expect(gmSheetOf("지성 마전")).toBe("지성 마전 PA");
    expect(gmSheetOf("지성마전")).toBe("지성 마전 PA");
    expect(gmSheetOf(" 아전 군포 ")).toBe("아전 군포 PA");
    expect(gmSheetOf("지성 군포 PA")).toBe("지성 군포 PA");
    expect(gmSheetOf("지성마전 RT")).toBeNull();
    expect(gmSheetOf("드림 합정주공")).toBeNull();
    expect(listBookOf({ sheetName: "지성 마전", site: "지성 마전" })).toBe("listGm");
    expect(listBookOf({ sheetName: " 드림 합정주공" })).toBe("list");
  });

  it("앱을 열면 모든 마전·아전·군포 기록을 그 PA 시트에 쓴다: 중복 없음, 날짜→라인→포인트, 기존 행 그대로, 평택 리스트는 건드리지 않음", async () => {
    const drive = realDrive();
    const before = { mj: rowsOf(drive, "지성 마전 PA"), aj: rowsOf(drive, "아전 군포 PA"), gp: rowsOf(drive, "지성 군포 PA") };
    const pyBefore = { ...drive.files(LIST_ID) };
    const recs = phone();
    expect(rerouteGmRecords(recs)).toBe(7);
    for (const r of recs) expect(listBookOf(r)).toBe("listGm");
    // 이미 끝난 작업일지·명세서 단계는 남는다(두 번 세지 않음)
    expect(recs[0].steps?.ilji?.status).toBe("done");
    expect(recs[0].steps?.list).toBeUndefined();

    const all = [...realListRecords(), ...realGmRecords(), ...recs];
    const out = await syncGraph({ client: client(drive), jobs: recs.filter(isJob), all, config: graphConfig(), meta: meta(), backfill: true, sleep: async () => undefined });
    expect(out.report.lists?.listGm).toMatchObject({ status: "ok", count: 7 });
    // 평택 리스트: Y열 확인(읽기)만, 쓰기 없음 · 내용 그대로
    expect(drive.requests((e) => e.item === LIST_ID && e.method === "PATCH")).toHaveLength(0);
    const py = drive.files(LIST_ID);
    for (const k of Object.keys(pyBefore)) expect(py[k] === pyBefore[k], k).toBe(true);

    const after = { mj: rowsOf(drive, "지성 마전 PA"), aj: rowsOf(drive, "아전 군포 PA"), gp: rowsOf(drive, "지성 군포 PA") };
    // 기존 행은 그대로
    expect(after.mj.slice(0, before.mj.length)).toEqual(before.mj);
    expect(after.aj.slice(0, before.aj.length)).toEqual(before.aj);
    expect(after.gp.slice(0, before.gp.length)).toEqual(before.gp);
    const fresh = (rows: typeof after.mj, n: number) => rows.slice(n).map((r) => `${r.row} ${r.date} ${r.point} ${r.result}`);
    expect(fresh(after.mj, before.mj.length)).toEqual(["314 2026-11-10 S01 합격", "315 2026-11-10 S02 합격", "316 2026-11-10 R01 불합격", "317 2026-11-11 R01 합격"]);
    expect(fresh(after.aj, before.aj.length)).toEqual(["40 2026-11-05 S17 합격", "41 2026-11-05 R17 합격"]);
    expect(fresh(after.gp, before.gp.length)).toEqual(["134 2026-11-06 S07 합격"]);
    // 10/27 R16 은 손으로 적힌 행(39)과 같아 새로 쓰지 않음
    expect(recs[6].excelRow).toBe(39);
    expect(recs[6].steps?.list?.status).toBe("done");
    // Y = 앱 ID, 비고(S)는 비움('앱' 없음)
    for (const r of [...after.mj, ...after.aj, ...after.gp].filter((x) => x.id.includes("-test-"))) {
      expect(r.note).toBe("");
    }
    const mjCells = readSheet(drive.files(LIST_GM_ID), "지성 마전 PA").cells;
    expect(mjCells.get("Y316")?.v).toBe(recs[1].id);
    expect(mjCells.get("E316")?.v).toBe(String(isoToExcelSerial("2026-11-10")));
    expect(mjCells.get("L316")?.v).toBe("200A");

    if (process.env.GM_BACKFILL_OUT) writeFileSync(process.env.GM_BACKFILL_OUT, drive.bytes(LIST_GM_ID));

    // 다시 열어도(다시 시도) 행이 늘지 않는다
    for (const r of recs) r.sync = "synced";
    const again = await syncGraph({ client: client(drive), jobs: [], all: [...realListRecords(), ...realGmRecords(), ...recs], config: graphConfig(), meta: meta(), backfill: true, scanList: true, sleep: async () => undefined });
    expect(again.report.lists?.listGm?.status ?? "ok").toBe("ok");
    expect(rowsOf(drive, "지성 마전 PA")).toHaveLength(after.mj.length);
    expect(rowsOf(drive, "아전 군포 PA")).toHaveLength(after.aj.length);
    expect(rowsOf(drive, "지성 군포 PA")).toHaveLength(after.gp.length);
  });
});
