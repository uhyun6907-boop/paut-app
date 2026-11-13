import { beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();
/** 저장(IndexedDB) 지연. 실제 기기처럼 저장이 끝나기 전에 다음 탭이 들어오게 한다. */
let storageDelay = 0;

vi.mock("./db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => {
    const copy = structuredClone(value);
    if (storageDelay) await new Promise((resolve) => setTimeout(resolve, storageDelay));
    mem.set(key, copy);
  },
  kvDel: async (key: string) => {
    mem.delete(key);
  },
}));

vi.stubGlobal("localStorage", {
  getItem: (key: string) => bag.get(key) ?? null,
  setItem: (key: string, value: string) => {
    bag.set(key, String(value));
  },
  removeItem: (key: string) => {
    bag.delete(key);
  },
});

import { FAIL, PASS, type InspectionRecord } from "../domain/types";
import { kvGet } from "./db";
import {
  getState,
  mirrorGreyDraft,
  noteDirectR,
  removeRecord,
  savePoint,
  setLineDrafts,
  setResult,
  setSrMirrorEnabled,
  settled,
  updateMeta,
} from "./session";
import { applyPaint } from "../ui/EntryView";
import { PaintGesture } from "../ui/paint";

const DATE = "2026-11-12";
const SITE = "테스트현장";
const SR_KEY = "paut-sr-mirror";
let LINE = "HJ-M-1";
let lineNo = 0;

function clearFollowFlags(): void {
  bag.set(SR_KEY, JSON.stringify({ enabled: true, independent: [] }));
}

function norm(point: string): string {
  return point.replace(/\s+/g, "").toUpperCase();
}

function today(point: string, line = LINE) {
  return getState()
    .records.filter((rec) => !rec.deleted && rec.date === DATE && rec.line.trim().toUpperCase() === line.toUpperCase() && norm(rec.point) === point.toUpperCase())
    .sort((a, b) => a.ts - b.ts);
}

function independentFlags(): string[] {
  return (JSON.parse(bag.get(SR_KEY) ?? "{}") as { independent?: string[] }).independent ?? [];
}

function rDrafts(line = LINE) {
  return getState().meta.draftCards?.[`${SITE}||${line}||R`] ?? [];
}

function pause(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

describe("S→R 자동", () => {
  beforeEach(async () => {
    await settled();
    storageDelay = 0;
    mem.clear();
    bag.clear();
    clearFollowFlags();
    getState().records.splice(0, getState().records.length);
    // 같은 번호·결과를 700ms 안에 두 번 저장하면 중복 탭으로 무시되므로 테스트마다 라인을 바꾼다.
    lineNo += 1;
    LINE = `HJ-M-${lineNo}`;
    await updateMeta({
      lastSite: SITE,
      activeLine: LINE,
      workDate: DATE,
      prefix: "S",
      insp1: "조우현",
      insp2: "심명보",
      draftCards: {},
    });
  });

  it("S23부터 S28을 섞어서 빨리 눌러도 같은 번호의 R이 따라가고, 저장을 다시 읽어도 남는다", async () => {
    await savePoint(FAIL, "R24");
    await savePoint(PASS, "R26");
    await savePoint(FAIL, "R27");
    const savedR27 = today("R27")[0];
    savedR27.point = "r27";
    savedR27.line = `${LINE} `;
    savedR27.sync = "synced";
    today("R24")[0].sync = "synced";
    today("R24")[0].result = FAIL;
    clearFollowFlags();

    const order = ["S26", "S23", "S28", "s24", "S27", "S25"];
    await Promise.all(order.map((point) => savePoint(PASS, point)));
    await savePoint(PASS, "S23-1");

    for (const point of ["R23", "R24", "R25", "R26", "R27", "R28"]) {
      const rows = today(point);
      expect(rows.map((rec) => rec.result)).toEqual([PASS]);
    }
    expect(today("R23-1").map((rec) => rec.result)).toEqual([PASS]);
    expect(today("R27")).toHaveLength(1);

    const r25 = today("R25")[0];
    await setResult(r25.id, FAIL);
    const s25 = today("S25")[0];
    await setResult(s25.id, PASS);
    expect(today("R25")[0].result).toBe(FAIL);
    expect(today("S25")[0].result).toBe(PASS);

    noteDirectR("R23");
    await setResult(today("S23")[0].id, FAIL);
    expect(today("R23")[0].result).toBe(PASS);

    const stored = (await kvGet<InspectionRecord[]>("records")) ?? [];
    const unsaved = JSON.parse(bag.get("paut-unsaved-v1") ?? "{}") as { records?: { point: string; result: string; date: string }[] };
    const merged = new Map<string, string>();
    for (const rec of stored) {
      if (rec.date === DATE) merged.set(norm(rec.point), rec.result);
    }
    for (const rec of unsaved.records ?? []) {
      if (rec.date === DATE) merged.set(norm(rec.point), rec.result);
    }
    for (const point of ["R23", "R24", "R26", "R27", "R28", "R23-1"]) expect(merged.get(point)).toBe(PASS);
    expect(merged.get("R25")).toBe(FAIL);
  });

  it("실제 손가락처럼 간격을 두고 S23~S28을 누르고 저장이 느려도 R23~R28이 모두 따라간다", async () => {
    // 예전 큐는 '실행 중이면 바로 실행'이라 저장 중에 들어온 탭이 미러를 건너뛰어 R 일부가 빠졌다.
    storageDelay = 3;
    const calls: Promise<unknown>[] = [];
    for (const point of ["S23", "S24", "S25", "S26", "S27", "S28"]) {
      calls.push(savePoint(PASS, point));
      await pause(2);
    }
    await Promise.all(calls);
    for (const point of ["R23", "R24", "R25", "R26", "R27", "R28"]) {
      expect(today(point).map((rec) => rec.result), point).toEqual([PASS]);
    }
    const stored = (await kvGet<InspectionRecord[]>("records")) ?? [];
    expect(stored.filter((rec) => rec.line === LINE && rec.point.startsWith("R")).map((rec) => rec.point).sort()).toEqual(["R23", "R24", "R25", "R26", "R27", "R28"]);
    expect(independentFlags()).toEqual([]);
  });

  it("이미 저장(동기화)된 R은 사용자 편집으로 보지 않고 따라간다", async () => {
    await savePoint(FAIL, "R30");
    const r30 = today("R30")[0];
    r30.sync = "synced";
    r30.excelRow = 120;
    clearFollowFlags(); // 지난번 저장으로 끝난 상태: 직접 편집 표시는 없다.

    await savePoint(PASS, "S30");
    expect(today("R30")).toHaveLength(1);
    expect(today("R30")[0].id).toBe(r30.id);
    expect(today("R30")[0].result).toBe(PASS);
    expect(today("R30")[0].sync).toBe("local"); // 다음 저장 때 엑셀 행을 고친다
    expect(independentFlags()).toEqual([]);
  });

  it("결과가 같으면 저장된 R을 미저장으로 되돌리지 않는다", async () => {
    await savePoint(PASS, "R31");
    today("R31")[0].sync = "synced";
    clearFollowFlags();
    await savePoint(PASS, "S31");
    expect(today("R31")[0].sync).toBe("synced");
  });

  it("사용자가 직접 고친 R은 덮어쓰지 않고, 옆 번호는 계속 따라간다", async () => {
    await savePoint(PASS, "S40");
    await savePoint(PASS, "S41");
    await setResult(today("R40")[0].id, FAIL); // R 카드를 직접 불합격으로

    await setResult(today("S40")[0].id, FAIL);
    await setResult(today("S40")[0].id, PASS);
    await setResult(today("S41")[0].id, FAIL);

    expect(today("R40").map((rec) => rec.result)).toEqual([FAIL]);
    expect(today("R41").map((rec) => rec.result)).toEqual([FAIL]);

    // R 을 직접 새로 입력한 번호도 S 를 따라가지 않는다.
    await savePoint(FAIL, "R42");
    await savePoint(PASS, "S42");
    expect(today("R42").map((rec) => rec.result)).toEqual([FAIL]);
  });

  it("짝 R 카드가 없으면 같은 라인·날짜에 만들고, 회색 R 카드는 기록으로 바뀐다", async () => {
    await mirrorGreyDraft("S50");
    expect(rDrafts().map((card) => card.point)).toEqual(["R50"]);

    await savePoint(PASS, "S50");
    await savePoint(FAIL, "S51");
    const r50 = today("R50");
    expect(r50).toHaveLength(1);
    expect(r50[0]).toMatchObject({ site: SITE, line: LINE, date: DATE, result: PASS, sync: "local" });
    expect(today("R51").map((rec) => rec.result)).toEqual([FAIL]);
    // 회색 R50 은 사라지고, R51 불합격에는 회색 재검사 카드가 붙는다.
    expect(rDrafts().map((card) => [card.point, !!card.reinspectOf])).toEqual([["R51", true]]);
  });

  it("재검사: S 불합격 뒤 재검사 합격이면 R도 불합격 + 재검사 합격 두 줄이 된다", async () => {
    await savePoint(FAIL, "S60");
    expect(today("R60").map((rec) => rec.result)).toEqual([FAIL]);
    const rFail = today("R60")[0];
    expect(rDrafts().some((card) => card.reinspectOf === rFail.id)).toBe(true);

    await pause(1); // ts 순서 보장
    await savePoint(PASS, "S60"); // 재검사 카드를 합격으로
    expect(today("S60").map((rec) => rec.result)).toEqual([FAIL, PASS]);
    expect(today("R60").map((rec) => rec.result)).toEqual([FAIL, PASS]);
    expect(rDrafts().some((card) => card.reinspectOf === rFail.id)).toBe(false);

    // 재검사 S 를 지우면 짝인 재검사 R 만 지워지고 원래 불합격 R 은 남는다.
    await removeRecord(today("S60")[1].id, true);
    expect(today("R60").map((rec) => rec.result)).toEqual([FAIL]);
  });

  it("다른 라인의 S 기록을 고치면 그 라인의 R 이 따라가고, 지금 라인은 건드리지 않는다", async () => {
    const other = `${LINE}-B`;
    await savePoint(PASS, "S70");
    await updateMeta({ activeLine: other });
    await savePoint(PASS, "S70");
    await updateMeta({ activeLine: LINE });

    await setResult(today("S70", other)[0].id, FAIL);
    expect(today("R70", other).map((rec) => rec.result)).toEqual([FAIL]);
    expect(today("R70", LINE).map((rec) => rec.result)).toEqual([PASS]);
  });

  it("S→R 자동을 끄면 R 을 만들지도 고치지도 않는다", async () => {
    setSrMirrorEnabled(false);
    await savePoint(PASS, "S80");
    expect(today("R80")).toHaveLength(0);
  });

  it("S29·S30·S31: 회색 S 를 탭(S29)·칠하기(S30~S31)하면 R29~R31 이 합격으로 켜지고 회색 R 은 사라진다", async () => {
    const key = `${SITE}||${LINE}||S`;
    await setLineDrafts(key, ["S29", "S30", "S31"].map((point) => ({ id: `d-${point}`, point })));
    for (const point of ["S29", "S30", "S31"]) await mirrorGreyDraft(point);
    expect(rDrafts().map((c) => c.point)).toEqual(["R29", "R30", "R31"]);
    await savePoint(PASS, "S29"); // 탭(applyTap → savePoint)
    await applyPaint("draft:d-S30", true); // 칠하기
    await applyPaint("draft:d-S31", false);
    await settled();
    for (const point of ["R29", "R30", "R31"]) expect(today(point).map((r) => r.result), point).toEqual([PASS]);
    expect(rDrafts()).toEqual([]);
  });

  it("R 탭 + 로 회색 R29~R31 을 먼저 만들어 둔 줄(예전 판이 '직접 고친 R' 표시를 남김)도 S29~S31 을 누르면 R 이 따라간다", async () => {
    // 예전 판: R 탭 + 가 noteDirectR 을 불러 회색 카드만 있어도 R29~R31 이 S 를 영영 안 따라갔다("s29 3031 눌러도 r 안 눌러짐").
    const rKey = `${SITE}||${LINE}||R`;
    await setLineDrafts(rKey, ["R29", "R30", "R31"].map((point) => ({ id: `r-${point}`, point })));
    for (const point of ["R29", "R30", "R31"]) noteDirectR(point);
    const key = `${SITE}||${LINE}||S`;
    await setLineDrafts(key, ["S29", "S30", "S31"].map((point) => ({ id: `d-${point}`, point })));
    await savePoint(PASS, "S29");
    await applyPaint("draft:d-S30", true);
    await applyPaint("draft:d-S31", false);
    await settled();
    for (const point of ["R29", "R30", "R31"]) expect(today(point).map((r) => r.result), point).toEqual([PASS]);
    expect(rDrafts()).toEqual([]); // 회색 R29~R31 은 기록으로 바뀜(같은 번호 카드 하나)
    // 표시는 지워졌으므로 S 를 바꾸면 계속 따라간다.
    await setResult(today("S30")[0].id, FAIL);
    expect(today("R30")[0].result).toBe(FAIL);
  });

  it("이미 합격인 S29~S31 인데 짝 R 이 없으면, 그 위를 칠하면 R29~R31 이 합격으로 생긴다(S 는 그대로, 다시 칠해도 중복 없음)", async () => {
    setSrMirrorEnabled(false);
    for (const point of ["S29", "S30", "S31"]) await savePoint(PASS, point);
    setSrMirrorEnabled(true);
    for (const point of ["S29", "S30", "S31"]) today(point)[0].sync = "synced";
    expect(today("R29")).toHaveLength(0);
    const keys = ["S29", "S30", "S31"].map((p) => `rec:${today(p)[0].id}`);
    for (let pass = 0; pass < 2; pass++) {
      await drag(keys[0], keys.map((k, i) => [cx(i) + 14, k] as [number, string]));
    }
    for (const point of ["S29", "S30", "S31"]) expect(today(point).map((r) => `${r.result}:${r.sync}`), point).toEqual([`${PASS}:synced`]);
    for (const point of ["R29", "R30", "R31"]) expect(today(point).map((r) => r.result), point).toEqual([PASS]);
    // 직접 고친 R 은 여전히 S 를 따라가지 않는다.
    await setResult(today("R30")[0].id, FAIL);
    await drag(keys[0], keys.map((k, i) => [cx(i) + 14, k] as [number, string]));
    expect(today("R30").map((r) => r.result)).toEqual([FAIL]);
  });

  /** 칸 i 의 가운데 x. 카드 폭 44px. 손가락은 x=0(첫 카드)에서 시작한다. */
  const cx = (i: number) => i * 44;

  /**
   * 화면 연결과 같은 흐름: 누름 → 이동(미리보기만) → 뗌 → 남은 길만 차례대로 저장.
   * moves 는 [x, 손가락 아래 카드 key] 목록.
   */
  async function drag(startKey: string | null, moves: [number, string | null][], beforeRelease?: () => void) {
    const gesture = new PaintGesture();
    gesture.down(0, 0, startKey);
    for (const [x, key] of moves) {
      gesture.move(x, 1, key);
      await pause(1);
    }
    beforeRelease?.();
    const release = gesture.up();
    let chain = Promise.resolve();
    for (const hit of release.commit) chain = chain.then(() => applyPaint(hit.key, hit.fromStart));
    await chain;
    await settled();
    return release;
  }

  /** S23~S28: S25 만 불합격 기록, 나머지는 회색. R 은 S 를 따라 회색/불합격. */
  async function setupRow() {
    const key = `${SITE}||${LINE}||S`;
    const grey = ["S23", "S24", "S26", "S27", "S28"];
    await setLineDrafts(key, grey.map((point) => ({ id: `d-${point}`, point })));
    for (const point of grey) await mirrorGreyDraft(point);
    await savePoint(FAIL, "S25");
    const s25 = today("S25")[0];
    const keys = ["draft:d-S23", "draft:d-S24", `rec:${s25.id}`, "draft:d-S26", "draft:d-S27", "draft:d-S28"];
    return { key, s25, keys };
  }

  function excelBound(point: string): InspectionRecord[] {
    return getState().records.filter((rec) => !rec.deleted && rec.line === LINE && norm(rec.point) === point && rec.result !== "없음");
  }

  it("칠하기: S23~S28 미입력 위를 가로로 지나가면 모두 합격, 지나가기만 한 불합격은 그대로, R 도 따라간다", async () => {
    storageDelay = 2;
    const { key, s25, keys } = await setupRow();
    expect(today("R25")[0].result).toBe(FAIL);

    await drag(keys[0], keys.map((k, i) => [cx(i) + 14, k] as [number, string]));

    for (const point of ["S23", "S24", "S26", "S27", "S28", "R23", "R24", "R26", "R27", "R28"]) {
      expect(today(point).map((rec) => rec.result), point).toEqual([PASS]);
    }
    expect(today("S25").map((rec) => rec.result)).toEqual([FAIL]);
    expect(today("R25").map((rec) => rec.result)).toEqual([FAIL]);
    expect(getState().meta.draftCards?.[key]?.filter((card) => !card.reinspectOf) ?? []).toEqual([]);
    expect(independentFlags()).toEqual([]);

    // 불합격 카드에서 시작하면 그 카드는 합격으로 바뀌고 R 도 따라간다.
    const release = await drag(`rec:${s25.id}`, [[20, `rec:${s25.id}`]]);
    expect(release.commit).toEqual([{ key: `rec:${s25.id}`, fromStart: true }]);
    expect(today("S25").map((rec) => rec.result)).toEqual([PASS]);
    expect(today("R25").map((rec) => rec.result)).toEqual([PASS]);
  });

  it("되돌리기: S23→S28 칠한 뒤 S25 로 돌아오면 S26~S28 과 R26~R28 은 드래그 전 그대로, 저장·엑셀 대기도 없다", async () => {
    storageDelay = 2;
    const { key, keys } = await setupRow();
    // R28 은 지난번에 저장된(동기화된) 불합격 R. 사용자 편집 표시는 없다.
    await savePoint(FAIL, "R28");
    const r28 = today("R28")[0];
    r28.sync = "synced";
    r28.excelRow = 88;
    clearFollowFlags();
    const before = getState().records.length;
    const greyR = rDrafts().map((card) => card.point).sort();

    const forward = keys.map((k, i) => [cx(i) + 14, k] as [number, string]);
    const back: [number, string][] = [[cx(4) + 14, keys[4]], [cx(3) + 14, keys[3]], [cx(2) + 14, keys[2]]];
    const release = await drag(keys[0], [...forward, ...back], () => {
      // 손을 떼기 전에는 미리보기일 뿐, 아무것도 저장되지 않는다.
      expect(getState().records.length).toBe(before);
    });
    expect(release.commit.map((hit) => hit.key)).toEqual(keys.slice(0, 3));

    for (const point of ["S23", "S24", "R23", "R24"]) expect(today(point).map((rec) => rec.result), point).toEqual([PASS]);
    // 지나가기만 한 S25 불합격은 그대로.
    expect(today("S25").map((rec) => rec.result)).toEqual([FAIL]);
    expect(today("R25").map((rec) => rec.result)).toEqual([FAIL]);
    // 되돌린 S26~S28: 기록 없음, 회색 카드 그대로.
    for (const point of ["S26", "S27", "S28", "R26", "R27"]) expect(excelBound(point), point).toEqual([]);
    expect((getState().meta.draftCards?.[key] ?? []).map((card) => card.point).sort()).toEqual(["S26", "S27", "S28"]);
    expect(rDrafts().map((card) => card.point).sort()).toEqual(greyR.filter((p) => !["R23", "R24"].includes(p)));
    // R28 은 드래그 전 상태(저장된 불합격) 그대로. 다시 보낼 대상도 아니다.
    expect(today("R28").map((rec) => [rec.result, rec.sync])).toEqual([[FAIL, "synced"]]);
    // R 을 되돌려도 '직접 고침' 표시가 생기지 않는다 → 이후에도 S 를 따라간다.
    expect(independentFlags()).toEqual([]);
    await savePoint(PASS, "S28");
    expect(today("R28").map((rec) => rec.result)).toEqual([PASS]);

    // 저장소(IndexedDB)와 미저장 백업에도 되돌린 카드가 없다.
    const stored = (await kvGet<InspectionRecord[]>("records")) ?? [];
    const unsaved = JSON.parse(bag.get("paut-unsaved-v1") ?? "{}") as { records?: InspectionRecord[] };
    for (const rows of [stored, unsaved.records ?? []]) {
      const points = rows.filter((rec) => rec.line === LINE).map((rec) => norm(rec.point));
      for (const point of ["S26", "S27", "R26", "R27"]) expect(points, point).not.toContain(point);
    }
  });

  it("되돌리기: 불합격 카드에서 시작했어도 처음 자리로 돌아와 떼면 불합격 그대로", async () => {
    const { s25, keys } = await setupRow();
    const start = `rec:${s25.id}`;
    const release = await drag(start, [
      [20, start],
      [cx(1) + 14, keys[3]],
      [cx(1) - 10, start],
      [4, start], // 처음 누른 자리(12px 안)로 돌아옴 → 시작 카드도 빠짐
    ]);
    expect(release.commit).toEqual([]);
    expect(release.swallowClick).toBe(true); // 드래그였으니 탭으로 치지 않는다
    expect(today("S25").map((rec) => rec.result)).toEqual([FAIL]);
    expect(today("R25").map((rec) => rec.result)).toEqual([FAIL]);
    expect(excelBound("S26")).toEqual([]);
    expect(independentFlags()).toEqual([]);
  });

  it("되돌리기: 불합격에서 시작해 옆으로 갔다가 시작 카드까지만 돌아오면 시작 카드는 합격", async () => {
    const { s25, keys } = await setupRow();
    const start = `rec:${s25.id}`;
    const release = await drag(start, [[20, start], [cx(1) + 14, keys[3]], [cx(1) - 18, start]]);
    expect(release.commit).toEqual([{ key: start, fromStart: true }]);
    expect(today("S25").map((rec) => rec.result)).toEqual([PASS]);
    expect(today("R25").map((rec) => rec.result)).toEqual([PASS]);
    expect(excelBound("S26")).toEqual([]);
  });
});
