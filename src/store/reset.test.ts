/** 초기화 버튼: 저장 안 한 카드 표시(합격/불합격·새 카드·회색 카드)를 마지막 저장 모습으로 되돌린다. 저장·엑셀 없음. */
import { beforeEach, describe, expect, it, vi } from "vitest";

const mem = new Map<string, unknown>();
const bag = new Map<string, string>();

vi.mock("./db", () => ({
  kvGet: async (key: string) => (mem.has(key) ? structuredClone(mem.get(key)) : undefined),
  kvSet: async (key: string, value: unknown) => {
    mem.set(key, structuredClone(value));
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

import { FAIL, PASS } from "../domain/types";
import { commitDay, getState, resetUnsaved, resettableCount, savePoint, setResult, settled, updateMeta } from "./session";
import { applyPaint } from "../ui/EntryView";
import { RESET_CONFIRM } from "../ui/EntryView";

const DATE = "2026-11-12";
const SITE = "초기화현장";
let LINE = "RS-1";
let n = 0;

const day = () => getState().records.filter((r) => r.site === SITE && r.date === DATE && !r.deleted);
const byPoint = (p: string) => day().filter((r) => r.point === p);
const snapshot = () =>
  day()
    .map((r) => `${r.point}:${r.result}:${r.sync === "local" ? "local" : "saved"}`)
    .sort();
const count = () => resettableCount(getState().records, getState().meta);

describe("초기화(저장 안 한 카드 표시 되돌리기)", () => {
  beforeEach(async () => {
    await settled();
    mem.clear();
    bag.clear();
    bag.set("paut-sr-mirror", JSON.stringify({ enabled: true, independent: [] }));
    getState().records.splice(0, getState().records.length);
    n += 1;
    LINE = `RS-${n}`;
    await updateMeta({ lastSite: SITE, activeLine: LINE, workDate: DATE, prefix: "S", insp1: "조우현", insp2: "심명보", draftCards: {} });
  });

  it("확인 문구", () => {
    expect(RESET_CONFIRM).toBe("저장 안 한 표시를 되돌릴까요?");
  });

  it("저장 뒤 잘못 누른 합격↔불합격·새 카드를 저장 때 모습으로 되돌리고(S→R 포함), 엑셀 대기는 늘지 않는다", async () => {
    await savePoint(PASS, "S1");
    await savePoint(PASS, "S2");
    await commitDay();
    await settled();
    const saved = snapshot();
    const savedDrafts = JSON.stringify(getState().meta.draftCards ?? {});
    expect(saved.every((x) => x.endsWith(":saved"))).toBe(true);
    expect(count()).toBe(0);
    const syncOf = new Map(day().map((r) => [r.id, r.sync]));

    // 잘못 누름: S1 불합격(R1 따라감·재검사 회색 카드), 새 S3 합격
    await setResult(byPoint("S01")[0].id, FAIL);
    await savePoint(PASS, "S3");
    await settled();
    expect(byPoint("S01")[0].result).toBe(FAIL);
    expect(byPoint("S03")).toHaveLength(1);
    expect(byPoint("R01")[0].result).toBe(FAIL); // S→R 따라감
    expect(JSON.stringify(getState().meta.draftCards ?? {})).not.toBe(savedDrafts); // 재검사 회색 카드
    expect(count()).toBeGreaterThan(0);

    await resetUnsaved();
    await settled();
    expect(snapshot()).toEqual(saved);
    expect(JSON.stringify(getState().meta.draftCards ?? {})).toBe(savedDrafts);
    expect(day().some((r) => r.sync === "local")).toBe(false);
    // 저장 때 상태 그대로(새로 엑셀에 보낼 것이 생기지 않음)
    for (const r of day()) expect(r.sync).toBe(syncOf.get(r.id));
    expect(count()).toBe(0);
    // 이 기기 저장소에도 되돌린 모습
    const stored = (mem.get("records") as { site: string; date: string; point: string; result: string }[]).filter((r) => r.site === SITE && r.date === DATE);
    expect(stored.find((r) => r.point === "S01")?.result).toBe(PASS);
    expect(stored.some((r) => r.point === "S03")).toBe(false);
  });

  it("한 번도 저장 안 한 카드는 모두 지운다. 바꾼 게 없으면 버튼이 꺼져 있다", async () => {
    expect(count()).toBe(0);
    await savePoint(PASS, "S5");
    await savePoint(FAIL, "S6");
    await settled();
    expect(count()).toBeGreaterThan(0);
    await resetUnsaved();
    await settled();
    expect(day()).toHaveLength(0);
    expect(getState().meta.draftCards?.[`${SITE}||${LINE}||R`] ?? []).toEqual([]);
    expect(count()).toBe(0);
  });

  it("저장한 S30·R30 에서 R30 을 직접 잘못 눌렀다가 초기화하면 R30 은 되돌아가고 다시 S30 을 따라간다", async () => {
    await savePoint(PASS, "S30");
    await commitDay();
    await settled();
    for (const r of day()) r.sync = "synced";
    await updateMeta({ activeLine: LINE }); // 저장된 모습 기억
    await setResult(byPoint("R30")[0].id, FAIL); // R 을 직접 누름 → R30 은 S 를 안 따라감
    await setResult(byPoint("S30")[0].id, FAIL);
    expect(byPoint("R30")[0].result).toBe(FAIL);
    await resetUnsaved();
    await settled();
    expect(snapshot()).toEqual([`R30:${PASS}:saved`, `S30:${PASS}:saved`]);
    // 초기화로 '직접 고친 R' 표시도 되돌렸다 → 칠하기/탭이 다시 R 을 움직인다.
    await setResult(byPoint("S30")[0].id, FAIL);
    expect(byPoint("R30")[0].result).toBe(FAIL);
    await setResult(byPoint("S30")[0].id, PASS);
    expect(byPoint("R30")[0].result).toBe(PASS);
    await applyPaint(`rec:${byPoint("S30")[0].id}`, false);
    expect(byPoint("R30")).toHaveLength(1);
  });
});
