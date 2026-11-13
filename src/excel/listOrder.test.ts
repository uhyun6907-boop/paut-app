import { describe, expect, it } from "vitest";
import { orderForAppend } from "./listBook";

const r = (point: string, ts: number, date = "2026-11-12", line = "L-1") => ({ id: `${point}-${ts}`, point, ts, date, line });

describe("orderForAppend", () => {
  it("숫자 순(S1·S2·S10), S 전부 다음 R, 입력 배열은 그대로", () => {
    const input = [r("S10", 1), r("S2", 2), r("R2", 3), r("S1", 4), r("R10", 5)];
    expect(orderForAppend(input).map((x) => x.point)).toEqual(["S1", "S2", "S10", "R2", "R10"]);
    expect(input.map((x) => x.point)).toEqual(["S10", "S2", "R2", "S1", "R10"]);
  });

  it("두 자리 표기와 하위 번호: S09 < S10 < S10-1 < S11", () => {
    expect(orderForAppend([r("S11", 1), r("S10-1", 2), r("S10", 3), r("S09", 4)]).map((x) => x.point)).toEqual(["S09", "S10", "S10-1", "S11"]);
  });

  it("날짜 → 라인(숫자 인식) → 포인트, 재검사(같은 번호)는 먼저 입력한 원본 바로 뒤", () => {
    const out = orderForAppend([
      r("S01", 1, "2026-11-13", "L-2"),
      r("S03", 9, "2026-11-12", "L-10"),
      r("S03", 3, "2026-11-12", "L-2"),
      r("S03", 2, "2026-11-12", "L-2"),
      r("R01", 4, "2026-11-12", "L-2"),
    ]);
    expect(out.map((x) => `${x.date} ${x.line} ${x.point}#${x.ts}`)).toEqual([
      "2026-11-12 L-2 S03#2",
      "2026-11-12 L-2 S03#3",
      "2026-11-12 L-2 R01#4",
      "2026-11-12 L-10 S03#9",
      "2026-11-13 L-2 S01#1",
    ]);
  });
});
