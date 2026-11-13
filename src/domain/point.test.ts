import { describe, expect, it } from "vitest";
import { advanceAfterSave, canonicalPoint, cycleCard, directionFromTrail, expandRange, fmtFullPoint, fmtPoint, nextDraftPoint, nextNumber, nextPointLabel, orderPointCards, mirrorPoint, parsePoint, planPrefixCopy, pointCells, reinspectLabel, stepNumber, suggestNumber } from "./point";

describe("포인트 번호", () => {
  it("두 자리로 맞춘다", () => {
    expect(fmtPoint("s", 1)).toBe("S01");
    expect(fmtPoint("R", 12)).toBe("R12");
    expect(fmtPoint("p", 100)).toBe("P100");
  });

  it("R05-1 의 기준 번호는 5, R27-4-1 은 27", () => {
    expect(parsePoint("r05-1")).toMatchObject({ prefix: "R", num: 5, suffix: "1", raw: "R05-1" });
    expect(parsePoint("R27-4-1")).toMatchObject({ prefix: "R", num: 27, suffix: "4-1" });
    expect(parsePoint("S01")).toMatchObject({ prefix: "S", num: 1, suffix: null });
    expect(parsePoint("")).toBeNull();
    expect(parsePoint("S")).toBeNull();
  });

  it("같은 접두어의 최대 기준 번호 + 1", () => {
    const pts = ["S01", "S05-1", "S03", "R21", "R05-1", "P02"];
    expect(nextNumber(pts, "S")).toBe(6);
    expect(nextPointLabel(pts, "S")).toBe("S06");
    expect(nextNumber(pts, "R")).toBe(22);
    expect(nextPointLabel(pts, "R")).toBe("R22");
    expect(nextNumber(pts, "P")).toBe(3);
    expect(nextNumber([], "S")).toBe(1);
  });

  it("저장 후에는 같은 접두어의 다음 번호. R05-1 다음은 R06", () => {
    const saved = parsePoint("R05-1")!;
    expect(advanceAfterSave("R", saved)).toEqual({ prefix: "R", num: 6 });
    expect(fmtPoint("R", 6)).toBe("R06");
  });

  it("하위 번호는 본번호 다음 제안에 넣지 않는다", () => {
    expect(fmtFullPoint("S", 1, "1")).toBe("S01-1");
    expect(fmtFullPoint("S", 1, null)).toBe("S01");
    expect(parsePoint("S01-2")).toMatchObject({ prefix: "S", num: 1, suffix: "2" });
    expect(nextPointLabel(["S01", "S01-1", "S01-2", "S02"], "S")).toBe("S03");
    const cells = pointCells(
      [
        { point: "S01", result: "합격" },
        { point: "S01-1", result: "합격" },
        { point: "S01-2", result: "불합격" },
        { point: "S02", result: "불합격" },
      ],
      "S",
    );
    expect(cells.map((cell) => [cell.label, cell.status, cell.subs.join("+")])).toEqual([
      ["S01", "pass", "1+2"],
      ["S02", "fail", ""],
    ]);
    const gapped = pointCells(
      [
        { point: "S14", result: "합격" },
        { point: "S16", result: "합격" },
        { point: "S15", result: "없음" },
      ],
      "S",
    );
    expect(gapped.map((cell) => cell.label)).toEqual(["S14", "S15", "S16"]);
    expect(gapped[1].status).toBe("skip");
    expect(nextPointLabel(["S14", "S16", "S15"], "S")).toBe("S17");
  });

  it("카드는 회색, 합격, 불합격 순으로 돌고 + 는 빈 번호를 만든다", () => {
    expect(cycleCard("draft")).toBe("pass");
    expect(cycleCard("pass")).toBe("fail");
    expect(cycleCard("fail")).toBe("draft");
    expect(reinspectLabel("S15")).toBe("S15");
    expect(reinspectLabel("r01-6")).toBe("R01-6");
    expect(canonicalPoint("s1")).toBe("S01");
    expect(nextDraftPoint({ direction: "up", prefix: "S", occupied: [14, 16], lastEntered: 16 })).toBe("S17");
    expect(nextDraftPoint({ direction: "down", prefix: "S", occupied: [29, 28], lastEntered: 28 })).toBe("S27");
    expect(nextDraftPoint({ direction: "up", prefix: "S", occupied: [16, 17], lastEntered: 16 })).toBe("S18");
  });

  it("S 번호는 같은 R 번호로 대응한다", () => {
    expect(mirrorPoint("S05")).toBe("R05");
    expect(mirrorPoint("s05-1")).toBe("R05-1");
    expect(mirrorPoint("R05")).toBeNull();
  });

  it("S 초안을 R로 복사할 때 있는 번호는 건너뛴다", () => {
    const planned = planPrefixCopy({
      from: "S",
      to: "R",
      drafts: [
        { point: "S05", result: "합격", ts: 1 },
        { point: "S05-1", result: "합격", ts: 2 },
        { point: "S06", result: "불합격", ts: 3 },
        { point: "S07", result: "합격", ts: 4 },
        { point: "R01", result: "합격", ts: 5 },
      ],
      existing: ["R07", "R08"],
    });
    expect(planned).toEqual([
      { point: "R05", result: "합격" },
      { point: "R05-1", result: "합격" },
      { point: "R06", result: "불합격" },
    ]);
  });

  it("카드 순서는 번호로 고정되고 재검사는 원본 바로 뒤다", () => {
    const cards = [
      { point: "S02", re: false },
      { point: "S01-1", re: false },
      { point: "S01", re: true },
      { point: "S01", re: false },
      { point: "R01", re: false },
    ];
    expect(orderPointCards(cards, "up").map((card) => `${card.point}${card.re ? "재" : ""}`)).toEqual(["S01", "S01재", "S01-1", "S02", "R01"]);
    expect(orderPointCards(cards, "down").map((card) => `${card.point}${card.re ? "재" : ""}`)).toEqual(["S02", "S01", "S01재", "S01-1", "R01"]);
    const tapped = cards.map((card) => ({ ...card, result: card.point === "S02" ? "불합격" : "합격" }));
    expect(orderPointCards(tapped, "up").map((card) => card.point)).toEqual(orderPointCards(cards, "up").map((card) => card.point));
  });

  it("내림이면 마지막 번호 아래의 빈 번호를 제안한다", () => {
    expect(directionFromTrail([29, 28])).toBe("down");
    expect(directionFromTrail([14, 16])).toBe("up");
    expect(suggestNumber({ direction: "down", recorded: [29, 28], absent: [], lastEntered: 28 })).toBe(27);
    expect(suggestNumber({ direction: "down", recorded: [29], absent: [28], lastEntered: 29 })).toBe(27);
    expect(stepNumber(27, "down", 2, [29, 28, 25], [])).toBe(24);
    expect(suggestNumber({ direction: "up", recorded: [14, 16], absent: [15], lastEntered: 16 })).toBe(17);
  });

  it("접두어를 R로 바꾸면 그 라인의 다음 R. 기록이 없으면 R01", () => {
    expect(nextPointLabel(["S01", "S02", "S10"], "R")).toBe("R01");
    expect(nextPointLabel(["S08", "R21"], "R")).toBe("R22");
  });

  it("범위 입력은 기본 합격, 적은 번호만 불합격", () => {
    const rows = expandRange("S05", "s10", ["S07", "s09"]);
    expect(rows).toEqual([
      { point: "S05", result: "합격" },
      { point: "S06", result: "합격" },
      { point: "S07", result: "불합격" },
      { point: "S08", result: "합격" },
      { point: "S09", result: "불합격" },
      { point: "S10", result: "합격" },
    ]);
    expect(expandRange("S10", "S05", [])).toBeNull();
    expect(expandRange("S01", "R03", [])).toBeNull();
  });
});
