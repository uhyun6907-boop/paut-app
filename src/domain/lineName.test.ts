import { describe, expect, it } from "vitest";
import { lineShape, suggestNextLines } from "./lineName";

describe("새 라인 이름 제안", () => {
  it("관경(뒤에 A) 숫자는 번호가 아니다", () => {
    expect(lineShape("군포산본 이원화 750A-5")).toMatchObject({ head: "군포산본 이원화 750A-", num: 5, tail: "" });
    expect(lineShape("HJ-M-005-350A")).toMatchObject({ head: "HJ-M-", num: 5, width: 3, tail: "-350A" });
    expect(lineShape("750A")).toBeNull();
  });

  it("같은 모양끼리 가장 큰 번호 + 1(자릿수 유지), 최근 묶음부터, 이미 있는 이름은 건너뜀", () => {
    const lines = ["군포산본 이원화 750A-5", "HJ-M-005-350A", "군포산본 이원화 750A-3", "HJ-M-007-350A", "HJ-M-008-350A"];
    expect(suggestNextLines(lines)).toEqual([
      { line: "군포산본 이원화 750A-6", from: "군포산본 이원화 750A-5" },
      { line: "HJ-M-009-350A", from: "HJ-M-005-350A" },
    ]);
  });

  it("마전(실제 리스트): MJDH-D-D-P-58-001~006-200A → 007, 58·200A 는 그대로", () => {
    const lines = ["MJDH-D-D-P-58-003-200A", "MJDH-D-D-P-58-005-200A", "MJDH-D-D-P-58-004-200A", "MJDH-D-D-P-58-006-200A", "MJDH-D-D-P-58-002-200A", "MJDH-D-D-P-58-001-200A"];
    expect(suggestNextLines(lines)).toEqual([{ line: "MJDH-D-D-P-58-007-200A", from: "MJDH-D-D-P-58-003-200A" }]);
  });
});
