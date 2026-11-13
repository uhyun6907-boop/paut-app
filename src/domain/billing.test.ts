import { describe, expect, it } from "vitest";
import { buildBill, countByDiameter, iljiDiameterLabel, isPassing } from "./billing";
import { defaultWeldLengths } from "./mapping";
import type { InspectionRecord } from "./types";

function rec(partial: Partial<InspectionRecord> & Pick<InspectionRecord, "result" | "diameter" | "point">): InspectionRecord {
  return {
    id: partial.id ?? partial.point,
    site: partial.site ?? "드림 합정주공",
    sheetName: " 드림 합정주공",
    date: partial.date ?? "2026-10-28",
    line: partial.line ?? "HJ-M-005-350A",
    point: partial.point,
    result: partial.result,
    diameter: partial.diameter,
    material: "SPPS380",
    sch: "",
    reportNo: "",
    reportDate: "",
    defectNo: "",
    drawingNo: "",
    weldNo: "",
    rtSheets: "",
    flawType: "",
    flawLen: "",
    note: "",
    issued: "",
    src: "user",
    sync: "pending",
    deleted: partial.deleted,
    ts: 1,
  };
}

describe("개소와 청구", () => {
  const lengths = defaultWeldLengths();

  it("합격만 개소. S01과 R01은 2개소, 불합격은 빠진다", () => {
    const rows = [
      rec({ point: "S01", result: "합격", diameter: "350A" }),
      rec({ point: "R01", result: "합격", diameter: "350A" }),
      rec({ point: "S02", result: "불합격", diameter: "350A" }),
      rec({ point: "R02", result: "합격", diameter: "200A" }),
      rec({ point: "S03", result: "합격", diameter: "350A", deleted: true }),
    ];
    expect(isPassing("불합격")).toBe(false);
    expect(countByDiameter(rows)).toEqual({ "350A": 2, "200A": 1 });
  });

  it("같은 번호의 재검사 합격만 1개소이고 불합격 원본은 빠진다", () => {
    const rows = [
      rec({ id: "fail", point: "S15", result: "불합격", diameter: "350A" }),
      rec({ id: "again", point: "S15", result: "합격", diameter: "350A" }),
    ];
    expect(countByDiameter(rows)).toEqual({ "350A": 1 });
  });

  it("350A 8·10·10개소 × 1.2m × 75,000원. 합계는 천 원 절사", () => {
    const rows: InspectionRecord[] = [];
    const push = (date: string, n: number) => {
      for (let i = 0; i < n; i++) rows.push(rec({ date, point: `${date}-S${i}`, result: "합격", diameter: "350A" }));
    };
    push("2026-10-15", 8);
    push("2026-10-21", 10);
    push("2026-10-28", 10);
    rows.push(rec({ date: "2026-10-28", point: "NG", result: "불합격", diameter: "350A" }));
    const bill = buildBill(rows, 75000, lengths);
    expect(bill.lines.map((l) => [l.date, l.count, l.qtyM, l.amount])).toEqual([
      ["2026-10-15", 8, 9.6, 720000],
      ["2026-10-21", 10, 12, 900000],
      ["2026-10-28", 10, 12, 900000],
    ]);
    expect(bill.totalM).toBeCloseTo(33.6, 5);
    expect(bill.supply).toBe(2520000);
    expect(bill.vat).toBe(252000);
    expect(bill.grand).toBe(2772000);
    expect(bill.missingLengths).toEqual([]);
  });

  it("길이표에 없는 관경은 금액에서 빠진다", () => {
    const bill = buildBill(
      [rec({ point: "S01", result: "합격", diameter: "150A" }), rec({ point: "S02", result: "합격", diameter: "350A" })],
      75000,
      lengths,
    );
    expect(bill.missingLengths).toEqual(["150A"]);
    expect(bill.supply).toBe(1.2 * 75000);
    expect(bill.lines.find((l) => l.diameter === "150A")?.amount).toBeNull();
  });

  it("그 외 열은 현장 이름을 관경에 붙인다", () => {
    expect(iljiDiameterLabel("150A", "그 외", "드림 화양")).toBe("150A(드림화양)");
    expect(iljiDiameterLabel("350A", "드림 합정주공", "드림 합정주공")).toBe("350A");
  });
});
