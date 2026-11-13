import { describe, expect, it } from "vitest";
import { JACKET_ROWS, jacketButtonLabel, jacketPaut, jacketScan } from "./jacket";

describe("이중보온관 사양", () => {
  it("28규격이고 200A는 내관 216.3 · 외관 315다", () => {
    expect(JACKET_ROWS).toHaveLength(28);
    const row = JACKET_ROWS.find((item) => item.a === 200);
    expect(row).toMatchObject({
      mat: "SPPS 38E",
      innerT: "6.4",
      innerD: "216.3",
      outerT: "5.6",
      outerD: "315",
      insul: "43.80",
      kg: "41.21",
      len: "6000 / 12000",
    });
    expect(jacketButtonLabel(row!)).toBe('200A / 8" · 710');
    expect(jacketScan(200)).toBe(710);
    expect(jacketScan(50)).toBe(220);
    expect(jacketScan(1000)).toBe(3250);
    expect(jacketScan(20)).toBeNull();
    expect(jacketScan(40)).toBeNull();
    expect(jacketScan(1100)).toBeNull();
    expect(jacketScan(1200)).toBeNull();
    expect(jacketPaut(200)?.ds).toBeUndefined();
    expect(jacketPaut(250)?.ds).toBe(3);
    expect(jacketPaut(250)?.dl).toBe(2);
  });

  it("작은 관은 6000, 125A부터는 둘 다 있다", () => {
    expect(JACKET_ROWS.find((item) => item.a === 100)?.len).toBe("6000");
    expect(JACKET_ROWS.find((item) => item.a === 125)?.len).toBe("6000 / 12000");
    expect(JACKET_ROWS.find((item) => item.a === 1200)?.outerD).toBe("1400");
    expect(jacketButtonLabel(JACKET_ROWS[0])).toBe('20A / 3/4" · —');
  });
});