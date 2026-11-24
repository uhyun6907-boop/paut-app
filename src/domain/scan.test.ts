import { describe, expect, it } from "vitest";
import { PAUT_OD } from "./reference";
import { sizeFromLine } from "./size";

/** 「위상배열초음파검사 기본 SETTING」 종이 표의 관경(A) → Scan 길이(mm). */
const SHEET: [number, number][] = [
  [50, 220], [65, 270], [80, 310], [100, 390], [125, 470], [150, 550], [200, 710], [250, 870], [300, 1030], [350, 1150],
  [400, 1310], [450, 1470], [500, 1650], [550, 1810], [600, 1970], [650, 2130], [700, 2300], [750, 2450], [800, 2600],
  [850, 2750], [900, 2930], [1000, 3250],
];

describe("스캔 길이", () => {
  it("앱의 표가 종이 SETTING 표와 같다", () => {
    expect(PAUT_OD.map((r) => [r.a, r.scan])).toEqual(SHEET);
  });

  it("라인 이름 끝의 관경으로 스캔 길이를 찾는다(750A-5 처럼 뒤에 번호가 붙어도)", () => {
    expect(sizeFromLine("군포산본 이원화 750A-5", PAUT_OD)).toMatchObject({ label: '30" / 750A', scan: 2450 });
    expect(sizeFromLine("HJ-M-005-350A", PAUT_OD)).toMatchObject({ label: '14" / 350A', scan: 1150 });
  });
});
