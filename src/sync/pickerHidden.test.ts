import { describe, expect, it } from "vitest";
import { hiddenInPicker } from "./graph";

describe("파일 고르기에서 숨기는 폴더", () => {
  it("문서·첨부 파일 폴더와 그 안 파일은 숨기고, 거래명세서 폴더는 보인다", () => {
    expect(hiddenInPicker({ id: "1", name: "문서", folder: {} })).toBe(true);
    expect(hiddenInPicker({ id: "2", name: "첨부 파일", folder: {} })).toBe(true);
    expect(hiddenInPicker({ id: "3", name: "Attachments", folder: {} })).toBe(true);
    expect(hiddenInPicker({ id: "4", name: "견적.xlsx", file: {}, parentReference: { path: "/drive/root:/%EC%B2%A8%EB%B6%80%20%ED%8C%8C%EC%9D%BC" } })).toBe(true);
    expect(hiddenInPicker({ id: "5", name: "a.xlsx", file: {}, parentReference: { path: "/drive/root:/문서/하위" } })).toBe(true);
    expect(hiddenInPicker({ id: "6", name: "00_PAUT 거래명세서", folder: {} })).toBe(false);
    expect(hiddenInPicker({ id: "7", name: "02_PAUT 거래명세서 (지성, 드림).xlsx", file: {}, parentReference: { path: "/drive/root:/00_PAUT 거래명세서" } })).toBe(false);
    expect(hiddenInPicker({ id: "8", name: "PAUT NDT 리스트(군포, 마전).xlsx", file: {}, parentReference: { path: "/drive/root:" } })).toBe(false);
  });
});
