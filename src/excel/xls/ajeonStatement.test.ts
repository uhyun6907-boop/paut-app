/**
 * 아전 .xls 명세서 쓰기(앱 동기화는 engine.ts ajeonStatementStep, ajeonSync.test.ts).
 * 로컬 fixture(LibreOffice 로 만든 .xls) 복사본에서만 시험한다. 실제 OneDrive 파일은 받지도 쓰지도 않는다.
 * 실제 파일(엑셀이 쓴 .xls)의 로컬 사본은 OLE FAT 이 깨져 있어 그 파일로는 왕복 확인을 못 했다(FAT 이 깨진 파일은 쓰지 않는다).
 */
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { isoToExcelSerial } from "../../domain/dates";
import { applyXlsEdits, expectedXlsSheetName, findPrice, openXlsBook, planAjeonFill, verifyXlsEdit, xlsSheetNames, xlsStatementRoute } from "./ajeonStatement";
import { cellKey, readXlsSheet } from "./biff";
import { CfbError, readCfb } from "./cfb";

const load = (name: string) => new Uint8Array(readFileSync(`src/excel/fixtures/${name}`));

describe("아전 .xls 명세서 쓰기", () => {
  it("현장 → 파일: 아전 군포만(마곡은 시트 있을 때만), 다른 현장은 .xlsx 명세서", () => {
    expect(xlsStatementRoute("아전 군포 PA")).toEqual({ prefix: "아전 군포", optional: false });
    expect(xlsStatementRoute("아전 마곡")).toEqual({ prefix: "아전 마곡", optional: true });
    expect(xlsStatementRoute("지성 군포 PA")).toBeNull();
    expect(xlsStatementRoute("승재영종")).toBeNull();
  });

  it("그 달 시트 이름은 파일 표기를 따른다(없으면 만들 이름): 아전 군포 2611", () => {
    const names = xlsSheetNames(openXlsBook(load("ajeon-real-sheets.xls")));
    expect(names).not.toContain("아전 군포 2611");
    expect(expectedXlsSheetName(names, "아전 군포", 2026, 11)).toBe("아전 군포 2611");
  });

  it("단가: 2611 에 없으면 2610 의 183,600", () => {
    const book = openXlsBook(load("ajeon-lo-styled.xls"));
    expect(findPrice(book, "아전 군포", "아전 군포 2611", "750A")).toEqual({ price: 183600, from: "아전 군포 2610" });
  });

  it("왕복(복사본): 2611 의 바꿀 칸만 바뀌고 승재·다른 시트는 바이트 그대로, 검증 통과", () => {
    const before = load("ajeon-lo-styled.xls");
    const book = openXlsBook(before);
    const plan = planAjeonFill(book, "아전 군포", "아전 군포 2611", [
      { date: "2026-11-03", diameter: "750A", count: 4 },
      { date: "2026-11-05", diameter: "750A", count: 2 },
    ]);
    expect(plan.written).toBe(2);
    expect(plan.edits).toHaveLength(12);
    const after = applyXlsEdits(before, [{ sheet: "아전 군포 2611", edits: plan.edits }]);
    expect(() => verifyXlsEdit(before, after, [{ sheet: "아전 군포 2611", edits: plan.edits }])).not.toThrow();
    const wb = openXlsBook(after).wb;
    const s = readXlsSheet(wb, "아전 군포 2611");
    expect(s.cells.get(cellKey(13, 0))?.value).toEqual({ kind: "number", n: isoToExcelSerial("2026-11-03") });
    expect(s.cells.get(cellKey(13, 5))?.value).toEqual({ kind: "number", n: 4 });
    expect(s.cells.get(cellKey(13, 6))?.value).toEqual({ kind: "number", n: 183600 });
    for (const c of s.cells.values()) if (c.value.kind === "string") expect(c.value.text).not.toMatch(/앱/);
    // 다시 보내도 바뀌는 칸 없음
    const again = planAjeonFill(openXlsBook(after), "아전 군포", "아전 군포 2611", [
      { date: "2026-11-03", diameter: "750A", count: 4 },
      { date: "2026-11-05", diameter: "750A", count: 2 },
    ], plan.owned);
    expect(again.edits).toHaveLength(0);
  });

  it("승재 시트에 쓰려 하면 검증에서 막는다", () => {
    const before = load("ajeon-lo-styled.xls");
    const sj = xlsSheetNames(openXlsBook(before)).find((n) => /승재/.test(n))!;
    expect(() => applyXlsEdits(before, [{ sheet: sj, edits: [{ row: 40, col: 0, value: { kind: "number", n: 1 } }] }])).toThrow();
  });

  it("FAT 이 깨진 파일은 읽지 않는다(쓰지도 않음)", () => {
    const bytes = load("ajeon-lo-styled.xls").slice();
    // 첫 FAT 섹터를 0 으로 채운다(로컬 실제 사본과 같은 손상)
    const fat0 = new DataView(bytes.buffer).getUint32(76, true);
    bytes.fill(0, (fat0 + 1) * 512, (fat0 + 2) * 512);
    expect(() => readCfb(bytes)).toThrow(CfbError);
  });
});
