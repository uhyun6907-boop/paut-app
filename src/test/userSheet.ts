/**
 * 테스트 전용: 사용자가 PC 엑셀에서 거래명세서 새 달 시트를 직접 만든 상황을 흉내 낸다.
 * 앱(운영 코드)은 시트를 만들지 않는다. 이 파일은 src/test 에만 있고 앱 번들에 들어가지 않는다.
 * (지난달 시트를 복사하고 PAUT 줄만 비운 '빈 새 달 시트')
 */
import { strToU8 } from "fflate";
import { findMonthSheet, latestSheet, suggestSheetName } from "../domain/mapping";
import { cellText, dropCalcChain, readSheet, setCell, setXml, workbookSheets, xmlText, type ZipFiles } from "../excel/ooxml";
import { readStatement } from "../excel/statementBook";

/** 같은 접두어 최근 시트를 복사해 그 달 빈 시트를 만든다(이미 있으면 그대로). 만든(또는 있던) 시트 이름. */
export function userCreatesMonthSheet(files: ZipFiles, prefix: string, year: number, month: number): string {
  const names = workbookSheets(files).map((s) => s.name);
  const existing = findMonthSheet(names, prefix, year, month);
  if (existing) return existing.name;
  const latest = latestSheet(names, prefix);
  if (!latest) throw new Error(`no sheet for ${prefix}`);
  const name = suggestSheetName(prefix, latest, year, month);
  copySheet(files, latest.name, name, names.indexOf(latest.name) + 1);
  const { cells, xml: original } = readSheet(files, name);
  const total = readStatement(files, name).totalRow;
  let xml = original;
  for (let r = 12; r < total; r++) {
    const method = cellText(cells, "B", r).replace(/\s/g, "");
    if (method.includes("방사선") || method.includes("RT") || cellText(cells, "E", r) === "매") continue;
    for (const col of ["A", "B", "C", "D", "E", "F", "G", "H", "I"]) xml = setCell(xml, col + r, { kind: "empty" });
  }
  const info = workbookSheets(files).find((s) => s.name === name)!;
  setXml(files, info.path, xml);
  return name;
}

function copySheet(files: ZipFiles, sourceName: string, newName: string, index: number): void {
  const sheets = workbookSheets(files);
  const source = sheets.find((s) => s.name === sourceName)!;
  const n = Math.max(...sheets.map((s) => Number(s.path.match(/sheet(\d+)\.xml$/)?.[1] || 0)), 0) + 1;
  const destPath = `xl/worksheets/sheet${n}.xml`;
  files[destPath] = strToU8(xmlText(files, source.path).replace(/\s+tabSelected="1"/g, ""));
  const ct = "[Content_Types].xml";
  setXml(files, ct, xmlText(files, ct).replace("</Types>", `<Override PartName="/${destPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`));
  const relPath = "xl/_rels/workbook.xml.rels";
  const relXml = xmlText(files, relPath);
  const rid = `rId${Math.max(...[...relXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1])), 0) + 1}`;
  setXml(
    files,
    relPath,
    relXml.replace(
      "</Relationships>",
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/></Relationships>`,
    ),
  );
  const wbPath = "xl/workbook.xml";
  const wb = xmlText(files, wbPath);
  const sheetId = Math.max(...[...wb.matchAll(/sheetId="(\d+)"/g)].map((m) => Number(m[1])), 0) + 1;
  const tag = `<sheet name="${newName}" sheetId="${sheetId}" r:id="${rid}"/>`;
  let pos = 0;
  let out = wb.replace(/<sheet\b[^>]*\/>/g, (t) => (pos++ === index ? tag + t : t));
  if (index >= pos) out = out.replace("</sheets>", tag + "</sheets>");
  // 시트 순서 번호를 쓰는 곳을 한 칸씩 민다.
  const shift = (v: number) => (v >= index ? v + 1 : v);
  out = out
    .replace(/localSheetId="(\d+)"/g, (_f, v: string) => `localSheetId="${shift(Number(v))}"`)
    .replace(/<workbookView\b[^>]*>/g, (t) => t.replace(/\b(activeTab|firstSheet)="(\d+)"/g, (_f, k: string, v: string) => `${k}="${shift(Number(v))}"`));
  setXml(files, wbPath, out);
  dropCalcChain(files);
}
