import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export type ZipFiles = Record<string, Uint8Array>;

export function unzipXlsx(data: Uint8Array): ZipFiles {
  return unzipSync(data);
}

export function zipXlsx(files: ZipFiles): Uint8Array {
  return zipSync(files, { level: 1 });
}

export function decodeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export function escapeXml(s: string): string {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function xmlText(files: ZipFiles, path: string): string {
  const bytes = files[path];
  if (!bytes) throw new Error(`xlsx 안에 ${path} 가 없습니다.`);
  return strFromU8(bytes);
}

export function setXml(files: ZipFiles, path: string, xml: string): void {
  files[path] = strToU8(xml);
}

export interface XCell {
  ref: string;
  col: string;
  row: number;
  v?: string;
  f?: string;
  t?: string;
  s?: string;
}

export function parseSharedStrings(xml: string): string[] {
  const out: string[] = [];
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const texts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => decodeXml(x[1]));
    out.push(texts.join(""));
  }
  return out;
}

/** 값이 있는 셀만. 65,536개의 빈 서식 셀은 건너뛴다. */
export function parseCells(xml: string, shared: string[]): Map<string, XCell> {
  const map = new Map<string, XCell>();
  // `(?<!/)>` so self-closing `<c .../>` is not treated as the start of a huge cell.
  const re = /<c r="([A-Z]+)(\d+)"([^>]*?)(?<!\/)>([\s\S]*?)<\/c>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const inner = m[4];
    if (!inner.includes("<v") && !inner.includes("<f") && !inner.includes("<t")) continue;
    const attrs = m[3] || "";
    const t = /(?:^|\s)t="([^"]+)"/.exec(attrs)?.[1];
    const s = /(?:^|\s)s="(\d+)"/.exec(attrs)?.[1];
    const fRaw = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(inner)?.[1];
    const vRaw = /<v>([\s\S]*?)<\/v>/.exec(inner)?.[1];
    let v = vRaw != null ? decodeXml(vRaw) : undefined;
    if (t === "s" && v != null) v = shared[Number(v)] ?? v;
    if (t === "inlineStr") {
      const it = /<t[^>]*>([\s\S]*?)<\/t>/.exec(inner)?.[1];
      v = it != null ? decodeXml(it) : "";
    }
    if (v == null && !fRaw) continue;
    const ref = m[1] + m[2];
    map.set(ref, { ref, col: m[1], row: Number(m[2]), v, f: fRaw != null ? decodeXml(fRaw) : undefined, t, s });
  }
  return map;
}

export function cell(cells: Map<string, XCell>, col: string, row: number): XCell | undefined {
  return cells.get(col + row);
}

export function cellText(cells: { get(key: string): { v?: string } | undefined }, col: string, row: number): string {
  return cells.get(col + row)?.v?.trim() ?? "";
}

export type CellSpec =
  | { kind: "empty" }
  | { kind: "inline"; text: string }
  | { kind: "number"; n: number }
  | { kind: "formula"; formula: string; cached?: string | number; cachedType?: "str" | "n" };

function renderCell(ref: string, spec: CellSpec, style?: string): string {
  const sAttr = style ? ` s="${style}"` : "";
  if (spec.kind === "empty") return `<c r="${ref}"${sAttr}/>`;
  if (spec.kind === "inline") {
    return `<c r="${ref}"${sAttr} t="inlineStr"><is><t>${escapeXml(spec.text)}</t></is></c>`;
  }
  if (spec.kind === "number") return `<c r="${ref}"${sAttr}><v>${spec.n}</v></c>`;
  const t = spec.cachedType === "str" ? ` t="str"` : "";
  const cached = spec.cached != null ? `<v>${escapeXml(String(spec.cached))}</v>` : "";
  return `<c r="${ref}"${sAttr}${t}><f>${escapeXml(spec.formula)}</f>${cached}</c>`;
}

/**
 * 기존 셀의 스타일(s)은 유지하고 값만 바꾼다.
 * 없는 셀은 엑셀처럼 행 서식(customFormat) → 열 기본 서식(<cols>) 순으로 서식을 물려받고(예: 군포·마전 리스트 E열 날짜 서식),
 * 행 안에서 열 순서 자리에 넣는다(순서가 어긋나면 엑셀이 '복구'를 묻는다). 행이 없으면 행 번호 순서 자리에 만든다.
 */
export function setCell(xml: string, ref: string, spec: CellSpec): string {
  const re = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`);
  const found = re.exec(xml);
  if (found) {
    const style = /(?:^|\s)s="(\d+)"/.exec(found[1])?.[1];
    const next = renderCell(ref, spec, style);
    return xml.slice(0, found.index) + next + xml.slice(found.index + found[0].length);
  }
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) throw new Error(`셀 주소 오류 ${ref}`);
  const [, colLetters, rowNum] = m;
  const col = columnIndex(colLetters);
  const rowRe = new RegExp(`<row r="${rowNum}"([^>]*?)(/>|>([\\s\\S]*?)</row>)`);
  const row = rowRe.exec(xml);
  const rowAttrs = row?.[1] ?? "";
  const rowStyle = /\bcustomFormat="(1|true)"/.test(rowAttrs) ? /(?:^|\s)s="(\d+)"/.exec(rowAttrs)?.[1] : undefined;
  const cellXml = renderCell(ref, spec, rowStyle ?? columnStyle(xml, col));
  if (!row) {
    const sheetData = /<sheetData\s*\/>/.exec(xml);
    if (sheetData) return xml.slice(0, sheetData.index) + `<sheetData><row r="${rowNum}">${cellXml}</row></sheetData>` + xml.slice(sheetData.index + sheetData[0].length);
    const end = xml.lastIndexOf("</sheetData>");
    if (end < 0) throw new Error("sheetData 없음");
    const rowsRe = /<row r="(\d+)"/g;
    let at = end;
    for (let r: RegExpExecArray | null; (r = rowsRe.exec(xml)); ) {
      if (Number(r[1]) > Number(rowNum)) {
        at = r.index;
        break;
      }
    }
    return xml.slice(0, at) + `<row r="${rowNum}">${cellXml}</row>` + xml.slice(at);
  }
  const open = `<row r="${rowNum}"${rowAttrs}>`;
  const body = row[2] === "/>" ? "" : row[3] ?? "";
  let insertAt = body.length;
  const cellsRe = /<c r="([A-Z]+)\d+"/g;
  for (let c: RegExpExecArray | null; (c = cellsRe.exec(body)); ) {
    if (columnIndex(c[1]) > col) {
      insertAt = c.index;
      break;
    }
  }
  const nextRow = open + body.slice(0, insertAt) + cellXml + body.slice(insertAt) + "</row>";
  return xml.slice(0, row.index) + nextRow + xml.slice(row.index + row[0].length);
}

/** <cols> 의 열 기본 서식 */
function columnStyle(xml: string, col: number): string | undefined {
  const cols = /<cols>([\s\S]*?)<\/cols>/.exec(xml)?.[1];
  if (!cols) return undefined;
  for (const c of cols.matchAll(/<col\b([^>]*)\/?>/g)) {
    const min = Number(/\bmin="(\d+)"/.exec(c[1])?.[1]);
    const max = Number(/\bmax="(\d+)"/.exec(c[1])?.[1]);
    if (col >= min && col <= max) {
      const s = /\bstyle="(\d+)"/.exec(c[1])?.[1];
      return s && s !== "0" ? s : undefined;
    }
  }
  return undefined;
}

export function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function columnLetter(index: number): string {
  let n = index;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export interface SheetInfo {
  name: string;
  path: string;
  rid: string;
}

function tagAttrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = decodeXml(m[2]);
  return out;
}

export function workbookSheets(files: ZipFiles): SheetInfo[] {
  const wb = xmlText(files, "xl/workbook.xml");
  const rels = xmlText(files, "xl/_rels/workbook.xml.rels");
  const ridToTarget = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const a = tagAttrs(m[0]);
    if (a.Id && a.Target) ridToTarget.set(a.Id, a.Target);
  }
  const sheets: SheetInfo[] = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*\/>/g)) {
    const a = tagAttrs(m[0]);
    const rid = a["r:id"];
    const name = a.name;
    if (!rid || !name) continue;
    let target = ridToTarget.get(rid) || "";
    target = target.replace(/^\//, "");
    if (target.startsWith("../")) target = target.replace(/^\.\.\//, "");
    const path = target.startsWith("xl/") ? target : `xl/${target}`;
    sheets.push({ name, path, rid });
  }
  return sheets;
}

export function sharedStrings(files: ZipFiles): string[] {
  if (!files["xl/sharedStrings.xml"]) return [];
  return parseSharedStrings(xmlText(files, "xl/sharedStrings.xml"));
}

export function readSheet(files: ZipFiles, sheetName: string): { info: SheetInfo; xml: string; cells: Map<string, XCell> } {
  const info = workbookSheets(files).find((s) => s.name === sheetName);
  if (!info) throw new Error(`시트를 찾지 못했습니다: ${sheetName}`);
  const xml = xmlText(files, info.path);
  return { info, xml, cells: parseCells(xml, sharedStrings(files)) };
}

export function dropCalcChain(files: ZipFiles): void {
  delete files["xl/calcChain.xml"];
  const relPath = "xl/_rels/workbook.xml.rels";
  if (files[relPath]) {
    const rels = xmlText(files, relPath).replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, "");
    setXml(files, relPath, rels);
  }
  if (files["[Content_Types].xml"]) {
    const ct = xmlText(files, "[Content_Types].xml").replace(/<Override\b[^>]*calcChain[^>]*\/>/g, "");
    setXml(files, "[Content_Types].xml", ct);
  }
}

/**
 * fromRow 이상인 행 번호를 delta 만큼 민다.
 * 셀 주소, 수식 안의 참조, 행 태그를 한 번씩만 바꾼다.
 */
export function shiftRowNumbers(xml: string, fromRow: number, delta: number): string {
  const bumped = xml.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (full, dol1: string, col: string, dol2: string, row: string) => {
    const r = Number(row);
    if (r >= fromRow) return `${dol1}${col}${dol2}${r + delta}`;
    return full;
  });
  return bumped.replace(/<row r="(\d+)"/g, (full, row: string) => {
    const r = Number(row);
    return r >= fromRow ? `<row r="${r + delta}"` : full;
  });
}

export function extendSumFormulas(xml: string, dataStart: number, dataEnd: number): string {
  return xml.replace(/SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)/g, (full, c1: string, r1: string, c2: string, r2: string) => {
    const a = Number(r1);
    const b = Number(r2);
    if (a <= dataStart && b >= dataStart && b < dataEnd) return `SUM(${c1}${r1}:${c2}${dataEnd})`;
    return full;
  });
}
