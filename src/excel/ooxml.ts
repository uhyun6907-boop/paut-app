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

const ROW_FORMAT_ATTRS = ["ht", "customHeight", "s", "customFormat"];

/** toRow 의 칸마다 fromRow 같은 열 칸의 서식(s)을, 행에는 fromRow 의 높이·행 서식을 입힌다. 값은 그대로. */
export function copyRowStyle(xml: string, fromRow: number, toRow: number): string {
  const from = new RegExp(`<row r="${fromRow}"([^>]*?)(?:/>|>([\\s\\S]*?)</row>)`).exec(xml);
  if (!from) return xml;
  let next = xml;
  for (const c of (from[2] ?? "").matchAll(/<c r="([A-Z]+)\d+"([^>]*?)\/?>/g)) {
    const style = /(?:^|\s)s="(\d+)"/.exec(c[2])?.[1];
    if (!style) continue;
    const ref = c[1] + toRow;
    const cellRe = new RegExp(`<c r="${ref}"([^>]*?)(/?>)`);
    if (!cellRe.test(next)) next = setCell(next, ref, { kind: "empty" });
    next = next.replace(cellRe, (_all, attrs: string, end: string) => `<c r="${ref}" s="${style}"${attrs.replace(/\s+s="\d+"/, "")}${end}`);
  }
  const copied = ROW_FORMAT_ATTRS.flatMap((name) => {
    const value = new RegExp(`\\s${name}="([^"]*)"`).exec(from[1])?.[1];
    return value === undefined ? [] : [` ${name}="${value}"`];
  }).join("");
  return next.replace(new RegExp(`<row r="${toRow}"([^>]*?)(/?>)`), (_all, attrs: string, end: string) => {
    const kept = ROW_FORMAT_ATTRS.reduce((a, name) => a.replace(new RegExp(`\\s${name}="[^"]*"`), ""), attrs);
    return `<row r="${toRow}"${kept}${copied}${end}`;
  });
}

/**
 * lo..hi 행의 firstCol..lastCol 칸을 sources 순서로 옮긴다(lo+i 행 ← sources[i] 행). 칸 서식(s)·행 높이도 따라가고,
 * 수식의 상대 참조는 옮긴 만큼 바꾼다. 그 행들에 걸친 공유 수식은 같은 뜻의 보통 수식으로 풀어 둔다. 그 밖의 칸은 그대로.
 */
export function moveRowCells(sheetXml: string, lo: number, sources: number[], firstCol: number, lastCol: number): string {
  const xml = unshareFormulas(sheetXml, new Set(sources));
  const rowRe = (r: number) => new RegExp(`<row r="${r}"([^>]*?)(?:/>|>([\\s\\S]*?)</row>)`);
  const original = new Map(sources.map((src) => [src, rowRe(src).exec(xml)]));
  const cellsOf = (body: string | undefined) =>
    [...(body ?? "").matchAll(/<c r="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((m) => ({ col: columnIndex(m[1]), letters: m[1], xml: m[0] }));
  const inside = (col: number) => col >= firstCol && col <= lastCol;
  let next = xml;
  sources.forEach((src, i) => {
    const target = lo + i;
    if (src === target) return;
    const from = original.get(src);
    const to = original.get(target) ?? rowRe(target).exec(xml);
    if (!from || !to) throw new Error(`${from ? target : src}행이 없습니다`);
    const moved = cellsOf(from[2])
      .filter((c) => inside(c.col))
      .map((c) => ({
        ...c,
        xml: c.xml
          .replace(`<c r="${c.letters}${src}"`, `<c r="${c.letters}${target}"`)
          .replace(/<f>([^<]*)<\/f>/, (_all, f: string) => `<f>${shiftRefs(f, target - src, 0)}</f>`),
      }));
    const kept = cellsOf(to[2]).filter((c) => !inside(c.col));
    const cells = [...kept, ...moved].sort((a, b) => a.col - b.col).map((c) => c.xml).join("");
    const attrs = ROW_FORMAT_ATTRS.reduce((a, name) => a.replace(new RegExp(`\\s${name}="[^"]*"`), ""), to[1]) +
      ROW_FORMAT_ATTRS.flatMap((name) => {
        const value = new RegExp(`\\s${name}="([^"]*)"`).exec(from[1])?.[1];
        return value === undefined ? [] : [` ${name}="${value}"`];
      }).join("");
    const current = rowRe(target).exec(next)!;
    next = next.slice(0, current.index) + `<row r="${target}"${attrs}>${cells}</row>` + next.slice(current.index + current[0].length);
  });
  return next;
}

/** rows 에 칸이 있는 공유 수식 묶음 전체를 칸마다 보통 수식(<f>…</f>)으로. */
function unshareFormulas(xml: string, rows: Set<number>): string {
  const sharedRe = /<f\b([^>]*\bt="shared"[^>]*?)(?:\/>|>([^<]*)<\/f>)/g;
  const siOf = (attrs: string) => /\bsi="(\d+)"/.exec(attrs)?.[1] ?? "";
  const cellAt = (index: number) => {
    const start = xml.lastIndexOf('<c r="', index);
    const m = /^<c r="([A-Z]+)(\d+)"/.exec(xml.slice(start, start + 16))!;
    return { col: columnIndex(m[1]), row: Number(m[2]) };
  };
  const masters = new Map<string, { col: number; row: number; formula: string }>();
  const groups = new Set<string>();
  for (const m of xml.matchAll(sharedRe)) {
    const at = cellAt(m.index);
    if (/\bref="/.test(m[1])) masters.set(siOf(m[1]), { ...at, formula: m[2] ?? "" });
    if (rows.has(at.row)) groups.add(siOf(m[1]));
  }
  if (!groups.size) return xml;
  return xml.replace(sharedRe, (all: string, attrs: string, text: string | undefined, offset: number) => {
    const si = siOf(attrs);
    if (!groups.has(si)) return all;
    const master = masters.get(si);
    if (!master) throw new Error(`공유 수식 ${si} 원본 없음`);
    const at = cellAt(offset);
    return `<f>${/\bref="/.test(attrs) ? (text ?? "") : shiftRefs(master.formula, at.row - master.row, at.col - master.col)}</f>`;
  });
}

/** 수식의 상대 참조를 dRow 행·dCol 열만큼(F152 → F151). $ 가 붙은 쪽·따옴표 안 글자는 그대로. */
function shiftRefs(formula: string, dRow: number, dCol: number): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) =>
      i % 2
        ? part
        : part.replace(/(?<![A-Za-z0-9_$.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (_all, cAbs: string, col: string, rAbs: string, row: string) =>
            `${cAbs}${cAbs ? col : columnLetter(columnIndex(col) + dCol)}${rAbs}${rAbs ? row : Number(row) + dRow}`,
          ),
    )
    .join("");
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
