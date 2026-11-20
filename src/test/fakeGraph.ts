/**
 * 테스트 전용 가짜 Microsoft Graph(OneDrive + Excel 통합 문서 API).
 * 실제 OneDrive 는 부르지 않는다. 파일은 메모리의 xlsx(ZipFiles)이고,
 * range PATCH/insert 와 content PUT 을 OOXML 로 그대로 흉내 낸다.
 */
import {
  columnIndex,
  columnLetter,
  parseCells,
  setCell,
  setXml,
  sharedStrings,
  shiftRowNumbers,
  unzipXlsx,
  workbookSheets,
  xmlText,
  zipXlsx,
  type XCell,
  type ZipFiles,
} from "../excel/ooxml";

const GRAPH = "https://graph.microsoft.com/v1.0";

export interface Fault {
  /** 예: (m, p) => m === "PATCH" && p.includes("ILJI") */
  match: (method: string, path: string) => boolean;
  status: number;
  /** 남은 횟수. Infinity 면 계속. */
  times: number;
  retryAfter?: number;
}

export interface LogEntry {
  method: string;
  path: string;
  item: string;
  status: number;
  session?: string;
  ifMatch?: string;
  body?: unknown;
}

/** 서식 요청(범위 서식·글꼴·채우기·테두리·표시 형식). 값 쓰기와 가른다. */
export function isFormatRequest(e: LogEntry): boolean {
  return e.path.includes("/format") || !!(e.body as { numberFormat?: unknown } | undefined)?.numberFormat;
}

interface Item {
  name: string;
  files: ZipFiles;
  version: number;
  /** 들어 있는 폴더 이름(없으면 드라이브 루트). 검색은 폴더와 상관없이 찾는다. */
  folder?: string;
  modified: string;
}

export class FakeDrive {
  /** 한 응답에 담을 수 있는 칸 수(넘으면 504). 실제 Graph 의 큰 범위 실패 흉내. */
  static maxCells = 200_000;
  items = new Map<string, Item>();
  log: LogEntry[] = [];
  faults: Fault[] = [];
  /** 이 요청은 응답하지 않는다(아이폰에서 멈춘 요청 흉내). signal 로 끊으면 AbortError. */
  hangs: Fault["match"][] = [];
  /** 이 토큰만 받는다. rotateToken() 으로 바꾸면 이전 토큰은 401. */
  validToken = "token-1";
  /** 순서 확인용 사건 목록(토큰 받기·요청). */
  events: string[] = [];
  private sessionNo = 0;
  private formatEdits = new Map<string, FormatEdit>();
  private rowHeightEdits = new Map<string, number>();
  private sharedCache = new WeakMap<Uint8Array, string[]>();

  put(id: string, name: string, bytes: Uint8Array, opts: { folder?: string; modified?: string } = {}): void {
    this.items.set(id, { name, files: unzipXlsx(bytes), version: 1, folder: opts.folder, modified: opts.modified ?? "2026-11-17T00:00:00Z" });
  }

  /** OneDrive 가 파일을 다시 올려 항목 ID 만 바뀐 것처럼 옮긴다(옛 ID 는 404). */
  reupload(from: string, to: string, folder?: string): void {
    const item = this.items.get(from);
    if (!item) throw new Error(`no item ${from}`);
    this.items.delete(from);
    this.items.set(to, { ...item, folder: folder ?? item.folder });
  }

  /** 다른 곳(PC 엑셀)에서 파일이 바뀐 것처럼 버전만 올린다(eTag 가 달라짐). */
  touch(id: string): void {
    const item = this.items.get(id);
    if (item) item.version += 1;
  }

  files(id: string): ZipFiles {
    const item = this.items.get(id);
    if (!item) throw new Error(`no item ${id}`);
    return item.files;
  }

  bytes(id: string): Uint8Array {
    return zipXlsx(this.files(id));
  }

  rotateToken(): string {
    const n = Number(this.validToken.split("-")[1]) + 1;
    this.validToken = `token-${n}`;
    return this.validToken;
  }

  fail(match: Fault["match"], status: number, times = 1, retryAfter?: number): Fault {
    const f = { match, status, times, retryAfter };
    this.faults.push(f);
    return f;
  }

  clearFaults(): void {
    this.faults = [];
    this.hangs = [];
  }

  requests(filter: (e: LogEntry) => boolean): LogEntry[] {
    return this.log.filter(filter);
  }

  /** Graph 가 돌려줄 칸 서식(파일의 s → styles.xml, 그 위에 서식 PATCH). */
  cellFormat(itemId: string, sheetName: string, ref: string): CellFormat {
    const files = this.files(itemId);
    const sheet = workbookSheets(files).find((s) => s.name === sheetName)!;
    return this.formatOf(itemId, files, sheet.path, ref);
  }

  rowHeight(itemId: string, sheetName: string, row: number): number {
    const files = this.files(itemId);
    const sheet = workbookSheets(files).find((s) => s.name === sheetName)!;
    return this.heightOf(itemId, files, sheet.path, row);
  }

  private formatOf(itemId: string, files: ZipFiles, path: string, ref: string): CellFormat {
    const base = xfFormat(files, xmlText(files, path), ref);
    const edit = this.formatEdits.get(`${itemId}|${path}|${ref}`) ?? {};
    return {
      numberFormat: edit.numberFormat ?? base.numberFormat,
      format: { ...base.format, ...edit.format },
      font: { ...base.font, ...edit.font },
      fill: edit.fill ?? base.fill,
      borders: { ...base.borders, ...edit.borders },
    };
  }

  private heightOf(itemId: string, files: ZipFiles, path: string, row: number): number {
    const edited = this.rowHeightEdits.get(`${itemId}|${path}|${row}`);
    if (edited !== undefined) return edited;
    const xml = xmlText(files, path);
    const ht = new RegExp(`<row r="${row}"[^>]*?\\sht="([\\d.]+)"`).exec(xml)?.[1] ?? /defaultRowHeight="([\d.]+)"/.exec(xml)?.[1];
    return Number(ht ?? 15);
  }

  private editFormat(itemId: string, path: string, ref: string, edit: (e: FormatEdit) => void): void {
    const key = `${itemId}|${path}|${ref}`;
    const e = this.formatEdits.get(key) ?? {};
    edit(e);
    this.formatEdits.set(key, e);
  }

  private async batch(init: RequestInit, headers: Headers): Promise<Response> {
    const { requests } = JSON.parse(String(init.body)) as { requests: BatchItem[] };
    if (requests.length > 20) return err(400, "TooManyRequestsInBatch");
    const responses: { id: string; status: number; body?: unknown }[] = [];
    for (const r of requests) {
      if (r.dependsOn?.some((id) => (responses.find((x) => x.id === id)?.status ?? 0) >= 300)) {
        responses.push({ id: r.id, status: 424 });
        continue;
      }
      const res = await this.fetch(GRAPH + r.url, {
        method: r.method,
        headers: { ...r.headers, Authorization: headers.get("Authorization") ?? "" },
        body: r.body === undefined ? undefined : JSON.stringify(r.body),
      });
      const text = await res.text();
      responses.push({ id: r.id, status: res.status, body: text ? JSON.parse(text) : undefined });
    }
    return json({ responses });
  }

  private rangeFormat(method: string, itemId: string, files: ZipFiles, path: string, address: string, part: string, side: string | undefined, body: unknown): Response {
    const box = parseAddress(address);
    const refs: string[] = [];
    for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) refs.push(columnLetter(c) + r);
    const first = this.formatOf(itemId, files, path, refs[0]);
    if (method === "GET") {
      if (part === "") return json({ ...first.format, rowHeight: this.heightOf(itemId, files, path, box.r1) });
      if (part === "/font") return json(first.font);
      if (part === "/fill") return json({ color: first.fill });
      if (part === "/borders") return json({ value: Object.entries(first.borders).map(([sideIndex, b]) => ({ sideIndex, ...b })) });
    }
    if (method === "POST" && part === "/fill/clear") {
      for (const ref of refs) this.editFormat(itemId, path, ref, (e) => (e.fill = "#FFFFFF"));
      return new Response(null, { status: 200 });
    }
    if (method === "PATCH" && part === "") {
      const { rowHeight, ...format } = body as CellFormat["format"] & { rowHeight?: number };
      for (const ref of refs) this.editFormat(itemId, path, ref, (e) => (e.format = { ...e.format, ...format }));
      if (rowHeight !== undefined) for (let r = box.r1; r <= box.r2; r++) this.rowHeightEdits.set(`${itemId}|${path}|${r}`, rowHeight);
      return json({});
    }
    if (method === "PATCH" && part === "/font") {
      for (const ref of refs) this.editFormat(itemId, path, ref, (e) => (e.font = { ...e.font, ...(body as CellFormat["font"]) }));
      return json({});
    }
    if (method === "PATCH" && part === "/fill") {
      for (const ref of refs) this.editFormat(itemId, path, ref, (e) => (e.fill = (body as { color: string }).color));
      return json({});
    }
    if (method === "PATCH" && part === "/borders" && side) {
      const border = body as Border;
      for (let r = box.r1; r <= box.r2; r++) {
        for (let c = box.c1; c <= box.c2; c++) {
          const touched: string[] = [];
          if (side === "EdgeTop" && r === box.r1) touched.push("EdgeTop");
          if (side === "EdgeBottom" && r === box.r2) touched.push("EdgeBottom");
          if (side === "EdgeLeft" && c === box.c1) touched.push("EdgeLeft");
          if (side === "EdgeRight" && c === box.c2) touched.push("EdgeRight");
          if (side === "InsideVertical") touched.push(...(c < box.c2 ? ["EdgeRight"] : []), ...(c > box.c1 ? ["EdgeLeft"] : []));
          if (side === "InsideHorizontal") touched.push(...(r < box.r2 ? ["EdgeBottom"] : []), ...(r > box.r1 ? ["EdgeTop"] : []));
          for (const t of touched) this.editFormat(itemId, path, columnLetter(c) + r, (e) => (e.borders = { ...e.borders, [t]: border }));
        }
      }
      return json({});
    }
    return err(400, `fake graph: unsupported ${method} format${part}`);
  }

  fetch = async (input: RequestInfo | URL, init: RequestInit = {}): Promise<Response> => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
    const method = (init.method ?? "GET").toUpperCase();
    const headers = new Headers(init.headers);
    const path = decodeURIComponent(url.replace(GRAPH, ""));
    const item = /\/me\/drive\/items\/([^/?]+)/.exec(path)?.[1] ?? "";
    const entry: LogEntry = {
      method,
      path,
      item,
      status: 0,
      session: headers.get("workbook-session-id") ?? undefined,
      ifMatch: headers.get("If-Match") ?? undefined,
    };
    this.log.push(entry);
    this.events.push(`req:${item}:${method}`);
    if (this.hangs.some((h) => h(method, path))) {
      return new Promise<Response>((_resolve, reject) => {
        init.signal?.addEventListener("abort", () => reject(Object.assign(new Error("aborted"), { name: "AbortError" })));
      });
    }
    const res = await this.handle(method, path, item, headers, init);
    entry.status = res.status;
    return res;
  };

  private async handle(method: string, path: string, itemId: string, headers: Headers, init: RequestInit): Promise<Response> {
    if (headers.get("Authorization") !== `Bearer ${this.validToken}`) return err(401, "InvalidAuthenticationToken");
    const fault = this.faults.find((f) => f.times > 0 && f.match(method, path));
    if (fault) {
      fault.times -= 1;
      const h: Record<string, string> = {};
      if (fault.retryAfter != null) h["Retry-After"] = String(fault.retryAfter);
      return err(fault.status, `fake ${fault.status}`, h);
    }
    if (path === "/$batch" && method === "POST") return this.batch(init, headers);
    const search = /^\/me\/drive\/root\/search\(q='(.*)'\)/.exec(path);
    if (search && method === "GET") {
      const words = search[1].replace(/''/g, "'").toLowerCase().split(/[\s(),_.]+/).filter(Boolean);
      const hits = [...this.items.entries()].filter(([, it]) => words.some((w) => it.name.toLowerCase().includes(w)));
      return json({
        value: [
          ...[...new Set(hits.map(([, it]) => it.folder).filter(Boolean))].map((name) => ({ id: `folder:${name}`, name, folder: {} })),
          ...hits.map(([id, it]) => ({ id, name: it.name, file: {}, lastModifiedDateTime: it.modified })),
        ],
      });
    }
    const item = this.items.get(itemId);
    if (!item) return err(404, "itemNotFound");
    const rest = path.slice(path.indexOf(itemId) + itemId.length);

    if (rest.startsWith("?") || rest === "") {
      return json({ id: itemId, name: item.name, eTag: etag(itemId, item.version) });
    }
    if (rest === "/content") {
      if (method === "GET") return new Response(Uint8Array.from(zipXlsx(item.files)), { status: 200 });
      if (method === "PUT") {
        const ifMatch = headers.get("If-Match");
        if (ifMatch && ifMatch !== etag(itemId, item.version)) return err(412, "resourceModified");
        const body = new Uint8Array(await new Response(init.body as BodyInit).arrayBuffer());
        item.files = unzipXlsx(body);
        item.version += 1;
        return json({ id: itemId, eTag: etag(itemId, item.version) });
      }
    }
    if (rest === "/workbook/createSession" && method === "POST") {
      this.sessionNo += 1;
      return json({ id: `session-${this.sessionNo}`, persistChanges: true });
    }
    if (rest === "/workbook/closeSession") return new Response(null, { status: 204 });
    if (rest.startsWith("/workbook/worksheets?") || rest === "/workbook/worksheets") {
      return json({ value: workbookSheets(item.files).map((s) => ({ id: sheetId(s.path), name: s.name })) });
    }
    const used = /^\/workbook\/worksheets\/([^/]+)\/usedRange(\(valuesOnly=(?:true|false)\))?(\?.*)?$/.exec(rest);
    if (used && method === "GET") {
      const sheet = workbookSheets(item.files).find((s) => sheetId(s.path) === used[1]);
      if (!sheet) return err(404, "ItemNotFound");
      const cells = this.cells(item.files, sheet.path);
      let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
      for (const c of cells.values()) {
        if (c.v == null && !c.f) continue;
        const ci = columnIndex(c.col);
        r1 = Math.min(r1, c.row); r2 = Math.max(r2, c.row); c1 = Math.min(c1, ci); c2 = Math.max(c2, ci);
      }
      // 실제 Graph: valuesOnly 가 아니면 서식만 있는 칸까지 넣는다(<dimension>, 예: 실제 리스트 A1:AE65536).
      if (used[2] !== "(valuesOnly=true)") {
        const dim = /<dimension ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"/.exec(xmlText(item.files, sheet.path))?.[1];
        if (dim) {
          const d = parseAddress(dim);
          if (d.r2 > r2) r2 = d.r2;
          if (d.c2 > c2) c2 = d.c2;
          if (!Number.isFinite(r1)) { r1 = d.r1; c1 = d.c1; }
        }
      }
      if (!r2) return json({ address: `${sheet.name}!A1`, values: [[""]], formulas: [[""]] });
      // 실제 Graph 는 너무 큰 범위(수백만 칸)를 한 번에 돌려주지 못한다(시간 초과·504).
      if (!/\$select=address$/.test(used[3] ?? "") && (r2 - r1 + 1) * (c2 - c1 + 1) > FakeDrive.maxCells) {
        return err(504, "GatewayTimeout: response payload too large");
      }
      if (/\$select=address$/.test(used[3] ?? "")) {
        const q0 = sheet.name.includes(" ") || /[^A-Za-z0-9_]/.test(sheet.name) ? `'${sheet.name}'` : sheet.name;
        return json({ address: `${q0}!${columnLetter(c1)}${r1}:${columnLetter(c2)}${r2}` });
      }
      const values: (string | number)[][] = [];
      const formulas: (string | number)[][] = [];
      for (let r = r1; r <= r2; r++) {
        const vr: (string | number)[] = [];
        const fr: (string | number)[] = [];
        for (let c = c1; c <= c2; c++) {
          const cell = cells.get(columnLetter(c) + r);
          const v = graphValue(cell);
          vr.push(v);
          fr.push(cell?.f ? `=${cell.f}` : v);
        }
        values.push(vr);
        formulas.push(fr);
      }
      const q = sheet.name.includes(" ") || /[^A-Za-z0-9_]/.test(sheet.name) ? `'${sheet.name}'` : sheet.name;
      return json({ address: `${q}!${columnLetter(c1)}${r1}:${columnLetter(c2)}${r2}`, values, formulas });
    }
    // 범위 안의 값 범위(range(...)/usedRange(valuesOnly=true)): 값 또는 수식이 있는 칸만 본다(시트 usedRange 와 같은 규칙).
    const rangeUsed = /^\/workbook\/worksheets\/([^/]+)\/range\(address='([^']+)'\)\/usedRange\(valuesOnly=true\)(\?.*)?$/.exec(rest);
    if (rangeUsed && method === "GET") {
      const sheet = workbookSheets(item.files).find((s) => sheetId(s.path) === rangeUsed[1]);
      if (!sheet) return err(404, "ItemNotFound");
      const box = parseAddress(rangeUsed[2]);
      let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
      for (const c of this.cells(item.files, sheet.path).values()) {
        if (c.v == null && !c.f) continue;
        const ci = columnIndex(c.col);
        if (c.row < box.r1 || c.row > box.r2 || ci < box.c1 || ci > box.c2) continue;
        r1 = Math.min(r1, c.row); r2 = Math.max(r2, c.row); c1 = Math.min(c1, ci); c2 = Math.max(c2, ci);
      }
      const q = sheet.name.includes(" ") || /[^A-Za-z0-9_]/.test(sheet.name) ? `'${sheet.name}'` : sheet.name;
      // 실제 Graph: 빈 범위면 왼쪽 위 칸 하나를 돌려준다.
      if (!r2) return json({ address: `${q}!${columnLetter(box.c1)}${box.r1}` });
      return json({ address: `${q}!${columnLetter(c1)}${r1}:${columnLetter(c2)}${r2}` });
    }
    const format = /^\/workbook\/worksheets\/([^/]+)\/range\(address='([^']+)'\)\/format(\/font|\/fill\/clear|\/fill|\/borders)?(?:\/(\w+))?(\?.*)?$/.exec(rest);
    if (format) {
      const sheet = workbookSheets(item.files).find((s) => sheetId(s.path) === format[1]);
      if (!sheet) return err(404, "ItemNotFound");
      const body = typeof init.body === "string" ? JSON.parse(init.body) : undefined;
      if (body) this.log[this.log.length - 1].body = body;
      return this.rangeFormat(method, itemId, item.files, sheet.path, format[2], format[3] ?? "", format[4], body);
    }
    const range = /^\/workbook\/worksheets\/([^/]+)\/range\(address='([^']+)'\)(\/insert)?(\?.*)?$/.exec(rest);
    if (range) {
      const sheet = workbookSheets(item.files).find((s) => sheetId(s.path) === range[1]);
      if (!sheet) return err(404, "ItemNotFound");
      if (range[3] && method === "POST") {
        const m = /^(\d+):(\d+)$/.exec(range[2]);
        if (!m) return err(400, "bad insert");
        const from = Number(m[1]);
        const count = Number(m[2]) - from + 1;
        setXml(item.files, sheet.path, shiftRowNumbers(xmlText(item.files, sheet.path), from, count));
        return json({});
      }
      const box = parseAddress(range[2]);
      if (method === "GET") {
        const cells = this.cells(item.files, sheet.path);
        if ((box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1) > FakeDrive.maxCells) return err(504, "GatewayTimeout: response payload too large");
        const values: (string | number)[][] = [];
        const formulas: (string | number)[][] = [];
        const formulasR1C1: (string | number)[][] = [];
        const valueTypes: string[][] = [];
        for (let r = box.r1; r <= box.r2; r++) {
          const row: (string | number)[] = [];
          const frow: (string | number)[] = [];
          const rrow: (string | number)[] = [];
          const trow: string[] = [];
          for (let c = box.c1; c <= box.c2; c++) {
            const cell = cells.get(columnLetter(c) + r);
            const v = graphValue(cell);
            row.push(v);
            frow.push(cell?.f ? `=${cell.f}` : v);
            rrow.push(cell?.f ? `=${toR1C1(cell.f, r, c)}` : v);
            trow.push(v === "" ? "Empty" : typeof v === "number" ? "Double" : "String");
          }
          values.push(row);
          formulas.push(frow);
          formulasR1C1.push(rrow);
          valueTypes.push(trow);
        }
        const numberFormat = /\$select=[^&]*numberFormat/.test(range[4] ?? "")
          ? values.map((row, i) => row.map((_, j) => this.formatOf(itemId, item.files, sheet.path, columnLetter(box.c1 + j) + (box.r1 + i)).numberFormat))
          : undefined;
        return json({ address: range[2], values, formulas, formulasR1C1, valueTypes, numberFormat });
      }
      if (method === "PATCH") {
        const ifMatch = headers.get("If-Match");
        if (ifMatch && ifMatch !== etag(itemId, item.version)) return err(412, "resourceModified");
        const body = JSON.parse(String(init.body)) as {
          formulas?: (string | number | null)[][];
          formulasR1C1?: (string | number | null)[][];
          values?: (string | number | null)[][];
          numberFormat?: string[][];
        };
        this.log[this.log.length - 1].body = body;
        if (body.numberFormat) {
          const nf = body.numberFormat;
          if (nf.length !== box.r2 - box.r1 + 1 || nf.some((row) => row.length !== box.c2 - box.c1 + 1)) return err(400, "numberFormat dimensions");
          nf.forEach((row, i) => row.forEach((f, j) => this.editFormat(itemId, sheet.path, columnLetter(box.c1 + j) + (box.r1 + i), (e) => (e.numberFormat = f))));
          return json({ address: range[2] });
        }
        const grid =
          body.formulasR1C1?.map((row, i) => row.map((v, j) => (typeof v === "string" && v.startsWith("=") ? `=${toA1(v.slice(1), box.r1 + i, box.c1 + j)}` : v))) ??
          body.formulas ??
          body.values ??
          [];
        if (grid.length !== box.r2 - box.r1 + 1 || grid.some((row) => row.length !== box.c2 - box.c1 + 1)) {
          return err(400, "The number of rows or columns in the input array doesn't match the size or dimensions of the range.");
        }
        let xml = xmlText(item.files, sheet.path);
        grid.forEach((row, i) =>
          row.forEach((v, j) => {
            if (v === null) return; // Graph: null 은 바꾸지 않음
            const ref = columnLetter(box.c1 + j) + (box.r1 + i);
            if (v === "") xml = setCell(xml, ref, { kind: "empty" });
            else if (typeof v === "number") xml = setCell(xml, ref, { kind: "number", n: v });
            else if (v.startsWith("=")) xml = setCell(xml, ref, { kind: "formula", formula: v.slice(1) });
            else xml = setCell(xml, ref, { kind: "inline", text: v.startsWith("'") ? v.slice(1) : v });
          }),
        );
        setXml(item.files, sheet.path, xml);
        return json({ address: range[2] });
      }
    }
    return err(400, `fake graph: unsupported ${method} ${rest}`);
  }

  private cells(files: ZipFiles, path: string): Map<string, XCell> {
    const key = files["xl/sharedStrings.xml"];
    let shared = key ? this.sharedCache.get(key) : [];
    if (!shared) {
      shared = sharedStrings(files);
      if (key) this.sharedCache.set(key, shared);
    }
    return parseCells(xmlText(files, path), shared);
  }
}

function sheetId(path: string): string {
  return `{ws-${path.match(/sheet(\d+)\.xml$/)?.[1] ?? path}}`;
}

function etag(id: string, version: number): string {
  return `"{${id}},${version}"`;
}

function parseAddress(address: string): { c1: number; r1: number; c2: number; r2: number } {
  const [a, b] = address.split(":");
  const pa = /^([A-Z]+)(\d+)$/.exec(a)!;
  const pb = /^([A-Z]+)(\d+)$/.exec(b ?? a)!;
  return { c1: columnIndex(pa[1]), r1: Number(pa[2]), c2: columnIndex(pb[1]), r2: Number(pb[2]) };
}

/** 수식의 칸 참조를 바꾼다. 따옴표 안 글자는 건드리지 않는다. */
function mapOutsideStrings(formula: string, re: RegExp, replace: (...m: string[]) => string): string {
  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part, i) => (i % 2 ? part : part.replace(re, replace as never)))
    .join("");
}

/** A1 → R1C1(row·col 칸 기준 상대). Graph formulasR1C1 흉내. */
function toR1C1(formula: string, row: number, col: number): string {
  return mapOutsideStrings(formula, /(?<![A-Za-z0-9_$.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (_all, cAbs, letters, rAbs, digits) => {
    const r = Number(digits);
    const c = columnIndex(letters);
    const R = rAbs ? `R${r}` : r === row ? "R" : `R[${r - row}]`;
    const C = cAbs ? `C${c}` : c === col ? "C" : `C[${c - col}]`;
    return R + C;
  });
}

/** R1C1(row·col 칸 기준) → A1. */
function toA1(formula: string, row: number, col: number): string {
  return mapOutsideStrings(formula, /(?<![A-Za-z0-9_])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])/g, (_all, rPart, cPart) => {
    const part = (p: string | undefined, base: number) => (!p ? { n: base, abs: false } : p.startsWith("[") ? { n: base + Number(p.slice(1, -1)), abs: false } : { n: Number(p), abs: true });
    const r = part(rPart, row);
    const c = part(cPart, col);
    return `${c.abs ? "$" : ""}${columnLetter(c.n)}${r.abs ? "$" : ""}${r.n}`;
  });
}

function graphValue(c: XCell | undefined): string | number {
  if (!c || c.v == null) return "";
  if (c.t === "s" || c.t === "inlineStr" || c.t === "str") return c.v;
  const n = Number(c.v);
  return Number.isFinite(n) ? n : c.v;
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

function err(status: number, message: string, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify({ error: { code: String(status), message } }), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });
}

interface BatchItem {
  id: string;
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: unknown;
  dependsOn?: string[];
}

interface Border {
  style: string;
  color?: string;
  weight?: string;
}

export interface CellFormat {
  numberFormat: string;
  format: { horizontalAlignment: string; verticalAlignment: string; wrapText: boolean };
  font: { name: string; size: number; bold: boolean; italic: boolean; color: string; underline: string };
  fill: string;
  borders: Record<string, Border>;
}

type FormatEdit = Partial<Omit<CellFormat, "borders">> & { borders?: Record<string, Border> };

const BORDER_STYLE: Record<string, Border> = {
  thin: { style: "Continuous", color: "#000000", weight: "Thin" },
  medium: { style: "Continuous", color: "#000000", weight: "Medium" },
  thick: { style: "Continuous", color: "#000000", weight: "Thick" },
  hair: { style: "Continuous", color: "#000000", weight: "Hairline" },
  dotted: { style: "Dot", color: "#000000", weight: "Thin" },
  dashed: { style: "Dash", color: "#000000", weight: "Thin" },
  double: { style: "Double", color: "#000000", weight: "Thick" },
};

const capitalize = (s: string | undefined, fallback: string) => (s ? s[0].toUpperCase() + s.slice(1) : fallback);

/** 엑셀처럼 칸 s → 행 서식(customFormat) → 열 기본 서식 순으로 정해지는 서식을 Graph 값 모양으로. */
function xfFormat(files: ZipFiles, sheetXml: string, ref: string): CellFormat {
  const [, colLetters, row] = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const cellAttrs = new RegExp(`<c r="${ref}"([^>]*?)/?>`).exec(sheetXml)?.[1];
  const rowAttrs = new RegExp(`<row r="${row}"([^>]*?)/?>`).exec(sheetXml)?.[1] ?? "";
  const col = columnIndex(colLetters);
  const colStyle = [...(/<cols>([\s\S]*?)<\/cols>/.exec(sheetXml)?.[1] ?? "").matchAll(/<col\b([^>]*)\/?>/g)]
    .find((c) => col >= Number(/\bmin="(\d+)"/.exec(c[1])?.[1]) && col <= Number(/\bmax="(\d+)"/.exec(c[1])?.[1]))?.[1];
  const sOf = (attrs: string | undefined) => (attrs ? /(?:^|\s)(?:s|style)="(\d+)"/.exec(attrs)?.[1] : undefined);
  const s = Number(
    (cellAttrs !== undefined ? sOf(cellAttrs) : undefined) ??
      (cellAttrs === undefined && /\bcustomFormat="(1|true)"/.test(rowAttrs) ? sOf(rowAttrs) : undefined) ??
      (cellAttrs === undefined ? sOf(colStyle) : undefined) ??
      0,
  );
  const styles = xmlText(files, "xl/styles.xml");
  const section = (name: string, item: string) =>
    [...(new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)</${name}>`).exec(styles)?.[1] ?? "").matchAll(new RegExp(`<${item}\\b[^>]*?(?:/>|>[\\s\\S]*?</${item}>)`, "g"))].map((m) => m[0]);
  const xf = section("cellXfs", "xf")[s] ?? "";
  const id = (name: string) => Number(new RegExp(`\\b${name}="(\\d+)"`).exec(xf)?.[1] ?? 0);
  const align = /<alignment\b([^>]*)\/>/.exec(xf)?.[1] ?? "";
  const font = section("fonts", "font")[id("fontId")] ?? "";
  const fill = section("fills", "fill")[id("fillId")] ?? "";
  const border = section("borders", "border")[id("borderId")] ?? "";
  const numFmtId = id("numFmtId");
  const side = (name: string): Border => BORDER_STYLE[new RegExp(`<${name}\\b[^>]*style="(\\w+)"`).exec(border)?.[1] ?? ""] ?? { style: "None" };
  const rgb = /patternType="solid"/.test(fill) ? /<fgColor\b[^>]*rgb="(\w+)"/.exec(fill)?.[1] : undefined;
  return {
    numberFormat: numFmtId === 0 ? "General" : new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"`).exec(styles)?.[1] ?? `builtin-${numFmtId}`,
    format: {
      horizontalAlignment: capitalize(/horizontal="(\w+)"/.exec(align)?.[1], "General"),
      verticalAlignment: capitalize(/vertical="(\w+)"/.exec(align)?.[1], "Bottom"),
      wrapText: /wrapText="1"/.test(align),
    },
    font: {
      name: /<name val="([^"]*)"/.exec(font)?.[1] ?? "",
      size: Number(/<sz val="([\d.]+)"/.exec(font)?.[1] ?? 11),
      bold: /<b\/>/.test(font),
      italic: /<i\/>/.test(font),
      color: "#000000",
      underline: /<u\b/.test(font) ? "Single" : "None",
    },
    fill: rgb ? `#${rgb.slice(-6)}` : "#FFFFFF",
    borders: { EdgeTop: side("top"), EdgeBottom: side("bottom"), EdgeLeft: side("left"), EdgeRight: side("right") },
  };
}
