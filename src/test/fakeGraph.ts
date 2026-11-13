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

interface Item {
  name: string;
  files: ZipFiles;
  version: number;
  /** .xls 처럼 zip 이 아닌 파일은 바이트 그대로 둔다. */
  raw?: Uint8Array;
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
  private sharedCache = new WeakMap<Uint8Array, string[]>();

  put(id: string, name: string, bytes: Uint8Array): void {
    this.items.set(id, { name, files: unzipXlsx(bytes), version: 1 });
  }

  /** zip 이 아닌 파일(.xls)을 바이트 그대로 올린다. */
  putRaw(id: string, name: string, bytes: Uint8Array): void {
    this.items.set(id, { name, files: {}, version: 1, raw: bytes.slice() });
  }

  raw(id: string): Uint8Array {
    const item = this.items.get(id);
    if (!item?.raw) throw new Error(`no raw item ${id}`);
    return item.raw;
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
    if (/^\/me\/drive\/root\/children(\?|$)/.test(path) && method === "GET") {
      return json({ value: [...this.items.entries()].map(([id, it]) => ({ id, name: it.name, file: {} })) });
    }
    const item = this.items.get(itemId);
    if (!item) return err(404, "itemNotFound");
    const rest = path.slice(path.indexOf(itemId) + itemId.length);

    if (rest.startsWith("?") || rest === "") {
      return json({ id: itemId, name: item.name, eTag: etag(itemId, item.version) });
    }
    if (rest === "/content") {
      if (method === "GET") return new Response(Uint8Array.from(item.raw ?? zipXlsx(item.files)), { status: 200 });
      if (method === "PUT") {
        const ifMatch = headers.get("If-Match");
        if (ifMatch && ifMatch !== etag(itemId, item.version)) return err(412, "resourceModified");
        const body = new Uint8Array(await new Response(init.body as BodyInit).arrayBuffer());
        if (item.raw) item.raw = body;
        else item.files = unzipXlsx(body);
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
        for (let r = box.r1; r <= box.r2; r++) {
          const row: (string | number)[] = [];
          const frow: (string | number)[] = [];
          for (let c = box.c1; c <= box.c2; c++) {
            const cell = cells.get(columnLetter(c) + r);
            const v = graphValue(cell);
            row.push(v);
            frow.push(cell?.f ? `=${cell.f}` : v);
          }
          values.push(row);
          formulas.push(frow);
        }
        return json({ address: range[2], values, formulas });
      }
      if (method === "PATCH") {
        const ifMatch = headers.get("If-Match");
        if (ifMatch && ifMatch !== etag(itemId, item.version)) return err(412, "resourceModified");
        const body = JSON.parse(String(init.body)) as { formulas?: (string | number | null)[][]; values?: (string | number | null)[][] };
        this.log[this.log.length - 1].body = body;
        const grid = body.formulas ?? body.values ?? [];
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
