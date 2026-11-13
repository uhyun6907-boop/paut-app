export interface ParsedPoint {
  prefix: string;
  /** 기준 번호. R05-1 과 R27-4-1 은 각각 5, 27. */
  num: number;
  /** '1' 또는 '4-1'. 없으면 null. */
  suffix: string | null;
  raw: string;
}

const POINT_RE = /^([A-Z]+)(\d{1,3})((?:-\d+)*)$/;

export function parsePoint(input: string): ParsedPoint | null {
  const raw = String(input ?? "").trim().toUpperCase();
  const m = raw.match(POINT_RE);
  if (!m) return null;
  const suffix = m[3] ? m[3].slice(1) : null;
  return { prefix: m[1], num: Number(m[2]), suffix, raw };
}

/** 두 자리. S1 → S01, S12 → S12. 하위 번호는 붙이지 않는다. */
export function fmtPoint(prefix: string, num: number): string {
  const p = prefix.toUpperCase().replace(/[^A-Z]/g, "").slice(0, 3) || "S";
  return p + String(Math.max(0, Math.trunc(num))).padStart(2, "0");
}

/** S01-1. suffix 가 없으면 본번호만. */
export function fmtFullPoint(prefix: string, num: number, suffix?: string | null): string {
  const base = fmtPoint(prefix, num);
  const sub = suffix?.replace(/^-/, "").trim();
  return sub ? `${base}-${sub}` : base;
}

export function nextNumber(points: string[], prefix: string): number {
  const pfx = prefix.toUpperCase();
  let max = 0;
  for (const pt of points) {
    const parsed = parsePoint(pt);
    if (parsed && parsed.prefix === pfx) max = Math.max(max, parsed.num);
  }
  return max + 1;
}

export function nextPointLabel(points: string[], prefix: string): string {
  return fmtPoint(prefix, nextNumber(points, prefix));
}

export function canonicalPoint(input: string): string | null {
  const parsed = parsePoint(input);
  if (!parsed) return null;
  return parsed.suffix ? parsed.raw : fmtPoint(parsed.prefix, parsed.num);
}

/**
 * 재검사 번호. 샘플 검사리스트는 접미어를 바꾸지 않고
 * 불합격과 같은 포인트 번호를 다음 행에 다시 적는다.
 */
export function reinspectLabel(point: string): string | null {
  return canonicalPoint(point);
}

export type CardPhase = "draft" | "pass" | "fail";

/** 회색(미입력) → 합격 → 불합격 → 회색. */
export function cycleCard(phase: CardPhase): CardPhase {
  if (phase === "draft") return "pass";
  if (phase === "pass") return "fail";
  return "draft";
}

/** + 카드가 만들 다음 본번호. 재검사(같은 번호)는 occupied 에 넣지 않는다. */
/** S05 → R05, S05-1 → R05-1. S가 아니면 null. */
export function mirrorPoint(point: string): string | null {
  const parsed = parsePoint(point);
  if (!parsed || parsed.prefix !== "S") return null;
  return fmtFullPoint("R", parsed.num, parsed.suffix);
}

/** 같은 날짜의 초안을 다른 접두어로 옮길 계획. 목적지에 이미 있는 번호는 뺀다. */
export function planPrefixCopy(args: {
  from: string;
  to: string;
  drafts: { point: string; result: string; ts: number }[];
  existing: string[];
}): { point: string; result: "합격" | "불합격" }[] {
  const from = args.from.toUpperCase();
  const existing = new Set(args.existing.map((point) => point.toUpperCase()));
  const latest = new Map<string, { point: string; result: "합격" | "불합격"; ts: number }>();
  for (const row of args.drafts) {
    const parsed = parsePoint(row.point);
    if (!parsed || parsed.prefix !== from) continue;
    if (row.result !== "합격" && row.result !== "불합격") continue;
    const point = parsed.suffix ? parsed.raw : fmtPoint(parsed.prefix, parsed.num);
    const prev = latest.get(point);
    if (!prev || row.ts >= prev.ts) latest.set(point, { point, result: row.result, ts: row.ts });
  }
  const planned: { point: string; result: "합격" | "불합격" }[] = [];
  for (const row of [...latest.values()].sort((a, b) => a.ts - b.ts)) {
    const dest = fmtFullPoint(args.to, parsePoint(row.point)!.num, parsePoint(row.point)!.suffix);
    if (existing.has(dest.toUpperCase())) continue;
    existing.add(dest.toUpperCase());
    planned.push({ point: dest, result: row.result });
  }
  return planned;
}

export function nextDraftPoint(args: {
  direction: WalkDirection;
  prefix: string;
  occupied: number[];
  lastEntered: number | null;
}): string {
  return fmtPoint(args.prefix, suggestNumber({
    direction: args.direction,
    recorded: args.occupied,
    absent: [],
    lastEntered: args.lastEntered,
  }));
}

export type WalkDirection = "up" | "down";

export interface OrderedCard {
  point: string;
  /** 같은 번호의 재검사 카드. 원본 바로 뒤에 둔다. */
  re?: boolean;
}

const PREFIX_RANK: Record<string, number> = { S: 0, R: 1, P: 2 };

/** 번호 순. 방향은 본번호만 뒤집고, 하위 번호와 재검사는 원본 바로 뒤에 고정한다. */
export function orderPointCards<T extends OrderedCard>(cards: T[], direction: WalkDirection = "up"): T[] {
  return cards
    .map((card, index) => ({ card, index }))
    .sort((a, b) => compareOrderedCard(a.card, b.card, direction) || a.index - b.index)
    .map((row) => row.card);
}

function compareOrderedCard(a: OrderedCard, b: OrderedCard, direction: WalkDirection): number {
  const pa = parsePoint(a.point);
  const pb = parsePoint(b.point);
  if (!pa || !pb) return a.point.localeCompare(b.point);
  const prefixOrder = (PREFIX_RANK[pa.prefix] ?? 9) - (PREFIX_RANK[pb.prefix] ?? 9);
  if (prefixOrder) return prefixOrder;
  if (pa.prefix !== pb.prefix) return pa.prefix.localeCompare(pb.prefix);
  const dir = direction === "down" ? -1 : 1;
  if (pa.num !== pb.num) return (pa.num - pb.num) * dir;
  const as = pa.suffix ?? "";
  const bs = pb.suffix ?? "";
  if (as !== bs) {
    if (!as) return -1;
    if (!bs) return 1;
    return as.localeCompare(bs, undefined, { numeric: true });
  }
  if (!!a.re !== !!b.re) return a.re ? 1 : -1;
  return 0;
}

/** 포인트 번호 순(오름차순). S 전부 → R 전부, 각 접두어 안에서 숫자 순(S2 < S10), 하위 번호는 원본 뒤. */
export function comparePointLabel(a: string, b: string): number {
  return compareOrderedCard({ point: a }, { point: b }, "up");
}

/** 최근 두 본번호가 내려갔으면 down. 범위 표시 순서와는 별개다. */
export function directionFromTrail(trail: number[]): WalkDirection | null {
  if (trail.length < 2) return null;
  const prev = trail[trail.length - 2];
  const last = trail[trail.length - 1];
  if (last < prev) return "down";
  if (last > prev) return "up";
  return null;
}

export function suggestNumber(args: {
  direction: WalkDirection;
  recorded: number[];
  absent: number[];
  lastEntered: number | null;
}): number {
  const recorded = new Set(args.recorded);
  const absent = new Set(args.absent);
  const blocked = (n: number) => recorded.has(n) || absent.has(n);
  if (args.direction === "down") {
    const anchor = args.lastEntered ?? Math.max(0, ...args.recorded, ...args.absent);
    let n = anchor - 1;
    while (n > 1 && blocked(n)) n -= 1;
    return Math.max(1, n);
  }
  const max = Math.max(0, ...args.recorded, ...args.absent, args.lastEntered ?? 0);
  let n = max + 1;
  while (blocked(n)) n += 1;
  return n;
}

/** +1/+2. 오름은 위로, 내림은 아래로. 이미 있거나 없음인 번호는 건너뛴다. */
export function stepNumber(from: number, direction: WalkDirection, steps: number, recorded: number[], absent: number[]): number {
  const blocked = new Set([...recorded, ...absent]);
  let n = from;
  let left = steps;
  const delta = direction === "up" ? 1 : -1;
  while (left > 0) {
    const next = n + delta;
    if (next < 1) return n;
    n = next;
    if (blocked.has(n)) continue;
    left -= 1;
  }
  return n;
}

/**
 * 저장 후 같은 접두어의 다음 번호.
 * R05-1 을 저장해도 다음은 R06 (기준 번호 + 1).
 * S에서 R로의 전환은 접두어 버튼으로 한다. R을 누르면 그 라인의 다음 R 번호(없으면 R01).
 */
export function advanceAfterSave(prefix: string, saved: ParsedPoint): { prefix: string; num: number } {
  return { prefix: saved.prefix || prefix, num: saved.num + 1 };
}

export interface RangePoint {
  point: string;
  result: "합격" | "불합격";
}

/** S05~S10, 불합격 목록. 접두어가 다르거나 역순이면 null. */
export function expandRange(fromRaw: string, toRaw: string, failList: string[]): RangePoint[] | null {
  const from = parsePoint(fromRaw);
  const to = parsePoint(toRaw);
  if (!from || !to) return null;
  if (from.prefix !== to.prefix) return null;
  if (to.num < from.num) return null;
  if (to.num - from.num > 200) return null;
  const fails = new Set(
    failList
      .map((s) => s.trim().toUpperCase())
      .filter(Boolean)
      .map((s) => {
        const p = parsePoint(s);
        return p ? fmtPoint(p.prefix, p.num) : s;
      }),
  );
  const out: RangePoint[] = [];
  for (let n = from.num; n <= to.num; n++) {
    const point = fmtPoint(from.prefix, n);
    out.push({ point, result: fails.has(point) ? "불합격" : "합격" });
  }
  return out;
}

export interface PointCell {
  num: number;
  label: string;
  status: "pass" | "fail" | "skip";
  subs: string[];
}

/** 기록된 본번호만. 빠진 번호는 빈 칸으로 만들지 않는다. 없음은 건너뜀. */
export function pointCells(records: { point: string; result: string }[], prefix: string): PointCell[] {
  const pfx = prefix.toUpperCase();
  const parsed = records
    .map((rec) => ({ rec, parsed: parsePoint(rec.point) }))
    .filter((row): row is { rec: { point: string; result: string }; parsed: ParsedPoint } => !!row.parsed && row.parsed.prefix === pfx);
  const nums = [...new Set(parsed.map((row) => row.parsed.num))].sort((a, b) => a - b);
  return nums.map((num) => {
    const here = parsed.filter((row) => row.parsed.num === num);
    const mains = here.filter((row) => !row.parsed.suffix);
    const lastMain = mains[mains.length - 1];
    const subs = [...new Set(here.filter((row) => row.parsed.suffix).map((row) => row.parsed.suffix as string))].sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true }),
    );
    const status: PointCell["status"] =
      !lastMain || lastMain.rec.result === "없음" ? "skip" : lastMain.rec.result === "불합격" ? "fail" : "pass";
    return { num, label: fmtPoint(pfx, num), status: lastMain ? status : "skip", subs };
  }).filter((cell) => cell.status !== "skip" || cell.subs.length > 0 || parsed.some((row) => row.parsed.num === cell.num && row.rec.result === "없음"));
}

export function splitFailList(text: string): string[] {
  return text
    .toUpperCase()
    .split(/[,\s]+/)
    .map((s) => s.trim())
    .filter(Boolean);
}
