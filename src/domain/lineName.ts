/**
 * 새 라인 이름을 쉽게: 기존 라인에서 '바뀌는 번호'를 찾아 다음 번호를 제안한다.
 * 관경 숫자(750A, 350A 처럼 뒤에 A)는 번호로 보지 않는다.
 *   군포산본 이원화 750A-5 → 군포산본 이원화 750A-6
 *   HJ-M-005-350A          → HJ-M-006-350A (자릿수 유지)
 * 같은 모양(번호 자리만 다름)의 라인끼리 묶어 가장 큰 번호 + 1, 이미 있으면 그다음.
 */

interface Shape {
  head: string;
  num: number;
  width: number;
  tail: string;
}

/** 라인의 '번호' 자리: 뒤에 A 가 붙지 않은 마지막 숫자 덩어리. */
export function lineShape(line: string): Shape | null {
  const runs = [...line.matchAll(/\d+/g)].filter((m) => !/^[Aa]/.test(line.slice(m.index! + m[0].length)));
  const last = runs.at(-1);
  if (!last) return null;
  const at = last.index!;
  return { head: line.slice(0, at), num: Number(last[0]), width: last[0].length, tail: line.slice(at + last[0].length) };
}

const key = (s: string) => s.replace(/\s+/g, "").toUpperCase();

/**
 * 제안할 새 라인(최근에 쓴 묶음부터, 최대 limit 개). lines 는 최근 순서.
 * 돌려주는 값: { line: 제안 이름, from: 바탕이 된 최근 라인 }
 */
export function suggestNextLines(lines: string[], limit = 4): { line: string; from: string }[] {
  const existing = new Set(lines.map(key));
  const families = new Map<string, { head: string; tail: string; width: number; max: number; from: string }>();
  for (const line of lines) {
    const s = lineShape(line);
    if (!s) continue;
    const fam = `${key(s.head)}#${key(s.tail)}`;
    const f = families.get(fam);
    if (!f) families.set(fam, { head: s.head, tail: s.tail, width: s.width, max: s.num, from: line });
    else if (s.num > f.max) f.max = s.num;
  }
  const out: { line: string; from: string }[] = [];
  for (const f of families.values()) {
    let n = f.max + 1;
    let next = `${f.head}${String(n).padStart(f.width, "0")}${f.tail}`;
    while (existing.has(key(next)) && n < f.max + 100) next = `${f.head}${String(++n).padStart(f.width, "0")}${f.tail}`;
    out.push({ line: next, from: f.from });
    if (out.length >= limit) break;
  }
  return out;
}
