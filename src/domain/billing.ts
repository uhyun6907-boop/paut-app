import type { Bill, BillLine, InspectionRecord } from "./types";
import { PASS } from "./types";
import { diameterNumber } from "./size";

export function isPassing(result: string): boolean {
  return result === PASS;
}

export function passingRecords(records: InspectionRecord[]): InspectionRecord[] {
  return records.filter((r) => !r.deleted && isPassing(r.result));
}

/** 포인트 1개 = 개소 1. 불합격은 제외. 관경별 개수. */
export function countByDiameter(records: InspectionRecord[]): Record<string, number> {
  const map: Record<string, number> = {};
  for (const r of passingRecords(records)) {
    const key = r.diameter || "미상";
    map[key] = (map[key] || 0) + 1;
  }
  return map;
}

export function recordsForDay(records: InspectionRecord[], date: string): InspectionRecord[] {
  return records.filter((r) => !r.deleted && r.date === date);
}

export function recordsForMonth(records: InspectionRecord[], ym: string): InspectionRecord[] {
  return records.filter((r) => !r.deleted && r.date.startsWith(ym));
}

/**
 * 공급가액 = Σ(개소 × 용접길이m × 단가).
 * 부가세 = 공급가액 × 10%.
 * 합계 = ROUNDDOWN(공급가액+부가세, -3) — 천 원 미만 절사.
 * 길이를 모르는 관경은 금액에서 빠지고 missingLengths 에 남는다.
 */
export function buildBill(
  records: InspectionRecord[],
  unitPrice: number,
  lengths: Record<string, number>,
): Bill {
  const grouped = new Map<string, { date: string; diameter: string; count: number }>();
  for (const r of passingRecords(records)) {
    const diameter = r.diameter || "미상";
    const key = `${r.date}||${diameter}`;
    const prev = grouped.get(key);
    if (prev) prev.count += 1;
    else grouped.set(key, { date: r.date, diameter, count: 1 });
  }
  const lines: BillLine[] = [];
  const missing = new Set<string>();
  let totalM = 0;
  let supply = 0;
  for (const row of [...grouped.values()].sort((a, b) => a.date.localeCompare(b.date) || a.diameter.localeCompare(b.diameter))) {
    const a = diameterNumber(row.diameter);
    const lengthM = a != null && lengths[String(a)] != null ? lengths[String(a)] : null;
    if (lengthM == null) missing.add(row.diameter);
    const qtyM = lengthM == null ? null : round3(row.count * lengthM);
    const amount = qtyM == null ? null : qtyM * unitPrice;
    if (qtyM != null) totalM += qtyM;
    if (amount != null) supply += amount;
    lines.push({ date: row.date, diameter: row.diameter, count: row.count, lengthM, qtyM, unitPrice, amount });
  }
  const vat = supply * 0.1;
  const grand = Math.floor((supply + vat) / 1000) * 1000;
  return { lines, totalM, supply, vat, grand, missingLengths: [...missing] };
}

function round3(n: number): number {
  return Math.round(n * 1000) / 1000;
}

/** 그 외 열은 여러 현장이 공유하므로 관경에 현장 이름을 붙인다. */
export function iljiDiameterLabel(diameter: string, column: string, site: string): string {
  if (column !== "그 외") return diameter;
  const short = site.replace(/\s+/g, "");
  return `${diameter}(${short})`;
}

export function won(n: number): string {
  return `${Math.round(n).toLocaleString("ko-KR")}원`;
}
