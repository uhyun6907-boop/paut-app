/**
 * 명세서 탭: 한 달 동안 한 작업을 현장별로 모아 얼마치인지 보여 준다(앱 기록만 읽음 — 엑셀에 쓰지 않음).
 * - 개소 = 합격 포인트 수(불합격은 따로 센다).
 * - 금액: 보통 현장은 개소 × 용접길이(m) × 단가(원/m) — 거래명세서와 같은 계산(buildBill).
 *   아전 군포는 01 명세서처럼 개소당 단가(750A 183,600원 = 75,000원 × 2.448m)로 센다.
 * - 현장별 합계 = 공급가액 + 부가세 10%, 천 원 미만 버림(명세서 합계와 같음). 달 합계는 현장별 합계를 더한 것.
 */
import { buildBill, isPassing } from "./billing";
import { defaultMappings, isRemovedSite, mappingFor } from "./mapping";
import { diameterNumber } from "./size";
import type { AppConfig, InspectionRecord } from "./types";

/** 01_PAUT 거래명세서 (아전, 승재) 아전 군포 시트의 개소당 단가(관경 A → 원). */
export const AJEON_POINT_PRICE: Record<string, number> = { "750": 183600 };

export interface MonthLine {
  date: string;
  diameter: string;
  count: number;
  /** 수량(m). 개소당 단가 현장은 null */
  qtyM: number | null;
  amount: number | null;
}

export interface SiteMonth {
  site: string;
  /** 합격 개소 */
  count: number;
  /** 불합격 포인트 수 */
  fail: number;
  totalM: number;
  supply: number;
  vat: number;
  /** 공급가액 + 부가세, 천 원 미만 버림 */
  grand: number;
  lines: MonthLine[];
  /** 단가 설명(예: 75,000원/m, 개소당 183,600원) */
  priceText: string;
  /** 금액을 셀 수 없는 관경(길이·단가 모름) */
  missing: string[];
  /** 이 현장은 단가가 없어 금액을 세지 않았다 */
  unpriced: boolean;
}

export interface MonthSummary {
  month: string;
  sites: SiteMonth[];
  count: number;
  fail: number;
  supply: number;
  grand: number;
}

function isAjeon(site: string): boolean {
  return site.replace(/\s+/g, "").replace(/PA$/i, "") === "아전군포";
}

/** 기록이 있는 달(최근 먼저) + 이번 달. */
export function monthsOf(records: InspectionRecord[], today: string): string[] {
  const set = new Set<string>([today.slice(0, 7)]);
  for (const r of records) if (!r.deleted && /^\d{4}-\d{2}/.test(r.date) && !isRemovedSite(r.site)) set.add(r.date.slice(0, 7));
  return [...set].sort().reverse();
}

export function monthSummary(records: InspectionRecord[], config: AppConfig, ym: string): MonthSummary {
  const bySite = new Map<string, InspectionRecord[]>();
  for (const r of records) {
    if (r.deleted || !r.date.startsWith(ym) || isRemovedSite(r.site) || r.result === "없음") continue;
    bySite.set(r.site, [...(bySite.get(r.site) ?? []), r]);
  }
  const sites: SiteMonth[] = [];
  for (const [site, recs] of bySite) {
    const passing = recs.filter((r) => isPassing(r.result));
    const fail = recs.length - passing.length;
    if (isAjeon(site)) {
      const grouped = new Map<string, MonthLine>();
      const missing = new Set<string>();
      for (const r of passing) {
        const key = `${r.date}|${r.diameter || "미상"}`;
        const line = grouped.get(key) ?? { date: r.date, diameter: r.diameter || "미상", count: 0, qtyM: null, amount: 0 };
        line.count += 1;
        grouped.set(key, line);
      }
      let supply = 0;
      const lines = [...grouped.values()].sort((a, b) => a.date.localeCompare(b.date) || a.diameter.localeCompare(b.diameter));
      for (const line of lines) {
        const price = AJEON_POINT_PRICE[String(diameterNumber(line.diameter) ?? "")];
        if (price == null) {
          missing.add(line.diameter);
          line.amount = null;
          continue;
        }
        line.amount = line.count * price;
        supply += line.amount;
      }
      const vat = supply * 0.1;
      sites.push({
        site,
        count: passing.length,
        fail,
        totalM: 0,
        supply,
        vat,
        grand: Math.floor((supply + vat) / 1000) * 1000,
        lines,
        priceText: "개소당 183,600원(750A)",
        missing: [...missing],
        unpriced: false,
      });
      continue;
    }
    const mapping = mappingFor(config.mappings, site);
    const unitPrice = mapping.unitPrice ?? defaultMappings().find((m) => m.site === site)?.unitPrice ?? null;
    if (unitPrice == null) {
      const bill = buildBill(recs, 0, config.weldLengths);
      sites.push({
        site,
        count: passing.length,
        fail,
        totalM: bill.totalM,
        supply: 0,
        vat: 0,
        grand: 0,
        lines: bill.lines.map((l) => ({ date: l.date, diameter: l.diameter, count: l.count, qtyM: l.qtyM, amount: null })),
        priceText: "단가 없음",
        missing: [],
        unpriced: true,
      });
      continue;
    }
    const bill = buildBill(recs, unitPrice, config.weldLengths);
    sites.push({
      site,
      count: passing.length,
      fail,
      totalM: bill.totalM,
      supply: bill.supply,
      vat: bill.vat,
      grand: bill.grand,
      lines: bill.lines.map((l) => ({ date: l.date, diameter: l.diameter, count: l.count, qtyM: l.qtyM, amount: l.amount })),
      priceText: `${unitPrice.toLocaleString("ko-KR")}원/m`,
      missing: bill.missingLengths,
      unpriced: false,
    });
  }
  sites.sort((a, b) => b.grand - a.grand || a.site.localeCompare(b.site, "ko"));
  return {
    month: ym,
    sites,
    count: sites.reduce((n, s) => n + s.count, 0),
    fail: sites.reduce((n, s) => n + s.fail, 0),
    supply: sites.reduce((n, s) => n + s.supply, 0),
    grand: sites.reduce((n, s) => n + s.grand, 0),
  };
}
