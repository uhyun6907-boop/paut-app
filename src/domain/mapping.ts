import type { SiteMapping } from "./types";
import { ILJI_COLUMNS } from "./types";

export interface StatementSheetRef {
  name: string;
  prefix: string;
  year: number;
  month: number;
  /** 이름에서 읽은 꼬리. 예: ' AM' */
  suffix: string;
  style: { gap: string; dot: string; monthWidth: number };
}

/**
 * `드림합정26.10`, `지성브레인2611`, `지성군포 26.4`, `지성 마전 26.5 AM`
 * 같은 시트 이름에서 연·월을 읽는다.
 */
export function matchStatementSheet(name: string, prefix: string): StatementSheetRef | null {
  if (!prefix) return null;
  if (!name.startsWith(prefix)) return null;
  const rest = name.slice(prefix.length);
  const m = rest.match(/^(\s*)(\d{2})(\.?)(\d{1,2})(.*)$/);
  if (!m) return null;
  const month = Number(m[4]);
  if (month < 1 || month > 12) return null;
  const yy = Number(m[2]);
  const year = yy >= 70 ? 1900 + yy : 2000 + yy;
  return {
    name,
    prefix,
    year,
    month,
    suffix: m[5] || "",
    style: { gap: m[1], dot: m[3], monthWidth: m[4].length },
  };
}

export function sheetSortKey(ref: StatementSheetRef): number {
  return ref.year * 12 + ref.month;
}

/** 가장 최근 시트의 표기(점 유무, 월 자릿수, 공백)를 따라 다음 이름을 만든다. AM 꼬리는 복사하지 않는다. */
export function suggestSheetName(prefix: string, latest: StatementSheetRef | null, year: number, month: number): string {
  const yy = String(year).slice(-2);
  if (!latest) return `${prefix}${yy}.${month}`;
  const mm = latest.style.monthWidth >= 2 ? String(month).padStart(2, "0") : String(month);
  return `${prefix}${latest.style.gap}${yy}${latest.style.dot}${mm}`;
}

export function findMonthSheet(names: string[], prefix: string, year: number, month: number): StatementSheetRef | null {
  const hits = names
    .map((n) => matchStatementSheet(n, prefix))
    .filter((x): x is StatementSheetRef => !!x && x.year === year && x.month === month && !/AM/i.test(x.suffix));
  if (!hits.length) return null;
  return hits.sort((a, b) => a.name.length - b.name.length)[0];
}

export function latestSheet(names: string[], prefix: string): StatementSheetRef | null {
  const hits = names
    .map((n) => matchStatementSheet(n, prefix))
    .filter((x): x is StatementSheetRef => !!x && !/AM/i.test(x.suffix));
  if (!hits.length) return null;
  return hits.sort((a, b) => sheetSortKey(a) - sheetSortKey(b))[hits.length - 1];
}

export function defaultMappings(): SiteMapping[] {
  return [
    {
      site: "드림 합정주공",
      iljiColumn: "드림 합정주공",
      statementPrefix: "드림합정",
      prefixConfirmed: true,
      unitPrice: 75000,
      priceConfirmed: true,
      priceNote: "드림합정 시트 행 단가 75,000원/M (헤더에도 75,000원/M).",
    },
    // 아래 셋은 'PAUT NDT 리스트(군포, 마전).xlsx'의 PA 시트(현장 이름 = 시트 이름).
    {
      site: "지성 군포 PA",
      iljiColumn: "지성 군포",
      statementPrefix: "지성군포",
      prefixConfirmed: true,
      unitPrice: 75000,
      priceConfirmed: true,
      priceNote: "지성군포 시트 행 단가 75,000원/M(750A 2.4M). 헤더의 압입구간 90,000원/M 은 앱이 구분하지 않으니 그런 날은 명세서에서 단가를 직접 고치세요.",
    },
    {
      site: "지성 마전 PA",
      iljiColumn: "지성 마전",
      statementPrefix: "지성 마전",
      prefixConfirmed: true,
      unitPrice: 65000,
      priceConfirmed: true,
      priceNote: "지성 마전 시트 PAUT 단가 65,000원/m(헤더·행 동일). 'AM' 시트는 쓰지 않습니다.",
    },
    {
      site: "아전 군포 PA",
      iljiColumn: "아전 군포",
      statementPrefix: "",
      prefixConfirmed: false,
      unitPrice: null,
      priceConfirmed: false,
      priceNote: "명세서는 아직 쓰지 않습니다(건너뜀). 작업일지는 '아전 군포' 열.",
    },
  ];
}

/**
 * 앱에서 지운 현장(지성 브레인, 택지 드림 = 드림 택지, 드림 화양, 브레인대지 인입). 현장 목록·선택·동기화 대상에서 뺀다.
 * 엑셀·OneDrive 의 시트와 행은 그대로 둔다(지우거나 비우지 않음). 이 기기의 기존 기록은 숨기기만 한다.
 */
const REMOVED_SITE_KEYS = new Set(["지성브레인", "택지드림", "드림택지", "드림화양", "브레인대지인입"]);

export function isRemovedSite(site: string | undefined | null): boolean {
  return !!site && REMOVED_SITE_KEYS.has(site.replace(/\s+/g, ""));
}

/** 지운 현장의 예전 설정. 같은 작업일지 열('그 외')을 쓰는 다른 현장을 다시 쓸 때 그 칸 값을 그대로 두려고 남긴다. */
export const REMOVED_SITE_MAPPINGS: readonly SiteMapping[] = [
    {
      site: "지성 브레인",
      iljiColumn: "지성 브레인연계",
      statementPrefix: "지성브레인",
      prefixConfirmed: true,
      unitPrice: 73000,
      priceConfirmed: true,
      priceNote: "지성브레인 시트 헤더에는 PAUT 단가가 없고, 행 단가가 73,000원입니다. 설정에서 바꿀 수 있습니다.",
    },
    {
      site: "택지 드림",
      iljiColumn: "그 외",
      statementPrefix: "",
      prefixConfirmed: false,
      unitPrice: null,
      priceConfirmed: false,
      priceNote: "작업일지 열이 없어 '그 외'로 둡니다. 거래명세서 시트는 파일에서 찾지 못했습니다. 접두어와 단가를 직접 입력하세요.",
    },
    {
      site: "드림 화양",
      iljiColumn: "그 외",
      statementPrefix: "드림화양",
      prefixConfirmed: true,
      unitPrice: null,
      priceConfirmed: false,
      priceNote: "작업일지는 '그 외'. 드림화양 시트는 RT 위주라 PAUT 단가를 찾지 못했습니다. 단가를 입력하기 전에는 거래명세서에 금액을 쓰지 않습니다.",
    },
    {
      site: "브레인대지 인입",
      iljiColumn: "그 외",
      statementPrefix: "",
      prefixConfirmed: false,
      unitPrice: null,
      priceConfirmed: false,
      priceNote: "작업일지 열·거래명세서 시트가 없습니다. 개소는 '그 외'에 넣습니다. 접두어는 비워 두었습니다.",
    },
];

/** 파일에서 읽은 길이가 있으면 그 값을 쓰고, 없으면 명세서에 반복된 기본 길이. */
export function defaultWeldLengths(): Record<string, number> {
  return {
    "200": 0.7,
    "250": 0.8,
    "300": 1.0,
    "350": 1.2,
    "400": 1.3,
    "450": 1.5,
    "500": 1.6,
    "550": 1.8,
    "600": 1.9,
    "750": 2.4,
  };
}

export function guessIljiColumn(site: string, columns: readonly string[] = ILJI_COLUMNS): string {
  const norm = site.replace(/\s+/g, "");
  for (const col of columns) {
    const cn = col.replace(/\s+/g, "");
    if (cn === "그외") continue;
    if (norm.includes(cn) || cn.includes(norm)) return col;
  }
  if (norm.includes("브레인") && !norm.includes("인입")) return "지성 브레인연계";
  return "그 외";
}

export function mappingFor(mappings: SiteMapping[], site: string): SiteMapping {
  return (
    mappings.find((m) => m.site === site) ??
    REMOVED_SITE_MAPPINGS.find((m) => m.site === site) ?? {
      site,
      iljiColumn: "그 외",
      statementPrefix: "",
      prefixConfirmed: false,
      unitPrice: null,
      priceConfirmed: false,
      priceNote: "새 현장입니다. 작업일지 열·거래명세서 접두어·단가를 지정하세요.",
    }
  );
}
