import { PAUT_OD, SCHEDULE } from "./reference";
import { inchLabel, type OdRow } from "./size";

/** 일산스틸 이중보온관 사양. KSD3562/KSD3583, 한국지역난방공사 제품사양 기준. */
export interface JacketRow {
  a: number;
  mat: string;
  innerT: string;
  innerD: string;
  outerT: string;
  outerD: string;
  insul: string;
  kg: string;
  /** 6000 또는 6000 / 12000. 단위 mm. */
  len: string;
}

export const JACKET_ROWS: JacketRow[] = [
  { a: 20, mat: "SPPS 38S/E", innerT: "2.9", innerD: "27.2", outerT: "2.5", outerD: "90", insul: "28.90", kg: "2.81", len: "6000" },
  { a: 25, mat: "SPPS 38S/E", innerT: "3.4", innerD: "34.0", outerT: "2.5", outerD: "90", insul: "25.50", kg: "3.61", len: "6000" },
  { a: 32, mat: "SPPS 38S/E", innerT: "3.6", innerD: "42.7", outerT: "3.0", outerD: "110", insul: "30.70", kg: "5.00", len: "6000" },
  { a: 40, mat: "SPPS 38S/E", innerT: "3.7", innerD: "48.6", outerT: "3.0", outerD: "110", insul: "27.70", kg: "5.60", len: "6000" },
  { a: 50, mat: "SPPS 38S/E", innerT: "3.9", innerD: "60.5", outerT: "3.2", outerD: "125", insul: "29.10", kg: "7.26", len: "6000" },
  { a: 65, mat: "SPPS 38S/E", innerT: "4.5", innerD: "76.3", outerT: "3.2", outerD: "140", insul: "28.70", kg: "10.05", len: "6000" },
  { a: 80, mat: "SPPS 38S/E", innerT: "4.5", innerD: "89.1", outerT: "3.2", outerD: "160", insul: "32.30", kg: "11.88", len: "6000" },
  { a: 100, mat: "SPPS 38S/E", innerT: "4.9", innerD: "114.3", outerT: "3.5", outerD: "200", insul: "39.30", kg: "16.81", len: "6000" },
  { a: 125, mat: "SPPS 38S/E", innerT: "5.1", innerD: "139.8", outerT: "3.5", outerD: "225", insul: "39.10", kg: "21.03", len: "6000 / 12000" },
  { a: 150, mat: "SPPS 38S/E", innerT: "5.5", innerD: "165.2", outerT: "4.0", outerD: "250", insul: "38.40", kg: "26.58", len: "6000 / 12000" },
  { a: 200, mat: "SPPS 38E", innerT: "6.4", innerD: "216.3", outerT: "5.6", outerD: "315", insul: "43.80", kg: "41.21", len: "6000 / 12000" },
  { a: 250, mat: "SPPS 38E", innerT: "6.4", innerD: "267.4", outerT: "6.3", outerD: "400", insul: "60.00", kg: "53.60", len: "6000 / 12000" },
  { a: 300, mat: "SPPS 38E", innerT: "6.4", innerD: "318.5", outerT: "7.8", outerD: "450", insul: "58.00", kg: "65.12", len: "6000 / 12000" },
  { a: 350, mat: "SPPS 38E", innerT: "6.4", innerD: "355.6", outerT: "7.8", outerD: "500", insul: "64.40", kg: "73.48", len: "6000 / 12000" },
  { a: 400, mat: "SPPS 38E", innerT: "6.4", innerD: "406.4", outerT: "9.8", outerD: "560", insul: "67.00", kg: "87.33", len: "6000 / 12000" },
  { a: 450, mat: "SPPS 38E", innerT: "6.4", innerD: "457.2", outerT: "9.8", outerD: "630", insul: "76.60", kg: "99.72", len: "6000 / 12000" },
  { a: 500, mat: "SPPS 38E", innerT: "6.4", innerD: "508.0", outerT: "11.5", outerD: "710", insul: "89.50", kg: "116.79", len: "6000 / 12000" },
  { a: 550, mat: "SPPS 38E", innerT: "9.5", innerD: "558.8", outerT: "11.5", outerD: "710", insul: "64.10", kg: "163.23", len: "6000 / 12000" },
  { a: 600, mat: "SPPS 38E", innerT: "9.5", innerD: "609.6", outerT: "12.5", outerD: "800", insul: "82.70", kg: "184.84", len: "6000 / 12000" },
  { a: 650, mat: "SPW 400", innerT: "8.7", innerD: "660.4", outerT: "12.5", outerD: "850", insul: "82.30", kg: "186.69", len: "6000 / 12000" },
  { a: 700, mat: "SPW 400", innerT: "8.7", innerD: "711.2", outerT: "15.0", outerD: "900", insul: "79.40", kg: "206.46", len: "6000 / 12000" },
  { a: 750, mat: "SPW 400", innerT: "8.7", innerD: "762.0", outerT: "15.0", outerD: "950", insul: "79.00", kg: "220.53", len: "6000 / 12000" },
  { a: 800, mat: "SPW 400", innerT: "9.5", innerD: "812.8", outerT: "17.6", outerD: "1000", insul: "76.00", kg: "257.21", len: "6000 / 12000" },
  { a: 850, mat: "SPW 400", innerT: "10.3", innerD: "863.6", outerT: "17.6", outerD: "1100", insul: "100.60", kg: "298.46", len: "6000 / 12000" },
  { a: 900, mat: "SPW 400", innerT: "10.3", innerD: "914.4", outerT: "17.6", outerD: "1100", insul: "75.20", kg: "305.68", len: "6000 / 12000" },
  { a: 1000, mat: "SPW 400", innerT: "11.1", innerD: "1016.0", outerT: "23.6", outerD: "1200", insul: "68.40", kg: "377.28", len: "6000 / 12000" },
  { a: 1100, mat: "SPW 400", innerT: "11.1", innerD: "1117.6", outerT: "25.5", outerD: "1300", insul: "65.70", kg: "420.24", len: "6000 / 12000" },
  { a: 1200, mat: "SPW 400", innerT: "11.1", innerD: "1219.2", outerT: "26.6", outerD: "1400", insul: "63.80", kg: "461.22", len: "6000 / 12000" },
];

const EXTRA_INCH: Record<number, number> = { 1100: 44, 1200: 48 };

export function inchForNominal(a: number): number | null {
  const known = PAUT_OD.find((row) => row.a === a);
  if (known) return known.inch;
  for (const [key, row] of Object.entries(SCHEDULE)) {
    if (row.a === a) return Number(key);
  }
  return EXTRA_INCH[a] ?? null;
}

/** 외경/스캔 표의 스캔길이. 없는 규격은 null. */
export function jacketScan(a: number): number | null {
  return PAUT_OD.find((row) => row.a === a)?.scan ?? null;
}

export function jacketSizeLabel(row: JacketRow): string {
  const inch = inchForNominal(row.a);
  return inch == null ? `${row.a}A` : `${row.a}A / ${inchLabel(inch)}`;
}

export function jacketButtonLabel(row: JacketRow): string {
  const scan = jacketScan(row.a);
  return `${jacketSizeLabel(row)} · ${scan ?? "—"}`;
}

export function jacketPaut(a: number): OdRow | undefined {
  return PAUT_OD.find((row) => row.a === a);
}
