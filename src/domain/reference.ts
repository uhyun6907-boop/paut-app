import type { OdRow } from "./size";

export const PAUT_OD: OdRow[] = [
  { inch: 2, a: 50, od: 60.5, circ: 190, scan: 220, rs: 1 },
  { inch: 2.5, a: 65, od: 76.3, circ: 240, scan: 270, rs: 1 },
  { inch: 3, a: 80, od: 89.1, circ: 280, scan: 310, rs: 2 },
  { inch: 4, a: 100, od: 114.3, circ: 359, scan: 390, rs: 2 },
  { inch: 5, a: 125, od: 139.8, circ: 439, scan: 470, rs: 4 },
  { inch: 6, a: 150, od: 165.2, circ: 519, scan: 550, rs: 5 },
  { inch: 8, a: 200, od: 216.3, circ: 680, scan: 710, rs: 8 },
  { inch: 10, a: 250, od: 267.4, circ: 840, scan: 870, rs: 10, ds: 3, dl: 2 },
  { inch: 12, a: 300, od: 318.5, circ: 1001, scan: 1030, rs: 10, rl: 1, ds: 3, dl: 3 },
  { inch: 14, a: 350, od: 355.6, circ: 1117, scan: 1150, rs: 2, rl: 4, ds: 2, dl: 4 },
  { inch: 16, a: 400, od: 406.4, circ: 1277, scan: 1310, rs: 2, rl: 5, ds: 1, dl: 5 },
  { inch: 18, a: 450, od: 457.2, circ: 1436, scan: 1470, rs: 4, rl: 5, ds: 4, dl: 5 },
  { inch: 20, a: 500, od: 508, circ: 1596, scan: 1650, rs: 4, rl: 6 },
  { inch: 22, a: 550, od: 558.8, circ: 1756, scan: 1810, rs: 7, rl: 6 },
  { inch: 24, a: 600, od: 609.6, circ: 1915, scan: 1970, rs: 9, rl: 6, ds: 2, dl: 8 },
  { inch: 26, a: 650, od: 660.4, circ: 2075, scan: 2130, rs: 2, rl: 9 },
  { inch: 28, a: 700, od: 711.2, circ: 2234, scan: 2300, rs: 2, rl: 10, ds: 1, dl: 10 },
  { inch: 30, a: 750, od: 762, circ: 2394, scan: 2450, rs: 2, rl: 11, ds: 2, dl: 10 },
  { inch: 32, a: 800, od: 812.8, circ: 2553, scan: 2600, rs: 4, rl: 11, ds: 3, dl: 11 },
  { inch: 34, a: 850, od: 863.6, circ: 2713, scan: 2750, rs: 4, rl: 12, ds: 3, dl: 12 },
  { inch: 36, a: 900, od: 914.4, circ: 2873, scan: 2930, rs: 6, rl: 12 },
  { inch: 40, a: 1000, od: 1016, circ: 3192, scan: 3250 },
];

export interface ThicknessRow {
  t: number;
  probe: string;
  elq: number;
  f1: number;
  f2: string | number;
  bead: number;
  idx: number;
}

export const PAUT_T: ThicknessRow[] = [
  { t: 3, probe: "7.5L-CCEV", elq: 16, f1: 5, f2: "-", bead: 12, idx: 6 },
  { t: 4, probe: "7.5L-CCEV", elq: 16, f1: 6, f2: "-", bead: 12, idx: 6 },
  { t: 5, probe: "7.5L-CCEV", elq: 16, f1: 8, f2: "-", bead: 12, idx: 7 },
  { t: 6.4, probe: "5L-64", elq: 16, f1: 10, f2: 16, bead: 12, idx: 7 },
  { t: 7, probe: "5L-64", elq: 16, f1: 11, f2: 18, bead: 11, idx: 8 },
  { t: 8.1, probe: "5L-64", elq: 16, f1: 12, f2: 20, bead: 12, idx: 8 },
  { t: 9.5, probe: "5L-64", elq: 32, f1: 14, f2: 24, bead: 13, idx: 8 },
  { t: 10, probe: "5L-64", elq: 32, f1: 15, f2: 25, bead: 14, idx: 8 },
  { t: 11, probe: "5L-64", elq: 32, f1: 17, f2: 28, bead: 15, idx: 9 },
  { t: 12, probe: "5L-64", elq: 32, f1: 18, f2: 30, bead: 16, idx: 9 },
  { t: 14, probe: "5L-64", elq: 32, f1: 21, f2: 35, bead: 19, idx: 11 },
  { t: 16, probe: "5L-64", elq: 32, f1: 24, f2: 40, bead: 21, idx: 12 },
  { t: 18, probe: "5L-64", elq: 32, f1: 27, f2: 45, bead: 23, idx: 13 },
  { t: 20, probe: "5L-64", elq: 32, f1: 30, f2: 50, bead: 26, idx: 14 },
  { t: 22, probe: "5L-64", elq: 32, f1: 33, f2: 55, bead: 28, idx: 15 },
  { t: 24, probe: "5L-64", elq: 32, f1: 36, f2: 60, bead: 30, idx: 16 },
  { t: 26, probe: "5L-64", elq: 32, f1: 39, f2: 65, bead: 33, idx: 18 },
  { t: 28, probe: "5L-64", elq: 32, f1: 42, f2: 70, bead: 35, idx: 19 },
  { t: 30, probe: "5L-64", elq: 32, f1: 45, f2: 75, bead: 37, idx: 20 },
  { t: 32, probe: "5L-64", elq: 32, f1: 48, f2: 80, bead: 39, idx: 21 },
];

type SchMap = Record<number, number>;
export interface ScheduleRow {
  a: number;
  ansi: number;
  jis: number | null;
  cs: SchMap;
  ss: SchMap;
}

export const SCHEDULE: Record<string, ScheduleRow> = {
  "0.375": { a: 10, ansi: 17.1, jis: 17.3, cs: { 40: 2.3, 80: 3.2 }, ss: { 10: 1.7, 20: 2.0, 40: 2.3, 80: 3.2 } },
  "0.5": { a: 15, ansi: 21.3, jis: 21.7, cs: { 40: 2.8, 80: 3.7, 160: 4.7 }, ss: { 10: 2.1, 20: 2.5, 40: 2.8, 80: 3.7 } },
  "0.75": { a: 20, ansi: 26.7, jis: 27.2, cs: { 40: 2.9, 80: 3.9, 160: 5.5 }, ss: { 10: 2.1, 20: 2.5, 40: 2.9, 80: 3.9 } },
  "1": { a: 25, ansi: 33.4, jis: 34.0, cs: { 40: 3.4, 80: 4.5, 160: 6.4 }, ss: { 10: 2.8, 20: 3.0, 40: 3.4, 80: 4.6 } },
  "1.25": { a: 32, ansi: 42.2, jis: 42.7, cs: { 40: 3.6, 80: 4.9, 160: 6.4 }, ss: { 10: 2.8, 20: 3.0, 40: 3.6, 80: 4.9 } },
  "1.5": { a: 40, ansi: 48.3, jis: 48.6, cs: { 40: 3.7, 80: 5.1, 160: 7.1 }, ss: { 10: 2.8, 20: 3.0, 40: 3.7, 80: 5.1 } },
  "2": { a: 50, ansi: 60.3, jis: 60.5, cs: { 40: 3.9, 80: 5.5, 160: 8.7 }, ss: { 10: 2.8, 20: 3.5, 40: 3.9, 80: 5.5 } },
  "2.5": { a: 65, ansi: 73.0, jis: 76.3, cs: { 40: 5.2, 80: 7.0, 160: 9.5 }, ss: { 10: 3.0, 20: 3.5, 40: 5.2, 80: 7.0 } },
  "3": { a: 80, ansi: 88.9, jis: 89.1, cs: { 40: 5.5, 80: 7.6, 160: 11.1 }, ss: { 10: 3.0, 20: 4.0, 40: 5.5, 80: 7.6 } },
  "3.5": { a: 90, ansi: 101.6, jis: 101.6, cs: { 40: 5.7, 80: 8.1 }, ss: { 10: 3.0, 20: 4.0, 40: 5.8, 80: 8.1 } },
  "4": { a: 100, ansi: 114.3, jis: 114.3, cs: { 40: 6.0, 80: 8.6, 120: 11.1, 160: 13.5 }, ss: { 10: 3.0, 20: 4.0, 40: 6.0, 80: 8.6 } },
  "5": { a: 125, ansi: 141.3, jis: 139.8, cs: { 40: 6.6, 80: 9.5, 120: 12.7, 160: 15.9 }, ss: { 10: 3.4, 20: 5.0, 40: 6.6, 80: 9.5 } },
  "6": { a: 150, ansi: 168.3, jis: 165.2, cs: { 40: 7.1, 80: 11.0, 120: 14.3, 160: 18.2 }, ss: { 10: 3.4, 20: 5.6, 40: 7.1, 80: 11.0 } },
  "8": { a: 200, ansi: 219.1, jis: 216.3, cs: { 20: 6.4, 30: 7.0, 40: 8.2, 60: 10.3, 80: 12.7, 100: 15.1, 120: 18.2, 160: 23.0 }, ss: { 10: 4.0, 20: 6.5, 40: 8.2, 80: 12.7 } },
  "10": { a: 250, ansi: 273.1, jis: 267.4, cs: { 20: 6.4, 30: 7.8, 40: 9.3, 60: 12.7, 80: 15.1, 100: 18.2, 120: 21.4, 160: 28.6 }, ss: { 10: 4.0, 20: 6.5, 40: 9.3, 80: 12.7 } },
  "12": { a: 300, ansi: 323.9, jis: 318.5, cs: { 20: 6.4, 30: 8.4, 40: 10.3, 60: 14.3, 80: 17.4, 100: 21.4, 120: 25.4, 160: 33.3 }, ss: { 10: 4.5, 20: 6.5, 40: 9.5, 80: 12.7 } },
  "14": { a: 350, ansi: 355.6, jis: 355.6, cs: { 10: 6.4, 20: 7.9, 30: 9.5, 40: 11.1, 60: 15.1, 80: 19.0, 100: 23.8, 120: 27.8, 160: 35.7 }, ss: {} },
  "16": { a: 400, ansi: 406.4, jis: 406.4, cs: { 10: 6.4, 20: 7.9, 30: 9.5, 40: 12.7, 60: 16.7, 80: 21.4, 100: 26.2, 120: 30.9, 160: 40.5 }, ss: {} },
  "18": { a: 450, ansi: 457.2, jis: 457.2, cs: { 10: 6.4, 20: 7.9, 30: 11.1, 40: 14.3, 60: 19.1, 80: 23.8, 100: 29.4, 120: 34.9, 160: 45.2 }, ss: {} },
  "20": { a: 500, ansi: 508.0, jis: 508.0, cs: { 10: 6.4, 20: 9.5, 30: 12.7, 40: 15.1, 60: 20.6, 80: 26.2, 100: 32.5, 120: 38.1, 160: 50.0 }, ss: {} },
  "22": { a: 550, ansi: 558.8, jis: 558.8, cs: { 10: 6.4, 20: 9.5, 40: 15.9, 80: 28.6, 160: 54.0 }, ss: {} },
  "24": { a: 600, ansi: 609.6, jis: 609.6, cs: { 10: 6.4, 20: 9.5, 30: 14.3, 40: 17.5, 60: 24.6, 80: 31.0, 100: 38.9, 120: 46.0, 160: 59.5 }, ss: {} },
  "26": { a: 650, ansi: 660.4, jis: 660.4, cs: { 10: 6.4, 20: 9.5, 40: 18.9, 80: 34.0, 160: 64.2 }, ss: {} },
  "30": { a: 750, ansi: 762.0, jis: null, cs: { 10: 7.9, 20: 12.7, 30: 15.9 }, ss: {} },
};

export const SCH_ORDER = [10, 20, 30, 40, 60, 80, 100, 120, 160];

export function nearestThickness(thk: number): ThicknessRow {
  let best = PAUT_T[0];
  let d = Math.abs(best.t - thk);
  for (const row of PAUT_T) {
    const x = Math.abs(row.t - thk);
    if (x < d) {
      best = row;
      d = x;
    }
  }
  return best;
}

export function inchKey(inch: number): string {
  return String(inch);
}

export interface SizeChoice {
  inch: number;
  a: number;
  hasScan: boolean;
  hasSch: boolean;
}

export function sizeChoices(): SizeChoice[] {
  const map = new Map<number, SizeChoice>();
  for (const [k, s] of Object.entries(SCHEDULE)) {
    const inch = Number(k);
    map.set(inch, { inch, a: s.a, hasScan: false, hasSch: true });
  }
  for (const p of PAUT_OD) {
    const prev = map.get(p.inch);
    if (!prev) map.set(p.inch, { inch: p.inch, a: p.a, hasScan: true, hasSch: false });
    else prev.hasScan = true;
  }
  return [...map.values()].sort((a, b) => a.inch - b.inch);
}
