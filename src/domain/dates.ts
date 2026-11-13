/** Excel 1900 date system. Epoch 1899-12-30 absorbs the fictitious 1900-02-29 leap day. */
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

export function excelSerialToISO(serial: number): string {
  const days = Math.round(serial);
  const ms = EXCEL_EPOCH + days * 86400000;
  const d = new Date(ms);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function isoToExcelSerial(iso: string): number {
  const [y, m, d] = iso.slice(0, 10).split("-").map(Number);
  const ms = Date.UTC(y, m - 1, d);
  return Math.round((ms - EXCEL_EPOCH) / 86400000);
}

export function excelTimeToHHMM(fraction: number): string {
  const mins = Math.round(fraction * 1440) % 1440;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

export function hhmmToExcelTime(hhmm: string): number {
  const [h, m] = hhmm.split(":").map(Number);
  return ((h || 0) * 60 + (m || 0)) / 1440;
}

export function todayISO(d = new Date()): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function parseFlexibleDate(raw: string | undefined): string {
  if (!raw) return "";
  const s = raw.trim();
  if (/^\d+(\.\d+)?$/.test(s)) return excelSerialToISO(Number(s));
  const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  return s.slice(0, 10);
}
