import ExcelJS from "exceljs";

// Target weekly-sheet column -> source "Form responses" column (1-based letters)
// Chiller uses the NEW Chiller WH1 column, falling back to Chiller WH2.
export const COLUMN_MAP: { target: string; source: string[]; label: string }[] = [
  { target: "C", source: ["E"], label: "Temperature (Ambient Shelves) Range 26C_30C" },
  { target: "D", source: ["F"], label: "Humidity Reading (Ambient Shelves)" },
  { target: "E", source: ["H"], label: "Temperature (AC1) Range 15-25" },
  { target: "F", source: ["I"], label: "Humidity Reading (AC1)" },
  { target: "G", source: ["J"], label: "AC2 (Range 15C_25C)" },
  { target: "H", source: ["K"], label: "Humidity Reading (AC2)" },
  { target: "I", source: ["L"], label: "\u00a0Ultralow Freezer (Range -80" },
  { target: "J", source: ["M"], label: "Temperature (Oxidizing Room)Range 26C_30C" },
  { target: "K", source: ["N"], label: "Humidity Reading (Oxidizing Room)" },
  { target: "L", source: ["AA"], label: "Temperature (Flammable Room) Range 26C_30C" },
  { target: "M", source: ["AB"], label: "Humidity Reading (Flammable Room)" },
  { target: "N", source: ["U"], label: "Temperature (Ambient Rack) Range 26C_30C" },
  { target: "O", source: ["V"], label: "Humidity Reading (Ambient Rack)" },
  { target: "P", source: ["Q"], label: "Temperature (Toxic/Corrosive Room 1) Range 26C_30C" },
  { target: "Q", source: ["R"], label: "Humidity Reading\u00a0(Toxic/Corrosive Room 1)" },
  { target: "R", source: ["W"], label: "Temperature (Bioref) Range -20" },
  { target: "S", source: ["AD", "G"], label: "Temperature (Chiller) Range 2 -8" },
];
const STAT_COLS = ["C", "E", "G", "I", "J", "L", "N", "P", "R", "S"];

// Plausible value ranges per target column (generous bounds around the expected operating range).
// Used to detect guard typos: missing decimal points and impossible values.
const RANGES: Record<string, { min: number; max: number }> = {
  C: { min: 10, max: 50 },   // Ambient Shelves 26-30
  D: { min: 0, max: 100 },   // Humidity
  E: { min: 0, max: 40 },    // AC1 15-25
  F: { min: 0, max: 100 },   // Humidity
  G: { min: 0, max: 40 },    // AC2 15-25
  H: { min: 0, max: 100 },   // Humidity
  I: { min: -100, max: 10 },  // Ultralow -80 (keep real warm excursions like -45)
  J: { min: 10, max: 50 },   // Oxidizing 26-30
  K: { min: 0, max: 100 },   // Humidity
  L: { min: 10, max: 50 },   // Flammable 26-30
  M: { min: 0, max: 100 },   // Humidity
  N: { min: 10, max: 50 },   // Ambient Rack 26-30
  O: { min: 0, max: 100 },   // Humidity
  P: { min: 10, max: 50 },   // Toxic/Corrosive 26-30
  Q: { min: 0, max: 100 },   // Humidity
  R: { min: -40, max: 5 },   // Bioref -20
  S: { min: -10, max: 20 },  // Chiller 2-8
};
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "August", "Sept", "Oct", "Nov", "Dec"];

const colNum = (l: string) => l.split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colLetter = (n: number) => {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
const dayKey = (d: Date) => `${d.getUTCFullYear()}-${d.getUTCMonth()}-${d.getUTCDate()}`;
const addDays = (d: Date, n: number) => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate() + n));

function cellRaw(v: ExcelJS.CellValue): unknown {
  if (v && typeof v === "object" && !(v instanceof Date)) {
    const o = v as { result?: unknown; text?: unknown; richText?: { text: string }[] };
    if ("result" in o) return o.result;
    if (o.richText) return o.richText.map((r) => r.text).join("");
    if ("text" in o) return o.text;
  }
  return v;
}
function toNumber(v: unknown): number | null {
  v = cellRaw(v as ExcelJS.CellValue);
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/[^\d.\-]/g, ""));
    return /\d/.test(v) && !isNaN(n) ? n : null;
  }
  return null;
}
function toDate(v: unknown): Date | null {
  v = cellRaw(v as ExcelJS.CellValue);
  if (v instanceof Date) return new Date(Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()));
  if (typeof v === "number") return addDays(new Date(Date.UTC(1899, 11, 30)), Math.floor(v));
  if (typeof v === "string") { const d = new Date(v); if (!isNaN(+d)) return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); }
  return null;
}
function toHour(v: unknown): number | null {
  v = cellRaw(v as ExcelJS.CellValue);
  if (v instanceof Date) return v.getUTCHours() + (v.getUTCMinutes() >= 30 ? 1 : 0);
  if (typeof v === "number") return Math.round((v % 1) * 24) % 24 || (v >= 1 ? 0 : Math.round(v * 24) % 24);
  if (typeof v === "string") {
    const m = v.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm)?/i);
    if (!m) return null;
    let h = parseInt(m[1] ?? "0");
    if (m[3]) { const pm = m[3].toLowerCase() === "pm"; if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
    return h % 24;
  }
  return null;
}

export type Reading = { ts: number; values: Record<string, number | null> };

export type Correction = {
  date: string;
  hour: string;
  label: string;
  raw: number;
  cleaned: number | null;
  action: "decimal" | "blank";
};

export type ParsedResponses = { map: Map<string, Reading>; minDate: Date | null; maxDate: Date | null; rows: number; corrections: Correction[] };

function cleanValue(val: number, target: string): { value: number | null; action: "decimal" | "blank" | null } {
  const range = RANGES[target];
  if (!range) return { value: val, action: null };
  if (val >= range.min && val <= range.max) return { value: val, action: null };
  // Missing decimal point: value is ~10x the expected range and dividing by 10 lands in range.
  // Only apply to positive values — negative out-of-range values are more likely wrong-column errors.
  // Negative values only when clearly 10x (e.g. -180 -> -18), so a wrong-field -80 in Bio Ref is not turned into -8.
  if ((val > 0 || val <= -100) && val / 10 >= range.min && val / 10 <= range.max) return { value: Math.round(val) / 10, action: "decimal" };
  return { value: null, action: "blank" };
}

export async function parseResponses(buf: ArrayBuffer): Promise<ParsedResponses> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const map = new Map<string, Reading>();
  const corrections: Correction[] = [];
  let minDate: Date | null = null, maxDate: Date | null = null, rows = 0;
  ws?.eachRow((row, i) => {
    if (i === 1) return;
    const date = toDate(row.getCell("C").value);
    const hour = toHour(row.getCell("D").value);
    if (!date || hour === null) return;
    rows++;
    const tsRaw = cellRaw(row.getCell("A").value);
    const ts = tsRaw instanceof Date ? +tsRaw : i;
    const key = `${dayKey(date)}|${hour}`;
    const slot = +date + hour * 3600000;
    const prev = map.get(key);
    // several submissions for the same hour: keep the one submitted closest to that hour
    if (prev && Math.abs(prev.ts - slot) <= Math.abs(ts - slot)) return;
    const hourLabel = hour === 0 ? "2400H" : `${String(hour).padStart(2, "0")}00H`;
    const dateStr = date.toISOString().slice(0, 10);
    const values: Record<string, number | null> = {};
    for (const m of COLUMN_MAP) {
      let val: number | null = null;
      for (const s of m.source) { val = toNumber(row.getCell(s).value); if (val !== null) break; }
      if (val !== null) {
        const cleaned = cleanValue(val, m.target);
        if (cleaned.action) {
          corrections.push({ date: dateStr, hour: hourLabel, label: m.label, raw: val, cleaned: cleaned.value, action: cleaned.action });
        }
        val = cleaned.value;
      }
      values[m.target] = val;
    }
    map.set(key, { ts, values });
    if (!minDate || date < minDate) minDate = date;
    if (!maxDate || date > maxDate) maxDate = date;
  });
  return { map, minDate, maxDate, rows, corrections };
}

export function sheetNameFor(start: Date) {
  const end = addDays(start, 6);
  const f = (d: Date) => `${MONTHS[d.getUTCMonth()]} ${d.getUTCDate()}`;
  return `Temp ${f(start)}-${f(end)}`;
}

