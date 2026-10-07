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
  target: string;
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
          corrections.push({ date: dateStr, hour: hourLabel, label: m.label, target: m.target, raw: val, cleaned: cleaned.value, action: cleaned.action });
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

export type ConvertResult = { blob: Blob; fileName: string; filled: number; missing: string[]; corrections: Correction[]; notes: string[] };

const MONTH_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const monthIdx = (v: unknown) => {
  const t = String(cellRaw(v as ExcelJS.CellValue) ?? "").trim().slice(0, 3).toLowerCase();
  const i = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(t);
  return i >= 0 ? i : null;
};
const copyStyle = (from: ExcelJS.Cell, to: ExcelJS.Cell) => { to.style = JSON.parse(JSON.stringify(from.style ?? {})); };
const numOf = (c: ExcelJS.Cell): number | null => { const v = cellRaw(c.value); return typeof v === "number" && isFinite(v) ? v : null; };

/** Checks the uploaded MKT workbook has the Summary + weekly formula structure needed for conversion. */
function validateMkt(wb: ExcelJS.Workbook) {
  const sum = wb.getWorksheet("Summary");
  if (!sum) throw new Error("The MKT workbook has no 'Summary' sheet.");
  let links = 0;
  sum.getRow(4).eachCell((cell) => { if (cell.type === ExcelJS.ValueType.Formula && cell.formula.includes("Temp ")) links++; });
  if (!links) throw new Error("The Summary sheet has no formulas linking to weekly 'Temp' sheets.");
  const hasWeekly = wb.worksheets.some((w) => w.name.startsWith("Temp ") && w.getCell("C178").type === ExcelJS.ValueType.Formula);
  if (!hasWeekly) throw new Error("No weekly 'Temp' sheet with the MKT formulas (row 178) was found.");
}

export async function buildWorkbook(parsed: ParsedResponses, start: Date, mktBuf: ArrayBuffer): Promise<ConvertResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(mktBuf);
  validateMkt(wb);
  const notes: string[] = [];
  const name = sheetNameFor(start);
  const existing = wb.getWorksheet(name);
  if (existing) { wb.removeWorksheet(existing.id); notes.push(`Replaced the existing '${name}' sheet.`); }
  const template = wb.worksheets.filter((w) => w.name.startsWith("Temp ") && w.getCell("C178").type === ExcelJS.ValueType.Formula)[0]!;
  const ws = wb.addWorksheet(name);

  // header, styled like the latest existing weekly sheet
  const header = ws.getRow(1);
  header.getCell("A").value = "Input date";
  header.getCell("B").value = "Time";
  COLUMN_MAP.forEach((m) => (header.getCell(m.target).value = m.label));
  for (let c = 1; c <= 19; c++) {
    const w = template.getColumn(c).width; if (w) ws.getColumn(c).width = w;
    copyStyle(template.getRow(1).getCell(c), header.getCell(c));
  }
  if (template.getRow(1).height) header.height = template.getRow(1).height;

  // week's corrections, keyed for highlighting
  const weekDays = new Set(Array.from({ length: 7 }, (_, d) => addDays(start, d).toISOString().slice(0, 10)));
  const corrections = parsed.corrections.filter((c) => weekDays.has(c.date));
  const corrMap = new Map(corrections.map((c) => [`${c.date}|${c.hour}|${c.target}`, c]));

  let filled = 0; const missing: string[] = [];
  for (let d = 0; d < 7; d++) {
    const date = addDays(start, d);
    const ds = date.toISOString().slice(0, 10);
    for (let h = 0; h < 24; h++) {
      const r = ws.getRow(2 + d * 24 + h);
      const hl = h === 0 ? "2400H" : `${String(h).padStart(2, "0")}00H`;
      r.getCell("A").value = date; r.getCell("A").numFmt = "m/d/yyyy";
      r.getCell("B").value = hl;
      const rd = parsed.map.get(`${dayKey(date)}|${h}`);
      if (rd) { filled++; COLUMN_MAP.forEach((m) => { if (rd.values[m.target] !== null) r.getCell(m.target).value = rd.values[m.target]; }); }
      else missing.push(`${ds} ${hl}`);
      for (const m of COLUMN_MAP) {
        const c = corrMap.get(`${ds}|${hl}|${m.target}`);
        if (!c) continue;
        const cell = r.getCell(m.target);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
        cell.note = c.action === "decimal" ? `Guard entered ${c.raw}; corrected to ${c.cleaned} (missing decimal point)` : `Guard entered ${c.raw}; removed (impossible value)`;
      }
    }
  }
  ws.getCell("B170").value = "2400H";

  // Same formulas as the existing weekly MKT sheets, with cached results so values show instantly
  ws.getCell("B173").value = "MIN"; ws.getCell("B174").value = "MAX"; ws.getCell("B175").value = "AVERAGE"; ws.getCell("B178").value = "MKT";
  const stats: Record<string, number> = {};
  for (const c of STAT_COLS) {
    const vals: number[] = [];
    for (let r = 2; r <= 170; r++) { const v = ws.getCell(`${c}${r}`).value; if (typeof v === "number") vals.push(v); }
    const cells: [number, string][] = [
      [173, `MIN(${c}2:${c}170)`], [174, `MAX(${c}2:${c}170)`], [175, `AVERAGE(${c}2:${c}170)`],
      [176, `${c}175+273.15`], [177, `-$B$181/($B$182*${c}176)`], [178, `$B$181/($B$182*(-${c}177))-$B$183`],
    ];
    if (vals.length) {
      const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
      const k = avg + 273.15, ln = -83.14472 / (0.008314472 * k);
      Object.assign(stats, { [`${c}173`]: Math.min(...vals), [`${c}174`]: Math.max(...vals), [`${c}175`]: avg, [`${c}176`]: k, [`${c}177`]: ln, [`${c}178`]: 83.14472 / (0.008314472 * -ln) - 273.15 });
    }
    for (const [n, f] of cells) {
      const res = stats[`${c}${n}`];
      ws.getCell(`${c}${n}`).value = (res !== undefined ? { formula: f, result: res } : { formula: f }) as ExcelJS.CellValue;
    }
    for (const n of [173, 174, 175, 178]) ws.getCell(`${c}${n}`).numFmt = "0.00";
  }
  ws.getCell("A181").value = "Delta H"; ws.getCell("B181").value = 83.14472; ws.getCell("C181").value = "kJ/mole";
  ws.getCell("A182").value = "R"; ws.getCell("B182").value = 0.008314472; ws.getCell("C182").value = "kJ/mole/degree";
  ws.getCell("A183").value = "°C to K"; ws.getCell("B183").value = 273.15;
  for (const n of [173, 174, 175, 178]) ws.getCell(`B${n}`).font = { bold: true };

  // place new sheet right after Summary
  const ordered = wb.worksheets.filter((w) => w !== ws);
  const sumIdx = ordered.findIndex((w) => w.name === "Summary");
  ordered.splice(sumIdx + 1, 0, ws);
  ordered.forEach((w, i) => ((w as unknown as { orderNo: number }).orderNo = i));
  updateSummary(wb, name, start, stats, notes);

  (wb as unknown as { calcProperties: { fullCalcOnLoad: boolean } }).calcProperties = { fullCalcOnLoad: true };
  const out = await wb.xlsx.writeBuffer();
  const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const end = addDays(start, 6);
  const up = (d: Date) => `${(MONTHS[d.getUTCMonth()] ?? "").toUpperCase()}_${d.getUTCDate()}`;
  return { blob, fileName: `${up(start)}-${up(end)}_MKT_Monitoring_ULC_Merck.xlsx`, filled, missing, corrections, notes };
}

function weekLinkCols(sum: ExcelJS.Worksheet) {
  const cols: number[] = [];
  sum.getRow(4).eachCell((cell, c) => { if (cell.type === ExcelJS.ValueType.Formula && cell.formula.includes("Temp ")) cols.push(c); });
  return cols;
}

/** Shift every Summary cell at column >= col one column to the right (used to make room for a new week). */
function shiftRight(sum: ExcelJS.Worksheet, col: number) {
  const merges: string[] = ((sum.model as unknown as { merges?: string[] }).merges ?? []).slice();
  const moved: [number, number, number, number][] = [];
  for (const m of merges) {
    const [a, b] = m.split(":");
    const ca = sum.getCell(a!), cb = sum.getCell(b ?? a!);
    if (+ca.col >= col) { sum.unMergeCells(m); moved.push([+ca.row, +ca.col + 1, +cb.row, +cb.col + 1]); }
  }
  const maxCol = sum.columnCount, maxRow = sum.rowCount;
  for (let c = maxCol; c >= col; c--) {
    const w = sum.getColumn(c).width; if (w) sum.getColumn(c + 1).width = w;
    for (let r = 1; r <= maxRow; r++) {
      const src = sum.getCell(r, c), dst = sum.getCell(r, c + 1);
      dst.value = src.type === ExcelJS.ValueType.Formula ? ({ formula: src.formula, result: src.result } as ExcelJS.CellValue) : src.value;
      copyStyle(src, dst);
      src.value = null; src.style = {};
    }
  }
  for (const [r1, c1, r2, c2] of moved) sum.mergeCells(r1, c1, r2, c2);
}

function updateSummary(wb: ExcelJS.Workbook, sheetName: string, start: Date, stats: Record<string, number>, notes: string[]) {
  const sum = wb.getWorksheet("Summary")!;
  let cols = weekLinkCols(sum);
  const last = cols[cols.length - 1]!;
  const lastF = sum.getRow(4).getCell(last).formula;
  const already = cols.find((c) => sum.getRow(4).getCell(c).formula.includes(`'${sheetName}'`));
  const target = already ?? last + 1;

  if (!already) {
    // make room if the "2026 WEEKLY MKT" table (or anything else) sits right where the new week goes
    let occupied = false;
    for (let r = 1; r <= sum.rowCount; r++) if (cellRaw(sum.getCell(r, target).value) != null && cellRaw(sum.getCell(r, target).value) !== "") { occupied = true; break; }
    if (occupied) { shiftRight(sum, target); notes.push("Moved the 2026 WEEKLY MKT table one column right to make room for the new week."); }

    sum.eachRow((row) => {
      const src = row.getCell(last);
      if (src.type !== ExcelJS.ValueType.Formula || !src.formula.includes("Temp ")) return;
      const formula = src.formula.replace(/'[^']*'!/g, `'${sheetName}'!`);
      const ref = formula.match(/^'[^']*'!\$?([A-Z]+)\$?(\d+)$/);
      const result = ref ? stats[`${ref[1]}${ref[2]}`] : undefined;
      const dst = row.getCell(target);
      dst.value = (result !== undefined ? { formula, result } : { formula }) as ExcelJS.CellValue;
      copyStyle(src, dst);
    });
    // fill styling of empty separator rows too
    for (let r = 3; r <= 52; r++) if (sum.getCell(r, target).value == null) copyStyle(sum.getCell(r, last), sum.getCell(r, target));
    const end = addDays(start, 6);
    sum.getCell(3, target).value = `${start.getUTCDate()}-${end.getUTCDate()}`;

    // month header in row 2: a week belongs to the month holding its Thursday (4+ of its 7 days)
    const m = addDays(start, 3).getUTCMonth();
    const lastMonth = monthIdx(sum.getCell(2, last).value);
    const merges: string[] = ((sum.model as unknown as { merges?: string[] }).merges ?? []).slice();
    const lastMerge = merges.find((g) => { const [a, b] = g.split(":"); const ca = sum.getCell(a!), cb = sum.getCell(b!); return +ca.row === 2 && +cb.row === 2 && +ca.col <= last && +cb.col >= last; });
    if (lastMonth === m) {
      let left = last;
      if (lastMerge) { left = +sum.getCell(lastMerge.split(":")[0]!).col; sum.unMergeCells(lastMerge); }
      sum.mergeCells(2, left, 2, target);
    } else {
      sum.getCell(2, target).value = MONTH_FULL[m]!;
      copyStyle(sum.getCell(2, last), sum.getCell(2, target));
    }
    void lastF;
    cols = weekLinkCols(sum);
  }

  // ---- 2026 WEEKLY MKT table: monthly average of the weekly MKT values ----
  let hr = 0, sc = 0;
  sum.eachRow((row, r) => row.eachCell((cell, c) => { if (!hr && String(cellRaw(cell.value) ?? "").trim() === "Storage") { hr = r; sc = c; } }));
  if (!hr) { notes.push("No '2026 WEEKLY MKT' table found — monthly table not updated."); return; }
  const monthCol: Record<number, number> = {};
  let avgCol = 0;
  for (let c = sc + 2; c <= sc + 15; c++) {
    const v = String(cellRaw(sum.getCell(hr, c).value) ?? "").trim();
    if (v.toLowerCase() === "average") avgCol = c;
    else { const i = monthIdx(v); if (i !== null) monthCol[i] = c; }
  }
  // which weekly columns belong to which month (from the row-2 month headers)
  const weeksByMonth: Record<number, number[]> = {};
  for (const c of cols) { const i = monthIdx(sum.getCell(2, c).value); if (i !== null) (weeksByMonth[i] ??= []).push(c); }
  const L = colLetter;
  for (let i = 0; i < 10; i++) {
    const r = hr + 1 + i;
    if (!cellRaw(sum.getCell(r, sc).value)) break;
    const mktRow = 7 + 5 * i; // Summary blocks: Chiller row 7, Ambient row 12, ... each block's MKT row
    const monthVals: number[] = [];
    for (const [mi, col] of Object.entries(monthCol)) {
      const wk = weeksByMonth[+mi] ?? [];
      const vals = wk.map((c) => numOf(sum.getCell(mktRow, c))).filter((v): v is number => v !== null);
      const cell = sum.getCell(r, col);
      if (!wk.length || !vals.length) { if (!wk.length) cell.value = null; continue; }
      const contiguous = wk.every((c, k) => k === 0 || c === wk[k - 1]! + 1);
      const formula = contiguous ? `AVERAGE(${L(wk[0]!)}${mktRow}:${L(wk[wk.length - 1]!)}${mktRow})` : `AVERAGE(${wk.map((c) => `${L(c)}${mktRow}`).join(",")})`;
      const result = vals.reduce((a, b) => a + b, 0) / vals.length;
      cell.value = { formula, result } as ExcelJS.CellValue;
      cell.numFmt = "0.00";
      monthVals.push(result);
    }
    if (avgCol && monthVals.length) {
      const cols12 = Object.values(monthCol);
      const a = Math.min(...cols12), b = Math.max(...cols12);
      sum.getCell(r, avgCol).value = { formula: `AVERAGE(${L(a)}${r}:${L(b)}${r})`, result: monthVals.reduce((x, y) => x + y, 0) / monthVals.length } as ExcelJS.CellValue;
      sum.getCell(r, avgCol).numFmt = "0.00";
    }
  }
  void colNum;
}
