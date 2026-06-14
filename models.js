// models.js — column mapping, value coercion, and the Excel write-plan builders.
// The app writes INPUT columns only; the workbook's pre-filled formulas
// (Time Log G/L/M/O/Q, Calendar B/D) compute the rest. Never write those cells.

import { CONFIG } from "./config.js";

// Excel 1900 date system: serial = whole days since 1899-12-30.
const EPOCH = Date.UTC(1899, 11, 30);

export function dateToSerial(isoDate) {
  const [y, m, d] = String(isoDate).split("-").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}
export function serialToISO(serial) {
  return new Date(EPOCH + Number(serial) * 86400000).toISOString().slice(0, 10);
}
export function todayISO() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth() + 1).padStart(2, "0")}-${String(n.getDate()).padStart(2, "0")}`;
}

const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };

// Replace smart punctuation so the workbook + QuickBooks stay plain ASCII.
export function clean(s) {
  if (s == null) return "";
  return String(s)
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[–—]/g, "-")
    .replace(/…/g, "...");
}

// Time Log PATCH ops for a given 1-based sheet row.
// Skips G/L/M/O/Q (formulas), R (Invoice# — blank = unbilled), S/T (blank in v1).
export function timeLogWrites(row, e) {
  return [
    { range: `A${row}:F${row}`, values: [[
      dateToSerial(e.date), clean(e.project), clean(e.client),
      clean(e.taskType), num(e.hours), e.billable ? "Y" : "N",
    ]] },
    { range: `H${row}:K${row}`, values: [[
      num(e.mileage), num(e.perDiemNights), clean(e.equipment), clean(e.notes),
    ]] },
    { range: `N${row}`, values: [[ e.reportSubmitted ? "Yes" : "No" ]] },
    { range: `P${row}`, values: [[ clean(e.wo) ]] },
  ];
}

// Calendar PATCH ops. Skips B (Day) and D (Client) — both formulas.
export function calendarWrites(row, c) {
  return [
    { range: `A${row}`, values: [[ dateToSerial(c.date) ]] },
    { range: `C${row}`, values: [[ clean(c.project) ]] },
    { range: `E${row}:F${row}`, values: [[ clean(c.task), clean(c.notes) ]] },
  ];
}

export const maxRowFor = (sheet) =>
  sheet === CONFIG.sheets.calendar ? CONFIG.calendarMaxRow : CONFIG.timeLogMaxRow;
