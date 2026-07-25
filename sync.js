// sync.js — push the offline outbox into the workbook and pull reference data.
import { CONFIG } from "./config.js";
import * as db from "./db.js";
import * as graph from "./graph.js";
import { getToken } from "./auth.js";
import { timeLogWrites, calendarWrites } from "./models.js";

// Excel's "=" comparison is case-insensitive, so Projects!M4's IF($D4="Active",...)
// matches "active" and "ACTIVE" too. Match that behaviour rather than being stricter
// than the workbook the list is mirroring.
const isActive = (status) => String(status ?? "").trim().toLowerCase() === "active";

// Pull the live Projects list (Name, Project #, Client) and cache it for offline use.
// Reads A:D so Status can be filtered — the app must show the same 12 projects the
// Time Log dropdown shows, not all 23. Anything not "Active" (Inactive, Dormant) is
// excluded here exactly as ActiveProjectList excludes it in the workbook.
export async function refreshProjects() {
  const token = await getToken();
  const projects = await graph.withWorkbook(token, async (id) => {
    const rows = await graph.readRange(token, id, CONFIG.sheets.projects, "A4:D200");
    return rows
      .filter((r) => r[0] !== "" && r[0] != null && isActive(r[3]))
      .map((r) => ({ name: r[0], number: r[1] ?? "", client: r[2] ?? "" }));
  });
  await db.setCache("projects", projects);
  return projects;
}

// Pull the Calendar (schedule) entries and cache them for the upcoming-30-days view.
export async function fetchUpcoming() {
  const token = await getToken();
  const items = await graph.withWorkbook(token, async (id) => {
    const rows = await graph.readRange(token, id, CONFIG.sheets.calendar, "A4:F204");
    return rows
      .map((r) => ({ date: r[0], day: r[1], project: r[2], client: r[3], task: r[4], notes: r[5] }))
      .filter((x) => x.date !== "" && x.date != null);
  });
  await db.setCache("calendar", items);
  return items;
}

// Flush queued entries in order. Each lands in the first empty row; stops on first failure
// so nothing is lost or duplicated.
//
// The outbox is read INSIDE withWorkbook so that if a stale file id forces a retry, the
// second pass sees the current queue rather than a stale snapshot. A stale id fails on
// firstEmptyRow — the first call of each iteration, before any PATCH — so a retry cannot
// double-write an entry that already landed.
export async function flushOutbox(onProgress) {
  const token = await getToken();
  return graph.withWorkbook(token, async (id) => {
    const items = await db.allOutbox();
    if (!items.length) return { done: 0, total: 0 };
    let done = 0;
    for (const item of items) {
      const isCal = item.kind === "calendar";
      const sheetName = isCal ? CONFIG.sheets.calendar : CONFIG.sheets.timeLog;
      const maxRow = isCal ? CONFIG.calendarMaxRow : CONFIG.timeLogMaxRow;
      const row = await graph.firstEmptyRow(token, id, sheetName, "A", 4, maxRow);
      if (row == null) throw new Error(`${sheetName} is full — extend it on the computer.`);
      const writes = isCal ? calendarWrites(row, item.payload) : timeLogWrites(row, item.payload);
      for (const w of writes) await graph.writeRange(token, id, sheetName, w.range, w.values);
      await db.deleteOutbox(item.id);
      done++;
      if (onProgress) onProgress(done, items.length);
    }
    return { done, total: items.length };
  });
}

// Most recent Time Log rows for the Status tab (read a small window near the end).
export async function recentTimeLog(limit = 20) {
  const token = await getToken();
  return graph.withWorkbook(token, async (id) => {
    const firstEmpty = await graph.firstEmptyRow(token, id, CONFIG.sheets.timeLog, "A", 4, CONFIG.timeLogMaxRow);
    const lastData = firstEmpty == null ? CONFIG.timeLogMaxRow : firstEmpty - 1;
    if (lastData < 4) return [];
    const start = Math.max(4, lastData - limit + 1);
    const rows = await graph.readRange(token, id, CONFIG.sheets.timeLog, `A${start}:R${lastData}`);
    return rows
      .map((r, i) => ({
        row: start + i,
        date: r[0], project: r[1], client: r[2], task: r[3], hours: r[4],
        billable: r[5], reportSubmitted: r[13], invoice: r[17],
      }))
      .filter((x) => x.date !== "" && x.date != null)
      .reverse();
  });
}

export async function setBillable(rowNum, yes) {
  const token = await getToken();
  return graph.withWorkbook(token, (id) =>
    graph.writeRange(token, id, CONFIG.sheets.timeLog, `F${rowNum}`, [[yes ? "Y" : "N"]]));
}

export async function setReportSubmitted(rowNum, yes) {
  const token = await getToken();
  return graph.withWorkbook(token, (id) =>
    graph.writeRange(token, id, CONFIG.sheets.timeLog, `N${rowNum}`, [[yes ? "Yes" : "No"]]));
}
