// sync.js — push the offline outbox into the workbook and pull reference data.
import { CONFIG } from "./config.js";
import * as db from "./db.js";
import * as graph from "./graph.js";
import { getToken } from "./auth.js";
import { timeLogWrites, calendarWrites } from "./models.js";

// Pull the live Projects list (Name, Project #, Client) and cache it for offline use.
export async function refreshProjects() {
  const token = await getToken();
  const id = await graph.fileId(token);
  const rows = await graph.readRange(token, id, CONFIG.sheets.projects, "A4:C200");
  const projects = rows
    .filter((r) => r[0] !== "" && r[0] != null)
    .map((r) => ({ name: r[0], number: r[1] ?? "", client: r[2] ?? "" }));
  await db.setCache("projects", projects);
  return projects;
}

// Flush queued entries in order. Each lands in the first empty row; stops on first failure
// so nothing is lost or duplicated.
export async function flushOutbox(onProgress) {
  const items = await db.allOutbox();
  if (!items.length) return { done: 0, total: 0 };
  const token = await getToken();
  const id = await graph.fileId(token);
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
}

// Most recent Time Log rows for the Status tab (read a small window near the end).
export async function recentTimeLog(limit = 20) {
  const token = await getToken();
  const id = await graph.fileId(token);
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
}

export async function setBillable(rowNum, yes) {
  const token = await getToken();
  const id = await graph.fileId(token);
  return graph.writeRange(token, id, CONFIG.sheets.timeLog, `F${rowNum}`, [[yes ? "Y" : "N"]]);
}

export async function setReportSubmitted(rowNum, yes) {
  const token = await getToken();
  const id = await graph.fileId(token);
  return graph.writeRange(token, id, CONFIG.sheets.timeLog, `N${rowNum}`, [[yes ? "Yes" : "No"]]);
}
