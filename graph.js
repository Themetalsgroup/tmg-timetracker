// graph.js — minimal Microsoft Graph workbook client for the Time Tracker file.
import { CONFIG } from "./config.js";
import * as db from "./db.js";

const BASE = "https://graph.microsoft.com/v1.0";

async function gfetch(token, url, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).error?.message || ""; } catch { /* ignore */ }
    const err = new Error(`Graph ${res.status}${detail ? `: ${detail}` : ""}`);
    err.status = res.status;
    throw err;
  }
  return res.status === 204 ? null : res.json();
}

// Resolve + cache the workbook's driveItem id (survives renames; handles the spaced filename).
export async function fileId(token) {
  let id = await db.getCache("fileId");
  if (id) return id;
  const path = CONFIG.workbookPath.split("/").map(encodeURIComponent).join("/");
  const item = await gfetch(token, `${BASE}/me/drive/root:${path}`);
  id = item.id;
  await db.setCache("fileId", id);
  return id;
}

const wb = (id) => `${BASE}/me/drive/items/${id}/workbook`;
const ws = (name) => `worksheets('${encodeURIComponent(name)}')`;

export async function readRange(token, id, sheetName, address) {
  const url = `${wb(id)}/${ws(sheetName)}/range(address='${address}')?$select=values`;
  const r = await gfetch(token, url);
  return (r && r.values) || [];
}

export function writeRange(token, id, sheetName, address, values) {
  const url = `${wb(id)}/${ws(sheetName)}/range(address='${address}')`;
  return gfetch(token, url, { method: "PATCH", body: JSON.stringify({ values }) });
}

// First row whose key column is blank, within [startRow, maxRow]; null if none (sheet full).
export async function firstEmptyRow(token, id, sheetName, keyCol, startRow, maxRow) {
  const values = await readRange(token, id, sheetName, `${keyCol}${startRow}:${keyCol}${maxRow}`);
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v === "" || v === null || v === undefined) return startRow + i;
  }
  return null;
}
