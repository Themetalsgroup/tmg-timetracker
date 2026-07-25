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

// Resolve the workbook's driveItem id from its path, and cache it.
//
// A cached id CAN go stale. If the workbook is restored from _backups/, re-uploaded,
// deleted and re-created, or moved between drives, OneDrive mints a brand new item id.
// The old id then 404s on every call — and because the cache lives in IndexedDB it
// survives force-quitting the app, so the failure looks permanent and unexplained.
// Pass { force: true } to bypass the cache and re-resolve from the path.
export async function fileId(token, { force = false } = {}) {
  if (!force) {
    const cached = await db.getCache("fileId");
    if (cached) return cached;
  }
  const path = CONFIG.workbookPath.split("/").map(encodeURIComponent).join("/");
  const item = await gfetch(token, `${BASE}/me/drive/root:${path}`);
  await db.setCache("fileId", item.id);
  return item.id;
}

// Drop the cached id so the next call re-resolves from the path.
export const invalidateFileId = () => db.delCache("fileId");

// Run a workbook operation with a self-healing file id.
//
// On a 404 the cached id is discarded, the path is re-resolved, and the operation is
// retried exactly once. If it 404s again the id was never the problem — most likely a
// sheet name in CONFIG.sheets doesn't match the workbook — so the error says so.
//
// NOTE: `fn` must be safe to run twice. A stale id fails on the first Graph call of an
// operation, before any PATCH lands, so retrying does not double-write. Callbacks that
// mutate local state should re-read it inside `fn` rather than closing over it.
export async function withWorkbook(token, fn) {
  const id = await fileId(token);
  try {
    return await fn(id);
  } catch (e) {
    if (!e || e.status !== 404) throw e;
    await invalidateFileId();
    const fresh = await fileId(token, { force: true });
    try {
      return await fn(fresh);
    } catch (e2) {
      if (e2 && e2.status === 404) {
        e2.message += " — workbook re-resolved, so check the sheet names in config.js against the file";
      }
      throw e2;
    }
  }
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
