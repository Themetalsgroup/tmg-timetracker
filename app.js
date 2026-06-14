// app.js — TMG Time Tracker. Offline capture + Microsoft 365 sync into the workbook.
import { CONFIG, isConfigured, SAMPLE_PROJECTS } from "./config.js";
import * as db from "./db.js";
import { todayISO, serialToISO } from "./models.js";
import { initAuth, getAccount, signIn, signOut } from "./auth.js";
import { flushOutbox, refreshProjects, recentTimeLog, setBillable, setReportSubmitted } from "./sync.js";

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
function el(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null) continue;
    if (k === "class") n.className = v;
    else if (k.startsWith("on") && typeof v === "function") n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v);
  }
  for (const kid of kids) if (kid != null) n.append(kid.nodeType ? kid : document.createTextNode(kid));
  return n;
}

let PROJECTS = [];
let PENDING = 0;
let AUTH_READY = false;
let SYNCING = false;
let SYNC_ERR = "";

async function boot() {
  PENDING = await db.countOutbox();
  PROJECTS = (await db.getCache("projects")) || (isConfigured() ? [] : SAMPLE_PROJECTS);
  wireTabs();
  renderBanner();
  renderLog();
  renderSchedule();
  renderStatus();

  window.addEventListener("online", () => { renderBanner(); runSync(); });
  window.addEventListener("offline", renderBanner);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) runSync(); });
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
  }

  if (isConfigured() && typeof msal !== "undefined") {
    try {
      await initAuth();
      AUTH_READY = true;
      renderBanner();
      if (getAccount()) runSync(); // pull projects + flush anything queued
    } catch (e) {
      console.error("auth init failed", e);
      SYNC_ERR = "Sign-in unavailable";
      renderBanner();
    }
  }
}

function wireTabs() {
  $("#tabs").addEventListener("click", (e) => {
    const btn = e.target.closest(".tab");
    if (!btn) return;
    const view = btn.dataset.view;
    $$(".tab").forEach((t) => t.setAttribute("aria-current", String(t === btn)));
    for (const v of ["log", "schedule", "status"]) $(`#view-${v}`).hidden = v !== view;
    if (view === "status") loadStatus();
  });
}

// ---- sync banner + actions ----
function renderBanner() {
  const b = $("#banner");
  b.hidden = false;
  const account = AUTH_READY ? getAccount() : null;
  let msg, cls = "warn", action = null;
  if (!isConfigured()) msg = "Demo mode — Microsoft not set up";
  else if (!AUTH_READY) msg = navigator.onLine ? "Starting…" : "Offline — entries are saved on device";
  else if (!account) { msg = "Not signed in"; action = { label: "Sign in", fn: doSignIn }; }
  else if (SYNCING) msg = "Syncing…";
  else if (SYNC_ERR) { msg = SYNC_ERR; cls = "err"; if (navigator.onLine) action = { label: "Retry", fn: doSync }; }
  else if (!navigator.onLine) msg = PENDING ? `Offline · ${PENDING} waiting` : "Offline — saved on device";
  else if (PENDING) { msg = `${PENDING} waiting to sync`; action = { label: "Sync now", fn: doSync }; }
  else { msg = "All synced"; cls = "ok"; }
  b.className = `banner ${cls}`;
  b.replaceChildren(
    el("span", { class: "banner-msg" }, msg),
    action ? el("button", { class: "banner-btn", onclick: action.fn }, action.label) : null,
  );
}

async function doSignIn() { try { await signIn(); } catch (e) { SYNC_ERR = e.message; renderBanner(); } }
async function doSignOut() { try { await signOut(); } catch (e) { console.error(e); } }
function doSync() { runSync(); }

async function runSync() {
  if (!isConfigured() || !AUTH_READY || !getAccount() || !navigator.onLine || SYNCING) return;
  SYNCING = true; SYNC_ERR = ""; renderBanner();
  try {
    await flushOutbox();
    PROJECTS = await refreshProjects();
    updateProjectSelects();
  } catch (e) {
    SYNC_ERR = friendly(e);
  } finally {
    SYNCING = false;
    PENDING = await db.countOutbox();
    renderBanner();
  }
}
function friendly(e) {
  const m = e && e.message ? e.message : String(e);
  if (/401|InteractionRequired|Not signed in/i.test(m)) return "Tap Sign in to reconnect";
  if (/423|locked/i.test(m)) return "File is open on the computer — will retry";
  return m;
}

async function bumpPending() { PENDING = await db.countOutbox(); renderBanner(); }

// ---- shared form bits ----
const projectOptions = () => [
  el("option", { value: "" }, PROJECTS.length ? "Select project…" : "Sign in to load projects"),
  ...PROJECTS.map((p) => el("option", { value: p.name }, p.name)),
];
const taskOptions = () => CONFIG.taskTypes.map((t) => el("option", { value: t }, t));
const field = (label, control) =>
  el("label", { class: "field" }, el("span", { class: "field-label" }, label), control);

function updateProjectSelects() {
  for (const id of ["#lt-project", "#sc-project"]) {
    const sel = $(id);
    if (!sel) continue;
    const cur = sel.value;
    sel.replaceChildren(...projectOptions());
    if ([...sel.options].some((o) => o.value === cur)) sel.value = cur;
  }
}

// ---- Log Time ----
function renderLog() {
  $("#view-log").replaceChildren(
    el("h2", {}, "Log Time"),
    field("Date", el("input", { type: "date", value: todayISO(), id: "lt-date" })),
    field("Project", el("select", { id: "lt-project" }, ...projectOptions())),
    field("Task", el("select", { id: "lt-task" }, ...taskOptions())),
    field("Hours", el("input", { type: "number", inputmode: "decimal", step: "0.25", min: "0", id: "lt-hours", placeholder: "0" })),
    field("Miles", el("input", { type: "number", inputmode: "decimal", step: "1", min: "0", id: "lt-miles", placeholder: "0" })),
    field("Per-diem nights", el("input", { type: "number", inputmode: "numeric", step: "1", min: "0", id: "lt-perdiem", placeholder: "0" })),
    el("label", { class: "field field-row" },
      el("span", { class: "field-label" }, "Billable"),
      el("input", { type: "checkbox", id: "lt-billable", checked: "checked" })),
    field("Equipment", el("input", { type: "text", id: "lt-equip", placeholder: "Optional" })),
    field("WO / PO #", el("input", { type: "text", id: "lt-wo", placeholder: "Optional" })),
    field("Notes", el("textarea", { id: "lt-notes", rows: "2", placeholder: "Optional" })),
    el("div", { class: "form-msg", id: "lt-msg" }),
    el("button", { class: "save", onclick: saveTime }, "Save entry"),
  );
}

async function saveTime() {
  const project = $("#lt-project").value;
  const hours = $("#lt-hours").value;
  if (!project) return flash($("#lt-msg"), "Pick a project.", false);
  if (!hours || Number(hours) <= 0) return flash($("#lt-msg"), "Enter hours.", false);
  const p = PROJECTS.find((x) => x.name === project);
  await db.addOutbox({ kind: "timelog", payload: {
    date: $("#lt-date").value, project, client: p ? p.client : "",
    taskType: $("#lt-task").value, hours, mileage: $("#lt-miles").value,
    perDiemNights: $("#lt-perdiem").value, billable: $("#lt-billable").checked,
    equipment: $("#lt-equip").value, wo: $("#lt-wo").value, notes: $("#lt-notes").value,
    reportSubmitted: false,
  } });
  await bumpPending();
  renderLog();
  flash($("#lt-msg"), "Saved.", true);
  runSync();
}

// ---- Schedule ----
function renderSchedule() {
  $("#view-schedule").replaceChildren(
    el("h2", {}, "Schedule"),
    el("p", { class: "hint" }, "Add an upcoming job to your calendar."),
    field("Date", el("input", { type: "date", value: todayISO(), id: "sc-date" })),
    field("Project", el("select", { id: "sc-project" }, ...projectOptions())),
    field("Task", el("select", { id: "sc-task" }, ...taskOptions())),
    field("Notes", el("textarea", { id: "sc-notes", rows: "2", placeholder: "Optional" })),
    el("div", { class: "form-msg", id: "sc-msg" }),
    el("button", { class: "save", onclick: saveSchedule }, "Add to schedule"),
  );
}

async function saveSchedule() {
  const project = $("#sc-project").value;
  if (!project) return flash($("#sc-msg"), "Pick a project.", false);
  await db.addOutbox({ kind: "calendar", payload: {
    date: $("#sc-date").value, project, task: $("#sc-task").value, notes: $("#sc-notes").value,
  } });
  await bumpPending();
  renderSchedule();
  flash($("#sc-msg"), "Added.", true);
  runSync();
}

// ---- Status (recent entries + quick toggles) ----
function renderStatus() {
  $("#view-status").replaceChildren(
    el("h2", {}, "Status"),
    el("p", { class: "hint" }, getAccount()
      ? "Open this tab to load recent entries."
      : "Recent entries and quick toggles appear here once you sign in."),
  );
}

async function loadStatus() {
  const v = $("#view-status");
  if (!isConfigured() || !AUTH_READY || !getAccount()) return renderStatus();
  if (!navigator.onLine) {
    return v.replaceChildren(el("h2", {}, "Status"), el("p", { class: "hint" }, "Offline — connect to see recent entries."));
  }
  v.replaceChildren(el("h2", {}, "Status"), el("p", { class: "hint" }, "Loading recent entries…"));
  try {
    renderStatusList(await recentTimeLog(20));
  } catch (e) {
    v.replaceChildren(el("h2", {}, "Status"), el("p", { class: "hint err" }, friendly(e)));
  }
}

function renderStatusList(rows) {
  const body = rows.length ? rows.map(rowCard) : [el("p", { class: "hint" }, "No entries yet.")];
  $("#view-status").replaceChildren(
    el("h2", {}, "Status"),
    el("p", { class: "hint" }, "Tap a chip to flip it — saves to the workbook."),
    el("div", { class: "cards" }, ...body),
    el("button", { class: "linkbtn", onclick: doSignOut }, "Sign out"),
  );
}

function rowCard(r) {
  const billOn = r.billable === "Y";
  const repOn = r.reportSubmitted === "Yes";
  const bill = el("button", { class: `chip ${billOn ? "on" : ""}` }, "Billable");
  bill.addEventListener("click", () => toggleChip(bill, r, "billable"));
  const rep = el("button", { class: `chip ${repOn ? "on" : ""}` }, "Report");
  rep.addEventListener("click", () => toggleChip(rep, r, "report"));
  return el("div", { class: "card" },
    el("div", { class: "card-top" },
      el("span", { class: "card-proj" }, r.project || "(no project)"),
      el("span", { class: "card-date" }, fmtDate(r.date))),
    el("div", { class: "card-sub" }, `${r.task || ""} · ${r.hours || 0} hr`),
    el("div", { class: "chips" }, bill, rep,
      r.invoice ? el("span", { class: "chip invoiced" }, `Invoiced ${r.invoice}`) : null),
  );
}

async function toggleChip(node, r, which) {
  const wasOn = node.classList.contains("on");
  const yes = !wasOn;
  node.classList.toggle("on", yes); // optimistic
  try {
    if (which === "billable") { await setBillable(r.row, yes); r.billable = yes ? "Y" : "N"; }
    else { await setReportSubmitted(r.row, yes); r.reportSubmitted = yes ? "Yes" : "No"; }
  } catch (e) {
    node.classList.toggle("on", wasOn); // revert
    alert(friendly(e));
  }
}

function fmtDate(serial) {
  if (serial == null || serial === "") return "";
  try { return serialToISO(serial); } catch { return String(serial); }
}

function flash(node, text, ok) {
  if (!node) return;
  node.textContent = text;
  node.className = "form-msg " + (ok ? "ok" : "err");
}

boot();
