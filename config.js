// config.js — central settings for the TMG Time Tracker PWA.
// Fill clientId/tenantId after the one-time Microsoft (Entra) app registration.

export const CONFIG = {
  // --- Microsoft 365 sign-in (from Entra "App registrations") ---
  clientId: "24df039c-58b0-40c3-98b0-568c3ca30acb",
  tenantId: "5f2ce972-43b9-44ba-8097-6cbcdd4de03e",
  scopes: ["User.Read", "Files.ReadWrite"],

  // --- Master workbook, relative to the signed-in user's OneDrive for Business root ---
  workbookPath: "/Documents/1_Scheduling/TMG_Time_Tracker_v2 1.xlsx",

  // --- Sheet names + the row ceilings baked into the workbook's formulas ---
  sheets: { timeLog: "Time Log", calendar: "Calendar", projects: "Projects" },
  timeLogMaxRow: 1003, // Invoice Export + formula range stop here — never write past it
  calendarMaxRow: 204,

  // Task types — mirror the workbook's Task Type validation list.
  taskTypes: [
    "Field Inspection", "Reinspection", "NDT", "WQT", "WPS/PQR",
    "Consulting", "Travel", "Office", "Training", "Other",
  ],
};

// True once the Microsoft IDs above have been filled in.
export const isConfigured = () =>
  !!CONFIG.clientId && !CONFIG.clientId.startsWith("PASTE_") &&
  !!CONFIG.tenantId && !CONFIG.tenantId.startsWith("PASTE_");

// Sample projects so the app is usable/demoable before Microsoft is wired up.
// Once signed in, the real Projects list replaces these (cached in IndexedDB).
export const SAMPLE_PROJECTS = [
  { name: "Tift County 7th & 8th", number: "TMG-26-014", client: "Cornerstone Engineering Consultants, Inc." },
  { name: "VLD General Aviation Terminal", number: "TMG-2026-009", client: "Cornerstone Engineering Consultants, Inc." },
  { name: "Archbold Orthopedics", number: "TMG-2026-006", client: "UES" },
  { name: "Jackson-Cook On-Call CSA", number: "TMG-2026-003", client: "Jackson-Cook, LC" },
];
