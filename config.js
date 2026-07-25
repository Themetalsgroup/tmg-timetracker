// config.js — central settings for the TMG Time Tracker PWA.
// Fill clientId/tenantId after the one-time Microsoft (Entra) app registration.

export const CONFIG = {
  // --- Microsoft 365 sign-in (from Entra "App registrations") ---
  clientId: "24df039c-58b0-40c3-98b0-568c3ca30acb",
  tenantId: "5f2ce972-43b9-44ba-8097-6cbcdd4de03e",
  scopes: ["User.Read", "Files.ReadWrite"],

  // --- Master workbook, relative to the signed-in user's OneDrive for Business root ---
  // The drive root IS the "Documents" library, so the leading /Documents/ below is the
  // real folder nested inside it — matching .../personal/<user>/Documents/Documents/1_Scheduling/.
  workbookPath: "/Documents/1_Scheduling/TMG_Time_Tracker_v2 1.xlsx",

  // --- Sheet names + the row ceilings baked into the workbook's formulas ---
  sheets: { timeLog: "Time Log", calendar: "Calendar", projects: "Projects" },
  timeLogMaxRow: 1003, // Invoice Export + formula range stop here — never write past it
  calendarMaxRow: 204,

  // Task types — mirrors the workbook's Task Type validation list on Time Log!D exactly,
  // in workbook order.
  //
  // These MUST match the workbook string-for-string. Graph PATCH bypasses Excel's data
  // validation, so a value that isn't on this list writes silently and lands invalid.
  // "NDT" was one such value: the workbook has no such entry, only the two Level rows —
  // and the Rates sheet keys NOVA's $140/$200 split on "NDT Level II" / "NDT Level III",
  // so a bare "NDT" would also miss the rate lookup.
  taskTypes: [
    "UT", "VT", "MT", "PT",
    "Shop Inspection", "Field Inspection", "Reinspection",
    "WQT", "WPS/PQR", "Consulting",
    "Structural Framing", "Bolting", "Skidmore",
    "Anchor Bolts", "Anchor Proof-Load", "Shear Studs", "Decking",
    "Drone/UAV", "Coatings",
    "NDT Level II", "NDT Level III",
    "Travel", "Office", "Training", "Other",
  ],
};

// True once the Microsoft IDs above have been filled in.
export const isConfigured = () =>
  !!CONFIG.clientId && !CONFIG.clientId.startsWith("PASTE_") &&
  !!CONFIG.tenantId && !CONFIG.tenantId.startsWith("PASTE_");

// Sample projects so the app is usable/demoable before Microsoft is wired up.
// Once signed in, the real (Active-only) Projects list replaces these.
export const SAMPLE_PROJECTS = [
  { name: "Tift County 7th & 8th", number: "TMG-26-014", client: "Cornerstone Engineering Consultants, Inc." },
  { name: "Archbold Ortho - Thomasville", number: "TMG-2026-006", client: "UES" },
  { name: "Jackson-Cook On-Call CSA", number: "TMG-2026-003", client: "Jackson-Cook, LC" },
  { name: "Crisp Regional - OR Suite", number: "TMG-2026-011", client: "UES" },
];
