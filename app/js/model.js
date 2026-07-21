// Data model: maps the "Master Tracker" sheet columns to app records,
// and infers a Building/Floor for navigation from whichever location
// clues are present on a row.

// 1-indexed column positions in the "Master Tracker" sheet (must match the
// workbook exactly, or export will write values into the wrong columns).
export const COLS = {
  tier: 1,
  subTier: 2,
  lpName: 3,
  friendlyName: 4,
  queueLocation: 5,
  ip: 6,
  ipSource: 7,
  protocol: 8,
  queueStatus: 9,
  usedSince: 10,
  foundInScan: 11,
  scanPrinterName: 12,
  scanLocation: 13,
  model: 14,
  serial: 15,
  mac: 16,
  lifetimePages: 17,
  consoleMsg: 18,
  tyroneBldg: 19,
  arpScan: 20,
  arpBldg: 21,
  arpLocationDetail: 22,
  arpMacObserved: 23,
  // Editable ("yellow") columns
  physicallyLocated: 24,
  verifiedBldgRoom: 25,
  assetTag: 26,
  verifiedUserDept: 27,
  businessNeedValidated: 28,
  decision: 29,
  migrationStatus: 30,
  date: 31,
  tech: 32,
  notes: 33,
};

export const SHEET_NAME = "Master Tracker";
export const HEADER_ROW = 1;

// Exact dropdown value sets from the workbook's data validation, so values
// written back round-trip cleanly through Excel's existing dropdowns.
export const OPTIONS = {
  physicallyLocated: ["", "Yes", "No", "Cannot Locate"],
  businessNeedValidated: ["", "Yes", "No", "Pending Dept Response"],
  decision: ["", "Migrate", "Decommission", "Replace", "Investigate"],
  migrationStatus: [
    "",
    "Not Started",
    "Compatibility Check",
    "AIS Setup",
    "PNAT Requested",
    "PNAT Complete",
    "Test Print OK",
    "Complete",
    "Removed",
  ],
};

// Known short building codes that show up in ARP Bldg / location text.
export const BUILDING_CODES = [
  "RCF", "HR", "DH", "CL", "AC", "KHH", "KC", "MH", "HW", "SET",
];

// Tyrone Bldg column uses full names; map to the same short codes where one
// exists so both sources land in the same bucket.
const TYRONE_TO_CODE = {
  "Academic Commons": "AC",
  "Duncan Hall": "DH",
  "Clark Library": "CL",
  "McCuan Hall": "MH",
  "Howard Hall": "HW",
  "RCF": "RCF",
  "SET": "SET",
};

const FLOOR_WORD_RE = /(\d+)(?:ST|ND|RD|TH)\s*FLOOR/i;

function findBuildingCodeIn(text) {
  if (!text) return null;
  const upper = String(text).toUpperCase();
  for (const code of BUILDING_CODES) {
    // word-ish boundary: code followed by digit, dash, space, or end of string
    const re = new RegExp(`\\b${code}(?=[\\s\\-\\d]|$)`);
    if (re.test(upper)) return code;
  }
  return null;
}

// When no known building code or Tyrone Bldg name is present, many Tier
// 2/3 queue names still encode a department/area code (e.g. "BO2 - Dr Ty
// Stone", "ADV13-Adm/Adv,J.Small", "Joan King - SSS2"). This is a much more
// useful walk grouping than dumping all of them into one "Unassigned"
// bucket - it's an unconfirmed guess, surfaced as such in the UI.
const BRAND_DENYLIST = new Set([
  "HP", "EPSON", "CANON", "XEROX", "LEXMARK", "BROTHER", "KYOCERA", "RICOH", "DELL", "OKI",
]);
const DEPT_CODE_RE = /\b([A-Z]{2,6})\d{1,4}\b/;

function findDeptCodeIn(text) {
  if (!text) return null;
  const m = String(text).toUpperCase().match(DEPT_CODE_RE);
  if (!m) return null;
  if (BRAND_DENYLIST.has(m[1])) return null;
  return m[1];
}

function guessFloorIn(text) {
  if (!text) return null;
  const m = String(text).match(FLOOR_WORD_RE);
  if (m) return m[1];
  // fall back to leading digit of a 3-digit room number, e.g. RCF-208 -> 2
  const roomMatch = String(text).match(/\b(\d)\d{2}\b/);
  if (roomMatch) return roomMatch[1];
  return null;
}

// Building/floor are inferred with this priority, and re-run any time a
// record's Verified Bldg/Room is edited so a manual correction always wins.
export function inferLocation(ref, field) {
  // 1. A manually verified room always wins - that's the whole point of the walk.
  if (field.verifiedBldgRoom) {
    const code = findBuildingCodeIn(field.verifiedBldgRoom);
    const building = code || field.verifiedBldgRoom;
    const floor = guessFloorIn(field.verifiedBldgRoom) || "?";
    return { building, floor, buildingConfirmed: true };
  }

  // 2. Known building code from ARP scan or location text (confirmed).
  const codeSources = [ref.arpBldg, ref.arpLocationDetail, ref.scanLocation, ref.queueLocation, ref.friendlyName];
  for (const s of codeSources) {
    if (!s) continue;
    const code = findBuildingCodeIn(s);
    if (code) {
      const floor = guessFloorIn(field.verifiedBldgRoom) || guessFloorIn(ref.arpLocationDetail)
        || guessFloorIn(ref.scanLocation) || guessFloorIn(ref.queueLocation) || "?";
      return { building: code, floor, buildingConfirmed: true };
    }
  }

  // 3. Tyrone Bldg full name, with no short code of its own (confirmed).
  if (ref.tyroneBldg) {
    const building = TYRONE_TO_CODE[ref.tyroneBldg] || ref.tyroneBldg;
    return { building, floor: "?", buildingConfirmed: true };
  }

  // 4. Department/area code buried in the queue name (unconfirmed guess -
  // this is the best lead available pre-walk for most Tier 2/3 rows).
  for (const s of [ref.queueLocation, ref.friendlyName, ref.scanLocation]) {
    const dept = findDeptCodeIn(s);
    if (dept) return { building: dept, floor: "?", buildingConfirmed: false };
  }

  return { building: "Unassigned", floor: "?", buildingConfirmed: false };
}

export function recordId(ref, rowIndex) {
  if (ref.lpName) return `lp:${ref.lpName}`;
  if (ref.ip || ref.mac) return `nm:${ref.ip || ""}|${ref.mac || ""}`;
  return `row:${rowIndex}`;
}

export function emptyField() {
  return {
    physicallyLocated: "",
    verifiedBldgRoom: "",
    assetTag: "",
    verifiedUserDept: "",
    businessNeedValidated: "",
    decision: "",
    migrationStatus: "",
    date: "",
    tech: "",
    notes: "",
  };
}

export function isTouched(field) {
  return Object.values(field).some((v) => v && String(v).trim() !== "");
}

export function tierLabel(tier) {
  if (tier === "NEW") return "NEW";
  return String(tier ?? "");
}

export function tierClass(tier) {
  const t = String(tier ?? "").toUpperCase();
  if (t === "1") return "tier-1";
  if (t === "2") return "tier-2";
  if (t === "3") return "tier-3";
  if (t === "NEW") return "tier-new";
  return "tier-unknown";
}
