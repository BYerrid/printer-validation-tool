// Import/export of the Master Tracker workbook, entirely in-browser via the
// vendored SheetJS build (window.XLSX). The original file bytes are kept
// as-is (base64 in localStorage) so export re-uses the same workbook -
// preserving the Approach Plan / Dashboard sheets, formulas, and formatting
// - and only overwrites the editable columns on Master Tracker.

import { COLS, SHEET_NAME, HEADER_ROW, recordId, inferLocation, emptyField } from "./model.js";

function arrayBufferToBase64(buf) {
  let binary = "";
  const bytes = new Uint8Array(buf);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function cell(ws, row1, col1) {
  const addr = XLSX.utils.encode_cell({ r: row1 - 1, c: col1 - 1 });
  const c = ws[addr];
  if (!c || c.v === undefined || c.v === null) return "";
  return c.v;
}

export async function parseWorkbookFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws["!ref"]);

  // Actual header text for each mapped column, e.g. "Found in 7/10/26 Scan"
  // or "ARP Scan 7/2026" - these dates are baked into the tracker's own
  // header row and change every time a fresh scan is merged in. Capturing
  // them here (instead of hardcoding generic labels) lets the UI always
  // show which dated scan/source a given lead came from without needing a
  // code change on every re-import.
  const headers = {};
  for (const [key, col] of Object.entries(COLS)) {
    headers[key] = cell(ws, HEADER_ROW, col) || "";
  }

  const records = [];
  for (let r = HEADER_ROW + 1; r <= range.e.r + 1; r++) {
    const rowHasData = Object.values(COLS).some((c) => cell(ws, r, c) !== "");
    if (!rowHasData) continue;

    const ref = {
      tier: cell(ws, r, COLS.tier),
      subTier: cell(ws, r, COLS.subTier),
      lpName: cell(ws, r, COLS.lpName),
      friendlyName: cell(ws, r, COLS.friendlyName),
      queueLocation: cell(ws, r, COLS.queueLocation),
      ip: cell(ws, r, COLS.ip),
      ipSource: cell(ws, r, COLS.ipSource),
      protocol: cell(ws, r, COLS.protocol),
      queueStatus: cell(ws, r, COLS.queueStatus),
      usedSince: cell(ws, r, COLS.usedSince),
      foundInScan: cell(ws, r, COLS.foundInScan),
      scanPrinterName: cell(ws, r, COLS.scanPrinterName),
      scanLocation: cell(ws, r, COLS.scanLocation),
      model: cell(ws, r, COLS.model),
      serial: cell(ws, r, COLS.serial),
      mac: cell(ws, r, COLS.mac),
      lifetimePages: cell(ws, r, COLS.lifetimePages),
      consoleMsg: cell(ws, r, COLS.consoleMsg),
      tyroneBldg: cell(ws, r, COLS.tyroneBldg),
      arpScan: cell(ws, r, COLS.arpScan),
      arpBldg: cell(ws, r, COLS.arpBldg),
      arpLocationDetail: cell(ws, r, COLS.arpLocationDetail),
      arpMacObserved: cell(ws, r, COLS.arpMacObserved),
    };

    const field = {
      physicallyLocated: cell(ws, r, COLS.physicallyLocated),
      verifiedBldgRoom: cell(ws, r, COLS.verifiedBldgRoom),
      assetTag: cell(ws, r, COLS.assetTag),
      verifiedUserDept: cell(ws, r, COLS.verifiedUserDept),
      businessNeedValidated: cell(ws, r, COLS.businessNeedValidated),
      decision: cell(ws, r, COLS.decision),
      migrationStatus: cell(ws, r, COLS.migrationStatus),
      date: cell(ws, r, COLS.date),
      tech: cell(ws, r, COLS.tech),
      notes: cell(ws, r, COLS.notes),
    };

    const id = recordId(ref, r);
    const computed = inferLocation(ref, field);
    records.push({ id, rowIndex: r, ref, field, computed });
  }

  const workbookB64 = arrayBufferToBase64(buf);
  return { records, workbookB64, sheetName, fileName: file.name, headers };
}

// Merge a freshly-imported set of records into existing ones: reference data
// is refreshed from the new file, but any field entries already made in the
// app (Physically Located, Notes, etc.) are preserved by id match.
export function mergeRecords(existingRecords, incomingRecords) {
  const existingById = new Map(existingRecords.map((r) => [r.id, r]));
  const merged = incomingRecords.map((incoming) => {
    const existing = existingById.get(incoming.id);
    if (!existing) return incoming;
    const field = { ...incoming.field };
    // keep any already-filled field values unless the new file also has a value
    for (const key of Object.keys(field)) {
      if (!field[key] && existing.field[key]) field[key] = existing.field[key];
    }
    const computed = incoming.computed.building === "Unassigned" && existing.computed.building !== "Unassigned"
      ? existing.computed
      : incoming.computed;
    return { ...incoming, field, computed };
  });
  return merged;
}

export function exportWorkbook(workbookB64, records) {
  const buf = base64ToArrayBuffer(workbookB64);
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const fieldColKeys = [
    "physicallyLocated", "verifiedBldgRoom", "assetTag", "verifiedUserDept",
    "businessNeedValidated", "decision", "migrationStatus", "date", "tech", "notes",
  ];

  for (const rec of records) {
    for (const key of fieldColKeys) {
      const col = COLS[key];
      const addr = XLSX.utils.encode_cell({ r: rec.rowIndex - 1, c: col - 1 });
      const value = rec.field[key] || "";
      if (value === "") {
        delete ws[addr];
      } else {
        ws[addr] = { t: "s", v: value };
      }
    }
  }

  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const stamp = new Date().toISOString().slice(0, 10);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Printer_Validation_Export_${stamp}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
