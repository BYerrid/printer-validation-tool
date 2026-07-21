// All persistence is local to this browser (localStorage). Nothing here
// ever leaves the device - there is no server component.

const KEYS = {
  workbook: "pvt.originalWorkbookB64",
  records: "pvt.records",
  meta: "pvt.meta",
  settings: "pvt.settings",
};

export function hasData() {
  return !!localStorage.getItem(KEYS.records);
}

export function saveWorkbookB64(b64) {
  localStorage.setItem(KEYS.workbook, b64);
}

export function loadWorkbookB64() {
  return localStorage.getItem(KEYS.workbook);
}

export function saveRecords(records) {
  localStorage.setItem(KEYS.records, JSON.stringify(records));
}

export function loadRecords() {
  const raw = localStorage.getItem(KEYS.records);
  return raw ? JSON.parse(raw) : [];
}

export function saveMeta(meta) {
  localStorage.setItem(KEYS.meta, JSON.stringify(meta));
}

export function loadMeta() {
  const raw = localStorage.getItem(KEYS.meta);
  return raw ? JSON.parse(raw) : null;
}

export function saveSettings(settings) {
  localStorage.setItem(KEYS.settings, JSON.stringify(settings));
}

export function loadSettings() {
  const raw = localStorage.getItem(KEYS.settings);
  return raw ? JSON.parse(raw) : { techInitials: "" };
}

export function clearAll() {
  Object.values(KEYS).forEach((k) => localStorage.removeItem(k));
}
