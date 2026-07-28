import * as storage from "./storage.js";
import { parseWorkbookFile, mergeRecords, exportWorkbook } from "./xlsxio.js";
import { OPTIONS, inferLocation, tierClass, tierLabel, isTouched } from "./model.js";

const state = {
  records: [],
  meta: null,
  settings: { techInitials: "" },
};

const root = document.getElementById("app");

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function statusOf(rec) {
  const f = rec.field;
  if (f.physicallyLocated === "Cannot Locate") return { cls: "st-red", label: "Cannot locate", mark: "✕", tint: "red" };
  if (f.physicallyLocated === "No") return { cls: "st-red", label: "Not located", mark: "✕", tint: "red" };
  if (f.physicallyLocated === "Yes" && f.businessNeedValidated) return { cls: "st-green", label: "Validated", mark: "✓", tint: "green" };
  if (f.physicallyLocated === "Yes") return { cls: "st-amber", label: "Located", mark: "✓", tint: "amber" };
  if (isTouched(f)) return { cls: "st-amber", label: "In progress", mark: "", tint: "amber" };
  return { cls: "st-grey", label: "Not checked", mark: "", tint: "" };
}

// Small fixed toast for immediate save confirmation. Lives on document.body
// (a sibling of #app) rather than inside root.innerHTML, so it survives the
// full re-render that happens when saveDetail() navigates back to the list.
function showToast(message, tint) {
  let el = document.getElementById("toast");
  if (!el) {
    el = document.createElement("div");
    el.id = "toast";
    document.body.appendChild(el);
  }
  el.textContent = message;
  el.className = `toast show ${tint ? `toast-${tint}` : ""}`;
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => el.classList.remove("show"), 1800);
}

function displayName(rec) {
  return rec.ref.friendlyName || rec.ref.lpName || rec.ref.model || rec.ref.ip || "(unnamed)";
}

// ---------- data loading ----------

function loadState() {
  state.records = storage.loadRecords();
  state.meta = storage.loadMeta();
  state.settings = storage.loadSettings();
}

function persistRecords() {
  storage.saveRecords(state.records);
}

// ---------- grouping ----------

const BUILDING_SORT_PRIORITY = ["RCF", "HR", "DH", "CL", "AC", "KHH", "KC", "MH", "HW", "SET"];

function sortBuildings(names) {
  return names.sort((a, b) => {
    const ai = BUILDING_SORT_PRIORITY.indexOf(a);
    const bi = BUILDING_SORT_PRIORITY.indexOf(b);
    if (a === "Unassigned") return 1;
    if (b === "Unassigned") return -1;
    if (ai !== -1 && bi !== -1) return ai - bi;
    if (ai !== -1) return -1;
    if (bi !== -1) return 1;
    return a.localeCompare(b);
  });
}

function groupByBuilding(records) {
  const map = new Map();
  for (const r of records) {
    const b = r.computed.building || "Unassigned";
    if (!map.has(b)) map.set(b, []);
    map.get(b).push(r);
  }
  return map;
}

function groupByFloor(records) {
  const map = new Map();
  for (const r of records) {
    const f = r.computed.floor || "?";
    if (!map.has(f)) map.set(f, []);
    map.get(f).push(r);
  }
  return new Map([...map.entries()].sort((a, b) => {
    if (a[0] === "?") return 1;
    if (b[0] === "?") return -1;
    return Number(a[0]) - Number(b[0]);
  }));
}

// ---------- routing ----------

function currentRoute() {
  const hash = location.hash.replace(/^#/, "") || "/";
  const parts = hash.split("/").filter(Boolean);
  return parts;
}

function navigate(path) {
  location.hash = path;
}

window.addEventListener("hashchange", render);
window.addEventListener("DOMContentLoaded", init);

function init() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }
  loadState();
  render();
}

function render() {
  if (!storage.hasData()) {
    renderImport();
    return;
  }
  loadState();
  const parts = currentRoute();

  if (parts[0] === "settings") return renderSettings();
  if (parts[0] === "search") return renderSearch();
  if (parts[0] === "printer" && parts[1]) return renderDetail(decodeURIComponent(parts[1]));
  if (parts[0] === "building" && parts[1] && parts[2] === "floor" && parts[3]) {
    return renderPrinterList(decodeURIComponent(parts[1]), decodeURIComponent(parts[3]));
  }
  if (parts[0] === "building" && parts[1]) return renderFloors(decodeURIComponent(parts[1]));
  return renderHome();
}

// ---------- screen: import ----------

function renderImport() {
  root.innerHTML = `
    <div class="screen import-screen">
      <div class="import-card">
        <div class="import-icon">🖨️</div>
        <h1>Printer Validation</h1>
        <p class="muted">Import your Master Tracker spreadsheet to get started. It's parsed right here in your browser &mdash; the file never leaves this device.</p>
        <label class="btn btn-primary file-btn">
          Choose spreadsheet (.xlsx)
          <input type="file" id="file-input" accept=".xlsx" hidden />
        </label>
        <p class="muted small">Expects a "Master Tracker" sheet matching your existing tracker's columns.</p>
      </div>
    </div>
  `;
  document.getElementById("file-input").addEventListener("change", onImportFile);
}

async function onImportFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  root.innerHTML = `<div class="screen"><p class="loading">Parsing ${escapeHtml(file.name)}&hellip;</p></div>`;
  try {
    const { records, workbookB64, sheetName, fileName, headers } = await parseWorkbookFile(file);
    const existing = storage.hasData() ? storage.loadRecords() : [];
    const merged = existing.length ? mergeRecords(existing, records) : records;
    storage.saveWorkbookB64(workbookB64);
    storage.saveRecords(merged);
    storage.saveMeta({ importedAt: new Date().toISOString(), sheetName, fileName, rowCount: merged.length, headers });
    navigate("/");
    render();
  } catch (err) {
    root.innerHTML = `
      <div class="screen">
        <p class="error">Couldn't read that file: ${escapeHtml(err.message || String(err))}</p>
        <button class="btn" onclick="location.reload()">Try again</button>
      </div>`;
  }
}

// ---------- screen: home ----------

function renderHome() {
  const total = state.records.length;
  const located = state.records.filter((r) => r.field.physicallyLocated === "Yes").length;
  const validated = state.records.filter((r) => r.field.businessNeedValidated === "Yes").length;
  const cannotLocate = state.records.filter((r) => r.field.physicallyLocated === "Cannot Locate").length;

  const byBuilding = groupByBuilding(state.records);
  const buildingNames = sortBuildings([...byBuilding.keys()]);

  const buildingCards = buildingNames.map((b) => {
    const recs = byBuilding.get(b);
    const bLocated = recs.filter((r) => r.field.physicallyLocated === "Yes" || r.field.physicallyLocated === "Cannot Locate").length;
    const pct = Math.round((bLocated / recs.length) * 100);
    const unconfirmed = b !== "Unassigned" && recs.some((r) => !r.computed.buildingConfirmed);
    return `
      <a class="card building-card" href="#/building/${encodeURIComponent(b)}">
        <div class="building-card-top">
          <span class="building-name">${escapeHtml(b)}${unconfirmed ? ' <span class="unconfirmed-tag">guess</span>' : ""}</span>
          <span class="building-count">${recs.length}</span>
        </div>
        <div class="progress-bar"><div class="progress-fill" style="width:${pct}%"></div></div>
        <div class="building-sub muted small">${bLocated} of ${recs.length} checked</div>
      </a>
    `;
  }).join("");

  root.innerHTML = `
    <div class="screen">
      ${renderTopBar("Printer Validation", true)}
      <div class="content">
        <div class="stat-grid">
          <div class="stat"><div class="stat-num">${total}</div><div class="stat-label">Total</div></div>
          <div class="stat"><div class="stat-num">${located}</div><div class="stat-label">Located</div></div>
          <div class="stat"><div class="stat-num">${validated}</div><div class="stat-label">Need validated</div></div>
          <div class="stat"><div class="stat-num">${cannotLocate}</div><div class="stat-label">Cannot locate</div></div>
        </div>
        <h2 class="section-title">Buildings</h2>
        <div class="card-list">${buildingCards}</div>
      </div>
      ${renderBottomNav("home")}
    </div>
  `;
}

// ---------- screen: floors within a building ----------

function renderFloors(building) {
  const recs = state.records.filter((r) => (r.computed.building || "Unassigned") === building);
  const byFloor = groupByFloor(recs);

  if (byFloor.size <= 1) {
    const only = [...byFloor.keys()][0] ?? "?";
    navigate(`/building/${encodeURIComponent(building)}/floor/${encodeURIComponent(only)}`);
    return;
  }

  const cards = [...byFloor.entries()].map(([floor, floorRecs]) => {
    const checked = floorRecs.filter((r) => r.field.physicallyLocated === "Yes" || r.field.physicallyLocated === "Cannot Locate").length;
    const label = floor === "?" ? "Unknown floor" : `Floor ${floor}`;
    return `
      <a class="card building-card" href="#/building/${encodeURIComponent(building)}/floor/${encodeURIComponent(floor)}">
        <div class="building-card-top">
          <span class="building-name">${escapeHtml(label)}</span>
          <span class="building-count">${floorRecs.length}</span>
        </div>
        <div class="building-sub muted small">${checked} of ${floorRecs.length} checked</div>
      </a>
    `;
  }).join("");

  root.innerHTML = `
    <div class="screen">
      ${renderTopBar(building, false, "#/")}
      <div class="content">
        <h2 class="section-title">Floors</h2>
        <div class="card-list">${cards}</div>
      </div>
      ${renderBottomNav()}
    </div>
  `;
}

// ---------- screen: printer list for building+floor ----------

function renderPrinterList(building, floor) {
  const recs = state.records
    .filter((r) => (r.computed.building || "Unassigned") === building && (r.computed.floor || "?") === floor)
    .sort((a, b) => displayName(a).localeCompare(displayName(b)));

  const backHref = `#/building/${encodeURIComponent(building)}`;
  const label = floor === "?" ? "Unknown floor" : `Floor ${floor}`;

  const rows = recs.map((r) => printerRowHtml(r)).join("") || `<p class="muted">No printers in this group.</p>`;

  root.innerHTML = `
    <div class="screen">
      ${renderTopBar(`${building} &middot; ${label}`, false, backHref)}
      <div class="content">
        <div class="printer-list">${rows}</div>
      </div>
      ${renderBottomNav()}
    </div>
  `;
}

function printerRowHtml(r) {
  const st = statusOf(r);
  const name = displayName(r);
  // IP is what techs actually need on a walk (it's what ties back to the
  // rest of the inventory via asset tag); it must never be the thing that
  // gets truncated. So it's the fixed-width primary line and the
  // name/model - which vary wildly in length - are the part that wraps to
  // the muted sub-line and truncates if it must.
  const ip = r.ref.ip || "";
  const sub = [name, r.ref.model].filter(Boolean).map(escapeHtml).join(" &middot; ");
  const rowTint = st.tint ? ` printer-row--${st.tint}` : "";
  return `
    <a class="printer-row${rowTint}" href="#/printer/${encodeURIComponent(r.id)}">
      <span class="status-dot ${st.cls}" title="${st.label}">${st.mark}</span>
      <span class="printer-row-main">
        <span class="printer-row-name${ip ? " printer-row-ip" : ""}">${ip ? escapeHtml(ip) : escapeHtml(name)}</span>
        ${ip ? `<span class="printer-row-sub muted small">${sub}</span>` : ""}
      </span>
      <span class="tier-chip ${tierClass(r.ref.tier)}">${escapeHtml(tierLabel(r.ref.tier))}</span>
    </a>
  `;
}

// ---------- screen: printer detail / form ----------

function renderDetail(id) {
  const rec = state.records.find((r) => r.id === id);
  if (!rec) {
    root.innerHTML = `<div class="screen"><p class="error">Printer not found.</p><button class="btn" onclick="location.hash='/'">Home</button></div>`;
    return;
  }

  const ref = rec.ref;
  const f = rec.field;

  // Use the tracker's actual header text where we have it (it carries the
  // scan date, e.g. "Found in 7/10/26 Scan", "ARP Scan 7/2026") so it's
  // obvious in the field whether a lead is from the latest walk-through or
  // a stale one - without hardcoding a date that goes wrong on next import.
  const hdr = state.meta?.headers || {};
  const label = (key, fallback) => hdr[key] || fallback;

  const leadRows = [
    ["Tier / Sub-tier", `${tierLabel(ref.tier)} &mdash; ${escapeHtml(ref.subTier)}`, false],
    // IP/MAC first and visually emphasized - this is the field that's
    // actually needed on a walk (asset tag ties everything else to it).
    ["IP / MAC", [ref.ip, ref.mac].filter(Boolean).map(escapeHtml).join(" &middot; "), true],
    ["Model / Serial", [ref.model, ref.serial].filter(Boolean).map(escapeHtml).join(" &middot; "), false],
    [label("queueLocation", "Queue location"), escapeHtml(ref.queueLocation), false],
    [label("scanLocation", "Scan location"), escapeHtml(ref.scanLocation), false],
    [label("foundInScan", "Found in scan"), escapeHtml(ref.foundInScan), false],
    ["ARP building / detail", [ref.arpBldg, ref.arpLocationDetail].filter(Boolean).map(escapeHtml).join(" &middot; "), false],
    [label("arpScan", "ARP scan"), escapeHtml(ref.arpScan), false],
    ["Tyrone building", escapeHtml(ref.tyroneBldg), false],
    ["Lifetime pages", escapeHtml(ref.lifetimePages), false],
    [label("consoleMsg", "Console message"), escapeHtml(ref.consoleMsg), false],
  ].filter(([, v]) => v && v !== "");

  const leadsHtml = leadRows.map(([k, v, emphasize]) => `
    <div class="lead-row${emphasize ? " lead-row-primary" : ""}"><span class="lead-key">${k}</span><span class="lead-val">${v}</span></div>
  `).join("");

  root.innerHTML = `
    <div class="screen">
      ${renderTopBar(displayName(rec), false, `#/building/${encodeURIComponent(rec.computed.building)}/floor/${encodeURIComponent(rec.computed.floor)}`)}
      <div class="content">
        <details class="leads" open>
          <summary>Location leads &amp; reference data</summary>
          ${leadsHtml || '<p class="muted small">No reference data on this row.</p>'}
        </details>

        <form id="detail-form" class="detail-form">
          <label>Physically Located
            <select name="physicallyLocated">${selectOptions(OPTIONS.physicallyLocated, f.physicallyLocated)}</select>
          </label>

          <div class="quick-actions">
            <button type="button" class="btn btn-small btn-danger" data-quick="cannot-locate">Cannot locate</button>
            <button type="button" class="btn btn-small btn-ok" data-quick="found">Found here</button>
          </div>

          <label>Verified Building / Room
            <input type="text" name="verifiedBldgRoom" value="${escapeHtml(f.verifiedBldgRoom)}" placeholder="${escapeHtml(guessLocationString(rec) || "e.g. RCF-208")}" />
          </label>

          <label>Asset Tag
            <input type="text" name="assetTag" value="${escapeHtml(f.assetTag)}" />
          </label>

          <label>Verified User / Dept
            <input type="text" name="verifiedUserDept" value="${escapeHtml(f.verifiedUserDept)}" />
          </label>

          <label>Business Need Validated
            <select name="businessNeedValidated">${selectOptions(OPTIONS.businessNeedValidated, f.businessNeedValidated)}</select>
          </label>

          <label>Decision
            <select name="decision">${selectOptions(OPTIONS.decision, f.decision)}</select>
          </label>

          <label>Migration Status
            <select name="migrationStatus">${selectOptions(OPTIONS.migrationStatus, f.migrationStatus)}</select>
          </label>

          <label>Date
            <input type="date" name="date" value="${escapeHtml(f.date || "")}" />
          </label>

          <label>Tech
            <input type="text" name="tech" value="${escapeHtml(f.tech || state.settings.techInitials || "")}" />
          </label>

          <label>Notes
            <textarea name="notes" rows="3">${escapeHtml(f.notes)}</textarea>
          </label>

          <button type="submit" class="btn btn-primary btn-save">Save</button>
        </form>
      </div>
      ${renderBottomNav()}
    </div>
  `;

  const form = document.getElementById("detail-form");
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    saveDetail(rec, new FormData(form));
  });
  form.querySelectorAll("[data-quick]").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.dataset.quick === "cannot-locate") {
        form.physicallyLocated.value = "Cannot Locate";
      } else if (btn.dataset.quick === "found") {
        form.physicallyLocated.value = "Yes";
        if (!form.date.value) form.date.value = todayISO();
      }
    });
  });
}

function guessLocationString(rec) {
  const b = rec.computed.building;
  const f = rec.computed.floor;
  if (!b || b === "Unassigned") return "";
  if (!f || f === "?") return `e.g. ${b}-208`;
  return `e.g. ${b}-${f}08 (floor ${f})`;
}

function selectOptions(options, current) {
  return options.map((o) => {
    const label = o === "" ? "&mdash; not set &mdash;" : escapeHtml(o);
    const selected = o === (current || "") ? "selected" : "";
    return `<option value="${escapeHtml(o)}" ${selected}>${label}</option>`;
  }).join("");
}

function saveDetail(rec, formData) {
  const fields = [
    "physicallyLocated", "verifiedBldgRoom", "assetTag", "verifiedUserDept",
    "businessNeedValidated", "decision", "migrationStatus", "date", "tech", "notes",
  ];
  for (const key of fields) {
    rec.field[key] = (formData.get(key) || "").toString().trim();
  }
  if (!rec.field.date && isTouched(rec.field)) rec.field.date = todayISO();
  if (rec.field.tech) {
    state.settings.techInitials = rec.field.tech;
    storage.saveSettings(state.settings);
  }
  rec.computed = inferLocation(rec.ref, rec.field);
  persistRecords();
  const st = statusOf(rec);
  showToast(`Saved — ${st.label}`, st.tint);
  history.back();
}

// ---------- screen: search ----------

function renderSearch() {
  root.innerHTML = `
    <div class="screen">
      ${renderTopBar("Search", false, "#/")}
      <div class="content">
        <input type="search" id="search-input" class="search-input" placeholder="Name, IP, serial, model, room&hellip;" autofocus />
        <div class="printer-list" id="search-results"></div>
      </div>
      ${renderBottomNav()}
    </div>
  `;
  const input = document.getElementById("search-input");
  const results = document.getElementById("search-results");
  input.addEventListener("input", () => {
    const q = input.value.trim().toLowerCase();
    if (!q) { results.innerHTML = ""; return; }
    const matches = state.records.filter((r) => {
      const hay = [
        r.ref.friendlyName, r.ref.lpName, r.ref.model, r.ref.serial, r.ref.ip,
        r.ref.mac, r.ref.queueLocation, r.ref.scanLocation, r.ref.arpLocationDetail,
        r.field.verifiedBldgRoom, r.field.assetTag, r.field.verifiedUserDept,
      ].filter(Boolean).join(" ").toLowerCase();
      return hay.includes(q);
    }).slice(0, 50);
    results.innerHTML = matches.map(printerRowHtml).join("") || `<p class="muted">No matches.</p>`;
  });
}

// ---------- screen: settings ----------

function renderSettings() {
  root.innerHTML = `
    <div class="screen">
      ${renderTopBar("Settings", false, "#/")}
      <div class="content">
        <label>Your initials (auto-filled as Tech)
          <input type="text" id="tech-initials" value="${escapeHtml(state.settings.techInitials || "")}" />
        </label>

        <h2 class="section-title">Data</h2>
        <p class="muted small">
          Imported ${escapeHtml(state.meta?.fileName || "")}${state.meta ? ` on ${new Date(state.meta.importedAt).toLocaleString()}` : ""}
          &mdash; ${state.records.length} rows.
        </p>
        <button class="btn btn-primary" id="export-btn">Export to Excel</button>

        <h2 class="section-title">Re-import spreadsheet</h2>
        <p class="muted small">Re-importing merges in new/updated reference data but keeps anything you've already filled in.</p>
        <label class="btn file-btn">
          Choose spreadsheet (.xlsx)
          <input type="file" id="reimport-input" accept=".xlsx" hidden />
        </label>

        <h2 class="section-title">Danger zone</h2>
        <button class="btn btn-danger" id="clear-btn">Clear all local data</button>
        <p class="muted small">This only clears data in this browser. Export first if you haven't already.</p>
      </div>
      ${renderBottomNav()}
    </div>
  `;

  document.getElementById("tech-initials").addEventListener("change", (e) => {
    state.settings.techInitials = e.target.value.trim();
    storage.saveSettings(state.settings);
  });
  document.getElementById("export-btn").addEventListener("click", () => {
    exportWorkbook(storage.loadWorkbookB64(), state.records);
  });
  document.getElementById("reimport-input").addEventListener("change", onImportFile);
  document.getElementById("clear-btn").addEventListener("click", () => {
    if (confirm("Clear all locally-stored data? This cannot be undone. Export first if you need a copy.")) {
      storage.clearAll();
      navigate("/");
      render();
    }
  });
}

// ---------- shared chrome ----------

function renderTopBar(title, showSearch, backHref) {
  const back = backHref ? `<a class="icon-btn" href="${backHref}">&larr;</a>` : `<span class="icon-btn spacer"></span>`;
  const search = showSearch ? `<a class="icon-btn" href="#/search">&#128269;</a>` : `<span class="icon-btn spacer"></span>`;
  return `
    <div class="top-bar">
      ${back}
      <h1 class="top-bar-title">${title}</h1>
      ${search}
    </div>
  `;
}

function renderBottomNav() {
  const parts = currentRoute();
  const isHome = parts.length === 0;
  const isSettings = parts[0] === "settings";
  return `
    <nav class="bottom-nav">
      <a class="nav-btn ${isHome ? "active" : ""}" href="#/">Buildings</a>
      <a class="nav-btn" href="#/search">Search</a>
      <a class="nav-btn ${isSettings ? "active" : ""}" href="#/settings">Settings</a>
    </nav>
  `;
}
