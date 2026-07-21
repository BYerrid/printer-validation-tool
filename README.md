# Printer Validation Tool

A mobile-first, installable web app (PWA) for walking buildings floor by
floor and validating a printer inventory against a spreadsheet — without
using the spreadsheet itself as the interface.

**Live app:** `https://<your-github-username>.github.io/printer-validation-tool/`

## How it works

1. **Import** — On first load, choose your "Master Tracker" spreadsheet
   (`.xlsx`). It's parsed entirely in your browser using
   [SheetJS](https://sheetjs.com) — the file is never uploaded anywhere.
2. **Walk** — The home screen groups printers by building, then by floor.
   Buildings/floors are inferred from whatever location data each row
   already has (ARP scan building, verified room, queue name, etc.). A
   "guess" tag marks buildings inferred from a department code rather than
   a confirmed building code — check those first.
3. **Validate** — Tap into a printer to see all its known "location leads"
   (queue location, scan location, ARP data, IP/MAC/serial) and fill in
   the same fields your spreadsheet tracks: Physically Located, Verified
   Building/Room, Asset Tag, Verified User/Dept, Business Need Validated,
   Decision, Migration Status, Date, Tech, Notes.
4. **Export** — From Settings, export an updated `.xlsx`. It's your
   original workbook with only the editable columns updated — the
   Approach Plan and Dashboard sheets (and its formulas) are untouched, so
   opening it in Excel recalculates the dashboard counts automatically.

## Data & privacy

- All data (your imported spreadsheet, every entry you make) is stored
  **only in your browser's local storage** on your device. Nothing is sent
  to a server — there is no backend.
- The GitHub repo backing this app contains **only application code**.
  Your spreadsheet, IPs, MACs, serials, and building/room data are never
  committed to the repo or included in the deployed site.
- Clearing your browser's site data (or using "Clear all local data" in
  Settings) permanently deletes your entries — export first.

## Installing on your phone

Open the live app URL in your phone's browser, then:

- **iPhone (Safari):** Share button → "Add to Home Screen"
- **Android (Chrome):** Menu (⋮) → "Add to Home screen" / "Install app"

Once installed, it launches full-screen like a regular app and keeps
working (using its last-loaded data) even with no signal.

## Re-importing an updated spreadsheet

If you get a refreshed export from the network/scan team, re-import it
from Settings. Reference data (IP, MAC, ARP scan results, etc.) is
refreshed from the new file; anything you've already filled in on the
device (Physically Located, Notes, etc.) is preserved.

## Local development

No build step — it's static HTML/CSS/JS. To run locally:

```
cd app
python -m http.server 8080
```

Then open `http://localhost:8080/`. (ES modules and the service worker
both require a real HTTP origin — opening `index.html` directly via
`file://` won't work.)

## Project structure

```
app/
  index.html            App shell
  manifest.webmanifest   PWA install metadata
  service-worker.js      Offline app-shell caching
  css/styles.css
  js/
    model.js             Column map, dropdown options, building/floor inference
    storage.js           localStorage persistence
    xlsxio.js            Import (.xlsx -> records) and export (records -> .xlsx)
    app.js               Screens/router/rendering
  vendor/xlsx.full.min.js  SheetJS (vendored, no CDN dependency)
  icons/
spreadsheet-review/      Drop spreadsheets here for review — gitignored, never committed
```
