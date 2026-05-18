# Geotab Dual Report Export Add-In

A clean, simple add-in that runs inside MyGeotab to:

1. **Pull two fixed reports** — HosLog and ExceptionsDetail
2. **Merge with an Excel template** — loads your template from GitHub Pages
3. **Replace data sheets** — Data1 and Data2 sheets get fresh data
4. **Export the final file** — download as a single Excel workbook

## How It Works

- Fetches your hosted Excel template from `https://mgranados-gl.github.io/geotab-report-merge/template/Gridline%20_%20Driver%20Events%20(Yesterday).xlsx`
- Pulls both reports using the same search criteria (date range, filters)
- Flattens all fields and replaces the `Data1` and `Data2` sheets in the template
- Keeps all other template sheets intact (summary, charts, etc.)
- Exports the merged workbook to download

## Setup

1. Host this folder on HTTPS (GitHub Pages recommended)
2. Register the hosted `index.html` as a MyGeotab add-in
3. Open the add-in from within MyGeotab
4. Configure filters or enter Search JSON
5. Click **Test** to verify, then **Pull Reports & Export Excel**

## Reports

- **Data1 sheet**: HosLog (or DutyStatusLog if HosLog unavailable)
- **Data2 sheet**: ExceptionsDetail

All fields from both reports are included.

## Customize

Edit `CONFIG` in [app.js](app.js):
- `templateUrl` — change the hosted template URL
- `reports[].typeName` — change report type names
- `reports[].sheetName` — change sheet names in output

## Notes

- Requires MyGeotab add-in session (no password needed)
- Buttons disabled if opened directly in browser
- HosLog auto-falls back to DutyStatusLog if unavailable
