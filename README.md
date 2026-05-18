# Geotab Dual Report Export Add-In

This add-in runs inside MyGeotab and uses the current signed-in user session (`api.call`) to:

1. Pull two fixed report datasets.
2. Export each to a separate Excel sheet.
3. No merging — just data placement on separate tabs.

## Reports

- **Report 1 (Data1 sheet)**: `HosLog` custom report (falls back to `DutyStatusLog` if unavailable)
- **Report 2 (Data2 sheet)**: `ExceptionsDetail` custom report

Both reports are pulled with the same search criteria (date range, filters, etc.) and placed on their own tabs with all fields included.

## Setup

1. Host this folder on HTTPS.
2. Register the hosted `index.html` as a MyGeotab add-in.
3. Open the add-in from MyGeotab (not directly in browser).
4. Pick date range and output filename.
5. Click **Test (1 row per report)**, then **Pull Reports & Export Excel**.

## Important

- This build does **not** use manual username/password fields.
- If you open `index.html` directly, buttons are disabled by design.
- If `HosLog` fails in your database, it automatically falls back to `DutyStatusLog`.

## Customize

Edit `REPORT_CONFIG` in [app.js](app.js) to:

- Change sheet names (`sheetName` properties)
- Change report type names (`typeName`, `fallbackTypeName`)
- Modify search criteria via the UI's filter builder
