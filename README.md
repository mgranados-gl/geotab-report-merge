# Geotab Report Merge Add-In

Clean restart version. This add-in runs inside MyGeotab and uses the current signed-in user session (`api.call`) to:

1. Pull two fixed report datasets.
2. Merge them on configured keys.
3. Export one Excel file.

## Fixed Reports (from your links)

- Report A: `HosLog` custom report `id:b13C`, `scheduleId:bAA8`, `source:HosLog`
- Report B: `ExceptionsDetail` custom report `id:b13B`, `scheduleId:bAA6`, `source:ExceptionsDetail`
- Join: `driver.id` to `driver.id` (`left` join)

## Setup

1. Host this folder on HTTPS.
2. Register the hosted `index.html` as a MyGeotab add-in.
3. Open the add-in from MyGeotab (not directly in browser).
4. Pick date range and output filename.
5. Click **Test Connection**, then **Run Merge & Export Excel**.

## Important

- This build does **not** use manual username/password fields.
- If you open `index.html` directly, buttons are disabled by design.
- If `HosLog` fails in your database, it falls back to `DutyStatusLog` automatically.

## Customize

Edit `REPORT_MERGE_CONFIG` in [app.js](app.js):

- `typeNames`
- `keyPath`
- `fields`
- `buildSearch(fromIso, toIso)`
- `joinType`
