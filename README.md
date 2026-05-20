# Gridline | Gap Report

MyGeotab add-in that pulls yesterday's HOS logs and exception events into a single Excel workbook with filtering capabilities.

## Features

- **HOS Log Extraction**: Fetches DutyStatusLog records from the previous day
- **Exception Event Filtering**: Pulls exception events and allows filtering by exception rules
- **Multi-level Filtering**: Filter results by drivers, vehicles, and groups
- **Excel Export**: Generates a two-sheet workbook (HOS logs and exceptions) with automatic download
- **Real-time Status**: Live status updates during data fetching and processing

## Files

- `index.html`: add-in page shell with filters and status display
- `app.js`: add-in registration (`window.geotab.addin.gap_report`) and business logic
- `styles.css`: responsive styling
- `addin.sample.json`: sample add-in manifest for MyGeotab configuration

## Host on GitHub Pages

1. Push this repository to GitHub.
2. In repository settings, enable GitHub Pages from the `main` branch (root).
3. Verify the site URL is:
   - `https://mgranados-gl.github.io/geotab-report-merge/`
4. Copy configuration from `addin.sample.json` into MyGeotab Add-In settings.

## Usage

1. Open the add-in in MyGeotab (must be authenticated)
2. Select exception rules, drivers, vehicles, and/or groups to filter by
3. Click "Run Report" to generate and download the Excel workbook
4. Workbook contains two sheets:
   - **Sheet 1**: HOS logs for the previous day
   - **Sheet 2**: Exception events for the previous day

## Dependencies

- MyGeotab API (via `window.geotab`)
- SheetJS (XLSX) library loaded from CDN for Excel generation
