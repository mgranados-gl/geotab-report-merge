# Gridline | Gap Report

MyGeotab add-in that pulls yesterday's HOS logs and exception events into a single Excel workbook with filtering capabilities.

## Features

- **Report Date**: Always yesterday
- **Time Zone**: Automatically applies user's profile time zone (EST/EDT, etc.)
- **HOS Log Extraction**: Fetches DutyStatusLog records for **On, Drive, and Login/logout states only**
- **All Drivers**: No driver filtering needed; all drivers included
- **Exception Events**: Automatically filters for **Entering Zone (Office)** and **Exiting Zone (Office)** rules only
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
2. The report is pre-configured to run for yesterday with your time zone
3. Click "Run Report" to generate and download the Excel workbook
4. Workbook contains two sheets:
   - **Sheet 1**: HOS logs (On, Drive, Login/logout states only) for all drivers
   - **Sheet 2**: Exception events (Entering/Exiting Zone Office rules) for all drivers

No manual filtering required—all settings are fixed.

## Dependencies

- MyGeotab API (via `window.geotab`)
- SheetJS (XLSX) library loaded from CDN for Excel generation
