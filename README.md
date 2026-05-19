# Dual Report Merge MyGeotab Add-In

This add-in lets you select a date range, groups, and exceptions, then pulls ExceptionsDetail and HOSLog reports and writes them into the "Data1" and "Data2" tabs of a provided Excel template. All other tabs in the template remain unchanged.

## Features
- MyGeotab-style UI (date range, group, exception selectors)
- Pulls ExceptionsDetail and HOSLog reports
- Writes data to "Data1" and "Data2" tabs in template.xlsx
- Leaves all other tabs untouched
- Exports a merged Excel file
- All files in repo root; GitHub Pages hosting

## Setup
1. Place your template as `template.xlsx` in the repo root (or use the sample provided).
2. Enable GitHub Pages (Settings > Pages > Source: GitHub Actions).
3. Register the add-in in MyGeotab using the URLs in `addin.sample.json`.

## Usage
- Open the add-in from MyGeotab.
- Select filters and settings.
- Click Generate Excel to export the merged file.

## Development
- All source files are in the repo root.
- To preview locally: `python -m http.server 8080`
- Open `http://localhost:8080` in your browser.

## License
MIT
