/* ── Geotab Add-In: Dual Report Export with Template ── */

const runtime = { api: null };

const CONFIG = {
  templateUrl: "https://mgranados-gl.github.io/geotab-report-merge/template/Gridline%20_%20Driver%20Events%20(Yesterday).xlsx",
  reports: [
    { typeName: "HosLog", fallback: "DutyStatusLog", sheetName: "Data1" },
    { typeName: "ExceptionsDetail", sheetName: "Data2" }
  ]
};

const ui = {};

// ── UI State Management ────────────────────────────────────────

function setStatus(message) {
  const lines = Array.isArray(message) ? message : [message];
  ui.statusEl.textContent = lines.join("\n");
}

function setBusy(busy) {
  ui.testBtn.disabled = busy;
  ui.exportBtn.disabled = busy;
  if (ui.loadFilterBtn) ui.loadFilterBtn.disabled = busy;
  if (ui.buildSearchBtn) ui.buildSearchBtn.disabled = busy;
  if (ui.clearFilterBtn) ui.clearFilterBtn.disabled = busy;
}

// ── Date/Time Utilities ────────────────────────────────────────

function toInputDateValue(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}T${pad(date.getUTCHours())}:${pad(date.getUTCMinutes())}`;
}

function toIsoDate(localDateTimeInput) {
  if (!localDateTimeInput) return null;
  const dt = new Date(localDateTimeInput);
  return isNaN(dt.getTime()) ? null : dt.toISOString();
}

function applyDatePreset() {
  const now = new Date();
  let start = new Date(now);
  let end = new Date(now);

  switch (ui.presetSelect.value) {
    case "today":
      start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 0, 0, 0));
      end = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 23, 59, 59));
      break;
    case "yesterday":
      start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - 1, 0, 0, 0));
      end = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - 1, 23, 59, 59));
      break;
    case "last7":
      start = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      break;
    case "last30":
      start = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
      break;
    default:
      return;
  }

  ui.fromDateInput.value = toInputDateValue(start);
  ui.toDateInput.value = toInputDateValue(end);
}

// ── Filter Dropdowns ───────────────────────────────────────────

function resetSelect(select, emptyLabel) {
  select.innerHTML = `<option value="">${emptyLabel}</option>`;
}

function fillSelect(select, items, emptyLabel) {
  resetSelect(select, emptyLabel);
  items.forEach((item) => {
    if (!item || !item.id) return;
    const opt = document.createElement("option");
    opt.value = item.id;
    opt.textContent = item.name || item.firstName || item.id;
    select.appendChild(opt);
  });
}

async function loadFilterData() {
  setBusy(true);
  try {
    setStatus("Loading filter options from MyGeotab...");
    
    const rulesData = await callApi("Get", { typeName: "Rule", resultsLimit: 500 });
    const deviceData = await callApi("Get", { typeName: "Device", resultsLimit: 500 });
    const driverData = await callApi("Get", { typeName: "User", search: { isDriver: true }, resultsLimit: 500 });

    const rules = Array.isArray(rulesData) ? rulesData : [];
    const devices = Array.isArray(deviceData) ? deviceData : [];
    const drivers = Array.isArray(driverData) ? driverData : [];

    fillSelect(ui.ruleSelect, rules, "Any Exception Rule");
    fillSelect(ui.deviceSelect, devices, "Any Asset");
    fillSelect(ui.driverSelect, drivers, "Any Driver");

    setStatus([
      "✓ Filter options loaded successfully.",
      `Rules: ${rules.length} | Assets: ${devices.length} | Drivers: ${drivers.length}`
    ]);
  } catch (err) {
    setStatus([
      "✗ Could not load filters.",
      `Error: ${err.message || err}`,
      "Check: Are you running this inside MyGeotab as an add-in?",
      "You can still type Search JSON manually."
    ]);
  } finally {
    setBusy(false);
  }
}

function getSelectedValues(select) {
  return Array.from(select.options)
    .filter((opt) => opt.selected && opt.value)
    .map((opt) => opt.value);
}

function buildSearch() {
  const fromIso = toIsoDate(ui.fromDateInput.value);
  const toIso = toIsoDate(ui.toDateInput.value);
  const rules = getSelectedValues(ui.ruleSelect);
  const devices = getSelectedValues(ui.deviceSelect);
  const drivers = getSelectedValues(ui.driverSelect);

  const search = {};
  if (fromIso) search.fromDate = fromIso;
  if (toIso) search.toDate = toIso;
  if (rules.length > 0) search.ruleSearch = { id: rules.length === 1 ? rules[0] : rules };
  if (devices.length > 0) search.deviceSearch = { id: devices.length === 1 ? devices[0] : devices };
  if (drivers.length > 0) search.driverSearch = { id: drivers.length === 1 ? drivers[0] : drivers };

  ui.searchJsonInput.value = JSON.stringify(search, null, 2);
  setStatus("Search JSON built from filters.");
}

function clearFilters() {
  ui.presetSelect.value = "last7";
  ui.fromDateInput.value = "";
  ui.toDateInput.value = "";
  Array.from(ui.ruleSelect.options).forEach((opt) => opt.selected = false);
  Array.from(ui.deviceSelect.options).forEach((opt) => opt.selected = false);
  Array.from(ui.driverSelect.options).forEach((opt) => opt.selected = false);
  ui.searchJsonInput.value = "";
  setStatus("Filters cleared.");
}

// ── API Calls ──────────────────────────────────────────────────

function callApi(method, params) {
  if (!runtime.api?.call) {
    throw new Error("MyGeotab API not available. Open as add-in inside MyGeotab.");
  }
  return new Promise((resolve, reject) => {
    runtime.api.call(method, params, resolve, reject);
  });
}

// ── Data Flattening ────────────────────────────────────────────

function flattenRow(record) {
  const out = {};
  function walk(obj, prefix) {
    Object.entries(obj || {}).forEach(([k, v]) => {
      const col = prefix ? `${prefix}.${k}` : k;
      if (v === null || v === undefined) {
        out[col] = "";
      } else if (typeof v === "object" && !Array.isArray(v)) {
        if (Object.prototype.hasOwnProperty.call(v, "id")) {
          out[`${col}.id`] = v.id;
        } else {
          walk(v, col);
        }
      } else if (Array.isArray(v)) {
        out[col] = JSON.stringify(v);
      } else {
        out[col] = v;
      }
    });
  }
  walk(record, "");
  return out;
}

// ── Excel Export with Template ─────────────────────────────────

async function exportToExcel(reportData, fileName) {
  // Fetch template
  setStatus("Loading template...");
  let response;
  try {
    response = await fetch(CONFIG.templateUrl);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
  } catch (err) {
    throw new Error(`Failed to fetch template: ${err.message}`);
  }

  const arrayBuffer = await response.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array" });

  // Replace sheets
  CONFIG.reports.forEach((report) => {
    const sheetName = report.sheetName;
    const rows = reportData[sheetName] || [];

    // Remove old sheet if exists
    const idx = wb.SheetNames.indexOf(sheetName);
    if (idx !== -1) {
      wb.SheetNames.splice(idx, 1);
      delete wb.Sheets[sheetName];
    }

    // Add new sheet with flattened data
    const flat = rows.map(flattenRow);
    const ws = XLSX.utils.json_to_sheet(flat);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  // Write file
  XLSX.writeFile(wb, fileName);
}

// ── Test: Pull 1 row from each report ──────────────────────────

async function testConnection() {
  setBusy(true);
  try {
    const raw = ui.searchJsonInput.value.trim();
    let search = {};
    if (raw) search = JSON.parse(raw);

    setStatus("Testing... pulling 1 row from each report.");

    const results = {};
    for (const report of CONFIG.reports) {
      try {
        const data = await callApi("Get", { typeName: report.typeName, search, resultsLimit: 1 });
        results[report.sheetName] = Array.isArray(data) ? data : [];
      } catch (err) {
        if (report.fallback) {
          setStatus(`${report.typeName} failed, trying ${report.fallback}...`);
          const data = await callApi("Get", { typeName: report.fallback, search, resultsLimit: 1 });
          results[report.sheetName] = Array.isArray(data) ? data : [];
        } else {
          throw err;
        }
      }
    }

    const lines = ["Test successful."];
    Object.entries(results).forEach(([sheet, rows]) => {
      if (rows.length > 0) {
        const cols = Object.keys(flattenRow(rows[0]));
        lines.push(`${sheet}: ${rows.length} row, ${cols.length} columns`);
      } else {
        lines.push(`${sheet}: 0 rows`);
      }
    });

    setStatus(lines);
  } catch (err) {
    setStatus(["Test failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Export: Pull all rows and merge with template ──────────────

async function runExport() {
  setBusy(true);
  try {
    const raw = ui.searchJsonInput.value.trim();
    let search = {};
    if (raw) search = JSON.parse(raw);

    const fileName = ui.fileNameInput.value.trim() || "geotab-export.xlsx";

    setStatus("Pulling reports...");

    const reportData = {};
    for (const report of CONFIG.reports) {
      try {
        const data = await callApi("Get", { typeName: report.typeName, search });
        reportData[report.sheetName] = Array.isArray(data) ? data : [];
      } catch (err) {
        if (report.fallback) {
          setStatus(`${report.typeName} not available, using ${report.fallback}...`);
          const data = await callApi("Get", { typeName: report.fallback, search });
          reportData[report.sheetName] = Array.isArray(data) ? data : [];
        } else {
          throw err;
        }
      }
    }

    const totalRows = Object.values(reportData).reduce((sum, rows) => sum + rows.length, 0);
    if (totalRows === 0) {
      setStatus("No data returned. Try broadening your search filters.");
      return;
    }

    setStatus(`Received ${totalRows} rows. Merging with template and exporting...`);
    await exportToExcel(reportData, fileName);

    const lines = ["Export complete."];
    Object.entries(reportData).forEach(([sheet, rows]) => {
      lines.push(`${sheet}: ${rows.length} rows`);
    });
    lines.push(`File: ${fileName}`);
    setStatus(lines);
  } catch (err) {
    setStatus(["Export failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Initialization ─────────────────────────────────────────────

function initUI() {
  ui.fileNameInput = document.getElementById("fileName");
  ui.presetSelect = document.getElementById("presetSelect");
  ui.fromDateInput = document.getElementById("fromDate");
  ui.toDateInput = document.getElementById("toDate");
  ui.ruleSelect = document.getElementById("ruleSelect");
  ui.deviceSelect = document.getElementById("deviceSelect");
  ui.driverSelect = document.getElementById("driverSelect");
  ui.searchJsonInput = document.getElementById("searchJson");
  ui.loadFilterBtn = document.getElementById("loadFilterBtn");
  ui.buildSearchBtn = document.getElementById("buildSearchBtn");
  ui.clearFilterBtn = document.getElementById("clearFilterBtn");
  ui.testBtn = document.getElementById("testBtn");
  ui.exportBtn = document.getElementById("exportBtn");
  ui.statusEl = document.getElementById("status");

  ui.presetSelect?.addEventListener("change", applyDatePreset);
  ui.loadFilterBtn?.addEventListener("click", loadFilterData);
  ui.buildSearchBtn?.addEventListener("click", buildSearch);
  ui.clearFilterBtn?.addEventListener("click", clearFilters);
  ui.testBtn?.addEventListener("click", testConnection);
  ui.exportBtn?.addEventListener("click", runExport);

  setBusy(false);
  applyDatePreset();
  setStatus("Ready. Load filters or enter Search JSON, then click Test or Export.");
}

// ── Geotab Add-in Entry Point ──────────────────────────────────

window.geotab = window.geotab || {};
window.geotab.addin = window.geotab.addin || {};
window.geotab.addin.dualReportExport = function (api) {
  return {
    initialize(freshApi, state, callback) {
      runtime.api = freshApi || api;
      initUI();
      if (runtime.api?.call) {
        loadFilterData();
      }
      callback();
    },
    focus(freshApi, state, callback) {
      if (freshApi || api) runtime.api = freshApi || api;
      callback();
    },
    blur() {},
    destroy() {
      runtime.api = null;
    }
  };
};

// Standalone fallback
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => {
    initUI();
    setStatus("Open this page as an add-in inside MyGeotab.");
  });
} else {
  initUI();
  setStatus("Open this page as an add-in inside MyGeotab.");
}
