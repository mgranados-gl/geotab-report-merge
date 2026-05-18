/* ── Geotab Add-In: dual report pull + multi-sheet Excel export ── */

const runtime = { api: null };

const REPORT_CONFIG = {
  templateUrl: "https://mgranados-gl.github.io/geotab-report-merge/template/Gridline%20_%20Driver%20Events%20(Yesterday).xlsx",
  report1: {
    typeName: "HosLog",
    fallbackTypeName: "DutyStatusLog",
    sheetName: "Data1",
    description: "HosLog"
  },
  report2: {
    typeName: "ExceptionsDetail",
    sheetName: "Data2",
    description: "ExceptionsDetail"
  }
};

const ui = {
  fileName: null,
  datePreset: null,
  fromDate: null,
  toDate: null,
  exceptionRuleSelect: null,
  deviceSelect: null,
  driverSelect: null,
  loadFilterDataBtn: null,
  buildSearchBtn: null,
  clearSearchBtn: null,
  testBtn: null,
  runBtn: null,
  status: null,
  searchJson: null
};

// ── UI helpers ────────────────────────────────────────────

function setStatus(lines) {
  ui.status.textContent = Array.isArray(lines) ? lines.join("\n") : String(lines);
}

function setBusy(busy) {
  ui.testBtn.disabled = busy;
  ui.runBtn.disabled = busy;
  if (ui.buildSearchBtn) {
    ui.buildSearchBtn.disabled = busy;
  }
  if (ui.clearSearchBtn) {
    ui.clearSearchBtn.disabled = busy;
  }
  if (ui.loadFilterDataBtn) {
    ui.loadFilterDataBtn.disabled = busy;
  }
}

function toInputDateTimeValue(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return [
    date.getUTCFullYear(),
    "-",
    pad(date.getUTCMonth() + 1),
    "-",
    pad(date.getUTCDate()),
    "T",
    pad(date.getUTCHours()),
    ":",
    pad(date.getUTCMinutes())
  ].join("");
}

function applyDatePreset() {
  const now = new Date();
  const end = new Date(now);
  let start = null;

  switch (ui.datePreset.value) {
    case "today": {
      start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 0, 0, 0));
      break;
    }
    case "yesterday": {
      start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - 1, 0, 0, 0));
      end.setUTCFullYear(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate());
      end.setUTCHours(23, 59, 59, 0);
      break;
    }
    case "last7": {
      start = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
      break;
    }
    case "last30": {
      start = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
      break;
    }
    case "monthToDate": {
      start = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, 0, 0));
      break;
    }
    default: {
      return;
    }
  }

  ui.fromDate.value = toInputDateTimeValue(start);
  ui.toDate.value = toInputDateTimeValue(end);
}

function resetSelectOptions(selectEl, emptyLabel) {
  selectEl.innerHTML = "";
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = emptyLabel;
  selectEl.appendChild(defaultOption);
}

function optionText(item, fallbackPrefix) {
  if (item.name) {
    return item.name;
  }
  if (item.firstName || item.lastName) {
    const fullName = [item.firstName || "", item.lastName || ""].join(" ").trim();
    if (fullName) {
      return fullName;
    }
  }
  return `${fallbackPrefix} ${item.id}`;
}

function fillSelect(selectEl, items, emptyLabel, fallbackPrefix) {
  resetSelectOptions(selectEl, emptyLabel);
  items.forEach((item) => {
    if (!item || !item.id) {
      return;
    }
    const opt = document.createElement("option");
    opt.value = item.id;
    opt.textContent = optionText(item, fallbackPrefix);
    selectEl.appendChild(opt);
  });
}

async function loadFilterData() {
  setBusy(true);
  try {
    setStatus("Loading filter dropdown data from Geotab...");

    const [rulesResult, devicesResult, driversResult] = await Promise.all([
      callApi("Get", { typeName: "Rule", resultsLimit: 500 }),
      callApi("Get", { typeName: "Device", resultsLimit: 500 }),
      callApi("Get", { typeName: "User", search: { isDriver: true }, resultsLimit: 500 })
    ]);

    const rules = Array.isArray(rulesResult) ? rulesResult : [];
    const devices = Array.isArray(devicesResult) ? devicesResult : [];
    const drivers = Array.isArray(driversResult) ? driversResult : [];

    fillSelect(ui.exceptionRuleSelect, rules, "Any Exception Rule", "Rule");
    fillSelect(ui.deviceSelect, devices, "Any Asset", "Asset");
    fillSelect(ui.driverSelect, drivers, "Any Driver", "Driver");

    setStatus([
      "Filter dropdowns loaded.",
      `Rules: ${rules.length}`,
      `Assets: ${devices.length}`,
      `Drivers: ${drivers.length}`
    ]);
  } catch (err) {
    setStatus([
      "Could not load one or more dropdowns.",
      err.message || String(err),
      "You can still paste Search JSON manually."
    ]);
  } finally {
    setBusy(false);
  }
}

function toIsoFromLocalDateTime(value) {
  if (!value) {
    return null;
  }

  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) {
    return null;
  }
  return dt.toISOString();
}

function getSelectedValues(selectEl) {
  const selected = [];
  Array.from(selectEl.options).forEach((opt) => {
    if (opt.selected && opt.value) {
      selected.push(opt.value);
    }
  });
  return selected;
}

function buildSearchFromFilters() {
  const fromIso = toIsoFromLocalDateTime(ui.fromDate.value.trim());
  const toIso = toIsoFromLocalDateTime(ui.toDate.value.trim());
  const ruleIds = getSelectedValues(ui.exceptionRuleSelect);
  const deviceIds = getSelectedValues(ui.deviceSelect);
  const driverIds = getSelectedValues(ui.driverSelect);

  if (!fromIso && !toIso && ruleIds.length === 0 && deviceIds.length === 0 && driverIds.length === 0) {
    ui.searchJson.value = "";
    setStatus([
      "Filters cleared.",
      "Search JSON is blank and no filter will be applied."
    ]);
    return;
  }

  const search = {};
  if (fromIso) {
    search.fromDate = fromIso;
  }
  if (toIso) {
    search.toDate = toIso;
  }
  if (ruleIds.length > 0) {
    search.ruleSearch = { id: ruleIds.length === 1 ? ruleIds[0] : ruleIds };
  }
  if (deviceIds.length > 0) {
    search.deviceSearch = { id: deviceIds.length === 1 ? deviceIds[0] : deviceIds };
  }
  if (driverIds.length > 0) {
    search.driverSearch = { id: driverIds.length === 1 ? driverIds[0] : driverIds };
  }

  ui.searchJson.value = JSON.stringify(search, null, 2);
  setStatus([
    "Search JSON updated from filters.",
    "You can edit it manually before running Test or Export."
  ]);
}

function clearFilters() {
  ui.datePreset.value = "custom";
  ui.fromDate.value = "";
  ui.toDate.value = "";
  Array.from(ui.exceptionRuleSelect.options).forEach((opt) => {
    opt.selected = false;
  });
  Array.from(ui.deviceSelect.options).forEach((opt) => {
    opt.selected = false;
  });
  Array.from(ui.driverSelect.options).forEach((opt) => {
    opt.selected = false;
  });
  ui.searchJson.value = "";

  setStatus([
    "Filters reset.",
    "Search JSON is now blank."
  ]);
}

// ── API call via Geotab add-in session ───────────────────

function callApi(method, params) {
  if (!runtime.api || typeof runtime.api.call !== "function") {
    throw new Error(
      "MyGeotab API is not available. " +
      "Open this page as a registered add-in inside MyGeotab."
    );
  }
  return new Promise((resolve, reject) => {
    runtime.api.call(method, params, resolve, reject);
  });
}

// ── Input helpers ────────────────────────────────────────

function getInputs() {
  const raw = ui.searchJson.value.trim();
  let search = {};
  if (raw) {
    try {
      search = JSON.parse(raw);
    } catch {
      throw new Error("Search JSON is not valid JSON. Check your syntax.");
    }
  }

  const fileName = ui.fileName.value.trim() || "geotab-dual-report.xlsx";
  return { search, fileName };
}

// ── Flatten a row for Excel export ───────────────────────

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

// ── Excel export with template ───────────────────────────

async function exportToExcel(reportData, fileName) {
  let wb;

  // Fetch the hosted template
  try {
    setStatus("Loading Excel template...");
    const response = await fetch(REPORT_CONFIG.templateUrl);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    wb = XLSX.read(arrayBuffer, { type: "array" });
  } catch (err) {
    throw new Error(`Failed to load template: ${err.message}`);
  }

  // Replace (or add) Data1 and Data2 sheets with report data
  Object.entries(reportData).forEach(([sheetName, rows]) => {
    const flat = rows.length > 0 ? rows.map(flattenRow) : [];
    const ws = XLSX.utils.json_to_sheet(flat);

    // Remove existing sheet with this name if present
    const existingIdx = wb.SheetNames.indexOf(sheetName);
    if (existingIdx !== -1) {
      wb.SheetNames.splice(existingIdx, 1);
      delete wb.Sheets[sheetName];
    }

    // Append the new sheet
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, fileName);
}

// ── Test: pull 1 row from each report ───────────────────

async function testConnection() {
  setBusy(true);
  try {
    const { search } = getInputs();
    setStatus("Testing — pulling 1 row from each report...");

    const tests = [];
    const results = {};

    // Test Report 1 (HosLog with fallback)
    try {
      let result = await callApi("Get", { typeName: REPORT_CONFIG.report1.typeName, search, resultsLimit: 1 });
      tests.push(`${REPORT_CONFIG.report1.description}: 1 row`);
      results.report1 = Array.isArray(result) ? result : [];
    } catch (err) {
      // Fallback to DutyStatusLog
      const result = await callApi("Get", { typeName: REPORT_CONFIG.report1.fallbackTypeName, search, resultsLimit: 1 });
      tests.push(`${REPORT_CONFIG.report1.fallbackTypeName}: 1 row (HosLog fallback)");
      results.report1 = Array.isArray(result) ? result : [];
    }

    // Test Report 2
    const result2 = await callApi("Get", { typeName: REPORT_CONFIG.report2.typeName, search, resultsLimit: 1 });
    tests.push(`${REPORT_CONFIG.report2.description}: 1 row`);
    results.report2 = Array.isArray(result2) ? result2 : [];

    const statusLines = ["Test OK — API calls succeeded.", ...tests];

    Object.entries(results).forEach(([key, rows]) => {
      if (rows.length > 0) {
        const cols = Object.keys(flattenRow(rows[0]));
        statusLines.push(`${key === "report1" ? "Report 1" : "Report 2"} columns (${cols.length}): ${cols.join(", ")}`);
      }
    });

    setStatus(statusLines);
  } catch (err) {
    setStatus(["Test failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Run: pull all rows from both reports and export ───────

async function runExport() {
  setBusy(true);
  try {
    const { search, fileName } = getInputs();
    const reportData = {};
    const rowCounts = {};

    // Pull Report 1 (HosLog with fallback)
    setStatus("Pulling Report 1 (HosLog)...");
    let report1TypeUsed = REPORT_CONFIG.report1.typeName;
    try {
      const result = await callApi("Get", { typeName: REPORT_CONFIG.report1.typeName, search });
      reportData[REPORT_CONFIG.report1.sheetName] = Array.isArray(result) ? result : [];
    } catch (err) {
      // Fallback to DutyStatusLog
      setStatus("HosLog not available, falling back to DutyStatusLog...");
      const result = await callApi("Get", { typeName: REPORT_CONFIG.report1.fallbackTypeName, search });
      reportData[REPORT_CONFIG.report1.sheetName] = Array.isArray(result) ? result : [];
      report1TypeUsed = REPORT_CONFIG.report1.fallbackTypeName;
    }
    rowCounts["Report 1"] = reportData[REPORT_CONFIG.report1.sheetName].length;

    // Pull Report 2
    setStatus("Pulling Report 2 (ExceptionsDetail)...");
    const result2 = await callApi("Get", { typeName: REPORT_CONFIG.report2.typeName, search });
    reportData[REPORT_CONFIG.report2.sheetName] = Array.isArray(result2) ? result2 : [];
    rowCounts["Report 2"] = reportData[REPORT_CONFIG.report2.sheetName].length;

    const totalRows = Object.values(rowCounts).reduce((a, b) => a + b, 0);
    if (totalRows === 0) {
      setStatus([
        "API calls succeeded but returned 0 rows total.",
        "Try broadening your Search JSON or leaving it blank."
      ]);
      return;
    }

    setStatus(`Received ${totalRows} total rows. Loading template and exporting to Excel...`);
    await exportToExcel(reportData, fileName);

    const statusLines = [
      "Export complete.",
      `Report 1 (${report1TypeUsed}): ${rowCounts["Report 1"]} rows`,
      `Report 2 (${REPORT_CONFIG.report2.typeName}): ${rowCounts["Report 2"]} rows`,
      `Total: ${totalRows} rows`,
      `File: ${fileName}`
    ];
    setStatus(statusLines);
  } catch (err) {
    setStatus(["Export failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Add-in lifecycle ─────────────────────────────────────

function onReady(api) {
  runtime.api = api;

  ui.searchJson = document.getElementById("searchJson");
  ui.fileName  = document.getElementById("fileName");
  ui.datePreset = document.getElementById("datePreset");
  ui.fromDate  = document.getElementById("fromDate");
  ui.toDate    = document.getElementById("toDate");
  ui.exceptionRuleSelect = document.getElementById("exceptionRuleSelect");
  ui.deviceSelect = document.getElementById("deviceSelect");
  ui.driverSelect = document.getElementById("driverSelect");
  ui.loadFilterDataBtn = document.getElementById("loadFilterDataBtn");
  ui.buildSearchBtn = document.getElementById("buildSearchBtn");
  ui.clearSearchBtn = document.getElementById("clearSearchBtn");
  ui.testBtn   = document.getElementById("testBtn");
  ui.runBtn    = document.getElementById("runBtn");
  ui.status    = document.getElementById("status");

  ui.datePreset.addEventListener("change", applyDatePreset);
  ui.loadFilterDataBtn.addEventListener("click", loadFilterData);
  ui.buildSearchBtn.addEventListener("click", buildSearchFromFilters);
  ui.clearSearchBtn.addEventListener("click", clearFilters);
  ui.testBtn.addEventListener("click", testConnection);
  ui.runBtn.addEventListener("click", runExport);

  setBusy(false);
  setStatus(["Ready. Configure filters and click Test or Pull Reports & Export.",
    "Report 1: HosLog (or DutyStatusLog fallback) → Data1 sheet",
    "Report 2: ExceptionsDetail → Data2 sheet"]);
  applyDatePreset();
  loadFilterData();
}

function onStandalone() {
  ui.searchJson = document.getElementById("searchJson");
  ui.fileName  = document.getElementById("fileName");
  ui.datePreset = document.getElementById("datePreset");
  ui.fromDate  = document.getElementById("fromDate");
  ui.toDate    = document.getElementById("toDate");
  ui.exceptionRuleSelect = document.getElementById("exceptionRuleSelect");
  ui.deviceSelect = document.getElementById("deviceSelect");
  ui.driverSelect = document.getElementById("driverSelect");
  ui.loadFilterDataBtn = document.getElementById("loadFilterDataBtn");
  ui.buildSearchBtn = document.getElementById("buildSearchBtn");
  ui.clearSearchBtn = document.getElementById("clearSearchBtn");
  ui.testBtn   = document.getElementById("testBtn");
  ui.runBtn    = document.getElementById("runBtn");
  ui.status    = document.getElementById("status");

  setBusy(true);
  setStatus(
    "No MyGeotab session detected.\n" +
    "Open this page as a registered add-in inside MyGeotab."
  );
}

// Standalone fallback (direct browser open)
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", onStandalone);
} else {
  onStandalone();
}

// Geotab Add-In entry point
window.geotab = window.geotab || {};
window.geotab.addin = window.geotab.addin || {};
window.geotab.addin.reportExport = function (api) {
  return {
    initialize(freshApi, state, callback) {
      onReady(freshApi || api);
      callback();
    },
    focus(freshApi, state, callback) {
      if (freshApi || api) {
        runtime.api = freshApi || api;
      }
      callback();
    },
    blur() {},
    destroy() {
      runtime.api = null;
    }
  };
};
