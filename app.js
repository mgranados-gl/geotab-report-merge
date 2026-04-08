/* ── Geotab Add-In: single report pull + Excel export ── */

const runtime = { api: null };

const ui = {
  typeName: null,
  searchJson: null,
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
  status: null
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
  const typeName = ui.typeName.value.trim();
  if (!typeName) {
    throw new Error("Type Name is required.");
  }

  const raw = ui.searchJson.value.trim();
  let search = {};
  if (raw) {
    try {
      search = JSON.parse(raw);
    } catch {
      throw new Error("Search JSON is not valid JSON. Check your syntax.");
    }
  }

  const fileName = ui.fileName.value.trim() || "geotab-report.xlsx";
  return { typeName, search, fileName };
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

// ── Excel export ─────────────────────────────────────────

function exportToExcel(rows, fileName) {
  const flat = rows.map(flattenRow);
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(flat);
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, fileName);
}

// ── Test: pull 1 row ─────────────────────────────────────

async function testConnection() {
  setBusy(true);
  try {
    const { typeName, search } = getInputs();
    setStatus(`Testing — pulling 1 row from "${typeName}"...`);

    const result = await callApi("Get", { typeName, search, resultsLimit: 1 });
    const rows = Array.isArray(result) ? result : [];

    if (rows.length === 0) {
      setStatus([
        "Test OK — API call succeeded.",
        `Type "${typeName}" returned 0 rows for the given search.`,
        "Try broadening your Search JSON or leaving it blank."
      ]);
    } else {
      const cols = Object.keys(flattenRow(rows[0]));
      setStatus([
        "Test OK — 1 row received.",
        `Type: ${typeName}`,
        `Columns: ${cols.join(", ")}`
      ]);
    }
  } catch (err) {
    setStatus(["Test failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Run: pull all rows and export ────────────────────────

async function runExport() {
  setBusy(true);
  try {
    const { typeName, search, fileName } = getInputs();
    setStatus(`Pulling all rows from "${typeName}"...`);

    const result = await callApi("Get", { typeName, search });
    const rows = Array.isArray(result) ? result : [];

    if (rows.length === 0) {
      setStatus([
        "API call succeeded but returned 0 rows.",
        "Try broadening your Search JSON or leaving it blank."
      ]);
      return;
    }

    setStatus(`Received ${rows.length} rows. Exporting to Excel...`);
    exportToExcel(rows, fileName);

    setStatus([
      "Export complete.",
      `Type: ${typeName}`,
      `Rows: ${rows.length}`,
      `File: ${fileName}`
    ]);
  } catch (err) {
    setStatus(["Export failed.", err.message || String(err)]);
  } finally {
    setBusy(false);
  }
}

// ── Add-in lifecycle ─────────────────────────────────────

function onReady(api) {
  runtime.api = api;

  ui.typeName  = document.getElementById("typeName");
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
  setStatus("Ready. Enter a Type Name and click Test or Pull Report.");
  applyDatePreset();
  loadFilterData();
}

function onStandalone() {
  ui.typeName  = document.getElementById("typeName");
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
