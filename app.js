/* ── Geotab Add-In: single report pull + Excel export ── */

const runtime = { api: null };

const ui = {
  typeName: null,
  searchJson: null,
  fileName: null,
  fromDate: null,
  toDate: null,
  deviceId: null,
  driverId: null,
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

function buildSearchFromFilters() {
  const fromIso = toIsoFromLocalDateTime(ui.fromDate.value.trim());
  const toIso = toIsoFromLocalDateTime(ui.toDate.value.trim());
  const deviceId = ui.deviceId.value.trim();
  const driverId = ui.driverId.value.trim();

  if (!fromIso && !toIso && !deviceId && !driverId) {
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
  if (deviceId) {
    search.deviceSearch = { id: deviceId };
  }
  if (driverId) {
    search.driverSearch = { id: driverId };
  }

  ui.searchJson.value = JSON.stringify(search, null, 2);
  setStatus([
    "Search JSON updated from filters.",
    "You can edit it manually before running Test or Export."
  ]);
}

function clearFilters() {
  ui.fromDate.value = "";
  ui.toDate.value = "";
  ui.deviceId.value = "";
  ui.driverId.value = "";
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
  ui.fromDate  = document.getElementById("fromDate");
  ui.toDate    = document.getElementById("toDate");
  ui.deviceId  = document.getElementById("deviceId");
  ui.driverId  = document.getElementById("driverId");
  ui.buildSearchBtn = document.getElementById("buildSearchBtn");
  ui.clearSearchBtn = document.getElementById("clearSearchBtn");
  ui.testBtn   = document.getElementById("testBtn");
  ui.runBtn    = document.getElementById("runBtn");
  ui.status    = document.getElementById("status");

  ui.buildSearchBtn.addEventListener("click", buildSearchFromFilters);
  ui.clearSearchBtn.addEventListener("click", clearFilters);
  ui.testBtn.addEventListener("click", testConnection);
  ui.runBtn.addEventListener("click", runExport);

  setBusy(false);
  setStatus("Ready. Enter a Type Name and click Test or Pull Report.");
}

function onStandalone() {
  ui.typeName  = document.getElementById("typeName");
  ui.searchJson = document.getElementById("searchJson");
  ui.fileName  = document.getElementById("fileName");
  ui.fromDate  = document.getElementById("fromDate");
  ui.toDate    = document.getElementById("toDate");
  ui.deviceId  = document.getElementById("deviceId");
  ui.driverId  = document.getElementById("driverId");
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
