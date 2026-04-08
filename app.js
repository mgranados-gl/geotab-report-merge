const REPORT_MERGE_CONFIG = {
  reportA: {
    name: "HOS Log",
    customReport: {
      id: "b13C",
      scheduleId: "bAA8",
      source: "HosLog"
    },
    typeNames: ["HosLog", "DutyStatusLog"],
    keyPath: "driver.id",
    prefix: "Hos",
    fields: ["id", "dateTime", "driver.id", "device.id", "state", "status"],
    buildSearch: (fromIso, toIso) => ({
      fromDate: fromIso,
      toDate: toIso
    })
  },
  reportB: {
    name: "Exceptions Detail",
    customReport: {
      id: "b13B",
      scheduleId: "bAA6",
      source: "ExceptionsDetail"
    },
    typeNames: ["ExceptionEvent"],
    keyPath: "driver.id",
    prefix: "Ex",
    fields: ["id", "dateTime", "driver.id", "device.id", "rule.id", "distance", "duration"],
    buildSearch: (fromIso, toIso) => ({
      fromDate: fromIso,
      toDate: toIso
    })
  },
  joinType: "left"
};

const runtime = {
  api: null,
  uiBound: false
};

const ui = {
  reportAInfo: null,
  reportBInfo: null,
  fromDate: null,
  toDate: null,
  joinTypeInfo: null,
  fileName: null,
  testBtn: null,
  runBtn: null,
  status: null
};

function bindUi() {
  if (runtime.uiBound) {
    return;
  }

  ui.reportAInfo = document.getElementById("reportAInfo");
  ui.reportBInfo = document.getElementById("reportBInfo");
  ui.fromDate = document.getElementById("fromDate");
  ui.toDate = document.getElementById("toDate");
  ui.joinTypeInfo = document.getElementById("joinTypeInfo");
  ui.fileName = document.getElementById("fileName");
  ui.testBtn = document.getElementById("testBtn");
  ui.runBtn = document.getElementById("runBtn");
  ui.status = document.getElementById("status");

  ui.testBtn.addEventListener("click", testConnection);
  ui.runBtn.addEventListener("click", runMergeExport);
  runtime.uiBound = true;
}

function setStatus(lines) {
  ui.status.textContent = Array.isArray(lines) ? lines.join("\n") : String(lines);
}

function setButtonsDisabled(disabled) {
  ui.testBtn.disabled = disabled;
  ui.runBtn.disabled = disabled;
}

function initializeUiDefaults() {
  ui.reportAInfo.value = `${REPORT_MERGE_CONFIG.reportA.name} | ${REPORT_MERGE_CONFIG.reportA.customReport.source} | id:${REPORT_MERGE_CONFIG.reportA.customReport.id}`;
  ui.reportBInfo.value = `${REPORT_MERGE_CONFIG.reportB.name} | ${REPORT_MERGE_CONFIG.reportB.customReport.source} | id:${REPORT_MERGE_CONFIG.reportB.customReport.id}`;
  ui.joinTypeInfo.value = REPORT_MERGE_CONFIG.joinType;

  const now = new Date();
  const prior = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
  const toLocalInput = (date) => {
    const offsetMs = date.getTimezoneOffset() * 60 * 1000;
    return new Date(date.getTime() - offsetMs).toISOString().slice(0, 16);
  };

  ui.fromDate.value = toLocalInput(prior);
  ui.toDate.value = toLocalInput(now);
}

function toIsoFromLocal(value) {
  if (!value) {
    return "";
  }
  const dt = new Date(value);
  if (Number.isNaN(dt.getTime())) {
    return "";
  }
  return dt.toISOString();
}

function getValueByPath(record, path) {
  return path.split(".").reduce((current, segment) => {
    if (current === null || current === undefined) {
      return undefined;
    }
    return current[segment];
  }, record);
}

function valueToCell(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "object") {
    if (Object.prototype.hasOwnProperty.call(value, "id")) {
      return value.id;
    }
    return JSON.stringify(value);
  }
  return value;
}

function normalizeFieldName(prefix, path) {
  return `${prefix}_${path.replaceAll(".", "_")}`;
}

function projectRow(record, fields, prefix) {
  const projected = {};
  fields.forEach((path) => {
    projected[normalizeFieldName(prefix, path)] = valueToCell(getValueByPath(record, path));
  });
  return projected;
}

function mergeReports(leftRows, rightRows, leftConfig, rightConfig, joinType) {
  const rightMap = new Map();
  rightRows.forEach((row) => {
    const key = getValueByPath(row, rightConfig.keyPath);
    if (key === undefined || key === null) {
      return;
    }
    if (!rightMap.has(key)) {
      rightMap.set(key, []);
    }
    rightMap.get(key).push(row);
  });

  const merged = [];
  const matchedRightIndexes = new Set();

  leftRows.forEach((leftRow) => {
    const leftValue = getValueByPath(leftRow, leftConfig.keyPath);
    const matches = rightMap.get(leftValue) || [];

    if (matches.length === 0) {
      if (joinType === "left" || joinType === "full") {
        merged.push({
          ...projectRow(leftRow, leftConfig.fields, leftConfig.prefix)
        });
      }
      return;
    }

    matches.forEach((rightRow) => {
      merged.push({
        ...projectRow(leftRow, leftConfig.fields, leftConfig.prefix),
        ...projectRow(rightRow, rightConfig.fields, rightConfig.prefix)
      });

      const idx = rightRows.indexOf(rightRow);
      if (idx >= 0) {
        matchedRightIndexes.add(idx);
      }
    });
  });

  if (joinType === "full") {
    rightRows.forEach((rightRow, idx) => {
      if (!matchedRightIndexes.has(idx)) {
        merged.push({
          ...projectRow(rightRow, rightConfig.fields, rightConfig.prefix)
        });
      }
    });
  }

  return merged;
}

function exportToExcel(mergedRows, details, fileName) {
  const workbook = XLSX.utils.book_new();
  const mergedSheet = XLSX.utils.json_to_sheet(mergedRows);
  const detailsSheet = XLSX.utils.json_to_sheet(details);

  XLSX.utils.book_append_sheet(workbook, mergedSheet, "Merged");
  XLSX.utils.book_append_sheet(workbook, detailsSheet, "RunDetails");
  XLSX.writeFile(workbook, fileName || "geotab-report-merge.xlsx");
}

function callApi(method, params) {
  if (!runtime.api || typeof runtime.api.call !== "function") {
    throw new Error("MyGeotab API session is not available. Open this page as a registered MyGeotab add-in.");
  }

  return new Promise((resolve, reject) => {
    runtime.api.call(method, params, resolve, reject);
  });
}

async function getReportDataByTypeFallback(typeNames, search, label, resultsLimit) {
  let lastError;

  for (const typeName of typeNames) {
    try {
      const params = {
        typeName,
        search
      };

      if (typeof resultsLimit === "number" && resultsLimit > 0) {
        params.resultsLimit = resultsLimit;
      }

      const result = await callApi("Get", params);
      return {
        typeName,
        rows: Array.isArray(result) ? result : []
      };
    } catch (error) {
      lastError = error;
    }
  }

  throw new Error(`Unable to pull ${label}. Tried: ${typeNames.join(", ")}. Last error: ${lastError?.message || "unknown"}`);
}

function getRunInputs() {
  const fromIso = toIsoFromLocal(ui.fromDate.value);
  const toIso = toIsoFromLocal(ui.toDate.value);

  if (!fromIso || !toIso) {
    throw new Error("From and To date are required.");
  }

  if (new Date(fromIso).getTime() > new Date(toIso).getTime()) {
    throw new Error("From date must be before To date.");
  }

  const fileName = ui.fileName.value.trim() || "geotab-report-merge.xlsx";

  return {
    fromIso,
    toIso,
    fileName,
    reportAConfig: REPORT_MERGE_CONFIG.reportA,
    reportBConfig: REPORT_MERGE_CONFIG.reportB,
    joinType: REPORT_MERGE_CONFIG.joinType
  };
}

async function testConnection() {
  setButtonsDisabled(true);

  try {
    const inputs = getRunInputs();

    setStatus([
      "Testing Report A...",
      "Testing Report B..."
    ]);

    const [reportAResult, reportBResult] = await Promise.all([
      getReportDataByTypeFallback(
        inputs.reportAConfig.typeNames,
        inputs.reportAConfig.buildSearch(inputs.fromIso, inputs.toIso),
        inputs.reportAConfig.name,
        1
      ),
      getReportDataByTypeFallback(
        inputs.reportBConfig.typeNames,
        inputs.reportBConfig.buildSearch(inputs.fromIso, inputs.toIso),
        inputs.reportBConfig.name,
        1
      )
    ]);

    setStatus([
      "Connection test passed.",
      `Report A OK: ${inputs.reportAConfig.name} via ${reportAResult.typeName}`,
      `Report B OK: ${inputs.reportBConfig.name} via ${reportBResult.typeName}`,
      "Ready to run full export."
    ]);
  } catch (error) {
    setStatus([
      "Connection test failed.",
      error.message || String(error)
    ]);
  } finally {
    setButtonsDisabled(false);
  }
}

async function runMergeExport() {
  setButtonsDisabled(true);

  try {
    const inputs = getRunInputs();

    setStatus([
      "Pulling Report A...",
      "Pulling Report B..."
    ]);

    const [reportAResult, reportBResult] = await Promise.all([
      getReportDataByTypeFallback(
        inputs.reportAConfig.typeNames,
        inputs.reportAConfig.buildSearch(inputs.fromIso, inputs.toIso),
        inputs.reportAConfig.name
      ),
      getReportDataByTypeFallback(
        inputs.reportBConfig.typeNames,
        inputs.reportBConfig.buildSearch(inputs.fromIso, inputs.toIso),
        inputs.reportBConfig.name
      )
    ]);

    const mergedRows = mergeReports(
      reportAResult.rows,
      reportBResult.rows,
      inputs.reportAConfig,
      inputs.reportBConfig,
      inputs.joinType
    );

    exportToExcel(
      mergedRows,
      [
        { key: "Run At", value: new Date().toISOString() },
        { key: "From Date", value: inputs.fromIso },
        { key: "To Date", value: inputs.toIso },
        { key: "Report A", value: `${inputs.reportAConfig.name} (${reportAResult.typeName})` },
        { key: "Report A Custom Report", value: `${inputs.reportAConfig.customReport.source} | id:${inputs.reportAConfig.customReport.id} | scheduleId:${inputs.reportAConfig.customReport.scheduleId}` },
        { key: "Report B", value: `${inputs.reportBConfig.name} (${reportBResult.typeName})` },
        { key: "Report B Custom Report", value: `${inputs.reportBConfig.customReport.source} | id:${inputs.reportBConfig.customReport.id} | scheduleId:${inputs.reportBConfig.customReport.scheduleId}` },
        { key: "Join Type", value: inputs.joinType },
        { key: "Report A Rows", value: reportAResult.rows.length },
        { key: "Report B Rows", value: reportBResult.rows.length },
        { key: "Merged Rows", value: mergedRows.length }
      ],
      inputs.fileName
    );

    setStatus([
      "Export complete.",
      `Report A rows: ${reportAResult.rows.length}`,
      `Report B rows: ${reportBResult.rows.length}`,
      `Merged rows: ${mergedRows.length}`,
      `Saved as: ${inputs.fileName}`
    ]);
  } catch (error) {
    setStatus([
      "Run failed.",
      error.message || String(error)
    ]);
  } finally {
    setButtonsDisabled(false);
  }
}

function onAddInReady(api) {
  runtime.api = api;
  bindUi();
  initializeUiDefaults();
  setButtonsDisabled(false);
  setStatus("Ready. Click Test Connection.");
}

function bootstrapStandaloneMode() {
  bindUi();
  initializeUiDefaults();
  setButtonsDisabled(true);
  setStatus("This page must be opened inside MyGeotab as a registered add-in.");
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", bootstrapStandaloneMode);
} else {
  bootstrapStandaloneMode();
}

window.geotab = window.geotab || {};
window.geotab.addin = window.geotab.addin || {};

window.geotab.addin.reportMerge = function (api) {
  return {
    initialize: function (freshApi, state, callback) {
      onAddInReady(freshApi || api);
      callback();
    },
    focus: function (freshApi, state, callback) {
      onAddInReady(freshApi || api);
      callback();
    },
    blur: function () {},
    destroy: function () {
      runtime.api = null;
      setButtonsDisabled(true);
      setStatus("Add-in destroyed.");
    }
  };
};
