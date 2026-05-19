(function () {
  "use strict";

  var state = {
    api: null,
    geotabState: null,
    logs: []
  };

  var controls = {};

  var DEFAULT_TEMPLATE_URL = "./template.xlsx";

  function $(id) {
    return document.getElementById(id);
  }

  function appendStatus(message, kind) {
    state.logs.push("[" + new Date().toLocaleTimeString() + "] " + message);
    if (state.logs.length > 200) {
      state.logs.shift();
    }
    var box = controls.statusBox;
    box.textContent = state.logs.join("\n");
    box.classList.remove("error", "success");
    if (kind === "error") box.classList.add("error");
    if (kind === "success") box.classList.add("success");
  }

  function setDateRangeFromPreset(preset) {
    var now = new Date();
    var toDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var fromDate = new Date(toDate);
    if (preset === "today") {
      fromDate = new Date(toDate);
    } else if (preset === "last7") {
      fromDate.setDate(fromDate.getDate() - 6);
    } else if (preset === "last30") {
      fromDate.setDate(fromDate.getDate() - 29);
    } else {
      return;
    }
    controls.fromDate.value = formatDateInput(fromDate);
    controls.toDate.value = formatDateInput(toDate);
  }

  function formatDateInput(dateObj) {
    var yyyy = dateObj.getFullYear();
    var mm = String(dateObj.getMonth() + 1).padStart(2, "0");
    var dd = String(dateObj.getDate()).padStart(2, "0");
    return yyyy + "-" + mm + "-" + dd;
  }

  function parseDateRange() {
    if (!controls.fromDate.value || !controls.toDate.value) {
      throw new Error("From and To date are required.");
    }
    var fromDate = new Date(controls.fromDate.value + "T00:00:00");
    var toDate = new Date(controls.toDate.value + "T23:59:59");
    if (isNaN(fromDate.getTime()) || isNaN(toDate.getTime())) {
      throw new Error("Invalid date range.");
    }
    if (fromDate > toDate) {
      throw new Error("From date cannot be after To date.");
    }
    return { fromDate: fromDate, toDate: toDate };
  }

  function callApi(method, params) {
    return new Promise(function (resolve, reject) {
      if (!state.api) {
        reject(new Error("Geotab API context is not available."));
        return;
      }
      try {
        state.api.call(method, params, function (result) {
          if (result && result.errors) {
            reject(new Error("API error: " + JSON.stringify(result.errors)));
            return;
          }
          resolve(result);
        }, function (error) {
          var msg = error && error.message ? error.message : String(error);
          reject(new Error("API call failed for " + method + ": " + msg));
        });
      } catch (err) {
        reject(new Error("API call exception: " + (err.message || String(err))));
      }
    });
  }

  function buildHosLogSearch(range) {
    var search = {
      fromDate: range.fromDate,
      toDate: range.toDate
    };
    return search;
  }

  async function fetchHosLog(range) {
    var search = buildHosLogSearch(range);
    appendStatus("Fetching HOSLog report...");
    var rows = [];
    try {
      var raw = await callApi("Get", {
        typeName: "HosLog",
        search: search,
        resultsLimit: 50000
      });
      rows = (Array.isArray(raw) ? raw : []);
    } catch (err) {
      appendStatus("Could not fetch HosLog: " + (err.message || String(err)) + ". Trying HosLogRecord...");
      try {
        var raw = await callApi("Get", {
          typeName: "HosLogRecord",
          search: search,
          resultsLimit: 50000
        });
        rows = (Array.isArray(raw) ? raw : []);
      } catch (err2) {
        appendStatus("HosLogRecord also failed: " + (err2.message || String(err2)), "error");
        throw err2;
      }
    }
    appendStatus("HOSLog rows: " + rows.length);
    return rows;
  }

  async function readTemplateArrayBuffer() {
    var selectedFile = controls.templateFile.files && controls.templateFile.files[0];
    if (selectedFile) {
      appendStatus("Using uploaded template: " + selectedFile.name);
      return selectedFile.arrayBuffer();
    }
    appendStatus("Fetching default template from " + DEFAULT_TEMPLATE_URL + "...");
    var response = await fetch(DEFAULT_TEMPLATE_URL);
    if (!response.ok) {
      throw new Error("Could not load default template. Upload a template file or host one at " + DEFAULT_TEMPLATE_URL);
    }
    return response.arrayBuffer();
  }

  function ensureSheet(workbook, sheetName) {
    var existing = workbook.sheet(sheetName);
    if (existing) return existing;
    return workbook.addSheet(sheetName);
  }

  function clearRows(sheet) {
    var usedRange = sheet.usedRange();
    if (usedRange) usedRange.clear();
  }

  function writeRows(sheet, rows) {
    if (!rows.length) return { writtenRows: 0, writtenCols: 0 };
    var headers = Object.keys(rows[0]);
    sheet.row(1).cell(1).value(headers);
    var values = rows.map(function (row) {
      return headers.map(function (header) {
        return row[header];
      });
    });
    sheet.cell(2, 1).value(values);
    return {
      writtenRows: rows.length + 1,
      writtenCols: headers.length
    };
  }

  async function exportWorkbook(hosRows) {
    var templateBuffer = await readTemplateArrayBuffer();
    var workbook = await XlsxPopulate.fromDataAsync(templateBuffer);
    var dataSheet = ensureSheet(workbook, "HOSLog");
    clearRows(dataSheet);
    var hosWrite = writeRows(dataSheet, hosRows);
    var outputName = (controls.outputFileName.value || "HOS-Log-Report.xlsx").trim();
    var exportBuffer = await workbook.outputAsync();
    var blob = new Blob([exportBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    var link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = outputName.toLowerCase().endsWith(".xlsx") ? outputName : (outputName + ".xlsx");
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
    appendStatus(
      "Workbook generated with " + hosWrite.writtenRows + " rows.",
      "success"
    );
  }

  async function runExport() {
    try {
      controls.runExportBtn.disabled = true;
      controls.runExportBtn.textContent = "Generating...";
      var range = parseDateRange();
      appendStatus("Fetching HOS log report for " + range.fromDate.toISOString() + " to " + range.toDate.toISOString());
      if (!state.api) {
        throw new Error("This add-in must run inside MyGeotab to fetch report data.");
      }
      var hosRows = await fetchHosLog(range);
      await exportWorkbook(hosRows);
    } catch (error) {
      var message = (error && error.message) ? error.message : String(error);
      appendStatus("Export failed: " + message, "error");
      console.error(error);
    } finally {
      controls.runExportBtn.disabled = false;
      controls.runExportBtn.textContent = "Generate Excel";
    }
  }

  function wireEvents() {
    controls.datePreset.addEventListener("change", function () {
      if (controls.datePreset.value !== "custom") {
        setDateRangeFromPreset(controls.datePreset.value);
      }
    });
    controls.runExportBtn.addEventListener("click", function () {
      runExport();
    });
  }

  function cacheControls() {
    controls.datePreset = $("datePreset");
    controls.fromDate = $("fromDate");
    controls.toDate = $("toDate");
    controls.templateFile = $("templateFile");
    controls.outputFileName = $("outputFileName");
    controls.runExportBtn = $("runExportBtn");
    controls.statusBox = $("statusBox");
    controls.envPill = $("envPill");
  }

  function initializeUi() {
    cacheControls();
    setDateRangeFromPreset("last7");
    wireEvents();
    controls.runExportBtn.disabled = false;
    if (state.api) {
      controls.envPill.textContent = "MyGeotab";
      appendStatus("Add-in ready. Select dates and export HOS log data.", "success");
    } else {
      controls.envPill.textContent = "Standalone";
      appendStatus("Running in standalone mode. Open inside MyGeotab add-in context to fetch report data.");
    }
  }

  function bootAddin(api, geotabState) {
    state.api = api;
    state.geotabState = geotabState;
    console.log("bootAddin called with api:", !!api, "geotabState:", !!geotabState);
    appendStatus("Initializing add-in...");
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeUi);
    } else {
      initializeUi();
    }
  }

  if (window.geotab && window.geotab.addin) {
    console.log("Registering HOSLogReport addin");
    window.geotab.addin.HOSLogReport = function (api, geotabState) {
      console.log("HOSLogReport addin function called");
      return {
        initialize: function () {
          console.log("HOSLogReport initialize called");
          bootAddin(api, geotabState);
        },
        focus: function () {
          appendStatus("Add-in focused.");
        },
        blur: function () {
          appendStatus("Add-in blurred.");
        }
      };
    };
  } else {
    console.log("geotab.addin namespace not available - add-in running in standalone mode");
  }

  // Standalone fallback for local UI preview.
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      if (!state.api) {
        initializeUi();
      }
    });
  } else if (!state.api) {
    initializeUi();
  }
})();
