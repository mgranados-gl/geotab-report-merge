(function () {
  "use strict";

  var state = {
    api: null,
    geotabState: null,
    logs: [],
    groups: [],
    exceptionRules: []
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

  function getSelectedGroupIds() {
    var values = [];
    for (var i = 0; i < controls.groupSelect.options.length; i += 1) {
      var option = controls.groupSelect.options[i];
      if (option.selected) values.push(option.value);
    }
    return values;
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

  async function loadMetadata() {
    if (!state.api) {
      appendStatus("No Geotab API context. Metadata load skipped.");
      return;
    }
    try {
      appendStatus("Loading groups...");
      var groups = null;
      var groupError = null;
      
      // Try Group first
      try {
        groups = await callApi("Get", {
          typeName: "Group",
          resultsLimit: 5000
        });
      } catch (err) {
        groupError = err;
        appendStatus("Group entity failed, trying CompanyGroup...");
        try {
          groups = await callApi("Get", {
            typeName: "CompanyGroup",
            resultsLimit: 5000
          });
        } catch (err2) {
          throw new Error("Could not fetch groups: " + err.message + " (also tried CompanyGroup)");
        }
      }
      
      state.groups = Array.isArray(groups) ? groups : [];
      controls.groupSelect.innerHTML = "";
      state.groups.forEach(function (group) {
        var option = document.createElement("option");
        option.value = group.id;
        option.textContent = group.name || group.id;
        controls.groupSelect.appendChild(option);
      });
      appendStatus("Loaded " + state.groups.length + " groups.");
      
      appendStatus("Loading exception rules...");
      var rules = null;
      try {
        rules = await callApi("Get", {
          typeName: "Rule",
          resultsLimit: 5000
        });
      } catch (err) {
        appendStatus("Could not load rules (non-critical): " + err.message);
        rules = [];
      }
      
      state.exceptionRules = Array.isArray(rules) ? rules : [];
      controls.exceptionRuleSelect.innerHTML = "<option value=\"\">All rules</option>";
      state.exceptionRules.forEach(function (rule) {
        var option = document.createElement("option");
        option.value = rule.id;
        option.textContent = rule.name || rule.id;
        controls.exceptionRuleSelect.appendChild(option);
      });
      
      appendStatus("Metadata loaded successfully. Groups: " + state.groups.length + ", Rules: " + state.exceptionRules.length, "success");
      return true;
    } catch (err) {
      var errMsg = (err && err.message) ? err.message : String(err);
      appendStatus("ERROR loading metadata: " + errMsg, "error");
      console.error("loadMetadata error:", err);
      throw err;
    }
  }

  function buildExceptionsDetailSearch(range, groupIds) {
    var search = {
      fromDate: range.fromDate,
      toDate: range.toDate
    };
    // Note: Some databases may not support these search parameters
    // Omit if causing API errors
    if (groupIds.length > 0) {
      try {
        search.deviceSearch = {
          groups: groupIds.map(function (id) { return ({ id: id }); })
        };
      } catch (e) {
        console.warn("Could not build deviceSearch:", e);
      }
    }
    if (controls.exceptionRuleSelect.value) {
      try {
        search.ruleSearch = { id: controls.exceptionRuleSelect.value };
      } catch (e) {
        console.warn("Could not build ruleSearch:", e);
      }
    }
    if (!controls.includeAcknowledged.checked) {
      search.isAcknowledged = false;
    }
    return search;
  }

  async function fetchExceptionsDetail(range, groupIds) {
    var search = buildExceptionsDetailSearch(range, groupIds);
    appendStatus("Fetching ExceptionsDetail report...");
    var rows = [];
    try {
      var raw = await callApi("Get", {
        typeName: "ExceptionDetail",
        search: search,
        resultsLimit: 50000
      });
      rows = (Array.isArray(raw) ? raw : []).map(function (event) {
        return Object.assign({}, event);
      });
    } catch (err) {
      appendStatus("Could not fetch ExceptionsDetail: " + (err.message || String(err)) + ". Trying ExceptionEvent...");
      try {
        var raw = await callApi("Get", {
          typeName: "ExceptionEvent",
          search: search,
          resultsLimit: 50000
        });
        rows = (Array.isArray(raw) ? raw : []).map(function (event) {
          return Object.assign({}, event);
        });
      } catch (err2) {
        appendStatus("ExceptionEvent also failed: " + (err2.message || String(err2)), "error");
        throw err2;
      }
    }
    appendStatus("ExceptionsDetail rows: " + rows.length);
    return rows;
  }

  function buildHosLogSearch(range, groupIds) {
    var search = {
      fromDate: range.fromDate,
      toDate: range.toDate
    };
    if (groupIds.length > 0) {
      search.userSearch = {
        companyGroups: groupIds.map(function (id) { return ({ id: id }); })
      };
    }
    return search;
  }

  async function fetchHosLog(range, groupIds) {
    var search = buildHosLogSearch(range, groupIds);
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
    var statusFilter = controls.hosStatusFilter.value;
    rows = rows.filter(function (row) {
      if (!statusFilter) return true;
      return (row.status || row.dutyStatus || "").toLowerCase() === statusFilter;
    });
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

  async function exportWorkbook(exceptionsRows, hosRows) {
    var templateBuffer = await readTemplateArrayBuffer();
    var workbook = await XlsxPopulate.fromDataAsync(templateBuffer);
    var data1Sheet = ensureSheet(workbook, "Data1");
    var data2Sheet = ensureSheet(workbook, "Data2");
    clearRows(data1Sheet);
    clearRows(data2Sheet);
    var exWrite = writeRows(data1Sheet, exceptionsRows);
    var hosWrite = writeRows(data2Sheet, hosRows);
    var outputName = (controls.outputFileName.value || "Geotab-Report-Merge.xlsx").trim();
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
      "Workbook generated. Data1 rows: " + exWrite.writtenRows + ", Data2 rows: " + hosWrite.writtenRows,
      "success"
    );
  }

  async function runExport() {
    try {
      controls.runExportBtn.disabled = true;
      controls.runExportBtn.textContent = "Generating...";
      var range = parseDateRange();
      var groupIds = getSelectedGroupIds();
      appendStatus("Starting report pull for " + range.fromDate.toISOString() + " to " + range.toDate.toISOString());
      appendStatus("Selected groups: " + (groupIds.length ? groupIds.join(", ") : "All"));
      if (!state.api) {
        throw new Error("This add-in must run inside MyGeotab to fetch report data.");
      }
      var exceptionsRows = await fetchExceptionsDetail(range, groupIds);
      var hosRows = await fetchHosLog(range, groupIds);
      await exportWorkbook(exceptionsRows, hosRows);
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
    controls.loadMetaBtn.addEventListener("click", function () {
      loadMetadata().catch(function (error) {
        appendStatus("Metadata load failed: " + (error.message || String(error)), "error");
      });
    });
    controls.runExportBtn.addEventListener("click", function () {
      runExport();
    });
  }

  function cacheControls() {
    controls.datePreset = $("datePreset");
    controls.fromDate = $("fromDate");
    controls.toDate = $("toDate");
    controls.groupSelect = $("groupSelect");
    controls.exceptionRuleSelect = $("exceptionRuleSelect");
    controls.includeAcknowledged = $("includeAcknowledged");
    controls.hosStatusFilter = $("hosStatusFilter");
    controls.templateFile = $("templateFile");
    controls.outputFileName = $("outputFileName");
    controls.loadMetaBtn = $("loadMetaBtn");
    controls.runExportBtn = $("runExportBtn");
    controls.statusBox = $("statusBox");
    controls.envPill = $("envPill");
  }

  function initializeUi() {
    cacheControls();
    setDateRangeFromPreset("last7");
    wireEvents();
    controls.runExportBtn.disabled = false;
    controls.loadMetaBtn.disabled = false;
    if (state.api) {
      controls.envPill.textContent = "MyGeotab";
      loadMetadata().then(function() {
        appendStatus("Add-in ready.", "success");
        controls.runExportBtn.disabled = false;
        controls.loadMetaBtn.disabled = false;
      }).catch(function (error) {
        appendStatus("Metadata load failed: " + (error.message || String(error)), "error");
        controls.runExportBtn.disabled = false;
        controls.loadMetaBtn.disabled = false;
      });
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
    console.log("Registering DualReportMerge addin");
    window.geotab.addin.DualReportMerge = function (api, geotabState) {
      console.log("DualReportMerge addin function called");
      return {
        initialize: function () {
          console.log("DualReportMerge initialize called");
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
