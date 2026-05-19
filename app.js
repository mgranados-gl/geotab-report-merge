(function () {
  "use strict";

  var _api = null;
  var _logs = [];
  var _rules = [];

  // ── Helpers ────────────────────────────────────────────────────────────────

  function qs(id) { return document.getElementById(id); }

  function getYesterdayRange() {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    return { fromDate: yesterday, toDate: today };
  }

  function formatDisplayDate(dateObj) {
    return dateObj.toLocaleDateString(undefined, {
      weekday: "short", year: "numeric", month: "short", day: "numeric"
    });
  }

  function log(msg, kind) {
    var ts = new Date().toLocaleTimeString();
    _logs.push("[" + ts + "] " + msg);
    if (_logs.length > 300) _logs.shift();
    var box = qs("statusBox");
    if (!box) return;
    box.textContent = _logs.join("\n");
    box.className = "status-box" + (kind ? " " + kind : "");
    box.scrollTop = box.scrollHeight;
  }

  // ── API ────────────────────────────────────────────────────────────────────

  function callApi(method, params) {
    return new Promise(function (resolve, reject) {
      if (!_api) { reject(new Error("No Geotab API context.")); return; }
      _api.call(method, params,
        function (result) { resolve(result); },
        function (err) { reject(new Error(String(err && err.message ? err.message : err))); }
      );
    });
  }

  // ── Data fetching ───────────────────────────────────────────────────────────

  async function fetchHosLogs(range) {
    log("Fetching HOS logs for " + formatDisplayDate(range.fromDate) + "…");
    var search = { fromDate: range.fromDate, toDate: range.toDate };
    var raw = await callApi("Get", { typeName: "DutyStatusLog", search: search, resultsLimit: 50000 });
    var rows = Array.isArray(raw) ? raw : [];
    log("HOS logs fetched: " + rows.length + " rows.");
    return rows;
  }

  async function fetchExceptions(range, selectedRuleIds) {
    var search = { fromDate: range.fromDate, toDate: range.toDate };
    var raw;
    try {
      raw = await callApi("Get", { typeName: "ExceptionEvent", search: search, resultsLimit: 50000 });
    } catch (e) {
      log("ExceptionEvent failed, trying ExceptionDetail: " + e.message);
      raw = await callApi("Get", { typeName: "ExceptionDetail", search: search, resultsLimit: 50000 });
    }
    var rows = Array.isArray(raw) ? raw : [];
    if (selectedRuleIds && selectedRuleIds.length > 0) {
      var idSet = {};
      selectedRuleIds.forEach(function (id) { idSet[id] = true; });
      rows = rows.filter(function (row) {
        var ruleId = row.rule && row.rule.id ? row.rule.id
          : row.ruleId ? row.ruleId : null;
        return ruleId && idSet[ruleId];
      });
      log("After rule filter: " + rows.length + " rows.");
    } else {
      log("Exceptions fetched: " + rows.length + " rows.");
    }
    return rows;
  }

  // ── Exception rules ──────────────────────────────────────────────────────────

  async function loadExceptionRules() {
    var container = qs("rulesContainer");
    var runBtn = qs("runBtn");
    if (container) container.innerHTML = '<p class="hint">Loading exception rules\u2026</p>';
    try {
      var raw = await callApi("Get", { typeName: "Rule", resultsLimit: 5000 });
      _rules = Array.isArray(raw) ? raw.filter(function (r) { return r.name; }) : [];
      _rules.sort(function (a, b) { return a.name.localeCompare(b.name); });
      populateRulesList();
      log("Loaded " + _rules.length + " exception rules.");
    } catch (e) {
      log("Could not load exception rules: " + e.message, "error");
      if (container) container.innerHTML = '<p class="hint error-text">Failed to load rules.</p>';
    }
    if (runBtn) runBtn.disabled = false;
  }

  function populateRulesList() {
    var container = qs("rulesContainer");
    if (!container) return;
    if (_rules.length === 0) {
      container.innerHTML = '<p class="hint">No exception rules found.</p>';
      return;
    }
    var select = document.createElement("select");
    select.id = "rulesSelect";
    select.multiple = true;
    select.size = Math.min(_rules.length, 8);
    _rules.forEach(function (rule) {
      var opt = document.createElement("option");
      opt.value = rule.id;
      opt.textContent = rule.name;
      opt.selected = true;
      select.appendChild(opt);
    });
    container.innerHTML = "";
    container.appendChild(select);
  }

  function getSelectedRuleIds() {
    var select = qs("rulesSelect");
    if (!select) return [];
    var ids = [];
    for (var i = 0; i < select.options.length; i++) {
      if (select.options[i].selected) ids.push(select.options[i].value);
    }
    return ids;
  }

  // ── Flattening ──────────────────────────────────────────────────────────────
  // Recursively flattens nested objects so each row is a plain key→value map.
  // Arrays are JSON-stringified; Date objects are left as-is for SheetJS.

  function flattenObject(obj, prefix) {
    prefix = prefix || "";
    var out = {};
    Object.keys(obj).forEach(function (k) {
      var val = obj[k];
      var key = prefix ? prefix + "_" + k : k;
      if (val === null || val === undefined) {
        out[key] = val;
      } else if (val instanceof Date) {
        out[key] = val;
      } else if (Array.isArray(val)) {
        out[key] = val.length ? JSON.stringify(val) : "";
      } else if (typeof val === "object") {
        Object.assign(out, flattenObject(val, key));
      } else {
        out[key] = val;
      }
    });
    return out;
  }

  function flattenRows(rows) {
    return rows.map(function (row) { return flattenObject(row); });
  }

  // ── Driver sort / filter ────────────────────────────────────────────────────

  var DRIVER_NAME_FIELDS = [
    "driver_lastName", "driver_name", "driver_firstName",
    "userName", "lastName", "driverName"
  ];

  function getDriverSortKey(row) {
    for (var i = 0; i < DRIVER_NAME_FIELDS.length; i++) {
      var v = row[DRIVER_NAME_FIELDS[i]];
      if (v && String(v).trim()) return String(v).trim().toLowerCase();
    }
    return null;
  }

  function filterAndSortByDriver(rows) {
    var copy = rows.slice();
    copy.sort(function (a, b) {
      var ka = getDriverSortKey(a) || "\uffff";
      var kb = getDriverSortKey(b) || "\uffff";
      return ka.localeCompare(kb);
    });
    return copy;
  }

  // ── Excel export ────────────────────────────────────────────────────────────

  function buildAndDownloadWorkbook(hosRows, exRows, dateLabel) {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS (XLSX) library not loaded.");
    }
    log("Building workbook…");
    var wb = XLSX.utils.book_new();

    var ws1 = hosRows.length
      ? XLSX.utils.json_to_sheet(hosRows)
      : XLSX.utils.aoa_to_sheet([["No HOS log data for this date."]]);
    XLSX.utils.book_append_sheet(wb, ws1, "Data1");

    var ws2 = exRows.length
      ? XLSX.utils.json_to_sheet(exRows)
      : XLSX.utils.aoa_to_sheet([["No exception data for this date."]]);
    XLSX.utils.book_append_sheet(wb, ws2, "Data2");

    var wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    var blob = new Blob([wbout], { type: "application/octet-stream" });
    var fileName = "GapReport-" + dateLabel + ".xlsx";
    var link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
    log("Download started: " + fileName, "success");
  }

  // ── Run ─────────────────────────────────────────────────────────────────────

  async function runReport() {
    var btn = qs("runBtn");
    btn.disabled = true;
    btn.textContent = "Running…";
    _logs = [];
    try {
      if (!_api) throw new Error("Add-in must run inside MyGeotab to access the API.");
      var range = getYesterdayRange();
      var dateLabel = range.fromDate.toISOString().slice(0, 10);
      var hosRows, exRows;
      var selectedRuleIds = getSelectedRuleIds();
      if (selectedRuleIds.length === 0) {
        throw new Error("Please select at least one exception rule before running.");
      }
      log("Running for " + selectedRuleIds.length + " selected rule(s)\u2026");
      [hosRows, exRows] = await Promise.all([
        fetchHosLogs(range),
        fetchExceptions(range, selectedRuleIds)
      ]);
      var flatHos = flattenRows(hosRows);
      var flatEx  = flattenRows(exRows);
      log("HOS rows: " + flatHos.length + " | Exception rows: " + flatEx.length);
      buildAndDownloadWorkbook(flatHos, flatEx, dateLabel);
    } catch (err) {
      log("ERROR: " + (err.message || String(err)), "error");
    } finally {
      btn.disabled = false;
      btn.textContent = "Run Report";
    }
  }

  // ── UI init ─────────────────────────────────────────────────────────────────

  function initUi() {
    var range = getYesterdayRange();
    var label = qs("dateLabel");
    if (label) label.textContent = formatDisplayDate(range.fromDate);
    var btn = qs("runBtn");
    if (btn) {
      btn.disabled = true;
      btn.addEventListener("click", runReport);
    }
    var selectAll = qs("selectAllBtn");
    var clearAll  = qs("clearAllBtn");
    if (selectAll) selectAll.addEventListener("click", function () {
      var sel = qs("rulesSelect");
      if (sel) for (var i = 0; i < sel.options.length; i++) sel.options[i].selected = true;
    });
    if (clearAll) clearAll.addEventListener("click", function () {
      var sel = qs("rulesSelect");
      if (sel) for (var i = 0; i < sel.options.length; i++) sel.options[i].selected = false;
    });
    if (_api) {
      log("Loading exception rules\u2026");
      loadExceptionRules();
    } else {
      log("Standalone preview \u2014 open inside MyGeotab to run reports.");
      if (btn) btn.disabled = false;
    }
  }

  // ── Geotab add-in registration ───────────────────────────────────────────────

  geotab.addin.gap_report = function (api) {
    _api = api;
    return {
      initialize: function (freshApi, freshState, callback) {
        _api = freshApi;
        if (document.readyState === "loading") {
          document.addEventListener("DOMContentLoaded", function () { initUi(); callback(); });
        } else {
          initUi();
          callback();
        }
      },
      focus: function () {},
      blur: function () {}
    };
  };

})();
