(function () {
  "use strict";

  var _api = null;
  var _logs = [];

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
    var raw;
    try {
      raw = await callApi("Get", { typeName: "HosLog", search: search, resultsLimit: 50000 });
    } catch (e) {
      log("HosLog failed, trying DutyStatusLog: " + e.message);
      raw = await callApi("Get", { typeName: "DutyStatusLog", search: search, resultsLimit: 50000 });
    }
    var rows = Array.isArray(raw) ? raw : [];
    log("HOS logs fetched: " + rows.length + " rows.");
    return rows;
  }

  async function fetchExceptions(range) {
    log("Fetching exceptions for " + formatDisplayDate(range.fromDate) + "…");
    var search = { fromDate: range.fromDate, toDate: range.toDate };
    var raw;
    try {
      raw = await callApi("Get", { typeName: "ExceptionEvent", search: search, resultsLimit: 50000 });
    } catch (e) {
      log("ExceptionEvent failed, trying ExceptionDetail: " + e.message);
      raw = await callApi("Get", { typeName: "ExceptionDetail", search: search, resultsLimit: 50000 });
    }
    var rows = Array.isArray(raw) ? raw : [];
    log("Exceptions fetched: " + rows.length + " rows.");
    return rows;
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
      [hosRows, exRows] = await Promise.all([
        fetchHosLogs(range),
        fetchExceptions(range)
      ]);
      buildAndDownloadWorkbook(flattenRows(hosRows), flattenRows(exRows), dateLabel);
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
    if (btn) btn.addEventListener("click", runReport);
    log(_api ? "Ready. Click Run Report to pull yesterday's data." : "Standalone preview — open inside MyGeotab to run reports.");
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
