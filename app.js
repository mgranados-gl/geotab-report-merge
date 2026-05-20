(function () {
  "use strict";

  var _api = null;
  var _logs = [];
  var _rules = [];
  var _drivers = [];
  var _vehicles = [];
  var _groups = [];

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
      opt.selected = false;
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

  function getSelectedIds(selectId) {
    var select = qs(selectId);
    if (!select) return [];
    var ids = [];
    for (var i = 0; i < select.options.length; i++) {
      if (select.options[i].selected) ids.push(select.options[i].value);
    }
    return ids;
  }

  function populateEntitySelect(containerId, selectId, items, noRowsText, labelField) {
    var container = qs(containerId);
    if (!container) return;
    if (!items || items.length === 0) {
      container.innerHTML = '<p class="hint">' + noRowsText + '</p>';
      return;
    }
    var select = document.createElement("select");
    select.id = selectId;
    select.multiple = true;
    select.size = Math.min(items.length, 8);
    items.forEach(function (item) {
      var opt = document.createElement("option");
      opt.value = item.id;
      opt.textContent = item[labelField] || item.name || item.id;
      opt.selected = false;
      select.appendChild(opt);
    });
    container.innerHTML = "";
    container.appendChild(select);
  }

  async function loadDrivers() {
    var container = qs("driversContainer");
    if (container) container.innerHTML = '<p class="hint">Loading drivers\u2026</p>';
    try {
      var raw = await callApi("Get", { typeName: "User", resultsLimit: 50000 });
      _drivers = Array.isArray(raw) ? raw.filter(function (u) {
        return !!u.id && (u.isDriver === true || u.isEULAAccepted === true || u.companyGroups);
      }) : [];
      _drivers.sort(function (a, b) {
        var an = (a.name || a.userName || "").toLowerCase();
        var bn = (b.name || b.userName || "").toLowerCase();
        return an.localeCompare(bn);
      });
      populateEntitySelect("driversContainer", "driversSelect", _drivers, "No drivers found.", "name");
      log("Loaded " + _drivers.length + " drivers.");
    } catch (e) {
      log("Could not load drivers: " + e.message, "error");
      if (container) container.innerHTML = '<p class="hint error-text">Failed to load drivers.</p>';
    }
  }

  async function loadVehicles() {
    var container = qs("vehiclesContainer");
    if (container) container.innerHTML = '<p class="hint">Loading vehicles\u2026</p>';
    try {
      var raw = await callApi("Get", { typeName: "Device", resultsLimit: 50000 });
      _vehicles = Array.isArray(raw) ? raw.filter(function (d) { return !!d.id; }) : [];
      _vehicles.sort(function (a, b) {
        var an = (a.name || "").toLowerCase();
        var bn = (b.name || "").toLowerCase();
        return an.localeCompare(bn);
      });
      populateEntitySelect("vehiclesContainer", "vehiclesSelect", _vehicles, "No vehicles found.", "name");
      log("Loaded " + _vehicles.length + " vehicles.");
    } catch (e) {
      log("Could not load vehicles: " + e.message, "error");
      if (container) container.innerHTML = '<p class="hint error-text">Failed to load vehicles.</p>';
    }
  }

  async function loadGroups() {
    var container = qs("groupsContainer");
    if (container) container.innerHTML = '<p class="hint">Loading groups\u2026</p>';
    try {
      var raw = await callApi("Get", { typeName: "Group", resultsLimit: 50000 });
      _groups = Array.isArray(raw) ? raw.filter(function (g) { return !!g.id && !!g.name; }) : [];
      _groups.sort(function (a, b) { return a.name.localeCompare(b.name); });
      populateEntitySelect("groupsContainer", "groupsSelect", _groups, "No groups found.", "name");
      log("Loaded " + _groups.length + " groups.");
    } catch (e) {
      log("Could not load groups: " + e.message, "error");
      if (container) container.innerHTML = '<p class="hint error-text">Failed to load groups.</p>';
    }
  }

  function getRefId(ref) {
    if (!ref) return null;
    if (typeof ref === "string") return ref;
    if (typeof ref === "object" && ref.id) return ref.id;
    return null;
  }

  function getRowDriverId(row) {
    return getRefId(row.driver) || getRefId(row.user) || getRefId(row.userId) || null;
  }

  function getRowVehicleId(row) {
    return getRefId(row.device) || getRefId(row.vehicle) || getRefId(row.deviceId) || null;
  }

  function getRowGroupIds(row) {
    var refs = [];
    function addGroupRefs(value) {
      if (!value) return;
      if (Array.isArray(value)) {
        value.forEach(addGroupRefs);
        return;
      }
      if (typeof value === "object") {
        var id = getRefId(value);
        if (id) refs.push(id);
      }
    }
    addGroupRefs(row.groups);
    addGroupRefs(row.companyGroups);
    if (row.driver) {
      addGroupRefs(row.driver.groups);
      addGroupRefs(row.driver.companyGroups);
    }
    if (row.user) {
      addGroupRefs(row.user.groups);
      addGroupRefs(row.user.companyGroups);
    }
    if (row.device) {
      addGroupRefs(row.device.groups);
      addGroupRefs(row.device.companyGroups);
    }
    return refs;
  }

  function filterRows(rows, filters) {
    var driverSet = {};
    var vehicleSet = {};
    var groupSet = {};
    var i;

    for (i = 0; i < filters.driverIds.length; i++) driverSet[filters.driverIds[i]] = true;
    for (i = 0; i < filters.vehicleIds.length; i++) vehicleSet[filters.vehicleIds[i]] = true;
    for (i = 0; i < filters.groupIds.length; i++) groupSet[filters.groupIds[i]] = true;

    var hasDriverFilter = filters.driverIds.length > 0;
    var hasVehicleFilter = filters.vehicleIds.length > 0;
    var hasGroupFilter = filters.groupIds.length > 0;

    return rows.filter(function (row) {
      if (hasDriverFilter) {
        var driverId = getRowDriverId(row);
        if (!driverId || !driverSet[driverId]) return false;
      }
      if (hasVehicleFilter) {
        var vehicleId = getRowVehicleId(row);
        if (!vehicleId || !vehicleSet[vehicleId]) return false;
      }
      if (hasGroupFilter) {
        var rowGroupIds = getRowGroupIds(row);
        var hasMatch = false;
        for (var g = 0; g < rowGroupIds.length; g++) {
          if (groupSet[rowGroupIds[g]]) {
            hasMatch = true;
            break;
          }
        }
        if (!hasMatch) return false;
      }
      return true;
    });
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

  function getTemplateUrl() {
    // Construct the template URL based on the current page location
    var currentUrl = window.location.href;
    var baseUrl = currentUrl.substring(0, currentUrl.lastIndexOf("/"));
    return baseUrl + "/templates/GapReportTemplate.xlsx";
  }

  async function buildAndDownloadWorkbook(hosRows, exRows, dateLabel) {
    if (typeof XLSX === "undefined") {
      throw new Error("SheetJS (XLSX) library not loaded.");
    }
    log("Building workbook…");
    var wb;
    var templateUrl = getTemplateUrl();

    try {
      // Fetch and load the template
      log("Loading template from " + templateUrl + "…");
      var response = await fetch(templateUrl);
      if (!response.ok) {
        throw new Error("Failed to load template: " + response.statusText);
      }
      var arrayBuffer = await response.arrayBuffer();
      wb = XLSX.read(arrayBuffer, { type: "array" });
      log("Template loaded successfully.");
    } catch (e) {
      log("Warning: Could not load template (" + e.message + "), creating blank workbook…");
      wb = XLSX.utils.book_new();
    }

    // Populate or create Data1 sheet with HOS logs
    var ws1 = hosRows.length
      ? XLSX.utils.json_to_sheet(hosRows)
      : XLSX.utils.aoa_to_sheet([["No HOS log data for this date."]]);
    if (wb.Sheets["Data1"]) {
      wb.Sheets["Data1"] = ws1;
    } else {
      XLSX.utils.book_append_sheet(wb, ws1, "Data1");
    }

    // Populate or create Data2 sheet with exceptions
    var ws2 = exRows.length
      ? XLSX.utils.json_to_sheet(exRows)
      : XLSX.utils.aoa_to_sheet([["No exception data for this date."]]);
    if (wb.Sheets["Data2"]) {
      wb.Sheets["Data2"] = ws2;
    } else {
      XLSX.utils.book_append_sheet(wb, ws2, "Data2");
    }

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
      var selectedDriverIds = getSelectedIds("driversSelect");
      var selectedVehicleIds = getSelectedIds("vehiclesSelect");
      var selectedGroupIds = getSelectedIds("groupsSelect");
      if (selectedRuleIds.length === 0) {
        throw new Error("Please select at least one exception rule before running.");
      }
      log("Running for " + selectedRuleIds.length + " selected rule(s)\u2026");
      [hosRows, exRows] = await Promise.all([
        fetchHosLogs(range),
        fetchExceptions(range, selectedRuleIds)
      ]);
      hosRows = filterRows(hosRows, {
        driverIds: selectedDriverIds,
        vehicleIds: selectedVehicleIds,
        groupIds: selectedGroupIds
      });
      exRows = filterRows(exRows, {
        driverIds: selectedDriverIds,
        vehicleIds: selectedVehicleIds,
        groupIds: selectedGroupIds
      });
      var flatHos = flattenRows(hosRows);
      var flatEx  = flattenRows(exRows);
      log("After filters \u2014 HOS rows: " + flatHos.length + " | Exception rows: " + flatEx.length);
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

    function bindSelectButtons(selectAllId, clearId, selectId) {
      var selectAllBtn = qs(selectAllId);
      var clearBtn = qs(clearId);
      if (selectAllBtn) selectAllBtn.addEventListener("click", function () {
        var sel = qs(selectId);
        if (sel) for (var i = 0; i < sel.options.length; i++) sel.options[i].selected = true;
      });
      if (clearBtn) clearBtn.addEventListener("click", function () {
        var sel = qs(selectId);
        if (sel) for (var i = 0; i < sel.options.length; i++) sel.options[i].selected = false;
      });
    }

    bindSelectButtons("selectAllDriversBtn", "clearDriversBtn", "driversSelect");
    bindSelectButtons("selectAllVehiclesBtn", "clearVehiclesBtn", "vehiclesSelect");
    bindSelectButtons("selectAllGroupsBtn", "clearGroupsBtn", "groupsSelect");

    if (_api) {
      log("Loading filters and exception rules\u2026");
      Promise.all([
        loadExceptionRules(),
        loadDrivers(),
        loadVehicles(),
        loadGroups()
      ]).then(function () {
        if (btn) btn.disabled = false;
      });
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
