(function () {
  "use strict";

  var APP_VERSION = "1.0.9";

  var _api = null;
  var _logs = [];
  var _rules = [];
  var _drivers = [];

  // ── Helpers ────────────────────────────────────────────────────────────────

  function qs(id) { return document.getElementById(id); }

  function setBannerStatus(status, message) {
    var banner = qs("accessBanner");
    var bannerText = qs("bannerText");
    if (banner && bannerText) {
      banner.className = "access-banner access-banner--" + status;
      bannerText.textContent = message;
    }
  }

  function getYesterdayRange(userTimeZone) {
    // Get today and yesterday in the user's time zone
    var today = new Date();
    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);

    // Format dates as UTC-equivalent for the user's timezone
    // This ensures the dates are interpreted in the user's local timezone
    var fromDate = new Date(yesterday.getFullYear(), yesterday.getMonth(), yesterday.getDate(), 0, 0, 0, 0);
    var toDate = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0, 0);

    return { fromDate: fromDate, toDate: toDate, userTimeZone: userTimeZone };
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
    
    // Log a sample of raw state values to verify what the API returns
    var sampleStates = {};
    rows.forEach(function (row) { sampleStates[row.status || row.state || row.dutystatus || "(none)"] = true; });
    log("Raw state values from API: " + Object.keys(sampleStates).join(", "));

    // Filter by allowed states (Geotab API enum values)
    var allowedStates = ["ON", "D", "INT_D", "Login", "Logoff"];
    rows = rows.filter(function (row) {
    var state = row.status || row.state || row.dutystatus || "";
      return allowedStates.indexOf(state) >= 0;
    });
    log("After state filter: " + rows.length + " rows.");
    
    return rows;
  }

  async function fetchExceptions(range) {
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
    
    // Filter by hardcoded rule IDs (Entering Zone Office and Exiting Zone Office)
    var targetRuleIds = ["aoL1tECTjFEqNGLqrZ_F8Ew", "a4cQ5-5kHtkOTFz8HKO2iXQ"];
    var idSet = {};
    targetRuleIds.forEach(function (id) { idSet[id] = true; });
    rows = rows.filter(function (row) {
      var ruleId = row.rule && row.rule.id ? row.rule.id
        : row.ruleId ? row.ruleId : null;
      return ruleId && idSet[ruleId];
    });
    log("After rule filter (Entering/Exiting Zone Office): " + rows.length + " rows.");
    
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

  function clearSelection(selectId) {
    var select = qs(selectId);
    if (!select) return;
    for (var i = 0; i < select.options.length; i++) {
      select.options[i].selected = false;
    }
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

  function getRefId(ref) {
    if (!ref) return null;
    if (typeof ref === "string") return ref;
    if (typeof ref === "object" && ref.id) return ref.id;
    return null;
  }

  function getRowDriverId(row) {
    return getRefId(row.driver) || getRefId(row.user) || getRefId(row.userId) || null;
  }

  function filterRows(rows, filters) {
    var driverSet = {};
    var i;

    for (i = 0; i < filters.driverIds.length; i++) driverSet[filters.driverIds[i]] = true;

    var hasDriverFilter = filters.driverIds.length > 0;

    return rows.filter(function (row) {
      if (hasDriverFilter) {
        var driverId = getRowDriverId(row);
        if (!driverId || !driverSet[driverId]) return false;
      }
      return true;
    });
  }

  // ── Lookup maps ─────────────────────────────────────────────────────────────

  async function fetchDriverMap() {
    try {
      log("Fetching driver details…");
      var raw = await callApi("Get", { typeName: "User", resultsLimit: 50000 });
      var map = {};
      if (Array.isArray(raw)) {
        raw.forEach(function (u) { if (u.id) map[u.id] = u; });
      }
      log("Driver details loaded: " + Object.keys(map).length + " users.");
      return map;
    } catch (e) {
      log("Warning: Could not fetch driver details: " + e.message);
      return {};
    }
  }

  async function fetchDeviceMap() {
    try {
      log("Fetching vehicle details…");
      var raw = await callApi("Get", { typeName: "Device", resultsLimit: 50000 });
      var map = {};
      if (Array.isArray(raw)) {
        raw.forEach(function (d) { if (d.id) map[d.id] = d; });
      }
      log("Vehicle details loaded: " + Object.keys(map).length + " vehicles.");
      return map;
    } catch (e) {
      log("Warning: Could not fetch vehicle details: " + e.message);
      return {};
    }
  }

  async function fetchRuleMap() {
    try {
      log("Fetching rule details…");
      var raw = await callApi("Get", { typeName: "Rule", resultsLimit: 5000 });
      var map = {};
      if (Array.isArray(raw)) {
        raw.forEach(function (r) { if (r.id) map[r.id] = r; });
      }
      log("Rule details loaded: " + Object.keys(map).length + " rules.");
      return map;
    } catch (e) {
      log("Warning: Could not fetch rule details: " + e.message);
      return {};
    }
  }

  // ── HOS row mapping ─────────────────────────────────────────────────────────
  // Maps raw DutyStatusLog API objects to clean, human-readable columns
  // matching the built-in Geotab HOS Log report format.

  function mapHosRow(row, driverMap, deviceMap) {
    var driverId = row.driver && row.driver.id ? row.driver.id : null;
    var driver = driverId ? (driverMap[driverId] || {}) : {};
    var deviceId = row.device && row.device.id ? row.device.id : null;
    var device = deviceId ? (deviceMap[deviceId] || {}) : {};
    var loc = row.location || {};
    var odometerKm = row.odometer ? Math.round(row.odometer / 1000 * 10) / 10 : "";
    var engineHrs = row.engineHours ? Math.round(row.engineHours / 3600 * 100) / 100 : "";
    return {
      "DateTime":        row.dateTime || "",
      "Status":          row.status || "",
      "DriverFirstName": driver.firstName || driver.name || "",
      "DriverLastName":  driver.lastName || "",
      "DriverUserName":  driver.userName || "",
      "Vehicle":         device.name || deviceId || "",
      "Location":        loc.address || (loc.y != null ? loc.y + ", " + loc.x : ""),
      "Odometer_km":     odometerKm,
      "EngineHours":     engineHrs,
      "Origin":          row.origin || "",
      "Sequence":        row.sequence || "",
      "Id":              row.id || ""
    };
  }

  function mapExceptionRow(row, driverMap, deviceMap, ruleMap) {
    var driverId = row.driver && row.driver.id ? row.driver.id : null;
    var driver = driverId ? (driverMap[driverId] || {}) : {};
    var deviceId = row.device && row.device.id ? row.device.id : null;
    var device = deviceId ? (deviceMap[deviceId] || {}) : {};
    var ruleId = row.rule && row.rule.id ? row.rule.id : null;
    var rule = ruleId ? (ruleMap[ruleId] || {}) : {};

    // Duration comes as ISO 8601 duration string e.g. "PT5M3S" — keep as-is; Excel can display it
    // Distance is already in km per API docs
    return {
      "ActiveFrom":       row.activeFrom || "",
      "ActiveTo":         row.activeTo || "",
      "Duration":         row.duration || "",
      "Distance_km":      row.distance != null ? Math.round(row.distance * 10) / 10 : "",
      "RuleName":         rule.name || ruleId || "",
      "DriverFirstName":  driver.firstName || driver.name || "",
      "DriverLastName":   driver.lastName || "",
      "DriverUserName":   driver.userName || "",
      "Vehicle":          device.name || deviceId || "",
      "DeviceComment":    device.comment || "",
      "Id":               row.id || ""
    };
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
    // Derive base URL from app.js script src (most reliable inside Geotab iframes)
    var scripts = document.getElementsByTagName("script");
    for (var i = 0; i < scripts.length; i++) {
      var src = scripts[i].src;
      if (src && src.indexOf("app.js") !== -1) {
        return src.substring(0, src.lastIndexOf("/")) + "/templates/GapReportTemplate.xlsx";
      }
    }
    // Fallback: strip query string and hash before deriving base
    var currentUrl = window.location.href.split("?")[0].split("#")[0];
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
        throw new Error("Failed to load template: HTTP " + response.status + (response.statusText ? " " + response.statusText : ""));
      }
      var arrayBuffer = await response.arrayBuffer();
      wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array", cellStyles: true, cellFormula: true, cellDates: true });
      log("Template loaded successfully.");
    } catch (e) {
      log("Warning: Could not load template (" + e.message + "), creating blank workbook (Notification and Summary tabs will be missing)…", "error");
      wb = XLSX.utils.book_new();
    }

    // Populate or create Data1 sheet with HOS logs
    // Write into the existing sheet (preserves template styles on other sheets)
    if (wb.Sheets["Data1"]) {
      // Clear existing data range then write new data starting at A1
      var ws1 = wb.Sheets["Data1"];
      if (hosRows.length) {
        XLSX.utils.sheet_add_json(ws1, hosRows, { origin: "A1", skipHeader: false });
      } else {
        XLSX.utils.sheet_add_aoa(ws1, [["No HOS log data for this date."]], { origin: "A1" });
      }
    } else {
      var ws1 = hosRows.length
        ? XLSX.utils.json_to_sheet(hosRows)
        : XLSX.utils.aoa_to_sheet([["No HOS log data for this date."]]);
      XLSX.utils.book_append_sheet(wb, ws1, "Data1");
    }

    // Populate or create Data2 sheet with exceptions
    if (wb.Sheets["Data2"]) {
      var ws2 = wb.Sheets["Data2"];
      if (exRows.length) {
        XLSX.utils.sheet_add_json(ws2, exRows, { origin: "A1", skipHeader: false });
      } else {
        XLSX.utils.sheet_add_aoa(ws2, [["No exception data for this date."]], { origin: "A1" });
      }
    } else {
      var ws2 = exRows.length
        ? XLSX.utils.json_to_sheet(exRows)
        : XLSX.utils.aoa_to_sheet([["No exception data for this date."]]);
      XLSX.utils.book_append_sheet(wb, ws2, "Data2");
    }

    var wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
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

      // Get user's session via the add-in session method
      var credentials = await new Promise(function (resolve) { _api.getSession(resolve); });
      var userName = credentials && credentials.userName ? credentials.userName.toLowerCase() : null;

      // Whitelist of allowed test users
      var allowedUsers = ["mgranados@gridline.com", "eborden@gridline.com"];
      var allowedUsersLower = allowedUsers.map(function (u) { return u.toLowerCase(); });

      if (!userName || allowedUsersLower.indexOf(userName) < 0) {
        setBannerStatus("denied", "✗ Access denied for " + (userName || "unknown user"));
        log("Access denied: " + (userName || "Unknown user") + " is not authorized to run this report.", "error");
        throw new Error("Unauthorized user.");
      }

      setBannerStatus("permitted", "✓ Access permitted for " + (credentials.userName || "user"));
      log("Access granted for " + userName + ".");

      // Time zone comes from the MyGeotab User record, fetched separately
      var userRecord = await callApi("Get", { typeName: "User", search: { name: credentials.userName }, resultsLimit: 1 });
      var userTimeZone = userRecord && userRecord[0] && userRecord[0].timeZoneId ? userRecord[0].timeZoneId : null;
      
      var range = getYesterdayRange(userTimeZone);
      var dateLabel = range.fromDate.toISOString().slice(0, 10);
      var hosRows, exRows;
      
      log("Running report for yesterday (" + dateLabel + "), using time zone: " + (userTimeZone || "default") + "…");
      log("Fetching HOS logs (ON, D, INT_D, Login, Logoff states only) for all drivers…");
      log("Fetching exception events (Entering/Exiting Zone Office) for all drivers…");
      
      var driverMap, deviceMap, ruleMap;
      [hosRows, exRows, driverMap, deviceMap, ruleMap] = await Promise.all([
        fetchHosLogs(range),
        fetchExceptions(range),
        fetchDriverMap(),
        fetchDeviceMap(),
        fetchRuleMap()
      ]);
      
      // Map rows to clean human-readable columns
      var flatHos = hosRows.map(function (row) { return mapHosRow(row, driverMap, deviceMap); });
      var flatEx  = exRows.map(function (row) { return mapExceptionRow(row, driverMap, deviceMap, ruleMap); });
      log("Final data — HOS rows: " + flatHos.length + " | Exception rows: " + flatEx.length);
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
    var versionEl = qs("appVersion");
    if (versionEl) versionEl.textContent = "v" + APP_VERSION;
    var range = getYesterdayRange(null);
    var label = qs("dateLabel");
    if (label) label.textContent = formatDisplayDate(range.fromDate);
    var btn = qs("runBtn");
    if (btn) {
      btn.disabled = true;
      btn.addEventListener("click", runReport);
    }

    if (_api) {
      // Check user access on initialization using the add-in session method
      _api.getSession(function (credentials) {
        var userName = credentials && credentials.userName ? credentials.userName.toLowerCase() : null;
        var allowedUsers = ["mgranados@gridline.com", "eborden@gridline.com"];
        var allowedUsersLower = allowedUsers.map(function (u) { return u.toLowerCase(); });

        if (userName && allowedUsersLower.indexOf(userName) >= 0) {
          setBannerStatus("permitted", "✓ Access permitted for " + (credentials.userName || "user"));          // Fetch actual user time zone and display it
          callApi("Get", { typeName: "User", search: { name: credentials.userName }, resultsLimit: 1 })
            .then(function (result) {
              var tz = result && result[0] && result[0].timeZoneId ? result[0].timeZoneId : "Unknown";
              var tzEl = qs("tzValue");
              if (tzEl) tzEl.textContent = tz;
            })
            .catch(function () {
              var tzEl = qs("tzValue");
              if (tzEl) tzEl.textContent = "Unavailable";
            });          log("Report configured: Yesterday's date, User's time zone, All drivers.");
          log("HOS: ON, D, INT_D, Login, Logoff states only.");
          log("Exceptions: Entering Zone (Office) and Exiting Zone (Office) rules.");
          if (btn) btn.disabled = false;
        } else {
          setBannerStatus("denied", "✗ Access denied for " + (userName || "unknown user"));
          log("ERROR: Access denied. Only authorized testers can run this report.", "error");
          if (btn) btn.disabled = true;
        }
      });
    } else {
      setBannerStatus("pending", "Opening in MyGeotab to check access…");
      log("Standalone preview - open inside MyGeotab to run reports.");
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
