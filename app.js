(function () {
  "use strict";

  function setMessage(text) {
    var node = document.getElementById("message");
    if (node) {
      node.textContent = text;
    }
  }

  function initializeHelloWorld() {
    setMessage("Hello world from Gridline | Gap Report.");
  }

  // Register with MyGeotab. The key "GapReport" must match the last segment
  // of the "path" field in the add-in manifest (e.g. "GapReport/").
  geotab.addin.GapReport = function (api, state) {
    return {
      initialize: function (freshApi, freshState, callback) {
        initializeHelloWorld();
        callback();
      },
      focus: function (freshApi, freshState) {},
      blur: function () {}
    };
  };

  // Standalone preview fallback (browser only, no MyGeotab context).
  if (typeof geotab === "undefined") {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeHelloWorld);
    } else {
      initializeHelloWorld();
    }
  }
})();
