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

  if (window.geotab && window.geotab.addin) {
    window.geotab.addin.GridlineGapReport = function () {
      return {
        initialize: function () {
          initializeHelloWorld();
        },
        focus: function () {},
        blur: function () {}
      };
    };
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeHelloWorld);
  } else {
    initializeHelloWorld();
  }
})();
