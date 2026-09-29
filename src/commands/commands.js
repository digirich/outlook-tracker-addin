/**
 * Tracker Tool — ribbon command handlers (function file)
 *
 * The primary "Add to Tracker" control uses ShowTaskpane in manifest.xml,
 * so most UX lives in the task pane. This file is ready for ExecuteFunction
 * commands (e.g. quick actions) if you add them later.
 */
(function () {
  "use strict";

  Office.onReady(function () {
    // Ready for Association of function names in the manifest.
  });

  /**
   * Example ExecuteFunction handler (not wired in manifest by default).
   * Associate via FunctionName in VersionOverrides if needed.
   */
  function actionOpenTracker(event) {
    // No-op placeholder — task pane is opened via ShowTaskpane.
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }

  // Register for potential ExecuteFunction association
  if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("actionOpenTracker", actionOpenTracker);
  } else {
    // Older hosts: global function name matching FunctionName
    window.actionOpenTracker = actionOpenTracker;
  }
})();
