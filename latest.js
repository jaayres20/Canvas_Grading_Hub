// Canvas Grading Hub master file
// You’ll replace this with the real code later.

var CanvasHub = (function () {
  'use strict';

  var VERSION = '0.0.1';

  function onOpen() {
    SpreadsheetApp.getUi().alert(
      'Canvas Hub master code loaded from GitHub, but real logic is not pasted yet.'
    );
  }

  function refreshSubmissions() {
    SpreadsheetApp.getUi().alert(
      'Placeholder: refreshSubmissions() is not implemented in latest.js yet.'
    );
  }

  function openMissingSubmissionsDialog() {
    SpreadsheetApp.getUi().alert(
      'Placeholder: openMissingSubmissionsDialog() is not implemented in latest.js yet.'
    );
  }

  function reloadSettings() {}
  function setupDailyTrigger() {}
  function showHelp() {}
  function startMissingSubmissions() {}
  function getVersion() { return VERSION; }

  return {
    getVersion: getVersion,
    onOpen: onOpen,
    refreshSubmissions: refreshSubmissions,
    openMissingSubmissionsDialog: openMissingSubmissionsDialog,
    reloadSettings: reloadSettings,
    setupDailyTrigger: setupDailyTrigger,
    showHelp: showHelp,
    startMissingSubmissions: startMissingSubmissions
  };
})();
