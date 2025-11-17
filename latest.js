/***************************************************************
 * CANVAS GRADING HUB - MASTER CODE (PREFIXED)
 * VERSION 2.1
 *
 * THIS FILE DEFINES:
 *   HUB_refreshSubmissions()
 *   HUB_openMissingSubmissionsDialog()
 *   HUB_startMissingSubmissions()
 *   HUB_reloadSettings()
 *   HUB_showHelp()
 *
 * All internal helpers remain unprefixed.
 ***************************************************************/

const CANVAS_HUB_VERSION = "2.1";

/***************************************************************
 * SETTINGS + HELPERS
 ***************************************************************/
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    throw new Error('Settings sheet missing. Create a "Settings" tab.');
  }

  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 2; i < values.length; i++) {
    const key = String(values[i][0] || "").trim();
    const val = String(values[i][1] || "").trim();
    if (key) map[key] = val;
  }

  if (!map["Canvas Base URL"]) throw new Error("Canvas Base URL missing.");
  if (!map["Canvas API Token"]) throw new Error("Canvas API Token missing.");
  if (!map["Course IDs (comma-separated)"]) {
    throw new Error("Course IDs missing.");
  }

  return {
    baseUrl: map["Canvas Base URL"].replace(/^https?:\/\//, ""),
    apiToken: map["Canvas API Token"],
    courseIds: map["Course IDs (comma-separated)"]
      .split(",")
      .map(s => s.trim())
      .filter(Boolean),
    hoursBack: parseInt(map["Hours to Look Back"], 10) || 24,
    showOnlyUngraded: normalizeYesNo_(map["Show Only Ungraded?"]),
    highlightLate: normalizeYesNo_(map["Highlight Late Submissions?"])
  };
}

function normalizeYesNo_(v) {
  if (!v) return false;
  const s = v.toString().trim().toLowerCase();
  return s === "yes" || s === "y" || s === "true";
}

function canvasHeaders_(settings) {
  return { Authorization: "Bearer " + settings.apiToken };
}

function canvasFetch_(settings, url, options, label) {
  const fullUrl = url.startsWith("http")
    ? url
    : "https://" + settings.baseUrl + url;

  const opts = {
    method: (options && options.method) || "get",
    muteHttpExceptions: true,
    headers: canvasHeaders_(settings)
  };

  const response = UrlFetchApp.fetch(fullUrl, opts);
  const code = response.getResponseCode();

  if (code === 401 || code === 403) {
    throw new Error(
      "Canvas denied access while " +
        label +
        ". Check your API token permissions."
    );
  }

  if (code < 200 || code >= 300) {
    throw new Error(
      "Canvas error " +
        code +
        " while " +
        label +
        ". Response: " +
        response.getContentText().slice(0, 400)
    );
  }

  return JSON.parse(response.getContentText());
}

/***************************************************************
 * TIME AGO HELPER
 ***************************************************************/
function getTimeAgo_(date) {
  const now = new Date();
  const diffMs = now - date;

  const mins = Math.floor(diffMs / 60000);
  const hours = Math.floor(diffMs / 3600000);
  const days = Math.floor(diffMs / 86400000);

  if (mins < 60) return mins + " min ago";
  if (hours < 24) return hours + " hr ago";
  return days + " days ago";
}
