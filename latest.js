/***************************************************************
 * CANVAS GRADING HUB – VERSION 2.1
 * Single-file system with:
 *  - Canvas API integration
 *  - Recent submissions dashboard (Day 1–5)
 *  - Missing submissions dashboard
 *  - Manual update checker using GitHub
 *  - Popup “Copy New Code” dialog for easy updates
 ***************************************************************/

const VERSION = "2.1";
const MASTER_CODE_URL =
  "https://raw.githubusercontent.com/jaayres20/Canvas_Grading_Hub/main/latest.js";

const VERSION_CELL = "B12";
const VERSION_LABEL_CELL = "A12";

/****************************************************
 * ON OPEN – Build Menu + Write Version
 ****************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Canvas Hub")
    .addItem("Refresh Now (Recent Submissions)", "refreshSubmissions")
    .addSeparator()
    .addItem("Check Missing Submissions", "openMissingSubmissionsDialog")
    .addSeparator()
    .addItem("Update Settings", "reloadSettings")
    .addItem("Set Up Auto-Run (5 AM Daily)", "setupDailyTrigger")
    .addSeparator()
    .addItem("Check for Updates", "checkForUpdates")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();

  writeLocalVersionToSettings_();
}

/****************************************************
 * UPDATE CHECKER
 ****************************************************/
function checkForUpdates() {
  const latestCode = fetchLatestFromGitHub_();
  if (!latestCode) {
    SpreadsheetApp.getUi().alert("Could not fetch latest code from GitHub.");
    return;
  }

  const remoteVersion = extractVersionFromCode_(latestCode);
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName("Settings");
  const localVersion = settings.getRange(VERSION_CELL).getValue().toString().trim();

  if (localVersion !== remoteVersion) {
    showUpdateAvailableDialog_(remoteVersion, latestCode);
  } else {
    SpreadsheetApp.getUi().alert("You are already running version " + remoteVersion);
  }
}

function showUpdateAvailableDialog_(remoteVersion, remoteCode) {
  const html = `
    <html>
      <head>
        <style>
          body { font-family: Arial; padding:16px; }
          textarea { width:100%; height:320px; font-family:monospace; }
          button { margin-top:10px; padding:8px 12px; background:#2F5597; color:white; border:none; cursor:pointer; }
        </style>
      </head>
      <body>
        <h2>New Version Available</h2>
        <p><strong>Your Version:</strong> ${VERSION}<br>
        <strong>Latest Version:</strong> ${remoteVersion}</p>

        <p>Copy the new code below and paste it into the Apps Script editor
        (replace everything).</p>

        <textarea id="code">${remoteCode.replace(/</g, "&lt;")}</textarea>

        <button onclick="copy()">Copy Code</button>

        <script>
          function copy() {
            const t = document.getElementById("code");
            t.select();
            document.execCommand("copy");
            alert("Copied!");
          }
        </script>
      </body>
    </html>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(600),
    "Update Available"
  );
}

function fetchLatestFromGitHub_() {
  try {
    const res = UrlFetchApp.fetch(MASTER_CODE_URL, { muteHttpExceptions: true });
    return res.getResponseCode() === 200 ? res.getContentText() : null;
  } catch (e) {
    Logger.log(e);
    return null;
  }
}

function extractVersionFromCode_(code) {
  const m = code.match(/VERSION\s*=\s*["']([0-9]+\.[0-9]+)["']/);
  return m ? m[1] : null;
}

function writeLocalVersionToSettings_() {
  const ss = SpreadsheetApp.getActive();
  const settings = ss.getSheetByName("Settings");
  if (settings) {
    settings.getRange(VERSION_LABEL_CELL).setValue("Local Version");
    settings.getRange(VERSION_CELL).setValue(VERSION);
  }
}

/***************************************************************
 * ================================
 *  MASTER LOGIC – v2.1
 * ================================
 ***************************************************************/

/****************************************************
 * READ SETTINGS
 ****************************************************/
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) throw new Error('Settings tab not found.');

  const data = settingsSheet.getDataRange().getValues();
  const settings = {};

  for (let i = 2; i < data.length; i++) {
    const key = (data[i][0] || '').toString().trim();
    const value = (data[i][1] || '').toString().trim();
    if (key) settings[key] = value;
  }

  if (!settings['Canvas Base URL']) {
    throw new Error('Canvas Base URL missing.');
  }
  if (!settings['Canvas API Token']) {
    throw new Error('Canvas API Token missing.');
  }

  return {
    baseUrl: settings['Canvas Base URL'].replace(/^https?:\/\//, ''),
    apiToken: settings['Canvas API Token'],
    courseIds: settings['Course IDs (comma-separated)']
      .split(',')
      .map(s => s.trim())
      .filter(Boolean),
    hoursBack: parseInt(settings['Hours to Look Back'], 10) || 48,
    showOnlyUngraded: normalizeYesNo_(settings['Show Only Ungraded?']),
    highlightLate: normalizeYesNo_(settings['Highlight Late Submissions?'])
  };
}

function normalizeYesNo_(v) {
  if (!v) return false;
  return ['yes', 'y', 'true'].includes(v.toString().trim().toLowerCase());
}

function canvasHeaders_(settings) {
  return { Authorization: 'Bearer ' + settings.apiToken };
}

/****************************************************
 * CANVAS API WRAPPER
 ****************************************************/
function canvasFetch_(settings, url, options, label) {
  const fullUrl = url.startsWith('http')
    ? url
    : 'https://' + settings.baseUrl + url;

  const opts = {
    method: options?.method || 'get',
    muteHttpExceptions: true,
    headers: canvasHeaders_(settings)
  };

  const res = UrlFetchApp.fetch(fullUrl, opts);
  const code = res.getResponseCode();

  if (code === 401 || code === 403) {
    throw new Error("Canvas refused access for: " + label);
  }
  if (code < 200 || code >= 300) {
    throw new Error("Canvas error " + code + " while " + label);
  }

  return JSON.parse(res.getContentText());
}

/****************************************************
 * RECENT SUBMISSIONS (Day 1 → Day 5)
 ****************************************************/
function refreshSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const settings = getSettings();
    const subs = fetchCanvasSubmissions_(settings);

    if (subs.length === 0) {
      SpreadsheetApp.getUi().alert("No new submissions.");
      return;
    }

    rotateDayTabs_();
    populateDayTab_(subs, settings);

    SpreadsheetApp.getUi().alert("Found " + subs.length + " new submissions!");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }
}

function fetchCanvasSubmissions_(settings) {
  const cutoff = new Date(Date.now() - settings.hoursBack * 3600000);
  const all = [];

  settings.courseIds.forEach(id => {
    const courseName = getCourseName_(settings, id);
    const assignments = getAssignments_(settings, id);

    assignments.forEach(a => {
      const subs = getAssignmentSubmissionsFiltered_(
        settings, id, a.id, a.name, courseName, cutoff
      );
      all.push(...subs);
    });
  });

  all.sort((a, b) => b.submittedDate - a.submittedDate);
  return all;
}

function getCourseName_(settings, id) {
  const c = canvasFetch_(settings, `/api/v1/courses/${id}`, {}, "course name");
  return c?.name || ("Course " + id);
}

function getAssignments_(settings, courseId) {
  return (
    canvasFetch_(
      settings,
      `/api/v1/courses/${courseId}/assignments?per_page=100`,
      {},
      "assignment list"
    ) || []
  );
}

function getAssignmentSubmissionsFiltered_(
  settings, courseId, asmtId, asmtName, courseName, cutoff
) {
  const arr =
    canvasFetch_(
      settings,
      `/api/v1/courses/${courseId}/assignments/${asmtId}/submissions?include[]=user&per_page=100`,
      {},
      "submissions"
    ) || [];

  const out = [];

  arr.forEach(s => {
    if (!s.submitted_at) return;

    const date = new Date(s.submitted_at);
    if (date < cutoff) return;

    if (settings.showOnlyUngraded && s.workflow_state === 'graded') return;

    const link =
      `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader?assignment_id=${asmtId}&student_id=${s.user_id}`;

    out.push({
      studentName: s.user?.name || "Unknown",
      courseName,
      assignmentName: asmtName,
      submittedDate: date,
      isLate: !!s.late,
      isGraded: s.workflow_state === "graded",
      speedGraderUrl: link
    });
  });

  return out;
}

/****************************************************
 * Day Tab Building
 ****************************************************/
function rotateDayTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const d5 = ss.getSheetByName("Day 5");
  if (d5) ss.deleteSheet(d5);

  for (let i = 4; i >= 1; i--) {
    const sh = ss.getSheetByName("Day " + i);
    if (sh) sh.setName("Day " + (i + 1));
  }

  ss.insertSheet("Day 1", 0);
  setupDayTabTemplate_(ss.getSheetByName("Day 1"));
}

function setupDayTabTemplate_(sh) {
  sh.getRange("A1:G1").merge().setValue("CANVAS GRADING DASHBOARD");
  sh.getRange("A1").setFontWeight("bold").setFontSize(14).setBackground("#4472C4").setFontColor("white").setHorizontalAlignment("center");

  sh.getRange("A2").setValue("Last Updated:").setFontWeight("bold");
  sh.getRange("B2").setValue(new Date());

  const headers = ["Graded?", "Student Name", "Class", "Assignment", "Submitted", "Link", "Notes"];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]);
  sh.getRange(4, 1, 1, headers.length).setBackground("#70AD47").setFontColor("white").setFontWeight("bold");
}

function populateDayTab_(subs, settings) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Day 1");
  const rows = subs.map(s => [
    false,
    s.studentName,
    s.courseName,
    s.assignmentName,
    getTimeAgo_(s.submittedDate),
    s.speedGraderUrl,
    ""
  ]);

  sh.getRange(5, 1, rows.length, 7).setValues(rows);
  sh.getRange(5, 1, rows.length, 1).insertCheckboxes();

  rows.forEach((r, i) => {
    const url = r[5];
    sh.getRange(5 + i, 6).setFormula(`=HYPERLINK("${url}","View")`);
  });
}

function getTimeAgo_(date) {
  const now = new Date();
  const diff = now - date;
  const mins = Math.floor(diff / 60000);
  const hrs = Math.floor(diff / 3600000);
  const days = Math.floor(diff / 86400000);

  if (mins < 60) return mins + " min ago";
  if (hrs < 24) return hrs + " hr ago";
  return days + " days ago";
}

/***************************************************************
 * MISSING SUBMISSIONS
 ***************************************************************/
function openMissingSubmissionsDialog() {
  const settings = getSettings();
  const courses = settings.courseIds.map(id => ({
    id,
    name: getCourseName_(settings, id)
  }));

  const html = buildMissingDialogHtml_(courses);
  SpreadsheetApp.getUi().showModalDialog(html, "Check Missing Submissions");
}

function buildMissingDialogHtml_(pairs) {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body { font-family: Arial; padding:16px; }
          select, button { width:100%; padding:8px; margin-bottom:12px; }
          button { background:#2F5597; color:white; border:none; cursor:pointer; }
        </style>
      </head>
      <body>
        <h3>Choose Course & Assignment Range</h3>

        <label>Course</label>
        <select id="course">
          <option value="ALL">All Courses</option>
          ${pairs.map(p => `<option value="${p.id}">${p.name}</option>`).join("")}
        </select>

        <label>Assignment Range</label>
        <select id="range">
          <option value="1">Latest</option>
          <option value="2">Last 2</option>
          <option value="3">Last 3</option>
          <option value="4">Last 4</option>
          <option value="5">Last 5</option>
          <option value="ALL">All Assignments (slow)</option>
        </select>

        <button onclick="run()">Run</button>

        <script>
          function run() {
            google.script.run.startMissingSubmissions(
              document.getElementById('course').value,
              document.getElementById('range').value
            );
            google.script.host.close();
          }
        </script>
      </body>
    </html>
  `);
  html.setWidth(350).setHeight(320);
  return html;
}

function startMissingSubmissions(courseChoice, rangeChoice) {
  const ss = SpreadsheetApp.getActive();
  const settings = getSettings();

  const courses = courseChoice === "ALL" ? settings.courseIds : [courseChoice];
  const max = rangeChoice === "ALL" ? "ALL" : parseInt(rangeChoice, 10);

  const out = prepareMissingSheet_();
  ss.toast("Checking Missing Submissions...", "Canvas Hub", 5);

  courses.forEach(id => {
    appendMissingForCourse_(settings, id, max, out);
  });

  SpreadsheetApp.getUi().alert("Missing Submissions Updated!");
}

function prepareMissingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Missing Submissions");
  if (!sh) sh = ss.insertSheet("Missing Submissions");

  sh.clear();
  sh.getRange("A1:E1")
    .merge()
    .setValue("MISSING SUBMISSIONS")
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center")
    .setBackground("#C00000")
    .setFontColor("white");

  sh.getRange("A2").setValue("Last Updated:");
  sh.getRange("B2").setValue(new Date());

  sh.getRange(4, 1, 1, 5).setValues([[
    "Student",
    "Assignment",
    "Created",
    "Course",
    "SpeedGrader"
  ]]).setFontWeight("bold").setBackground("#2F5597").setFontColor("white");

  return { sheet: sh, nextRow: 5 };
}

function appendMissingForCourse_(settings, courseId, max, out) {
  const sh = out.sheet;
  const courseName = getCourseName_(settings, courseId);

  // Divider
  const r = out.nextRow++;
  sh.getRange(r, 1).setValue(courseName).setFontWeight("bold");

  const students = getCourseStudents_(settings, courseId);
  let as = getAssignments_(settings, courseId);

  as.sort((a, b) =>
    new Date(b.created_at || 0) - new Date(a.created_at || 0)
  );

  if (max !== "ALL") as = as.slice(0, max);

  as.forEach(a => {
    const subs = getAssignmentSubmissionsRaw_(settings, courseId, a.id);
    const submittedIds = new Set(
      subs.filter(s => s.submitted_at).map(s => s.user_id)
    );

    const missing = students.filter(s => !submittedIds.has(s.id));

    const asmtRow = out.nextRow++;
    sh.getRange(asmtRow, 1, 1, 5)
      .merge()
      .setValue(a.name)
      .setBackground("#FFF2CC")
      .setFontWeight("bold");

    const created = a.created_at
      ? new Date(a.created_at).toLocaleString()
      : "—";

    missing.forEach(stu => {
      const link =
        `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader?assignment_id=${a.id}&student_id=${stu.id}`;

      sh.getRange(out.nextRow, 1, 1, 5).setValues([[
        stu.name,
        a.name,
        created,
        courseName,
        link
      ]]);

      sh.getRange(out.nextRow, 5).setFormula(`=HYPERLINK("${link}", "Open")`);

      out.nextRow++;
    });
  });

  out.nextRow++;
}

function getCourseStudents_(settings, id) {
  return canvasFetch_(settings, `/api/v1/courses/${id}/users?enrollment_type[]=student&per_page=100`, {}, "students")
    .map(u => ({ id: u.id, name: u.name }));
}

function getAssignmentSubmissionsRaw_(settings, courseId, asmtId) {
  return canvasFetch_(
    settings,
    `/api/v1/courses/${courseId}/assignments/${asmtId}/submissions?per_page=100`,
    {},
    "raw submissions"
  );
}

/***************************************************************
 * SETTINGS UI, HELP, TRIGGERS
 ***************************************************************/
function reloadSettings() {
  const s = getSettings();
  SpreadsheetApp.getUi().alert(
    "Settings Loaded:\n" +
      "Courses: " + s.courseIds.length + "\n" +
      "Hours Back: " + s.hoursBack + "\n" +
      "Show Only Ungraded: " + s.showOnlyUngraded
  );
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "refreshSubmissions") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("refreshSubmissions").timeBased().everyDays(1).atHour(5).create();
  SpreadsheetApp.getUi().alert("Auto-run set for 5 AM daily.");
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    "Canvas Grading Hub Help:\n\n" +
    "- Refresh to view recent submissions\n" +
    "- Use the Missing Submissions tool\n" +
    "- API token required\n" +
    "- Supports multiple courses"
  );
}
