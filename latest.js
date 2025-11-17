/**
 * CANVAS GRADING HUB - MASTER CODE
 * VERSION 2.1
 *
 * This file contains ONLY the Canvas logic:
 *  - Recent Submissions Dashboard (Day 1–5)
 *  - Missing Submissions (latest created assignments)
 *  - Settings helpers
 *  - Triggers + Help
 *
 * Bootstrap.gs (inside Apps Script) loads this file automatically from GitHub.
 */

const CANVAS_HUB_VERSION = '2.1';

// =====================================================
// SETTINGS & SHARED HELPERS
// =====================================================

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Settings tab not found. Please create a "Settings" sheet.');
  }

  const data = settingsSheet.getDataRange().getValues();
  const settings = {};

  for (let i = 2; i < data.length; i++) {
    const key = (data[i][0] || '').toString().trim();
    const value = (data[i][1] || '').toString().trim();
    if (key) settings[key] = value;
  }

  if (!settings['Canvas Base URL'] || settings['Canvas Base URL'] === 'yourschool.instructure.com') {
    throw new Error('Please enter your Canvas Base URL in the Settings tab.');
  }
  if (!settings['Canvas API Token'] || settings['Canvas API Token'] === 'PASTE_YOUR_TOKEN_HERE') {
    throw new Error('Please enter your Canvas API Token in the Settings tab.');
  }
  if (!settings['Course IDs (comma-separated)']) {
    throw new Error('Please enter one or more Course IDs in the Settings tab.');
  }

  return {
    baseUrl: settings['Canvas Base URL']
      .replace(/^https?:\/\//, '')
      .replace(/\/$/, ''),
    apiToken: settings['Canvas API Token'],
    courseIds: settings['Course IDs (comma-separated)']
      .split(',')
      .map(s => s.trim())
      .filter(Boolean),
    hoursBack: parseInt(settings['Hours to Look Back'], 10) || 24,
    runTime: settings['Run Time'] || '5:00 AM',
    showOnlyUngraded: normalizeYesNo_(settings['Show Only Ungraded?']),
    highlightLate: normalizeYesNo_(settings['Highlight Late Submissions?'])
  };
}

function normalizeYesNo_(val) {
  if (!val) return false;
  const s = val.toString().trim().toLowerCase();
  return s === 'yes' || s === 'y' || s === 'true';
}

function canvasHeaders_(settings) {
  return { Authorization: 'Bearer ' + settings.apiToken };
}

function canvasFetch_(settings, url, options, contextLabel) {
  const fullUrl = url.startsWith('http') ? url : 'https://' + settings.baseUrl + url;
  const opts = options || {};
  opts.muteHttpExceptions = true;
  opts.headers = Object.assign({}, canvasHeaders_(settings), opts.headers || {});

  const res = UrlFetchApp.fetch(fullUrl, opts);
  const code = res.getResponseCode();

  if (code === 401 || code === 403) {
    throw new Error(
      `Canvas returned ${code} for ${contextLabel}.
This usually means:
• The API token is invalid or expired
• OR the token does not have permission to view this course.

Try generating a new token in Canvas (Account → Settings → New Access Token).`
    );
  }

  if (code < 200 || code >= 300) {
    throw new Error(
      `Canvas API error (${code}) while ${contextLabel}.
Response: ${res.getContentText().slice(0, 500)}`
    );
  }

  const text = res.getContentText();
  if (!text) return null;

  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('Failed to parse Canvas response while ' + contextLabel);
  }
}

// =====================================================
// RECENT SUBMISSIONS DASHBOARD
// =====================================================

function refreshSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== Starting Canvas Grading Hub Refresh ===');

  try {
    const settings = getSettings();
    const submissions = fetchCanvasSubmissions_(settings);

    if (submissions.length === 0) {
      SpreadsheetApp.getUi().alert(
        `No new submissions found in the last ${settings.hoursBack} hours.
You can adjust "Hours to Look Back" in the Settings tab.`
      );
      return;
    }

    rotateDayTabs_();
    populateDayTab_(submissions, settings);

    SpreadsheetApp.getUi().alert(`Found ${submissions.length} new submission(s)! Check Day 1.`);
    Logger.log('=== Refresh Complete ===');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
    Logger.log('ERROR in refreshSubmissions: ' + err.message);
  }
}

function fetchCanvasSubmissions_(settings) {
  const all = [];
  const cutoff = new Date(Date.now() - settings.hoursBack * 3600 * 1000);

  const courseNames = {};
  settings.courseIds.forEach(id => {
    try {
      courseNames[id] = getCourseName_(settings, id);
    } catch {
      courseNames[id] = 'Course ' + id;
    }
  });

  settings.courseIds.forEach(courseId => {
    try {
      const subs = fetchCourseSubmissions_(settings, courseId, courseNames[courseId], cutoff);
      all.push(...subs);
      Utilities.sleep(200);
    } catch (e) {
      Logger.log(`fetchCourseSubmissions error ${courseId}: ${e.message}`);
    }
  });

  return all.sort((a, b) => b.submittedDate - a.submittedDate);
}

function getCourseName_(settings, courseId) {
  const course = canvasFetch_(
    settings,
    `/api/v1/courses/${courseId}`,
    { method: 'get' },
    'fetching course'
  );
  return course?.name || `Course ${courseId}`;
}

function fetchCourseSubmissions_(settings, courseId, courseName, cutoff) {
  const out = [];
  const assignments = getAssignments_(settings, courseId);

  assignments.forEach(asmt => {
    try {
      const subs = getAssignmentSubmissionsFiltered_(
        settings,
        courseId,
        asmt.id,
        asmt.name,
        courseName,
        cutoff
      );
      out.push(...subs);
      Utilities.sleep(120);
    } catch (e) {
      Logger.log(`Assignment ${asmt.id} error: ${e.message}`);
    }
  });

  return out;
}

function getAssignments_(settings, courseId) {
  try {
    return (
      canvasFetch_(
        settings,
        `/api/v1/courses/${courseId}/assignments?per_page=100`,
        { method: 'get' },
        'fetching assignments'
      ) || []
    );
  } catch (e) {
    Logger.log(`Failed to fetch assignments for course ${courseId}: ${e.message}`);
    return [];
  }
}

function getAssignmentSubmissionsFiltered_(
  settings,
  courseId,
  assignmentId,
  assignmentName,
  courseName,
  cutoff
) {
  let arr;
  try {
    arr =
      canvasFetch_(
        settings,
        `/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions?include[]=user&per_page=100`,
        { method: 'get' },
        'fetching submissions'
      ) || [];
  } catch {
    return [];
  }

  const out = [];

  arr.forEach(s => {
    if (!s.submitted_at || s.workflow_state === 'unsubmitted') return;

    const submittedDate = new Date(s.submitted_at);
    if (submittedDate < cutoff) return;

    if (settings.showOnlyUngraded && s.workflow_state === 'graded') return;

    const studentName = s.user?.name || s.user?.sortable_name || 'Unknown Student';

    const link = `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader?assignment_id=${assignmentId}&student_id=${s.user_id}`;

    out.push({
      studentName,
      courseName,
      assignmentName,
      submittedDate,
      isLate: !!s.late,
      isGraded: s.workflow_state === 'graded',
      speedGraderUrl: link
    });
  });

  return out;
}

function rotateDayTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const day5 = ss.getSheetByName('Day 5');
  if (day5) ss.deleteSheet(day5);

  for (let i = 4; i >= 1; i--) {
    const sh = ss.getSheetByName(`Day ${i}`);
    if (sh) sh.setName(`Day ${i + 1}`);
  }

  const day1 = ss.insertSheet('Day 1', 0);
  setupDayTabTemplate_(day1);
}

function setupDayTabTemplate_(sheet) {
  sheet.getRange('A1:G1').merge();
  sheet.getRange('A1')
    .setValue('CANVAS GRADING DASHBOARD')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#FFF')
    .setBackground('#4472C4')
    .setHorizontalAlignment('center');

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a');

  sheet.getRange('A2').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B2').setValue(dateStr);

  sheet.getRange('D2').setValue('Ungraded:').setFontWeight('bold');
  sheet.getRange('E2').setValue(0).setFontWeight('bold').setFontColor('#F00');

  sheet.getRange('F2').setValue('Graded:').setFontWeight('bold');
  sheet.getRange('G2').setValue(0).setFontWeight('bold').setFontColor('#0A0');

  const headers = ['Graded?', 'Student Name', 'Class', 'Assignment', 'Submitted', 'Link', 'Notes'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length)
    .setFontWeight('bold')
    .setFontColor('#FFF')
    .setBackground('#70AD47')
    .setHorizontalAlignment('center');
}

function populateDayTab_(submissions, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const day1 = ss.getSheetByName('Day 1');
  if (!day1) throw new Error('Day 1 tab missing.');

  const rows = submissions.map(s => {
    const ago = getTimeAgo_(s.submittedDate);
    const submittedText = s.isLate ? `${ago} (LATE)` : ago;
    return [false, s.studentName, s.courseName, s.assignmentName, submittedText, s.speedGraderUrl, ''];
  });

  const startRow = 5;
  if (rows.length > 0) {
    day1.getRange(startRow, 1, rows.length, 7).setValues(rows);

    day1.getRange(startRow, 1, rows.length, 1).insertCheckboxes();

    rows.forEach((r, i) => {
      const url = r[5];
      day1.getRange(startRow + i, 6)
        .setFormula(`=HYPERLINK("${url}","View Submission")`)
        .setHorizontalAlignment('center')
        .setFontColor('#0B5394');
    });

    if (settings.highlightLate) {
      submissions.forEach((s, i) => {
        if (s.isLate) {
          day1.getRange(startRow + i, 5)
            .setBackground('#FFFF00')
            .setFontColor('#FF0000')
            .setFontWeight('bold');
        }
      });
    }
  }

  const ungraded = submissions.filter(s => !s.isGraded).length;
  const graded = submissions.filter(s => s.isGraded).length;

  day1.getRange('E2').setValue(ungraded);
  day1.getRange('G2').setValue(graded);
}

function getTimeAgo_(date) {
  const now = new Date();
  const diff = now - date;
  const mins = Math.floor(diff / 60000);
  const hrs = Math.floor(diff / 3600000);
  const days = Math.floor(diff / 86400000);

  if (mins < 60) return mins === 1 ? '1 minute ago' : `${mins} minutes ago`;
  if (hrs < 24) return hrs === 1 ? '1 hour ago' : `${hrs} hours ago`;
  if (days === 1) return '1 day ago';
  return `${days} days ago`;
}

// =====================================================
// MISSING SUBMISSIONS
// =====================================================

function openMissingSubmissionsDialog() {
  const ui = SpreadsheetApp.getUi();
  try {
    const settings = getSettings();

    const coursePairs = settings.courseIds.map(id => {
      let name;
      try {
        name = getCourseName_(settings, id);
      } catch {
        name = 'Course ' + id;
      }
      return { id, name };
    });

    const html = buildMissingDialogHtml_(coursePairs);
    ui.showModalDialog(html, 'Check Missing Submissions');
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

function buildMissingDialogHtml_(coursePairs) {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial; margin: 16px; }
          label { font-weight:bold; display:block; margin-top:10px; }
          select, button { width:100%; padding:8px; }
          button { background:#2F5597; color:white; border:none; margin-top:14px; }
        </style>
      </head>
      <body>
        <label>Course</label>
        <select id="course">
          <option value="ALL">All Classes</option>
          ${coursePairs.map(c => `<option value="${c.id}">${c.name} (${c.id})</option>`).join('')}
        </select>

        <label>Assignment Range</label>
        <select id="range">
          <option value="1">Latest Assignment Only</option>
          <option value="2">Last 2 Assignments</option>
          <option value="3">Last 3 Assignments</option>
          <option value="4">Last 4 Assignments</option>
          <option value="5">Last 5 Assignments</option>
          <option value="ALL">All Assignments (Slow)</option>
        </select>

        <button id="runBtn">Run Missing Submissions</button>

        <script>
          document.getElementById('runBtn').onclick = function() {
            const c = document.getElementById('course').value;
            const r = document.getElementById('range').value;
            this.disabled = true;
            this.innerText = "Working...";
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler(err => alert(err.message || err))
              .startMissingSubmissions(c, r);
          };
        </script>
      </body>
    </html>
  `);

  return html.setWidth(360).setHeight(310);
}

function startMissingSubmissions(courseChoice, rangeChoice) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  let settings;
  try {
    settings = getSettings();
  } catch (e) {
    ui.alert('Settings error: ' + e.message);
    return;
  }

  const courseList = courseChoice === 'ALL' ? settings.courseIds : [courseChoice];
  const maxAssignments = rangeChoice === 'ALL' ? 'ALL' : parseInt(rangeChoice, 10);

  const out = prepareMissingSheet_();
  ss.toast('Starting Missing Submissions…', 'Canvas Hub', 5);

  const START = Date.now();
  const LIMIT = 5.5 * 60 * 1000;

  let processed = 0;

  for (let i = 0; i < courseList.length; i++) {
    const id = courseList[i];

    ss.toast(`Processing ${i + 1} of ${courseList.length}…`, 'Canvas Hub', 8);

    try {
      appendMissingForCourse_(settings, id, maxAssignments, out);
    } catch (e) {
      Logger.log(`MissingSubmissions error: ${e.message}`);
    }

    processed++;

    if (Date.now() - START > LIMIT && i < courseList.length - 1) {
      const btn = ui.alert(
        'Continue?',
        `Processed ${processed} of ${courseList.length} courses.\nContinue?`,
        ui.ButtonSet.OK_CANCEL
      );
      if (btn !== ui.Button.OK) return;
      processed = 0;
    }
  }

  ui.alert('Missing Submissions complete!');
  ss.toast('Missing Submissions complete!', 'Done', 5);
}

function prepareMissingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Missing Submissions');
  if (!sh) sh = ss.insertSheet('Missing Submissions');

  sh.clear();
  sh.setFrozenRows(2);

  sh.getRange('A1:E1').merge().setValue('MISSING SUBMISSIONS DASHBOARD')
    .setFontSize(14).setFontWeight('bold').setFontColor('#FFF').setBackground('#C00000')
    .setHorizontalAlignment('center');

  sh.getRange('A2').setValue('Last Updated:').setFontWeight('bold');
  sh.getRange('B2').setValue(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy h:mm a')
  );

  const headers = ['Student Name', 'Assignment', 'Created Date', 'Course', 'SpeedGrader Link'];
  sh.getRange(4, 1, 1, 5).setValues([headers])
    .setFontWeight('bold').setFontColor('#FFF').setBackground('#2F5597')
    .setHorizontalAlignment('center');

  return { sheet: sh, nextRow: 5 };
}

function appendMissingForCourse_(settings, courseId, maxAssignments, out) {
  const sh = out.sheet;

  let courseName;
  try {
    courseName = getCourseName_(settings, courseId);
  } catch {
    courseName = 'Course ' + courseId;
  }

  const hdr = out.nextRow++;
  sh.getRange(hdr, 1, 1, 5).merge()
    .setValue(`${courseName} (${courseId})`)
    .setBackground('#FBE5D6').setFontWeight('bold');

  const students = getCourseStudents_(settings, courseId);
  let assignments = getAssignments_(settings, courseId);

  assignments.sort((a, b) =>
    new Date(b.created_at).getTime() - new Date(a.created_at).getTime()
  );

  if (maxAssignments !== 'ALL') assignments = assignments.slice(0, maxAssignments);

  if (assignments.length === 0) {
    const row = out.nextRow++;
    sh.getRange(row, 1, 1, 5).merge()
      .setValue('No assignments found for this selection.')
      .setFontStyle('italic').setFontColor('#666');
    return;
  }

  assignments.forEach((asmt, idx) => {
    const subs = getAssignmentSubmissionsRaw_(settings, courseId, asmt.id);
    const submittedIds = new Set(
      subs
        .filter(s => s?.submitted_at && s.workflow_state !== 'unsubmitted')
        .map(s => s.user_id)
    );

    const missing = students.filter(stu => !submittedIds.has(stu.id));

    const createdText = asmt.created_at
      ? Utilities.formatDate(new Date(asmt.created_at), Session.getScriptTimeZone(), 'M/d/yyyy h:mm a')
      : '—';

    const row = out.nextRow++;
    sh.getRange(row, 1, 1, 5).merge()
      .setValue(asmt.name)
      .setFontWeight('bold').setBackground('#FFF2CC');

    if (missing.length === 0) {
      const r = out.nextRow++;
      sh.getRange(r, 1, 1, 5).merge().setValue(`✓ No missing submissions for ${asmt.name}`)
        .setFontColor('#198754');
      return;
    }

    const rows = missing.map(m => {
      const link = `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader?assignment_id=${asmt.id}&student_id=${m.id}`;
      return [m.name, asmt.name, createdText, courseName, link];
    });

    sh.getRange(out.nextRow, 1, rows.length, 5).setValues(rows);

    for (let i = 0; i < rows.length; i++) {
      sh.getRange(out.nextRow + i, 5)
        .setFormula(`=HYPERLINK("${rows[i][4]}","Open SpeedGrader")`)
        .setHorizontalAlignment('center')
        .setFontColor('#0B5394');
    }

    out.nextRow += rows.length;
  });

  out.nextRow++;
}

function getCourseStudents_(settings, courseId) {
  const arr =
    canvasFetch_(
      settings,
      `/api/v1/courses/${courseId}/users?enrollment_type[]=student&per_page=100`,
      { method: 'get' },
      'fetching students'
    ) || [];

  return arr.map(u => ({
    id: u.id,
    name: u.name || u.sortable_name || 'Unknown Student'
  }));
}

function getAssignmentSubmissionsRaw_(settings, courseId, assignmentId) {
  return (
    canvasFetch_(
      settings,
      `/api/v1/courses/${courseId}/assignments/${assignmentId}/submissions?per_page=100`,
      { method: 'get' },
      'fetching raw submissions'
    ) || []
  );
}

// =====================================================
// LOCAL VERSION EXPORT
// =====================================================

function getLocalVersion() {
  return CANVAS_HUB_VERSION;
}
