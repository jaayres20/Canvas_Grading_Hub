/**
 * CANVAS GRADING HUB - MASTER CODE
 * VERSION 2.1
 *
 * This file contains all of the Canvas logic:
 *  - Day 1–5 "Recent Submissions" dashboard
 *  - Missing Submissions by latest created assignments
 *  - Settings helpers (Hours Back, Only Ungraded, Highlight Late)
 *  - Triggers + Help
 *
 * The menu, update system, and bootstrap loader live in Bootstrap.gs.
 */

const CANVAS_HUB_VERSION = '2.1'; // used by the bootstrap/update system

// =====================================================
// SETTINGS & SHARED HELPERS
// =====================================================

/**
 * Read the Settings sheet.
 * Expects rows like:  Setting | Value
 * Required:
 *   - Canvas Base URL
 *   - Canvas API Token
 *   - Course IDs (comma-separated)
 * Optional:
 *   - Hours to Look Back
 *   - Run Time (info only)
 *   - Show Only Ungraded?
 *   - Highlight Late Submissions?
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    throw new Error('Settings tab not found. Please create a "Settings" sheet.');
  }

  const data = settingsSheet.getDataRange().getValues();
  const settings = {};

  // Start at row 3 (index 2): row 1 title, row 2 column headers.
  for (let i = 2; i < data.length; i++) {
    const key = (data[i][0] || '').toString().trim();
    const value = (data[i][1] || '').toString().trim();
    if (key) settings[key] = value;
  }

  if (!settings['Canvas Base URL'] || settings['Canvas Base URL'] === 'yourschool.instructure.com') {
    throw new Error(
      'Please enter your Canvas Base URL in the Settings tab (e.g. yourschool.instructure.com).'
    );
  }
  if (!settings['Canvas API Token'] || settings['Canvas API Token'] === 'PASTE_YOUR_TOKEN_HERE') {
    throw new Error('Please enter your Canvas API Token in the Settings tab.');
  }
  if (!settings['Course IDs (comma-separated)']) {
    throw new Error('Please enter one or more Course IDs in the Settings tab (comma-separated).');
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
    hoursBack: parseInt(settings['Hours to Look Back'], 10) || 24, // default 24 hours
    runTime: settings['Run Time'] || '5:00 AM', // informational only
    showOnlyUngraded: normalizeYesNo_(settings['Show Only Ungraded?']),
    highlightLate: normalizeYesNo_(settings['Highlight Late Submissions?'])
  };
}

/**
 * Normalize Yes/No dropdowns from Settings.
 */
function normalizeYesNo_(val) {
  if (!val) return false;
  const s = val.toString().trim().toLowerCase();
  return s === 'yes' || s === 'y' || s === 'true';
}

/**
 * Canvas auth header.
 */
function canvasHeaders_(settings) {
  return { Authorization: 'Bearer ' + settings.apiToken };
}

/**
 * Centralized Canvas API fetch helper with better error messages.
 */
function canvasFetch_(settings, url, options, contextLabel) {
  const fullUrl = url.startsWith('http') ? url : 'https://' + settings.baseUrl + url;
  const opts = options || {};
  opts.muteHttpExceptions = true;
  opts.headers = Object.assign({}, canvasHeaders_(settings), opts.headers || {});

  const res = UrlFetchApp.fetch(fullUrl, opts);
  const code = res.getResponseCode();

  if (code === 401 || code === 403) {
    throw new Error(
      'Canvas returned ' +
        code +
        ' for ' +
        contextLabel +
        '.\n\nThis usually means:\n' +
        '• The API token is invalid or expired, OR\n' +
        '• The token does not have permission to view this course.\n\n' +
        'Try generating a new token in Canvas (Account → Settings → New Access Token),\n' +
        'update it in the Settings tab, and re-run.'
    );
  }

  if (code < 200 || code >= 300) {
    throw new Error(
      'Canvas API error (' +
        code +
        ') while ' +
        contextLabel +
        '.\nResponse: ' +
        res.getContentText().slice(0, 500)
    );
  }

  const text = res.getContentText();
  if (!text) return null;

  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('Failed to parse Canvas response while ' + contextLabel + '.');
  }
}

// =====================================================
// RECENT SUBMISSIONS DASHBOARD (DAY 1–5)
// =====================================================

/**
 * Main entry: refresh recent submissions into Day 1 (rotating Day 1–5).
 * Called from the menu in Bootstrap.gs.
 */
function refreshSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('=== Starting Canvas Grading Hub Refresh ===');

  try {
    const settings = getSettings();
    const submissions = fetchCanvasSubmissions_(settings);

    if (submissions.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No new submissions found in the last ' +
          settings.hoursBack +
          ' hours.\n\nYou can adjust "Hours to Look Back" in the Settings tab.'
      );
      return;
    }

    rotateDayTabs_();
    populateDayTab_(submissions, settings);

    Logger.log('=== Refresh Complete ===');
    SpreadsheetApp.getUi().alert(
      'Found ' + submissions.length + ' new submission(s)! Check the Day 1 tab.'
    );
  } catch (err) {
    Logger.log('ERROR in refreshSubmissions: ' + err.message);
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
  }
}

/**
 * Gather recent submissions across all configured courses within hoursBack window.
 */
function fetchCanvasSubmissions_(settings) {
  const all = [];
  const cutoff = new Date(Date.now() - settings.hoursBack * 60 * 60 * 1000);

  // Preload course names for nice display.
  const courseNames = {};
  settings.courseIds.forEach(id => {
    try {
      courseNames[id] = getCourseName_(settings, id);
    } catch (e) {
      Logger.log('Course name fallback for ' + id + ': ' + e.message);
      courseNames[id] = 'Course ' + id;
    }
  });

  settings.courseIds.forEach(courseId => {
    try {
      const subs = fetchCourseSubmissions_(
        settings,
        courseId,
        courseNames[courseId],
        cutoff
      );
      all.push.apply(all, subs);
      Utilities.sleep(200); // small breather between courses
    } catch (e) {
      Logger.log('fetchCourseSubmissions error ' + courseId + ': ' + e.message);
    }
  });

  all.sort((a, b) => b.submittedDate - a.submittedDate);
  return all;
}

/**
 * Get Canvas course name.
 */
function getCourseName_(settings, courseId) {
  const course = canvasFetch_(
    settings,
    '/api/v1/courses/' + courseId,
    { method: 'get' },
    'fetching course ' + courseId
  );
  return course && course.name ? course.name : 'Course ' + courseId;
}

/**
 * Get submissions for all assignments in a single course.
 */
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
      out.push.apply(out, subs);
      Utilities.sleep(120); // avoid hammering Canvas
    } catch (e) {
      Logger.log('Assignment ' + asmt.id + ' error: ' + e.message);
    }
  });

  return out;
}

/**
 * Assignments for a course.
 */
function getAssignments_(settings, courseId) {
  try {
    return (
      canvasFetch_(
        settings,
        '/api/v1/courses/' + courseId + '/assignments?per_page=100',
        { method: 'get' },
        'fetching assignments for course ' + courseId
      ) || []
    );
  } catch (e) {
    Logger.log('Failed to fetch assignments for course ' + courseId + ': ' + e.message);
    return [];
  }
}

/**
 * Submissions for a specific assignment, filtered for recency and (optionally) ungraded only.
 */
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
        '/api/v1/courses/' +
          courseId +
          '/assignments/' +
          assignmentId +
          '/submissions?include[]=user&per_page=100',
        { method: 'get' },
        'fetching submissions for assignment ' + assignmentId
      ) || [];
  } catch (e) {
    Logger.log('Failed to fetch submissions for assignment ' + assignmentId + ': ' + e.message);
    return [];
  }

  const out = [];
  arr.forEach(s => {
    if (!s.submitted_at || s.workflow_state === 'unsubmitted') return;

    const submittedDate = new Date(s.submitted_at);
    if (submittedDate < cutoff) return;

    if (settings.showOnlyUngraded && s.workflow_state === 'graded') return;

    const studentName = s.user
      ? s.user.name || s.user.sortable_name || 'Unknown Student'
      : 'Unknown Student';

    const isLate = !!s.late;

    const link =
      'https://' +
      settings.baseUrl +
      '/courses/' +
      courseId +
      '/gradebook/speed_grader?assignment_id=' +
      assignmentId +
      '&student_id=' +
      s.user_id;

    out.push({
      studentName: studentName,
      courseName: courseName,
      assignmentName: assignmentName,
      submittedDate: submittedDate,
      isLate: isLate,
      isGraded: s.workflow_state === 'graded',
      speedGraderUrl: link
    });
  });

  return out;
}

/**
 * Rotate Day 1–5 tabs and create a fresh Day 1.
 */
function rotateDayTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const day5 = ss.getSheetByName('Day 5');
  if (day5) ss.deleteSheet(day5);

  for (let i = 4; i >= 1; i--) {
    const sh = ss.getSheetByName('Day ' + i);
    if (sh) sh.setName('Day ' + (i + 1));
  }

  const day1 = ss.insertSheet('Day 1', 0);
  setupDayTabTemplate_(day1);
}

/**
 * Configure header + columns for a Day tab.
 */
function setupDayTabTemplate_(sheet) {
  sheet.getRange('A1:G1').merge();
  sheet
    .getRange('A1')
    .setValue('CANVAS GRADING DASHBOARD')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#4472C4')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 25);

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a');

  sheet.getRange('A2').setValue('Last Updated:').setFontWeight('bold');
  sheet.getRange('B2').setValue(dateStr);

  sheet.getRange('D2').setValue('Ungraded:').setFontWeight('bold');
  sheet.getRange('E2').setValue(0).setFontWeight('bold').setFontColor('#FF0000');

  sheet.getRange('F2').setValue('Graded:').setFontWeight('bold');
  sheet.getRange('G2').setValue(0).setFontWeight('bold').setFontColor('#008000');

  const headers = ['Graded?', 'Student Name', 'Class', 'Assignment', 'Submitted', 'Link', 'Notes'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet
    .getRange(4, 1, 1, headers.length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#70AD47')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 250);
}

/**
 * Fill Day 1 with rows + checkboxes + late highlighting + stats.
 */
function populateDayTab_(submissions, settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const day1 = ss.getSheetByName('Day 1');
  if (!day1) throw new Error('Day 1 tab not found!');

  const rows = submissions.map(s => {
    const ago = getTimeAgo_(s.submittedDate);
    const submittedText = s.isLate ? ago + ' (LATE)' : ago;
    return [false, s.studentName, s.courseName, s.assignmentName, submittedText, s.speedGraderUrl, ''];
  });

  const startRow = 5;

  if (rows.length > 0) {
    day1.getRange(startRow, 1, rows.length, 7).setValues(rows);

    // Column A: checkboxes
    day1.getRange(startRow, 1, rows.length, 1).insertCheckboxes();

    // Column F: SpeedGrader links
    for (let i = 0; i < rows.length; i++) {
      const cell = day1.getRange(startRow + i, 6);
      const url = rows[i][5];
      cell
        .setFormula('=HYPERLINK("' + url + '","View Submission")')
        .setHorizontalAlignment('center')
        .setFontColor('#0B5394');
    }

    // Late highlighting (optional)
    if (settings.highlightLate) {
      for (let i = 0; i < submissions.length; i++) {
        if (submissions[i].isLate) {
          const cell = day1.getRange(startRow + i, 5);
          cell.setBackground('#FFFF00').setFontColor('#FF0000').setFontWeight('bold');
        }
      }
    }
  }

  const ungradedCount = submissions.filter(s => !s.isGraded).length;
  const gradedCount = submissions.filter(s => s.isGraded).length;

  day1.getRange('E2').setValue(ungradedCount);
  day1.getRange('G2').setValue(gradedCount);
}

/**
 * Human-friendly "time ago" string.
 */
function getTimeAgo_(date) {
  const now = new Date();
  const diffMs = now - date;
  const mins = Math.floor(diffMs / 60000);
  const hrs = Math.floor(diffMs / 3600000);
  const days = Math.floor(diffMs / 86400000);

  if (mins < 60) return mins === 1 ? '1 minute ago' : mins + ' minutes ago';
  if (hrs < 24) return hrs === 1 ? '1 hour ago' : hrs + ' hours ago';
  if (days === 1) return '1 day ago';
  return days + ' days ago';
}

// =====================================================
// MISSING SUBMISSIONS (LATEST CREATED ASSIGNMENTS)
// =====================================================

/**
 * Show missing-submissions dialog.
 * Called from the Canvas Hub menu (Bootstrap.gs).
 */
function openMissingSubmissionsDialog() {
  try {
    const settings = getSettings();
    const coursePairs = settings.courseIds.map(id => {
      let name;
      try {
        name = getCourseName_(settings, id);
      } catch (e) {
        name = 'Course ' + id;
      }
      return { id: id, name: name };
    });

    const html = buildMissingDialogHtml_(coursePairs);
    SpreadsheetApp.getUi().showModalDialog(html, 'Check Missing Submissions');
  } catch (e) {
    Logger.log('openMissingSubmissionsDialog error: ' + e);
    SpreadsheetApp.getUi().alert('Error: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Backwards-compatible name, in case any existing script calls it.
 */
function buildMissingDialogHtml(coursePairs) {
  return buildMissingDialogHtml_(coursePairs);
}

/**
 * HTML dialog:
 *  - Course dropdown (All or single course)
 *  - Assignment range dropdown (latest, last 2, …, all)
 *  - Button disables on click, shows "Working…", and closes when server starts.
 */
function buildMissingDialogHtml_(coursePairs) {
  const html = HtmlService.createHtmlOutput(
    `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, sans-serif; margin: 16px; }
          label { display:block; margin: 8px 0 4px; font-weight: bold; }
          select, button { width: 100%; padding: 8px; font-size: 13px; }
          .row { margin-bottom: 12px; }
          .btn { margin-top: 12px; background-color: #2F5597; color: white; border: none; cursor: pointer; }
          .btn[disabled] { background-color: #9FA9C3; cursor: default; }
          .btn:hover:not([disabled]) { background-color: #1D3F73; }
          small { color:#666; display:block; margin-top:4px; }
        </style>
      </head>
      <body>
        <div class="row">
          <label>Course</label>
          <select id="course">
            <option value="ALL">All Classes</option>
            ${coursePairs
              .map(c => `<option value="${c.id}">${c.name} (${c.id})</option>`)
              .join('')}
          </select>
        </div>
        <div class="row">
          <label>Assignment Range</label>
          <select id="range">
            <option value="1">Latest Assignment Only</option>
            <option value="2">Last 2 Assignments</option>
            <option value="3">Last 3 Assignments</option>
            <option value="4">Last 4 Assignments</option>
            <option value="5">Last 5 Assignments</option>
            <option value="ALL">All Assignments (slow)</option>
          </select>
          <small>Assignments are ordered by <strong>created date</strong> (newest first), not due date.</small>
        </div>
        <button class="btn" id="runBtn">Run Missing Submissions</button>
        <small id="statusMsg"></small>

        <script>
          window.addEventListener('load', function () {
            const btn = document.getElementById('runBtn');
            const status = document.getElementById('statusMsg');

            btn.addEventListener('click', function () {
              if (btn.disabled) return;

              const course = document.getElementById('course').value;
              const range = document.getElementById('range').value;

              btn.disabled = true;
              btn.textContent = 'Working...';
              status.textContent =
                'Running Missing Submissions in the background. This may take a moment. ' +
                'Watch the toast messages at the bottom of your sheet.';

              google.script.run
                .withSuccessHandler(function () {
                  google.script.host.close();
                })
                .withFailureHandler(function (err) {
                  alert('Server error: ' + (err && err.message ? err.message : JSON.stringify(err)));
                  btn.disabled = false;
                  btn.textContent = 'Run Missing Submissions';
                  status.textContent = 'Something went wrong. Please try again.';
                })
                .startMissingSubmissions(course, range);
            });
          });
        </script>
      </body>
    </html>
  `
  );
  html.setWidth(380).setHeight(300);
  return html;
}

/**
 * Server entry for Missing Submissions.
 * @param {string} courseChoice 'ALL' or specific courseId.
 * @param {string} rangeChoice  '1' | '2' | '3' | '4' | '5' | 'ALL'.
 */
function startMissingSubmissions(courseChoice, rangeChoice) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  if (!courseChoice || !rangeChoice) {
    ui.alert('Please use Canvas Hub → Check Missing Submissions to run this feature.');
    return;
  }

  let settings;
  try {
    settings = getSettings();
  } catch (e) {
    ui.alert('Error loading settings: ' + e.message);
    return;
  }

  const courseList = courseChoice === 'ALL' ? settings.courseIds : [courseChoice];
  const maxAssignments = rangeChoice === 'ALL' ? 'ALL' : parseInt(rangeChoice, 10);

  const out = prepareMissingSheet_();

  ss.toast('Starting Missing Submissions…', 'Missing Submissions', 5);

  // Safety guard against hitting Apps Script time limits.
  const START = Date.now();
  const MAX_MS = 5.5 * 60 * 1000; // ~5.5 minutes

  let processed = 0;

  for (let i = 0; i < courseList.length; i++) {
    const id = courseList[i];

    ss.toast(
      'Processing course ' + (i + 1) + ' of ' + courseList.length + '…',
      'Missing Submissions',
      10
    );

    try {
      appendMissingForCourse_(settings, id, maxAssignments, out);
    } catch (e) {
      Logger.log('Error in appendMissingForCourse_ for course ' + id + ': ' + e.message);
      // continue with next course
    }

    processed++;

    if (Date.now() - START > MAX_MS && i < courseList.length - 1) {
      const btn = ui.alert(
        'Partial run complete',
        'Finished ' +
          processed +
          ' of ' +
          courseList.length +
          ' courses.\nClick “OK” to continue with the remaining courses.',
        ui.ButtonSet.OK_CANCEL
      );

      if (btn !== ui.Button.OK) {
        ss.toast('Missing Submissions cancelled by user.', 'Missing Submissions', 5);
        return;
      }

      ss.toast('Continuing Missing Submissions…', 'Missing Submissions', 5);
      processed = 0;
    }
  }

  ss.toast('Missing Submissions complete ✔', 'Done', 5);
  ui.alert('Missing Submissions updated.\nCheck the "Missing Submissions" tab.');
}

/**
 * Prepare / clear the "Missing Submissions" sheet.
 */
function prepareMissingSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Missing Submissions');
  if (!sh) sh = ss.insertSheet('Missing Submissions');

  sh.clear();
  sh.setFrozenRows(2);

  sh.getRange('A1:E1').merge();
  sh
    .getRange('A1')
    .setValue('MISSING SUBMISSIONS DASHBOARD')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#C00000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 24);

  const now = new Date();
  sh.getRange('A2').setValue('Last Updated:').setFontWeight('bold');
  sh
    .getRange('B2')
    .setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a'));

  const headers = ['Student Name', 'Assignment', 'Created Date', 'Course', 'SpeedGrader Link'];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]);
  sh
    .getRange(4, 1, 1, headers.length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#2F5597')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 260);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 170);

  return { sheet: sh, nextRow: 5 };
}

/**
 * Append one course’s missing-submissions list.
 */
function appendMissingForCourse_(settings, courseId, maxAssignments, out) {
  const ss = SpreadsheetApp.getActive();
  const sh = out.sheet;

  // Course name for section header.
  let courseName;
  try {
    courseName = getCourseName_(settings, courseId);
  } catch (e) {
    courseName = 'Course ' + courseId;
  }

  // Divider row per course.
  const hdr = out.nextRow++;
  sh.getRange(hdr, 1, 1, 5).merge();
  sh
    .getRange(hdr, 1)
    .setValue(courseName + ' (' + courseId + ')')
    .setBackground('#FBE5D6')
    .setFontWeight('bold');
  sh.setRowHeight(hdr, 22);

  // Data fetch.
  const students = getCourseStudents_(settings, courseId);
  let assignments = getAssignments_(settings, courseId);

  // Sort by created_at DESC (newest first).
  assignments.sort((a, b) => {
    const ad = a.created_at ? new Date(a.created_at).getTime() : -Infinity;
    const bd = b.created_at ? new Date(b.created_at).getTime() : -Infinity;
    return bd - ad;
  });

  if (maxAssignments !== 'ALL') {
    assignments = assignments.slice(0, maxAssignments);
  }

  if (assignments.length === 0) {
    const r = out.nextRow++;
    sh.getRange(r, 1, 1, 5).merge();
    sh
      .getRange(r, 1)
      .setValue('No assignments found for this selection.')
      .setFontStyle('italic')
      .setFontColor('#666666');
    return;
  }

  assignments.forEach((asmt, idx) => {
    ss.toast(
      'Checking "' + asmt.name + '" (' + (idx + 1) + ' of ' + assignments.length + ')…',
      'Missing Submissions',
      10
    );

    const subs = getAssignmentSubmissionsRaw_(settings, courseId, asmt.id);
    const submittedIds = new Set(
      subs
        .filter(s => s && s.submitted_at && s.workflow_state !== 'unsubmitted')
        .map(s => s.user_id)
    );

    const missing = students.filter(stu => !submittedIds.has(stu.id));

    if (missing.length === 0) {
      const r = out.nextRow++;
      sh.getRange(r, 1, 1, 5).merge();
      sh
        .getRange(r, 1)
        .setValue('✓ No missing submissions for: ' + asmt.name)
        .setFontColor('#198754');
      return;
    }

    const createdText = asmt.created_at
      ? Utilities.formatDate(
          new Date(asmt.created_at),
          Session.getScriptTimeZone(),
          'M/d/yyyy h:mm a'
        )
      : '—';

    // Assignment title row.
    const titleRow = out.nextRow++;
    sh.getRange(titleRow, 1, 1, 5).merge();
    sh
      .getRange(titleRow, 1)
      .setValue(asmt.name)
      .setFontWeight('bold')
      .setBackground('#FFF2CC');

    // Missing rows.
    const rows = missing.map(m => {
      const link =
        'https://' +
        settings.baseUrl +
        '/courses/' +
        courseId +
        '/gradebook/speed_grader?assignment_id=' +
        asmt.id +
        '&student_id=' +
        m.id;

      return [m.name, asmt.name, createdText, courseName, link];
    });

    sh.getRange(out.nextRow, 1, rows.length, 5).setValues(rows);

    // Hyperlinks.
    for (let i = 0; i < rows.length; i++) {
      const cell = sh.getRange(out.nextRow + i, 5);
      const url = rows[i][4];
      cell
        .setFormula('=HYPERLINK("' + url + '","Open SpeedGrader")')
        .setHorizontalAlignment('center')
        .setFontColor('#0B5394');
    }

    out.nextRow += rows.length;
  });

  // Blank spacer row after each course section.
  out.nextRow++;
}

/**
 * Course roster (students only).
 */
function getCourseStudents_(settings, courseId) {
  let arr;
  try {
    arr =
      canvasFetch_(
        settings,
        '/api/v1/courses/' + courseId + '/users?enrollment_type[]=student&per_page=100',
        { method: 'get' },
        'fetching students for course ' + courseId
      ) || [];
  } catch (e) {
    Logger.log('Failed to fetch students for ' + courseId + ': ' + e.message);
    return [];
  }

  return arr.map(u => ({
    id: u.id,
    name: u.name || u.sortable_name || 'Unknown Student'
  }));
}

/**
 * Raw submissions for an assignment (no filtering).
 */
function getAssignmentSubmissionsRaw_(settings, courseId, assignmentId) {
  try {
    return (
      canvasFetch_(
        settings,
        '/api/v1/courses/' +
          courseId +
          '/assignments/' +
          assignmentId +
          '/submissions?per_page=100',
        { method: 'get' },
        'fetching raw submissions for assignment ' + assignmentId
      ) || []
    );
  } catch (e) {
    Logger.log(
      'Failed to fetch submissions for assignment ' + assignmentId + ': ' + e.message
    );
    return [];
  }
}

// =====================================================
// SETTINGS UTILITIES / TRIGGERS / HELP
// (Bootstrap.gs owns onOpen + update system)
// =====================================================

/**
 * Reload settings and show a quick summary.
 * Called from Canvas Hub → Update Settings.
 */
function reloadSettings() {
  try {
    const s = getSettings();
    SpreadsheetApp.getUi().alert(
      'Settings Loaded Successfully!\n\n' +
        'Canvas URL: ' +
        s.baseUrl +
        '\n' +
        'Courses: ' +
        s.courseIds.length +
        ' course(s)\n' +
        'Hours Back: ' +
        s.hoursBack +
        '\n' +
        'Show Only Ungraded: ' +
        (s.showOnlyUngraded ? 'Yes' : 'No') +
        '\n' +
        'Highlight Late Submissions: ' +
        (s.highlightLate ? 'Yes' : 'No') +
        '\n\n' +
        'Ready to refresh!'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error loading settings: ' + e.message);
  }
}

/**
 * Create or replace a daily trigger for refreshSubmissions at ~5 AM.
 * Called from Canvas Hub → Set Up Auto-Run (5 AM Daily).
 */
function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshSubmissions') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('refreshSubmissions')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Auto-run set up!\n\nThe script will check for new submissions daily at 5 AM.'
  );
}

/**
 * Simple helper so Bootstrap / future code can read the declared version.
 */
function getLocalVersion() {
  return CANVAS_HUB_VERSION;
}

/**
 * Help dialog, called from Canvas Hub → Help.
 */
function showHelp() {
  const helpText =
    'CANVAS GRADING HUB - QUICK HELP\n\n' +
    'DAILY WORKFLOW:\n' +
    '1) Use the Day 1 tab to see recent submissions.\n' +
    '2) Click "View Submission" to open SpeedGrader in Canvas.\n' +
    '3) Grade in Canvas, then check the "Graded?" box in column A.\n\n' +
    'MISSING SUBMISSIONS:\n' +
    '- Canvas Hub → Check Missing Submissions.\n' +
    '- Choose class (or All) and assignment range.\n' +
    '- Click Run and wait; the dialog will close and toasts will show progress.\n' +
    '- Results are grouped by class and assignment on the "Missing Submissions" tab.\n\n' +
    'SETTINGS:\n' +
    '- Canvas Base URL (e.g., yourschool.instructure.com).\n' +
    '- Canvas API Token (Account → Settings → New Access Token).\n' +
    '- Course IDs (comma-separated Canvas course IDs).\n' +
    '- Hours to Look Back, Show Only Ungraded, Highlight Late Submissions.\n\n' +
    'PERMISSIONS / REVOKE ACCESS:\n' +
    '- You can revoke or review Google permissions at:\n' +
    '  https://myaccount.google.com/permissions\n';

  SpreadsheetApp.getUi().alert(helpText);
}
