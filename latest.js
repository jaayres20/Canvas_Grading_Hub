/****************************************************
 * CANVAS GRADING HUB — VERSION 2.1 (Standalone)
 * COMPLETE SINGLE-FILE VERSION
 *
 * No bootstrap, no GitHub loading.
 * Everything runs locally for speed & reliability.
 *
 * Features:
 *  - Day 1–5 Recent Submissions Dashboard
 *  - Missing Submissions (latest assignment or range)
 *  - Settings loader
 *  - Daily 5 AM trigger setup
 *  - Help dialog
 ****************************************************/

const CANVAS_HUB_VERSION = '2.1';

/****************************************************
 * MENU
 ****************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Canvas Hub')
    .addItem('Refresh Now (Recent Submissions)', 'refreshSubmissions')
    .addSeparator()
    .addItem('Check Missing Submissions', 'openMissingSubmissionsDialog')
    .addSeparator()
    .addItem('Update Settings', 'reloadSettings')
    .addItem('Set Up Auto-Run (5 AM Daily)', 'setupDailyTrigger')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/****************************************************
 * SETTINGS
 ****************************************************/
function getSettings() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Settings');
  if (!sh) throw new Error('Settings sheet missing.');

  const data = sh.getDataRange().getValues();
  const settings = {};

  for (let i = 2; i < data.length; i++) {
    const key = data[i][0]?.toString().trim();
    const value = data[i][1]?.toString().trim();
    if (key) settings[key] = value;
  }

  if (!settings['Canvas Base URL'])
    throw new Error('Enter Canvas Base URL in Settings.');
  if (!settings['Canvas API Token'])
    throw new Error('Enter Canvas API Token in Settings.');
  if (!settings['Course IDs (comma-separated)'])
    throw new Error('Enter Course IDs in Settings.');

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
    showOnlyUngraded: normalizeYesNo(settings['Show Only Ungraded?']),
    highlightLate: normalizeYesNo(settings['Highlight Late Submissions?'])
  };
}

function normalizeYesNo(v) {
  if (!v) return false;
  const s = v.toString().toLowerCase().trim();
  return s === 'yes' || s === 'true';
}

function canvasHeaders(settings) {
  return { Authorization: 'Bearer ' + settings.apiToken };
}

function canvasFetch(settings, url, context) {
  const fullUrl = url.startsWith('http')
    ? url
    : 'https://' + settings.baseUrl + url;

  const res = UrlFetchApp.fetch(fullUrl, {
    method: 'get',
    muteHttpExceptions: true,
    headers: canvasHeaders(settings)
  });

  const code = res.getResponseCode();
  if (code === 401 || code === 403) {
    throw new Error(
      'Canvas API ' + code + ' while ' + context +
        '\nCheck token permissions or course access.'
    );
  }
  if (code < 200 || code >= 300) {
    throw new Error(
      'Canvas API error ' + code + ' for ' + context +
        '\nResponse: ' + res.getContentText().slice(0, 400)
    );
  }

  return JSON.parse(res.getContentText());
}

/****************************************************
 * RECENT SUBMISSIONS (DAY 1–5)
 ****************************************************/
function refreshSubmissions() {
  const settings = getSettings();
  const subs = fetchRecentSubmissions(settings);

  if (!subs.length) {
    SpreadsheetApp.getUi().alert(
      'No submissions in the last ' + settings.hoursBack + ' hours.'
    );
    return;
  }

  rotateDayTabs();
  populateDay1(subs, settings);

  SpreadsheetApp.getUi().alert(
    'Found ' + subs.length + ' new submissions.\nCheck Day 1.'
  );
}

function fetchRecentSubmissions(settings) {
  const out = [];
  const cutoff = new Date(Date.now() - settings.hoursBack * 3600 * 1000);

  settings.courseIds.forEach(courseId => {
    let assignments;
    try {
      assignments = canvasFetch(
        settings,
        `/api/v1/courses/${courseId}/assignments?per_page=100`,
        'fetching assignments'
      );
    } catch {
      assignments = [];
    }

    assignments.forEach(asmt => {
      const subs = canvasFetch(
        settings,
        `/api/v1/courses/${courseId}/assignments/${asmt.id}/submissions?include[]=user&per_page=100`,
        'fetching submissions'
      );

      subs.forEach(s => {
        if (!s.submitted_at) return;
        const d = new Date(s.submitted_at);
        if (d < cutoff) return;

        if (settings.showOnlyUngraded && s.workflow_state === 'graded') return;

        out.push({
          student: s.user?.name || 'Unknown',
          assignment: asmt.name,
          course: courseId,
          date: d,
          isLate: !!s.late,
          isGraded: s.workflow_state === 'graded',
          url:
            `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader` +
            `?assignment_id=${asmt.id}&student_id=${s.user_id}`
        });
      });
    });
  });

  out.sort((a, b) => b.date - a.date);
  return out;
}

function rotateDayTabs() {
  const ss = SpreadsheetApp.getActive();

  const day5 = ss.getSheetByName('Day 5');
  if (day5) ss.deleteSheet(day5);

  for (let i = 4; i >= 1; i--) {
    const sh = ss.getSheetByName('Day ' + i);
    if (sh) sh.setName('Day ' + (i + 1));
  }

  const day1 = ss.insertSheet('Day 1', 0);
  setupDay1Header(day1);
}

function setupDay1Header(sh) {
  sh.getRange('A1:G1').merge();
  sh.getRange('A1')
    .setValue('CANVAS GRADING DASHBOARD')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#4472C4')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  sh.getRange('A2').setValue('Last Updated:').setFontWeight('bold');
  sh.getRange('B2').setValue(new Date());

  const headers = [
    'Graded?',
    'Student',
    'Class',
    'Assignment',
    'Submitted',
    'Link',
    'Notes'
  ];
  sh.getRange(4, 1, 1, headers.length).setValues([headers]);
  sh.getRange(4, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#70AD47')
    .setFontColor('white');
}

function populateDay1(subs, settings) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Day 1');

  const rows = subs.map(s => {
    const ago = getTimeAgo(s.date) + (s.isLate ? ' (LATE)' : '');
    return [false, s.student, s.course, s.assignment, ago, s.url, ''];
  });

  if (rows.length) {
    sh.getRange(5, 1, rows.length, 7).setValues(rows);
    sh.getRange(5, 1, rows.length, 1).insertCheckboxes();

    rows.forEach((r, i) => {
      sh.getRange(5 + i, 6).setFormula(
        `=HYPERLINK("${r[5]}", "View")`
      );
    });
  }
}

function getTimeAgo(d) {
  const diff = (new Date() - d) / 1000;
  if (diff < 3600) return Math.floor(diff / 60) + ' minutes ago';
  if (diff < 86400) return Math.floor(diff / 3600) + ' hours ago';
  return Math.floor(diff / 86400) + ' days ago';
}

/****************************************************
 * MISSING SUBMISSIONS
 ****************************************************/
function openMissingSubmissionsDialog() {
  const settings = getSettings();

  const courseOptions = settings.courseIds
    .map(id => `<option value="${id}">Course ${id}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
    <html><body>
      <h3>Missing Submissions</h3>
      <label>Course:</label>
      <select id="course">
        <option value="ALL">All Courses</option>${courseOptions}
      </select>
      <label>Assignments:</label>
      <select id="range">
        <option value="1">Latest Assignment</option>
        <option value="2">Last 2</option>
        <option value="3">Last 3</option>
        <option value="5">Last 5</option>
        <option value="ALL">All Assignments</option>
      </select>
      <button onclick="run()">Run</button>

      <script>
        function run() {
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .startMissingSubmissions(
              document.getElementById('course').value,
              document.getElementById('range').value
            );
        }
      </script>
    </body></html>
  `).setWidth(300).setHeight(240);

  SpreadsheetApp.getUi().showModalDialog(html, 'Missing Submissions');
}

function startMissingSubmissions(courseChoice, rangeChoice) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();

  const sh = ss.getSheetByName('Missing Submissions') || ss.insertSheet('Missing Submissions');
  sh.clear();

  sh.getRange('A1:D1').merge();
  sh.getRange('A1')
    .setValue('MISSING SUBMISSIONS')
    .setFontWeight('bold')
    .setBackground('#C00000')
    .setFontColor('white')
    .setHorizontalAlignment('center');

  let nextRow = 3;

  const courseList = courseChoice === 'ALL'
    ? settings.courseIds
    : [courseChoice];

  courseList.forEach(courseId => {
    sh.getRange(nextRow, 1).setValue('Course ' + courseId).setFontWeight('bold');
    nextRow++;

    const students = canvasFetch(
      settings,
      `/api/v1/courses/${courseId}/users?enrollment_type[]=student&per_page=100`,
      'fetching students'
    );

    const assignments = canvasFetch(
      settings,
      `/api/v1/courses/${courseId}/assignments?per_page=100`,
      'fetching assignments'
    );

    assignments.sort((a, b) =>
      new Date(b.created_at) - new Date(a.created_at)
    );

    const subset =
      rangeChoice === 'ALL'
        ? assignments
        : assignments.slice(0, parseInt(rangeChoice, 10));

    subset.forEach(asmt => {
      sh.getRange(nextRow, 1).setValue(asmt.name).setFontWeight('bold');
      nextRow++;

      const subs = canvasFetch(
        settings,
        `/api/v1/courses/${courseId}/assignments/${asmt.id}/submissions?per_page=100`,
        'fetching subs'
      );

      const submittedIds = new Set(
        subs.filter(s => s.submitted_at).map(s => s.user_id)
      );

      const missing = students.filter(stu => !submittedIds.has(stu.id));

      missing.forEach(stu => {
        const link =
          `https://${settings.baseUrl}/courses/${courseId}/gradebook/speed_grader` +
          `?assignment_id=${asmt.id}&student_id=${stu.id}`;

        sh.getRange(nextRow, 1, 1, 4).setValues([
          [stu.name, asmt.name, new Date(asmt.created_at), link]
        ]);

        sh.getRange(nextRow, 4).setFormula(
          `=HYPERLINK("${link}", "SpeedGrader")`
        );

        nextRow++;
      });

      nextRow++;
    });

    nextRow++;
  });

  ui.alert('Missing Submissions updated.');
}

/****************************************************
 * TRIGGERS & HELP
 ****************************************************/
function reloadSettings() {
  const s = getSettings();
  SpreadsheetApp.getUi().alert(
    `Settings loaded:\n\nBase URL: ${s.baseUrl}\nCourses: ${s.courseIds.length}`
  );
}

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'refreshSubmissions') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('refreshSubmissions')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  SpreadsheetApp.getUi().alert('Daily trigger set for ~5 AM.');
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    'Canvas Hub Help\n\n' +
    '• Refresh Now → pulls recent submissions\n' +
    '• Check Missing Submissions → find missing work\n' +
    '• Update Settings → reload settings\n' +
    '• Daily trigger → automate daily checks\n'
  );
}
