/**
 * TEST PORTAL BACKEND — Google Apps Script
 * ------------------------------------------------------------
 * This script turns a Google Sheet (which lives in your Drive)
 * into a free API for the test portal. Every submission gets
 * written as a row, one column per answer block (no JSON blobs).
 *
 * SETUP:
 * 1. Open your Google Sheet, go to Extensions > Apps Script.
 * 2. Select all existing code, delete it, paste in this entire file.
 * 3. Confirm ADMIN_TOKEN below matches CONFIG.ADMIN_PASSWORD in the HTML.
 * 4. In the Sheet, delete every row of the "Submissions" tab (including
 *    the header row) — the column layout changed again, so old rows
 *    won't line up with the new headers otherwise. The script recreates
 *    a correct header automatically on the next write.
 * 5. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *    (Editing code alone does NOT update the live /exec endpoint — you
 *    must create a new version. The URL itself does not change.)
 */

const ADMIN_TOKEN = "changeme123"; // must match CONFIG.ADMIN_PASSWORD in the HTML

const HEADERS = [
  "email", "name", "pre_test_link", "startTime", "durationMinutes",
  "submittedTime", "status", "violations", "violationLog",
  "q1a", "q1b", "q2", "q3_link", "q3_text"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Submissions");
  if (!sheet) {
    sheet = ss.insertSheet("Submissions");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function findRowByEmail_(sheet, email) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(email).toLowerCase()) {
      return i + 1; // 1-indexed sheet row
    }
  }
  return -1;
}

function formatViolationLog_(log) {
  if (!log || !log.length) return "";
  return log.map(v => {
    const time = v.at ? new Date(v.at).toLocaleTimeString() : "";
    return `${v.type} @ ${time}`;
  }).join("\n");
}

function parseViolationLog_(str) {
  if (!str) return [];
  // Reconstructed only for resuming an in-progress test; exact timestamps
  // aren't recoverable from the readable format, but the count is preserved
  // so a resumed session keeps accumulating flags correctly.
  return str.split("\n").filter(Boolean).map(line => ({ type: line, at: null }));
}

function recordToRow_(record) {
  const answers = record.answers || {};
  return [
    record.email || "",
    record.name || "",
    record.preTestLink || "",
    record.startTime || "",
    record.durationMinutes || "",
    record.submittedTime || "",
    record.status || "",
    record.violations || 0,
    formatViolationLog_(record.violationLog),
    answers.q1a || "",
    answers.q1b || "",
    answers.q2 || "",
    answers.q3_link || "",
    answers.q3_text || ""
  ];
}

function rowToRecord_(row) {
  return {
    email: row[0],
    name: row[1],
    preTestLink: row[2],
    startTime: row[3],
    durationMinutes: row[4],
    submittedTime: row[5],
    status: row[6],
    violations: row[7],
    violationLog: parseViolationLog_(row[8]),
    answers: { q1a: row[9], q1b: row[10], q2: row[11], q3_link: row[12], q3_text: row[13] }
  };
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet_();

  if (body.action === "save") {
    const record = body.record;
    const rowIndex = findRowByEmail_(sheet, record.email);
    const rowValues = recordToRow_(record);
    if (rowIndex === -1) {
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    }
    return jsonOut_({ ok: true });
  }

  return jsonOut_({ ok: false, error: "unknown action" });
}

function doGet(e) {
  const params = e.parameter;
  const sheet = getSheet_();

  if (params.action === "get" && params.email) {
    const rowIndex = findRowByEmail_(sheet, params.email);
    if (rowIndex === -1) return jsonOut_({ found: false });
    const row = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
    return jsonOut_({ found: true, record: rowToRecord_(row) });
  }

  if (params.action === "list") {
    if (params.token !== ADMIN_TOKEN) {
      return jsonOut_({ error: "unauthorized" });
    }
    const data = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < data.length; i++) {
      records.push(rowToRecord_(data[i]));
    }
    return jsonOut_({ records });
  }

  return jsonOut_({ error: "unknown action" });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
