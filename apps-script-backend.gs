/**
 * TEST PORTAL BACKEND — Google Apps Script
 * ------------------------------------------------------------
 * This script turns a Google Sheet (which lives in your Drive)
 * into a free API for the test portal. Every submission gets
 * written as a row you can view, sort, filter, or move into any
 * Drive folder like a normal spreadsheet.
 *
 * SETUP (about 5 minutes):
 * 1. Go to https://sheet.new to create a fresh Google Sheet.
 *    Rename it something like "Test Portal Submissions".
 *    (Move it into whatever Drive folder you want — the script
 *    doesn't care where the sheet lives.)
 * 2. In the Sheet, go to Extensions > Apps Script.
 * 3. Delete the placeholder code and paste in this entire file.
 * 4. Change ADMIN_TOKEN below to match CONFIG.ADMIN_PASSWORD in
 *    the HTML file (or just leave both as "changeme123" and
 *    change them together).
 * 5. Click Deploy > New deployment.
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 *    (This does NOT make your sheet public — it only exposes
 *    this script's doGet/doPost functions, and reading the full
 *    list still requires the ADMIN_TOKEN.)
 * 6. Click Deploy, authorize the permissions it asks for, and
 *    copy the Web app URL it gives you (ends in /exec).
 * 7. Paste that URL into CONFIG.BACKEND_URL in test-portal.html.
 *
 * That's it — submissions will now append/update as rows in
 * this Sheet in real time as candidates take the test.
 */

const ADMIN_TOKEN = "changeme123"; // must match CONFIG.ADMIN_PASSWORD in the HTML

const HEADERS = [
  "email", "name", "startTime", "durationMinutes", "submittedTime", "status", "violations", "violationLog",
  "q1_github_link", "q2", "q3", "q4"
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
    record.startTime || "",
    record.durationMinutes || "",
    record.submittedTime || "",
    record.status || "",
    record.violations || 0,
    formatViolationLog_(record.violationLog),
    answers.q1 || "",
    answers.q2 || "",
    answers.q3 || "",
    answers.q4 || ""
  ];
}

function rowToRecord_(row) {
  return {
    email: row[0],
    name: row[1],
    startTime: row[2],
    durationMinutes: row[3],
    submittedTime: row[4],
    status: row[5],
    violations: row[6],
    violationLog: parseViolationLog_(row[7]),
    answers: { q1: row[8], q2: row[9], q3: row[10], q4: row[11] }
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
