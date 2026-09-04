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
 * 4. Deploy > Manage deployments > pencil icon > Version: New version > Deploy.
 *    (Editing code alone does NOT update the live /exec endpoint — you
 *    must create a new version. The URL itself does not change.)
 *
 * You no longer need to clear the sheet by hand when the columns change.
 * Every read and write now looks columns up BY NAME from the header row, and
 * the sheet is migrated automatically on the next request: columns are
 * reordered to match HEADERS, new ones are added blank, and existing answers
 * follow their header wherever it moves.
 *
 * Previously the header row was only written when the sheet was completely
 * empty. After a column change the script therefore kept writing the new,
 * wider set of values positionally underneath a stale header — so a Part B
 * answer was saved, but landed under a column labelled something else and
 * looked like it had never been submitted at all.
 */

const ADMIN_TOKEN = "changeme123"; // must match CONFIG.ADMIN_PASSWORD in the HTML

const HEADERS = [
  "email", "name", "pre_test_link", "startTime", "durationMinutes",
  "submittedTime", "status", "violations", "violationLog",
  "q1a", "q1b", "q1_revealed", "q2a", "q2b", "q2_revealed",
  "q3_link", "q3_text"
];

/**
 * Headers used by earlier versions, mapped to their current name, so answers
 * written under an old layout are carried across instead of dropped.
 */
const RENAMED_HEADERS = {
  "q2": "q2a"
};

/** Every answer field the portal can send, keyed by its column name. */
const ANSWER_FIELDS = [
  "q1a", "q1b", "q1_revealed",
  "q2a", "q2b", "q2_revealed",
  "q3_link", "q3_text"
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Submissions");
  if (!sheet) {
    sheet = ss.insertSheet("Submissions");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return sheet;
  }

  const width = Math.max(sheet.getLastColumn(), HEADERS.length);
  const current = sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(function (h) { return String(h || "").trim(); });

  const matches = HEADERS.every(function (h, i) { return current[i] === h; }) &&
    current.slice(HEADERS.length).every(function (h) { return h === ""; });

  if (!matches) {
    migrateSheet_(sheet, current);
  }
  return sheet;
}

/**
 * Rebuilds the sheet in the current HEADERS order, keeping each value with its
 * header. Anything under a header we no longer recognise is dropped; anything
 * newly added starts blank.
 */
function migrateSheet_(sheet, oldHeaders) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Where does each current header live in the old layout?
  const indexOf = {};
  oldHeaders.forEach(function (h, i) {
    const name = RENAMED_HEADERS[h] || h;
    if (name && indexOf[name] === undefined) indexOf[name] = i;
  });

  const rebuilt = [HEADERS];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const empty = row.every(function (cell) { return String(cell || "") === ""; });
    if (empty) continue;
    rebuilt.push(HEADERS.map(function (h) {
      const i = indexOf[h];
      return i === undefined ? "" : row[i];
    }));
  }

  sheet.clear();
  sheet.getRange(1, 1, rebuilt.length, HEADERS.length).setValues(rebuilt);
  SpreadsheetApp.flush();
}

function headerIndex_(sheet) {
  const header = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0]
    .map(function (h) { return String(h || "").trim(); });
  const index = {};
  header.forEach(function (h, i) { if (h) index[h] = i; });
  return index;
}

function findRowByEmail_(sheet, email) {
  const index = headerIndex_(sheet);
  const emailCol = index["email"];
  if (emailCol === undefined) return -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
      return i + 1; // 1-indexed sheet row
    }
  }
  return -1;
}

function formatViolationLog_(log) {
  if (!log || !log.length) return "";
  return log.map(function (v) {
    const time = v.at ? new Date(v.at).toLocaleTimeString() : "";
    return v.type + " @ " + time;
  }).join("\n");
}

function parseViolationLog_(str) {
  if (!str) return [];
  // Reconstructed only for resuming an in-progress test; exact timestamps
  // aren't recoverable from the readable format, but the count is preserved
  // so a resumed session keeps accumulating flags correctly.
  return String(str).split("\n").filter(Boolean).map(function (line) {
    return { type: line, at: null };
  });
}

function recordToRow_(record) {
  const answers = record.answers || {};
  const values = {
    email: record.email || "",
    name: record.name || "",
    pre_test_link: record.preTestLink || "",
    startTime: record.startTime || "",
    durationMinutes: record.durationMinutes || "",
    submittedTime: record.submittedTime || "",
    status: record.status || "",
    violations: record.violations || 0,
    violationLog: formatViolationLog_(record.violationLog)
  };
  ANSWER_FIELDS.forEach(function (f) { values[f] = answers[f] || ""; });

  // Built in HEADERS order, so adding or moving a column can never shift data.
  return HEADERS.map(function (h) {
    return values[h] === undefined ? "" : values[h];
  });
}

function rowToRecord_(row, index) {
  function get(name) {
    const i = index[name];
    return i === undefined ? "" : row[i];
  }
  const answers = {};
  ANSWER_FIELDS.forEach(function (f) { answers[f] = get(f); });

  return {
    email: get("email"),
    name: get("name"),
    preTestLink: get("pre_test_link"),
    startTime: get("startTime"),
    durationMinutes: get("durationMinutes"),
    submittedTime: get("submittedTime"),
    status: get("status"),
    violations: get("violations"),
    violationLog: parseViolationLog_(get("violationLog")),
    answers: answers
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
  const index = headerIndex_(sheet);

  if (params.action === "get" && params.email) {
    const rowIndex = findRowByEmail_(sheet, params.email);
    if (rowIndex === -1) return jsonOut_({ found: false });
    const row = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
    return jsonOut_({ found: true, record: rowToRecord_(row, index) });
  }

  if (params.action === "list") {
    if (params.token !== ADMIN_TOKEN) {
      return jsonOut_({ error: "unauthorized" });
    }
    const data = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < data.length; i++) {
      records.push(rowToRecord_(data[i], index));
    }
    return jsonOut_({ records });
  }

  // Quick check that the live /exec endpoint is running THIS version of the
  // code and that the sheet columns are what the portal expects.
  if (params.action === "ping") {
    return jsonOut_({ ok: true, expected: HEADERS, sheet: Object.keys(index) });
  }

  return jsonOut_({ error: "unknown action" });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
