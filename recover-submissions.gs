/**
 * ONE-OFF RECOVERY SCRIPT
 * ------------------------------------------------------------
 * Run this BEFORE you clear the Submissions sheet or redeploy the new
 * backend. It reads the existing rows, works out which column layout each
 * one was written under, and writes them into a new "Recovered" sheet with
 * the correct labels. Nothing is deleted or overwritten.
 *
 * HOW TO RUN:
 * 1. Open your Google Sheet > Extensions > Apps Script.
 * 2. Add a new file (the + next to Files) and paste this in. Leave the
 *    existing backend code alone.
 * 3. In the function dropdown at the top, choose  recoverSubmissions
 *    and press Run. Approve permissions if asked.
 * 4. Go back to the Sheet. You'll have two new tabs:
 *      Submissions_backup_<date>  — an untouched copy of what was there
 *      Recovered                  — the same rows, correctly labelled
 * 5. Read the Recovered tab and check the answers sit under sensible
 *    headers. Then clear Submissions and deploy the new backend.
 *
 * WHY THE ANSWERS MOVED:
 * Rows are written positionally. Each time the column list changed, the
 * portal started sending more values while the sheet kept its old header
 * row, so everything after q1b shifted left relative to the labels. The
 * values themselves were saved correctly every time.
 */

/** The column layouts this portal has used, oldest first. */
const LAYOUT_V1 = [
  "email", "name", "pre_test_link", "startTime", "durationMinutes",
  "submittedTime", "status", "violations", "violationLog",
  "q1a", "q1b", "q2", "q3_link", "q3_text"
];

const LAYOUT_V2 = [
  "email", "name", "pre_test_link", "startTime", "durationMinutes",
  "submittedTime", "status", "violations", "violationLog",
  "q1a", "q1b", "q2a", "q2b", "q3_link", "q3_text"
];

const LAYOUT_V3 = [
  "email", "name", "pre_test_link", "startTime", "durationMinutes",
  "submittedTime", "status", "violations", "violationLog",
  "q1a", "q1b", "q1_revealed", "q2a", "q2b", "q2_revealed",
  "q3_link", "q3_text"
];

/** What the recovered sheet will look like. */
const CANONICAL = LAYOUT_V3;

function recoverSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Submissions");
  if (!sheet) throw new Error('No sheet named "Submissions" was found.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error("Submissions has no data rows to recover.");

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // 1. Back up first, so this is always reversible.
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd-HHmm");
  const backupName = "Submissions_backup_" + stamp;
  if (!ss.getSheetByName(backupName)) {
    const backup = ss.insertSheet(backupName);
    backup.getRange(1, 1, data.length, lastCol).setValues(data);
  }

  // 2. Rebuild each data row under the layout it was actually written with.
  const out = [CANONICAL.concat(["_detected_layout"])];
  const notes = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row)) continue;

    const layout = detectLayout_(row);
    const values = {};
    layout.names.forEach(function (name, i) {
      // The old single-answer q2 column becomes q2a.
      const target = (name === "q2") ? "q2a" : name;
      values[target] = row[i];
    });

    out.push(CANONICAL.map(function (name) {
      return values[name] === undefined ? "" : values[name];
    }).concat([layout.label]));

    notes.push(
      "Row " + (r + 1) + ": " + layout.label +
      " | email=" + String(row[0]) +
      " | q2b length=" + String(values["q2b"] || "").length
    );
  }

  // 3. Write the result to a fresh tab.
  let dest = ss.getSheetByName("Recovered");
  if (dest) dest.clear(); else dest = ss.insertSheet("Recovered");
  dest.getRange(1, 1, out.length, out[0].length).setValues(out);
  dest.setFrozenRows(1);

  Logger.log("Backed up to: " + backupName);
  Logger.log("Recovered " + (out.length - 1) + " row(s):");
  notes.forEach(function (n) { Logger.log("  " + n); });
  Logger.log("Open the Recovered tab and check the answers sit under sensible headers.");
}

function isBlankRow_(row) {
  return row.every(function (cell) { return String(cell || "").trim() === ""; });
}

/**
 * Each save rewrote the whole row at the width of the code version running at
 * the time, and the column list only ever grew. So the right-most cell that
 * holds anything tells us which layout the row was last written with.
 */
function detectLayout_(row) {
  let last = -1;
  for (let i = 0; i < row.length; i++) {
    if (String(row[i] || "").trim() !== "") last = i;
  }

  if (last >= LAYOUT_V3.length - 1) return { names: LAYOUT_V3, label: "v3 (17 cols)" };
  if (last >= LAYOUT_V2.length - 1) return { names: LAYOUT_V2, label: "v2 (15 cols)" };
  if (last >= LAYOUT_V1.length - 1) return { names: LAYOUT_V1, label: "v1 (14 cols)" };

  // Trailing fields were left empty (e.g. an abandoned attempt). Fall back to
  // the widest layout that could contain what is present, and flag it.
  if (last >= LAYOUT_V2.length) return { names: LAYOUT_V3, label: "v3 assumed - CHECK" };
  if (last >= LAYOUT_V1.length) return { names: LAYOUT_V2, label: "v2 assumed - CHECK" };
  return { names: LAYOUT_V1, label: "v1 assumed - CHECK" };
}
