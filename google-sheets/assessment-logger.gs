/**
 * Lingo Chaps — Assessment Results Logger
 *
 * SETUP (one time):
 * 1. Open your Google Sheet (manager database sheet).
 * 2. Extensions → Apps Script → paste this entire file → Save.
 * 3. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web App URL into pm_onboarding_quiz.html (GOOGLE_SCRIPT_URL).
 *
 * Sheet tabs created automatically:
 *   - "PM Quiz Logs"        — PM Onboarding Quiz results
 *   - "Reviewer Quiz Logs"  — Reviewer Certification results
 */

var HEADERS = [
  "Timestamp",
  "Quiz Type",
  "Name",
  "Email",
  "Score (%)",
  "Correct",
  "Total",
  "Result",
  "Time Taken",
  "Feedback"
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var quizType = data.quizType || "Unknown Quiz";
    var sheetName = quizType.indexOf("PM") !== -1 ? "PM Quiz Logs" : "Reviewer Quiz Logs";
    var sheet = getOrCreateSheet_(sheetName);

    var row = [
      new Date(),
      quizType,
      data.name || "",
      data.email || "",
      data.score !== undefined ? data.score : "",
      data.correctCount !== undefined ? data.correctCount : "",
      data.totalQuestions !== undefined ? data.totalQuestions : "",
      data.resultStatus || "",
      data.timeTaken || "",
      data.feedback || ""
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}
