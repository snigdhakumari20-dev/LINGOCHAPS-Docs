/**
 * Lingo Chaps — PM Onboarding Quiz Logger
 *
 * Attach this script to your PM Google Sheet only (separate from Reviewer sheet).
 *
 * SETUP:
 * 1. Create a NEW Google Sheet for PM quiz results (or use an existing PM-only sheet)
 * 2. Extensions → Apps Script → paste this file → Save
 * 3. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web App URL into pm_onboarding_quiz.html → GOOGLE_SCRIPT_URL
 */

var SHEET_NAME = "PM Quiz Logs";
var HEADERS = [
  "Timestamp",
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
    var sheet = getOrCreateSheet_(SHEET_NAME);

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.email || "",
      data.score !== undefined ? data.score : "",
      data.correctCount !== undefined ? data.correctCount : "",
      data.totalQuestions !== undefined ? data.totalQuestions : "",
      data.resultStatus || "",
      data.timeTaken || "",
      data.feedback || ""
    ]);

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
