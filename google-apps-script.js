// Google Apps Script — deploy as Web App
// 1. Open the "ICC Lesson Log Worksheet for App" Google Sheet
// 2. Go to Extensions > Apps Script
// 3. Paste this code
// 4. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL and paste it into the app's Settings
//
// Tabs are created automatically on first use — no manual sheet setup:
// - "Pros" tab: the pro list (seeded with the current roster). To add or
//   remove a pro, just edit the names in column A — the app's dropdown
//   updates on next load.
// - "Lesson Log" tab: one row per lesson, all pros in one list, with a
//   "Charged" checkbox the shop manager ticks after entering it in Jonas.

const HEADERS = ['Date', 'Pro', 'Client Name', 'Member/Guest', 'Duration', 'People', 'Notes', 'Charged'];
const LOG_SHEET_NAME = 'Lesson Log';
const PROS_SHEET_NAME = 'Pros';
const CHARGED_COL = 8;
const GUEST_MEMBER_COL = 4;

const DEFAULT_PROS = [
  'J.C. Freeman',
  'Joey Francis',
  'A.B. Hill',
  'Will Davidson',
  'Matt Kendrick',
  'Lisa Webb',
  'Stephanie Heckler'
];

function getProsSheet_(ss) {
  var prosSheet = ss.getSheetByName(PROS_SHEET_NAME);
  if (!prosSheet) {
    prosSheet = ss.insertSheet(PROS_SHEET_NAME);
    prosSheet.appendRow(['Pro Names (edit anytime — the app updates on next load)']);
    prosSheet.getRange(1, 1).setFontWeight('bold');
    prosSheet.setFrozenRows(1);
    for (var i = 0; i < DEFAULT_PROS.length; i++) {
      prosSheet.appendRow([DEFAULT_PROS[i]]);
    }
    prosSheet.setColumnWidth(1, 400);
  }
  return prosSheet;
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'getPros') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prosSheet = getProsSheet_(ss);

    var pros = [];
    var lastRow = prosSheet.getLastRow();
    if (lastRow > 1) {
      var values = prosSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < values.length; i++) {
        var name = values[i][0].toString().trim();
        if (name) pros.push(name);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', pros: pros }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput('ICC Lesson Log API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME, 0);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#1a1a2e');
      sheet.getRange(1, 1, 1, HEADERS.length).setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.date,
      data.pro || 'Unknown',
      data.client,
      data.guestMember,
      data.duration,
      data.people || 1,
      data.notes || '',
      false
    ]);

    var newRow = sheet.getLastRow();
    sheet.getRange(newRow, CHARGED_COL).insertCheckboxes();

    var gmCell = sheet.getRange(newRow, GUEST_MEMBER_COL);
    if (data.guestMember === 'GUEST') {
      gmCell.setBackground('#e74c3c');
      gmCell.setFontColor('#ffffff');
      gmCell.setFontWeight('bold');
    } else if (data.guestMember === 'MEMBER') {
      gmCell.setBackground('#2ecc71');
      gmCell.setFontColor('#ffffff');
      gmCell.setFontWeight('bold');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
