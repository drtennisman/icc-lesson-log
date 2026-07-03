// Google Apps Script — deploy as Web App
// 1. Create a new Google Sheet
// 2. Create a tab called "Pros" and list pro names in column A (one per row)
// 3. Go to Extensions > Apps Script
// 4. Paste this code
// 5. Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Copy the deployment URL and paste it into the app's Settings
//
// All lessons land in a single "Lesson Log" tab so the shop manager can
// work down one list and tick the "Charged" checkbox after entering each
// lesson in Jonas.

const HEADERS = ['Date', 'Pro', 'Client Name', 'Member/Guest', 'Duration', 'People', 'Notes', 'Charged'];
const LOG_SHEET_NAME = 'Lesson Log';
const CHARGED_COL = 8;
const GUEST_MEMBER_COL = 4;

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  if (action === 'getPros') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prosSheet = ss.getSheetByName('Pros');

    var pros = [];
    if (prosSheet) {
      var lastRow = prosSheet.getLastRow();
      if (lastRow > 0) {
        var values = prosSheet.getRange(1, 1, lastRow, 1).getValues();
        for (var i = 0; i < values.length; i++) {
          var name = values[i][0].toString().trim();
          if (name) pros.push(name);
        }
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
