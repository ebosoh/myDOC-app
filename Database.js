/**
 * myDoc - Database Setup Functions
 */

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  for (const [sheetName, headers] of Object.entries(SHEET_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      // Basic formatting
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2370b8').setFontColor('#ffffff');
    }
  }
}
