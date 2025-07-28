/**
 * V.10.2 — Clear Partial Flat File Utility
 * Last updated: 2025-07-05 @ 17:06
 * ✅ Clears data from row 4 downward
 * ✅ Leaves headers intact
 * ✅ Feedback alert after completion
 */


function clearPartialFlatFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Partial Flat File");


  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 'Partial Flat File' sheet not found.");
    return;
  }


  const lastRow = sheet.getLastRow();
  if (lastRow <= 3) {
    return;  // Skip message
  }


  const numRowsToClear = lastRow - 3;
  sheet.getRange(4, 1, numRowsToClear, sheet.getLastColumn()).clearContent();
}
