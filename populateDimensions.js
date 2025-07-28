/**
 * V.10.2 — Populate Dimensions (Manual Tool)
 * Last updated: 2025-07-05 @ 17:06
 * ✅ Matches products by title + size
 * ✅ Injects AS–AV (dimensions) for 1-pack and non-multipack
 */


function populateDimensionsForWallOrDoorSign() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const attrSheet = ss.getSheetByName("Attributes");




  if (!flatSheet || !attrSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets missing: 'Partial Flat File' or 'Attributes'.");
    return;
  }




  const flatData = flatSheet.getRange(4, 1, flatSheet.getLastRow() - 3, 48).getValues();
  const attrData = attrSheet.getRange(2, 1, attrSheet.getLastRow() - 1, 12).getValues();
  let injectedCount = 0;




  for (let i = 0; i < flatData.length; i++) {
    const row = flatData[i];
    const title = row[9];
    const sizeText = row[38];




    if (!title || !sizeText) continue;




    const match = attrData.find(attr =>
      attr[0]?.toLowerCase() === "wall or door sign (lasered)" &&
      title.toLowerCase().includes(attr[1]?.toLowerCase()) &&
      sizeText.toLowerCase().includes(attr[2]?.toLowerCase())
    );




    if (match) {
      flatSheet.getRange(i + 4, 45).setValue(match[3]); // AS: Length
      flatSheet.getRange(i + 4, 46).setValue(match[4]); // AT: Length Unit
      flatSheet.getRange(i + 4, 47).setValue(match[5]); // AU: Width
      flatSheet.getRange(i + 4, 48).setValue(match[6]); // AV: Width Unit
      injectedCount++;
    }
  }




  SpreadsheetApp.getUi().alert(`✅ Dimensions populated for ${injectedCount} rows.`);
}




function populateDimensionsForRubberStamp() {
  SpreadsheetApp.getUi().alert("ℹ️ Rubber Stamp dimension logic not yet implemented.");
}
