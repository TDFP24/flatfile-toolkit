/**
 * V.10.2 — Pack Size Variant Generator
 * ✅ Branded & unbranded SKUs
 * ✅ Appends '2' to full parent SKU including '-MAIN'
 * ✅ Injects final parent SKU into AA5:AA
 * ✅ Clears validation from AA5:AA
 * ✅ Includes price + dimension injectors
 */


function runWallOrDoorSignsLasered() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Partial Flat File");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 'Partial Flat File' not found.");
    return;
  }
  const sku = sheet.getRange(4, 2).getValue();
  if (!sku) {
    SpreadsheetApp.getUi().alert("❌ No SKU found in B4.");
    return;
  }


  try {
    if (sku.includes("-")) {
      generatePackVariants_WallOrDoorSigns_BrandPrefix();
    } else {
      generatePackVariants_WallOrDoorSigns_UnbrandPrefix();
    }
    SpreadsheetApp.getUi().alert("✅ Multipack generation completed.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Error: " + e.message);
  }
}


function generatePackVariants_WallOrDoorSigns_UnbrandPrefix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Partial Flat File");
  const parentSku = source.getRange(4, 2).getValue();
  const newSheet = ss.insertSheet(`${parentSku} (Multipack)`);
  const lastCol = source.getLastColumn();
  const packs = [2, 5, 10];


  source.getRange(1, 1, 3, lastCol).copyTo(newSheet.getRange(1, 1));
  source.getRange(4, 1, 1, lastCol).copyTo(newSheet.getRange(4, 1));


  const b5 = `${parentSku}2`;
  newSheet.getRange(4, 2).setValue(b5);
  newSheet.getRange(4, 10).setValue(`${newSheet.getRange(4, 10).getValue()} (Multipack)`);
  newSheet.getRange(4, 39).setValue("");


  let count = 0;
  while (source.getRange(5 + count, 2).getValue()) count++;
  const childRange = source.getRange(5, 1, count, lastCol);
  childRange.copyTo(newSheet.getRange(5, 1));


  newSheet.getRange(5, 39, count).setValues(
    newSheet.getRange(5, 39, count).getValues().map(r => [`${r[0]} (1 Pack)`])
  );
  newSheet.getRange(5, 3, count).setValue("Partial Update");
  newSheet.getRange(5, 12, count).clearContent();
  newSheet.getRange(5, 13, count).setValues(source.getRange(5, 13, count).getValues());
  newSheet.getRange(5, 127, count).setValues(source.getRange(5, 127, count).getValues());


  newSheet.getRange(5, 27, count).clearDataValidations();
  newSheet.getRange(5, 27, count).setValues(Array(count).fill([b5]));


  let row = 5 + count;
  packs.forEach(p => {
    childRange.copyTo(newSheet.getRange(row, 1));
    newSheet.getRange(row, 2, count).setValues(
      newSheet.getRange(row, 2, count).getValues().map(r => [`${p}-${r[0]}`])
    );


    newSheet.getRange(row, 27, count).clearDataValidations();
    newSheet.getRange(row, 27, count).setValues(Array(count).fill([b5]));


    newSheet.getRange(row, 10, count).setValues(
      newSheet.getRange(row, 10, count).getValues().map(r => [`${r[0]} (${p} Pack)`])
    );
    newSheet.getRange(row, 39, count).setValues(
      newSheet.getRange(row, 39, count).getValues().map(r => {
        const base = r[0].replace(/\(\s*\d+\s*Pack\)/, "").trim();
        return [`${base} (${p} Pack)`];
      })
    );
    newSheet.getRange(row, 3, count).setValue("Update");
    newSheet.getRange(row, 5, count).clearContent();
    newSheet.getRange(row, 12, count).setValue(30);
    newSheet.getRange(row, 13, count).clearContent();
    newSheet.getRange(row, 128, count).clearContent();
    row += count;
  });


  injectPricesFromAttributes();
  injectDimensionsFromSizeColumn(newSheet, 5);
  newSheet.getRange(5, 129, row - 5, 1).setValue("Free Shipping Template");
}


function generatePackVariants_WallOrDoorSigns_BrandPrefix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Partial Flat File");
  const parentSku = source.getRange(4, 2).getValue();
  const newSheet = ss.insertSheet(`${parentSku} (Multipack)`);
  const lastCol = source.getLastColumn();
  const packs = [2, 5, 10];


  source.getRange(1, 1, 3, lastCol).copyTo(newSheet.getRange(1, 1));
  source.getRange(4, 1, 1, lastCol).copyTo(newSheet.getRange(4, 1));


  const b5 = `${parentSku}2`;
  newSheet.getRange(4, 2).setValue(b5);
  newSheet.getRange(4, 10).setValue(`${newSheet.getRange(4, 10).getValue()} (Multipack)`);
  newSheet.getRange(4, 39).setValue("");


  let count = 0;
  while (source.getRange(5 + count, 2).getValue()) count++;
  const childRange = source.getRange(5, 1, count, lastCol);
  childRange.copyTo(newSheet.getRange(5, 1));


  newSheet.getRange(5, 39, count).setValues(
    newSheet.getRange(5, 39, count).getValues().map(r => [`${r[0]} (1 Pack)`])
  );
  newSheet.getRange(5, 12, count).clearContent();


  newSheet.getRange(5, 27, count).clearDataValidations();
  newSheet.getRange(5, 27, count).setValues(Array(count).fill([b5]));


  let row = 5 + count;
  packs.forEach(p => {
    childRange.copyTo(newSheet.getRange(row, 1));
    newSheet.getRange(row, 2, count).setValues(
      newSheet.getRange(row, 2, count).getValues().map(r => {
        const sku = r[0];
        return [sku.slice(0, sku.indexOf("-")) + p + sku.slice(sku.indexOf("-"))];
      })
    );


    newSheet.getRange(row, 27, count).clearDataValidations();
    newSheet.getRange(row, 27, count).setValues(Array(count).fill([b5]));


    newSheet.getRange(row, 10, count).setValues(
      newSheet.getRange(row, 10, count).getValues().map(r => [`${r[0]} (${p} Pack)`])
    );
    newSheet.getRange(row, 39, count).setValues(
      newSheet.getRange(row, 39, count).getValues().map(r => {
        const base = r[0].replace(/\(\s*\d+\s*Pack\)/, "").trim();
        return [`${base} (${p} Pack)`];
      })
    );
    newSheet.getRange(row, 3, count).setValue("Update");
    newSheet.getRange(row, 5, count).clearContent();
    newSheet.getRange(row, 13, count).clearContent();
    newSheet.getRange(row, 128, count).clearContent();
    newSheet.getRange(row, 12, count).setValue(30);
    row += count;
  });


  injectPricesFromAttributes();
  injectDimensionsFromSizeColumn(newSheet, 5);
  newSheet.getRange(5, 129, row - 5, 1).setValue("Free Shipping Template");
}


function injectPricesFromAttributes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const attrSheet = ss.getSheetByName("Attributes");


  if (!attrSheet) {
    SpreadsheetApp.getUi().alert("❌ 'Attributes' sheet not found.");
    return;
  }


  const attrData = attrSheet.getRange(2, 1, attrSheet.getLastRow() - 1, 13).getValues();
  const data = sheet.getRange(5, 1, sheet.getLastRow() - 4, 139).getValues();


  let injectedCount = 0;


  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const title = row[9];
    const size = row[38];
    if (!title || !size) continue;


    const match = attrData.find(attr =>
      attr[0]?.toLowerCase() === "wall or door sign (lasered)" &&
      title.toLowerCase().includes(attr[1]?.toLowerCase()) &&
      size.toLowerCase().includes(attr[2]?.toLowerCase())
    );
    if (!match) continue;


    const packMatch = title.match(/\((\d+)\s*Pack\)/i);
    if (!packMatch) continue;


    const pack = parseInt(packMatch[1], 10);
    let price = "";
    if (pack === 1) price = match[7];
    else if (pack === 2) price = match[8];
    else if (pack === 5) price = match[9];
    else if (pack === 10) price = match[10];
    if (!price) continue;


    const rowNum = i + 5;


    sheet.getRange(rowNum, 13).setValue(price);
    sheet.getRange(5, 13).copyFormatToRange(sheet, 13, 13, rowNum, rowNum);
    sheet.getRange(rowNum, 127).setValue(price);
    sheet.getRange(5, 127).copyFormatToRange(sheet, 127, 127, rowNum, rowNum);
    sheet.getRange(rowNum, 12).setValue(30);
    sheet.getRange(5, 12).copyFormatToRange(sheet, 12, 12, rowNum, rowNum);
    sheet.getRange(rowNum, 129).setValue(match[11] || "Free Shipping Template");
    sheet.getRange(5, 129).copyFormatToRange(sheet, 129, 129, rowNum, rowNum);


    injectedCount++;
  }


  SpreadsheetApp.getUi().alert(`✅ Prices injected for ${injectedCount} multipack rows.`);
}


function injectDimensionsFromSizeColumn(sheet, startRow = 5) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attrSheet = ss.getSheetByName("Attributes");
  if (!attrSheet) {
    SpreadsheetApp.getUi().alert("❌ 'Attributes' sheet not found.");
    return;
  }


  const attrData = attrSheet.getRange(2, 1, attrSheet.getLastRow() - 1, 7).getValues();
  const data = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, 48).getValues();


  let injectedCount = 0;


  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const title = row[9];
    const size = row[38];


    if (!title || !size) continue;


    const match = attrData.find(attr =>
      attr[0]?.toLowerCase() === "wall or door sign (lasered)" &&
      title.toLowerCase().includes(attr[1]?.toLowerCase()) &&
      size.toLowerCase().includes(attr[2]?.toLowerCase())
    );
    if (!match) continue;


    const rowNum = i + startRow;
    sheet.getRange(rowNum, 45).setValue(match[3]).setFontFamily("Arial").setFontSize(11);
    sheet.getRange(rowNum, 46).setValue(match[4]).setFontFamily("Arial").setFontSize(11);
    sheet.getRange(rowNum, 47).setValue(match[5]).setFontFamily("Arial").setFontSize(11);
    sheet.getRange(rowNum, 48).setValue(match[6]).setFontFamily("Arial").setFontSize(11);


    injectedCount++;
  }


  SpreadsheetApp.getUi().alert(`✅ Dimensions injected for ${injectedCount} multipack rows.`);
}
