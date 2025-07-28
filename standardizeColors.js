/**
 * V.11.2 — Auto-Suggest Standardized Colors from Alias Groups
 * ✅ Replaces values in Column J (Color) with standardized name from Color Swatch
 * ✅ Uses first alias in matching swatch row as the canonical suggestion
 * ✅ Fully integrated into "Partial Flat File" workflow
 */

function suggestStandardizedColorNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const swatchSheet = ss.getSheetByName("Color Swatch");

  if (!flatSheet || !swatchSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets missing.");
    return;
  }

  const startRow = 4;
  const colColor = 37; // Column J (zero-indexed)
  const totalRows = flatSheet.getLastRow() - 3;
  const colorValues = flatSheet.getRange(startRow, colColor, totalRows).getValues();
  const aliasGroups = getAliasGroupsFromSwatchFull(swatchSheet);

  const suggested = [];

  for (let i = 0; i < colorValues.length; i++) {
    const original = colorValues[i][0];
    const normalized = normalizeColor(original);
    let matchedStandard = "";

    for (const group of aliasGroups) {
      if (group.includes(normalized)) {
        matchedStandard = group[0]; // Use first entry as canonical
        break;
      }
    }

    suggested.push([matchedStandard || original]);
  }

  flatSheet.getRange(startRow, colColor, totalRows, 1).setValues(suggested);

  SpreadsheetApp.getUi().alert(`✅ Color column standardized using swatch aliases.`);
}
