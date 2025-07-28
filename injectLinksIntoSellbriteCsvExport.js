function injectLinksIntoSellbriteCsvExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const csvSheet = ss.getSheetByName("Sellbrite CSV Export");
  const previewSheet = ss.getSheetByName("Image Match Preview");
  const linkSheet = ss.getSheetByName("Image Link Generator");

  if (!csvSheet || !previewSheet || !linkSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets not found.");
    return;
  }

  const previewData = previewSheet.getRange(2, 1, previewSheet.getLastRow() - 1, previewSheet.getLastColumn()).getValues();
  const csvData = csvSheet.getRange(2, 1, csvSheet.getLastRow() - 1, 50).getValues();
  const linkData = linkSheet.getRange(1, 1, linkSheet.getLastRow(), 2).getValues();

  const linkMap = {};
  for (const [filename, url] of linkData) {
    if (filename && url) {
      const normalized = filename.toLowerCase().trim();
      linkMap[normalized] = url;
    }
  }

  let unmatchedCount = 0;

  for (let i = 0; i < csvData.length; i++) {
    const csvRow = csvData[i];
    const title = normalizeText(csvRow[2]); // MATCH using Column C (index 2)
    const matchRow = previewData.find(r => normalizeText(r[0]) === title); // Compare to Match Preview Col A

    if (!matchRow) {
      unmatchedCount++;
      continue;
    }

    const suggestedFilename = matchRow[1];
    const fallbackMainUrl = matchRow[2];
    const dimensionUrl = matchRow[3];
    const lifestyleUrls = matchRow.slice(4, 11);

    const resolvedMainUrl = (() => {
      const normalized = suggestedFilename?.toLowerCase().trim();
      if (normalized && linkMap[normalized]) return linkMap[normalized];
      if (fallbackMainUrl) return fallbackMainUrl;
      return "";
    })();

    // Inject to Sellbrite CSV Export:
    if (resolvedMainUrl) csvSheet.getRange(i + 2, 36).setValue(resolvedMainUrl);  // AJ
    if (dimensionUrl)    csvSheet.getRange(i + 2, 37).setValue(dimensionUrl);     // AK

    for (let j = 0; j < 7; j++) {
      if (lifestyleUrls[j]) {
        csvSheet.getRange(i + 2, 38 + j).setValue(lifestyleUrls[j]);              // AL–AR
      }
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (unmatchedCount > 0) {
    ui.alert(`⚠️ Injection complete. ${unmatchedCount} rows had no preview match.`);
  } else {
    ui.alert("✅ Injection into Sellbrite CSV Export complete.");
  }
}

function normalizeText(val) {
  return (val || "").toString().toLowerCase().trim().replace(/\s+/g, " ");
}
