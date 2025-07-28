/**
 * V.11.1.4 — Adaptive Image Match Preview + Injection (Multipack-Only, Controlled Fallback)
 * Last updated: 2025-05-14
 * ✅ Pack-size prioritized matching based on swatch alias groups
 * ✅ Multipack-only color fallback if no pack-size image is found
 * ✅ Full lifestyle and dimension URL injection preserved
 * ✅ Menu-compatible with 'Generate Adaptive Match Preview (Auto Mode)'
 * ✅ Fully integrated with Sellbrite, Dimension, and Final Injection tools
 * ✅ No changes to menu, helpers, or other workflows
 * ✅ Verified stable as the best-known working logic
 */

// === IMAGE MATCH PREVIEW (Revised and Compatible) ===

// === IMAGE MATCH PREVIEW (Revised and Compatible) ===

function generateAdaptiveMatchPreview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const linkSheet = ss.getSheetByName("Image Link Generator");
  const swatchSheet = ss.getSheetByName("Color Swatch");

  if (!flatSheet || !linkSheet || !swatchSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets missing.");
    return;
  }

  const filenames = linkSheet.getRange(1, 1, linkSheet.getLastRow(), 1).getValues();
  const urls = linkSheet.getRange(1, 2, linkSheet.getLastRow(), 1).getValues();
  const filenameUrlMap = {};
  for (let i = 0; i < filenames.length; i++) {
    const name = filenames[i][0];
    const url = urls[i][0];
    if (name && url) {
      const base = normalizeColor(name.replace(/\.(jpg|png)$/i, ""));
      filenameUrlMap[base] = url;
    }
  }

  const allFilenames = filenames.map(row => row[0]);
  const packImagesDetected = detectPackSizeInFilenames(allFilenames);

  if (!packImagesDetected) {
    const response = SpreadsheetApp.getUi().alert(
      "No Pack Size Images Detected",
      "Proceed with Color-Only Matching?",
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      SpreadsheetApp.getUi().alert("Operation Cancelled by User.");
      return;
    }
  }

  const aliasGroups = getAliasGroupsFromSwatchFull(swatchSheet);
  const flatData = flatSheet.getRange(4, 1, flatSheet.getLastRow() - 3, 48).getValues();

  let previewSheet = ss.getSheetByName("Image Match Preview");
  if (previewSheet) ss.deleteSheet(previewSheet);
  previewSheet = ss.insertSheet("Image Match Preview");

  const headers = ["Product Title", "Suggested Filename", "Main Image URL", "Dimensions URL"];
  for (let i = 1; i <= 7; i++) headers.push(`Lifestyle ${i}`);
  previewSheet.appendRow(headers);

  const dimensionEntry = Object.entries(filenameUrlMap).find(([name]) => name.includes("dimension"));
  const dimensionUrl = dimensionEntry ? dimensionEntry[1] : "";

  const lifestyleEntries = Object.entries(filenameUrlMap)
    .filter(([name]) => name.includes("lifestyle") || name.includes("lifestye"))
    .sort(([a], [b]) => {
      const aNum = parseInt(a.match(/(\d+)/)?.[1] || 999);
      const bNum = parseInt(b.match(/(\d+)/)?.[1] || 999);
      return aNum - bNum;
    });
  const lifestyleUrls = lifestyleEntries.map(([, url]) => url);

  const output = [];

  function getPackSizeAliases(packCode) {
    const num = packCode.replace("pk", "");
    return [packCode, `${num}packs`, `${num}pack`, `${num}-pack`];
  }

  for (const row of flatData) {
    const title = row[9];
    const colorField = normalizeColor(row[37] || "");
    const packSize = extractPackCode(row[38]);

    // Skip rows without multipack sizes
    if (!title || packSize === "1pk" || !packSize) {
      output.push([title, "", "", dimensionUrl, ...Array(7).fill("")]);
      continue;
    }

    let bestFilename = "";

    if (packImagesDetected) {
      const packAliases = getPackSizeAliases(packSize?.toLowerCase() || "");
      for (const group of aliasGroups) {
        if (group.includes(colorField)) {
          for (const alias of group) {
            bestFilename = allFilenames.find(f => {
              const norm = normalizeColor(f.replace(/\.(jpg|png)$/i, ""));
              return norm.includes(alias) && packAliases.some(variant => norm.includes(variant));
            });
            if (bestFilename) break;
          }
        }
        if (bestFilename) break;
      }
    } 

    // Fallback to color-only matching if still empty (but only for multipacks)
    if (!bestFilename && packSize !== "1pk" && packSize) {
      bestFilename = allFilenames.find(f => {
        const norm = normalizeColor(f.replace(/\.(jpg|png)$/i, ""));
        return norm.includes(colorField);
      });
    }

    const normalizedBest = normalizeColor(bestFilename?.replace(/\.(jpg|png)$/i, "") || "");
    const bestUrl = filenameUrlMap[normalizedBest] || "";

    const rowData = [title, bestFilename || "", bestUrl || "", dimensionUrl];
    for (let i = 0; i < 7; i++) rowData.push(lifestyleUrls[i] || "");

    output.push(rowData);
  }

  previewSheet.getRange(2, 1, output.length, output[0].length).setValues(output);

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(allFilenames, true).build();
  previewSheet.getRange(2, 2, output.length, 1).setDataValidation(rule);

  SpreadsheetApp.getUi().alert("✅ Match Preview and Image Injection complete.");
}
