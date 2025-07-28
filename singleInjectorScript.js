/**
 * V.11.2.0 — Fully Unified Image Matching & Injection System
 * ✅ Unified Match Preview (1-pack + Multipack)
 * ✅ Color alias matching using Color Swatch
 * ✅ Injection into both Partial Flat File & Sellbrite CSV Export
 * ✅ Compatible with Autofill Menu and Dimension tools
 * ✅ Stable base version for "Sherryl" pipeline
 */

// === MENU ===

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Autofill")
    .addSubMenu(
      ui.createMenu("Generate Pack Size Variants")
        .addItem("Wall or Door Sign (lasered)", "runWallOrDoorSignsLasered")
    )
    .addSubMenu(
      ui.createMenu("Image Links and Match Preview")
        .addItem("Import Image Links (URL input)", "showUrlInputDialog")
        .addItem("Generate Unified Match Preview", "generateUnifiedImageMatchPreview")
        .addSeparator()
        .addItem("Inject Links into Partial Flat File", "injectFinalImageLinks")
        .addItem("Inject into Sellbrite CSV Export", "injectLinksIntoSellbriteCsvExport")
    )
    .addSubMenu(
      ui.createMenu("Populate Dimensions")
        .addItem("Wall or Door Sign", "populateDimensionsForWallOrDoorSign")
    )
    .addSubMenu(
      ui.createMenu("Standardize Data")
        .addItem("Suggest Standardized Colors (Auto)", "suggestStandardizedColorNames")
    )
    .addItem("Generate Sellbrite CSV from Partial Flat File", "generateSellbriteCsvFromPartialFlatFile")
    .addItem("🧹 Clear Partial Flat File", "clearPartialFlatFile")
    .addToUi();  // ✅ Last method in chain
}


// === IMAGE LINK FETCHING ===

function showUrlInputDialog() {
  const html = HtmlService.createHtmlOutputFromFile("UrlInputDialog")
    .setWidth(520)
    .setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, "📂 Enter Image Directory URL");
}

function fetchAndPopulateImageFilenames(url) {
  try {
    const html = UrlFetchApp.fetch(url).getContentText();
    const matches = [...html.matchAll(/href="([^"]+\.(jpg|png))"/gi)];
    const filenames = matches.map(m => decodeURIComponent(m[1].split("/").pop()));
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Image Link Generator");
    if (!sheet) throw new Error("Image Link Generator sheet not found");

    sheet.getRange(1, 1, sheet.getMaxRows(), 2).clearContent();
    if (filenames.length > 0) {
      sheet.getRange(1, 1, filenames.length, 1).setValues(filenames.map(name => [name]));
      const urls = filenames.map(name => [url.endsWith("/") ? url + name : url + "/" + name]);
      sheet.getRange(1, 2, urls.length, 1).setValues(urls);
    }

    return `✅ ${filenames.length} image filenames added to Image Link Generator.`;
  } catch (err) {
    return `❌ Error fetching images: ${err.message}`;
  }
}

// === UNIFIED IMAGE MATCH PREVIEW ===

function generateUnifiedImageMatchPreview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const linkSheet = ss.getSheetByName("Image Link Generator");
  const swatchSheet = ss.getSheetByName("Color Swatch");

  if (!flatSheet || !linkSheet || !swatchSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets missing.");
    return;
  }

  const filenames = linkSheet.getRange(1, 1, linkSheet.getLastRow(), 1).getValues().flat();
  const urls = linkSheet.getRange(1, 2, linkSheet.getLastRow(), 1).getValues().flat();
  const filenameUrlMap = {};
  filenames.forEach((name, i) => {
    if (name && urls[i]) {
      const norm = normalizeColor(name.replace(/\.(jpg|png)$/i, ""));
      filenameUrlMap[norm] = urls[i];
    }
  });

  const flatData = flatSheet.getRange(4, 1, flatSheet.getLastRow() - 3, 48).getValues();
  const aliasGroups = getAliasGroupsFromSwatchFull(swatchSheet);

  const dimensionEntry = Object.entries(filenameUrlMap).find(([name]) => name.includes("dimension"));
  const dimensionUrl = dimensionEntry ? dimensionEntry[1] : "";

  const lifestyleUrls = Object.entries(filenameUrlMap)
    .filter(([name]) => name.includes("lifestyle") || name.includes("lifestye"))
    .sort(([a], [b]) => parseInt(a.match(/\d+/)?.[0] || "999") - parseInt(b.match(/\d+/)?.[0] || "999"))
    .map(([, url]) => url);

  const output = [["Product Title", "Suggested Filename", "Main Image URL", "Dimensions URL", ...Array.from({ length: 7 }, (_, i) => `Lifestyle ${i + 1}`)]];

  const isOnePackTitle = (title) =>
    /\b(1[\s-]?pack|1pk|single)\b/i.test(title);

  const isMultipackFilename = (filename) =>
    /\b(2pk|3pk|4pk|5pk|10pk|2-pack|3-pack|5-pack|10-pack|packs?)\b/i.test(filename);

  const getPackSizeAliases = (packCode) => {
    const num = packCode.replace("pk", "");
    return [packCode, `${num}packs`, `${num}pack`, `${num}-pack`];
  };

  for (const row of flatData) {
    const title = row[9] || "";
    const color = normalizeColor(row[37] || "");
    const size = normalizeColor(row[38] || "");
    const packCode = extractPackCode(row[38])?.toLowerCase();
    const isOnePack = isOnePackTitle(title);
    let suggestedFilename = "";
    let matchUrl = "";

    const searchImage = (filterFn) => {
      for (const group of aliasGroups) {
        if (!group.includes(color)) continue;
        for (const alias of group) {
          const match = filenames.find(f =>
            normalizeColor(f).includes(alias) && filterFn(f)
          );
          if (match) {
            suggestedFilename = match;
            matchUrl = filenameUrlMap[normalizeColor(match.replace(/\.(jpg|png)$/i, ""))] || "";
            return true;
          }
        }
      }
      return false;
    };

    if (isOnePack) {
      searchImage(f => !isMultipackFilename(f));
    } else if (packCode) {
      const packAliases = getPackSizeAliases(packCode);
      const found = searchImage(f =>
        packAliases.some(p => normalizeColor(f).includes(p))
      );
      if (!found) {
        searchImage(f => normalizeColor(f).includes(color));
      }
    }

    output.push([
      title,
      suggestedFilename,
      matchUrl,
      dimensionUrl,
      ...lifestyleUrls.slice(0, 7)
    ]);
  }

  let previewSheet = ss.getSheetByName("Image Match Preview");
  if (previewSheet) ss.deleteSheet(previewSheet);
  previewSheet = ss.insertSheet("Image Match Preview");

  previewSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(filenames.filter(f => !!f), true)
    .build();
  previewSheet.getRange(2, 2, output.length - 1, 1).setDataValidation(rule);

  previewSheet.autoResizeColumns(1, output[0].length);

  SpreadsheetApp.getUi().alert("✅ Unified Image Match Preview generated.");
}

// === INJECTION INTO PARTIAL FLAT FILE ===

function injectFinalImageLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const previewSheet = ss.getSheetByName("Image Match Preview");
  const linkSheet = ss.getSheetByName("Image Link Generator");

  if (!flatSheet || !previewSheet || !linkSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets missing.");
    return;
  }

  const previewData = previewSheet.getRange(2, 1, previewSheet.getLastRow() - 1, previewSheet.getLastColumn()).getValues();
  const linkData = linkSheet.getRange(1, 1, linkSheet.getLastRow(), 2).getValues();

  const linkMap = {};
  for (const [filename, url] of linkData) {
    if (filename && url) {
      const normalized = filename.toLowerCase().trim();
      linkMap[normalized] = url;
    }
  }

  for (let i = 0; i < previewData.length; i++) {
    const rowOffset = i + 4;
    const suggestedFilename = previewData[i][1];
    const mainUrlDirect = previewData[i][2];
    const dimensionsUrl = previewData[i][3];
    const lifestyles = previewData[i].slice(4, 11);

    if (suggestedFilename) {
      const normalized = suggestedFilename.toLowerCase().trim();
      const resolvedUrl = linkMap[normalized] || mainUrlDirect || "";
      if (resolvedUrl) flatSheet.getRange(rowOffset, 14).setValue(resolvedUrl); // Column N
    }

    if (dimensionsUrl) flatSheet.getRange(rowOffset, 15).setValue(dimensionsUrl); // Column O

    for (let j = 0; j < lifestyles.length; j++) {
      if (lifestyles[j]) {
        flatSheet.getRange(rowOffset, 16 + j).setValue(lifestyles[j]); // P–V
      }
    }
  }

  SpreadsheetApp.getUi().alert("✅ Final Injection into Partial Flat File complete.");
}

// === INJECTION INTO SELLBRITE CSV EXPORT ===

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

  for (let i = 0; i < csvData.length; i++) {
    const csvRow = csvData[i];
    const title = csvRow[0]?.toString().trim();
    const matchRow = previewData.find(r => r[0]?.toString().trim() === title);
    if (!matchRow) continue;

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

    if (resolvedMainUrl) csvSheet.getRange(i + 2, 36).setValue(resolvedMainUrl);     // AJ
    if (dimensionUrl)    csvSheet.getRange(i + 2, 37).setValue(dimensionUrl);        // AK

    for (let j = 0; j < 7; j++) {
      if (lifestyleUrls[j]) {
        csvSheet.getRange(i + 2, 38 + j).setValue(lifestyleUrls[j]);                // AL–AR
      }
    }
  }

  SpreadsheetApp.getUi().alert("✅ Injection into Sellbrite CSV Export complete.");
}

// === HELPERS ===

function extractPackCode(rowSize, titleText) {
  // Priority 1: From size field if present
  let match = rowSize?.match(/\b(2|3|4|5|6|10)\s*(pk|pack)\b/i);
  if (match) return `${match[1]}pk`;

  // Priority 2: From product title
  match = titleText?.match(/\b(2|3|4|5|6|10)\s*(pk|pack)\b/i);
  return match ? `${match[1]}pk` : null;
}

function normalizeColor(colorStr) {
  if (!colorStr) return "";
  return String(colorStr)
    .toLowerCase()
    .replace(/[\s/_]+/g, "-")
    .replace(/[()]/g, "")
    .replace(/[^a-z0-9\-]/g, "")
    .trim();
}

function getAliasGroupsFromSwatchFull(swatchSheet) {
  const values = swatchSheet.getRange(1, 1, swatchSheet.getLastRow(), swatchSheet.getLastColumn()).getValues();
  return values.map(row => row.map(cell => normalizeColor(cell)).filter(Boolean));
}
function normalizeColor(str) {
  return (str || "")
    .toLowerCase()
    .replace(/[\s/_]+/g, "-")
    .replace(/[()]/g, "")
    .replace(/[^a-z0-9\-]/g, "")
    .trim();
}

function getPackSizeAliases(packCode) {
  const num = packCode.replace("pk", "");
  return [packCode, `${num}packs`, `${num}pack`, `${num}-pack`];
}

function resolveAliasGroup(colorField, swatchSheet) {
  const values = swatchSheet.getRange(1, 1, swatchSheet.getLastRow(), swatchSheet.getLastColumn()).getValues();
  const normalizedColor = normalizeColor(colorField);

  for (let i = 0; i < values.length; i++) {
    const row = values[i].filter(Boolean).map(cell => normalizeColor(cell));
    if (row.includes(normalizedColor)) {
      return row;
    }
  }

  return [normalizedColor];  // fallback as a group of 1
}