/**
 * V.11.2.1 — Fully Unified Image Matching & Injection System (Optimized)
 * ✅ Unified Match Preview (1-pack + Multipack)
 * ✅ Color alias matching using Color Swatch
 * ✅ Injection into both Partial Flat File & Sellbrite CSV Export
 * ✅ Compatible with Autofill Menu and Dimension tools
 * ✅ Stable base version for "Sherryl" pipeline
 * ✅ Improved 1-pack detection with '-main.jpg' fallback
 * ✅ Batches dropdowns for all rows without slowdown
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
    .addItem("Generate Sellbrite CSV from Partial Flat File", "generateSellbriteCsvFromPartialFlatFile")
    .addItem("🧹 Clear Partial Flat File", "clearPartialFlatFile")
    .addToUi();
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

  const normalizedFilenameMap = filenames.map((filename, index) => ({
    original: filename,
    normalized: normalizeColor(filename.replace(/\.(jpg|png)$/i, "")),
    url: urls[index]
  }));

  const filenameUrlMap = Object.fromEntries(
    normalizedFilenameMap.map(obj => [obj.normalized, obj.url])
  );

  const flatData = flatSheet.getRange(4, 1, flatSheet.getLastRow() - 3, 48).getValues();
  const aliasGroups = getAliasGroupsFromSwatchFull(swatchSheet);

  const dimensionEntry = normalizedFilenameMap.find(obj => obj.normalized.includes("dimension"));
  const dimensionUrl = dimensionEntry ? dimensionEntry.url : "";

  const lifestyleUrls = normalizedFilenameMap
    .filter(obj => obj.normalized.includes("lifestyle") || obj.normalized.includes("lifestye"))
    .sort((a, b) =>
      parseInt(a.normalized.match(/\d+/)?.[0] || "999") -
      parseInt(b.normalized.match(/\d+/)?.[0] || "999")
    )
    .map(obj => obj.url);

  const output = [["Product Title", "Suggested Filename", "Main Image URL", "Dimensions URL", ...Array.from({ length: 7 }, (_, i) => `Lifestyle ${i + 1}`)]];

  const isOnePackTitle = (title) => {
  return /\(1\s*pack\)/i.test(title) || /\b1pk\b/i.test(title) || /\bsingle\b/i.test(title);
  };

  const isMultipackFilename = (filename) => {
    const norm = normalizeColor(filename);
    return /\b(2pk|3pk|4pk|5pk|10pk|2-pack|3-pack|5-pack|10-pack|packs?)\b/.test(norm) || !norm.endsWith("-main");
  };

  const getPackSizeAliases = (packCode) => {
    const num = packCode.replace("pk", "");
    return [packCode, `${num}packs`, `${num}pack`, `${num}-pack`];
  };

  for (const row of flatData) {
    const title = row[9] || "";
    const color = normalizeColor(row[37] || "");
    const size = normalizeColor(row[38] || "");
    const packCode = extractPackCode(row[38], row[9])?.toLowerCase();
    const isOnePack = isOnePackTitle(title);
    let suggestedFilename = "";
    let matchUrl = "";

    const searchImage = (filterFn) => {
      for (const group of aliasGroups) {
        if (!group.includes(color)) continue;
        for (const alias of group) {
          const matchObj = normalizedFilenameMap.find(obj =>
            obj.normalized.includes(alias) && filterFn(obj.original)
          );
          if (matchObj) {
            suggestedFilename = matchObj.original;
            matchUrl = matchObj.url || "";
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
      ...Array.from({ length: 7 }, (_, i) => lifestyleUrls[i] || "")
    ]);
  }

  let previewSheet = ss.getSheetByName("Image Match Preview");
  if (previewSheet) ss.deleteSheet(previewSheet);
  previewSheet = ss.insertSheet("Image Match Preview");

  previewSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(filenames.filter(f => !!f), true)
    .build();

  // ✅ Apply dropdown to every row in column B (Suggested Filename)
  previewSheet.getRange(2, 2, output.length - 1, 1).setDataValidation(rule);

  previewSheet.autoResizeColumns(1, output[0].length);

  SpreadsheetApp.getUi().alert("✅ Unified Image Match Preview generated.");
}

// === HELPERS ===

function extractPackCode(rowSize, titleText) {
  let match = rowSize?.match(/\b(2|3|4|5|6|10)\s*(pk|pack)\b/i);
  if (match) return `${match[1]}pk`;
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
