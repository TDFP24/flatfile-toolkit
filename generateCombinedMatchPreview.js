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

  // --- STRICT SWATCH MAPPING LOGIC ---
  const swatchValues = swatchSheet.getRange(1, 1, swatchSheet.getLastRow(), swatchSheet.getLastColumn()).getValues();
  const swatchRows = swatchValues.map(row => ({
    canonical: normalizeColor(row[0]),
    alternative: normalizeColor(row[1]),
    aliases: row.slice(2).map(cell => normalizeColor(cell)).filter(Boolean)
  }));

  function findSwatchRow(productColor) {
    const normColor = normalizeColor(productColor);
    return swatchRows.find(
      row => row.canonical === normColor || row.alternative === normColor
    );
  }

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
    // If the title contains any multipack identifier, it's NOT a 1-pack
    if (/\b(2|3|4|5|6|10)\s*(pk|pack)\b/i.test(title)) return false;
    // If it says 1 pack, 1pk, or single, it's a 1-pack
    if (/\(1\s*pack\)/i.test(title) || /\b1pk\b/i.test(title) || /\bsingle\b/i.test(title)) return true;
    // If it doesn't mention any pack size, treat as 1-pack
    return true;
  };

  const multipackPattern = /\b(2pk|3pk|4pk|5pk|10pk|2-pack|3-pack|5-pack|10-pack|packs?)\b/i;

  const getPackSizeAliases = (packCode) => {
    const num = packCode.replace("pk", "");
    return [packCode, `${num}packs`, `${num}pack`, `${num}-pack`];
  };

  // Helper to extract color from product title
  function extractColorFromTitle(title) {
    const match = title.match(/\(([^)]+)\)/);
    if (match) {
      return match[1].trim();
    }
    return "";
  }

  // Helper to expand single-word colors to their full combinations
  function expandSingleWordColor(color) {
    const singleWordExpansions = {
      'gold': 'gold-black',
      'red': 'red-white',
      'yellow': 'yellow-black',
      'silver': 'silver-black',
      'blue': 'blue-white',
      'black': 'black-white',
      'white': 'white-black',
      // Add other single-word expansions as needed
    };
    return singleWordExpansions[color.toLowerCase()] || color;
  }

  // Helper to match aliases in filenames with flexible formatting but strict content/order
  function matchesAliasInFilename(alias, filename) {
    const normAlias = normalizeColor(alias);
    const normFilename = normalizeColor(filename);
    
    // Check if the normalized alias appears in the normalized filename
    // But prioritize exact matches over substring matches
    if (normFilename === normAlias) return true;
    if (normFilename.includes(normAlias)) {
      // For single-word aliases, ensure they're not part of a larger combination
      if (normAlias.split('-').length === 1) {
        // Single word like 'black', 'gold', 'red' - check for word boundaries
        const wordBoundaryPattern = new RegExp(`(^|[-_])${normAlias}([-_]|$)`);
        const hasWordBoundary = wordBoundaryPattern.test(normFilename);
        
        // Additional check: if this is a single-word alias like 'black', 
        // make sure it's not part of a combination like 'black-gold'
        if (hasWordBoundary && normAlias === 'black') {
          // For 'black', only match if it's 'black' or 'black-white', not 'black-gold' or 'black-silver'
          if (normFilename.includes('black-gold') || normFilename.includes('black-silver')) return false;
        }
        if (hasWordBoundary && normAlias === 'gold') {
          // For 'gold', only match if it's 'gold' or 'gold-black', not 'black-gold'
          if (normFilename.includes('black-gold')) return false;
        }
        if (hasWordBoundary && normAlias === 'silver') {
          // For 'silver', only match if it's 'silver' or 'silver-black', not 'black-silver'
          if (normFilename.includes('black-silver')) return false;
        }
        
        return hasWordBoundary;
      }
      return true;
    }
    return false;
  }

  for (const row of flatData) {
    const title = row[9] || "";
    const colorField = extractColorFromTitle(title) || row[37] || "";
    const expandedColor = expandSingleWordColor(colorField);
    const color = normalizeColor(expandedColor);
    const size = normalizeColor(row[38] || "");
    const packCode = extractPackCode(row[38], row[9])?.toLowerCase();
    const isOnePack = isOnePackTitle(title);
    let suggestedFilename = "";
    let matchUrl = "";
    let debugSwatchMatch = "";

    // --- STRICT COLOR/ALIAS MATCHING ---
    const swatchRow = swatchRows.find(
      row => row.canonical === color || row.alternative === color
    );
    if (swatchRow) {
      debugSwatchMatch = swatchRow.canonical + (swatchRow.alternative ? ` / ${swatchRow.alternative}` : "");
      if (isOnePack) {
        // 1-pack: match color/alias, but NOT any multipack identifier
        let matchedAlias = "";
        const matchObj = normalizedFilenameMap.find(obj => {
          // First try exact matches, but with filtering
          const exactAlias = swatchRow.aliases.find(alias => {
            const normAlias = normalizeColor(alias);
            const normObj = obj.normalized;
            
            // Skip problematic single-word aliases that cause cross-matches
            if (normAlias === 'black' && normObj.includes('black-gold')) return false;
            if (normAlias === 'black' && normObj.includes('black-silver')) return false;
            if (normAlias === 'gold' && normObj.includes('black-gold')) return false;
            if (normAlias === 'silver' && normObj.includes('black-silver')) return false;
            
            return normObj === normAlias || normObj.includes(normAlias);
          });
          if (exactAlias && !multipackPattern.test(obj.normalized)) {
            matchedAlias = exactAlias;
            return true;
          }
          
          // Then try partial matches, but filter out problematic partial aliases
          const foundAlias = swatchRow.aliases.find(alias => {
            // Skip aliases that start with dash (like "-gold") as they cause false matches
            if (alias.startsWith('-') || alias.endsWith('-')) return false;
            
            // Skip problematic aliases that cause cross-matches
            const normAlias = normalizeColor(alias);
            const normObj = obj.normalized;
            
            // Skip problematic single-word aliases that cause cross-matches
            if (normAlias === 'black' && normObj.includes('black-gold')) return false;
            if (normAlias === 'black' && normObj.includes('black-silver')) return false;
            if (normAlias === 'gold' && normObj.includes('black-gold')) return false;
            if (normAlias === 'silver' && normObj.includes('black-silver')) return false;
            
            // Check if this alias would match the wrong color combination
            // For example, if we're in the "Brush Gold" row, we don't want "gold-black" 
            // to match "black-gold" filenames
            if (normAlias.includes('-')) {
              const parts = normAlias.split('-');
              // If this is a two-part color like "gold-black", check if it's in the wrong order
              // for the current swatch row
              if (parts.length === 2) {
                const [first, second] = parts;
                // If we're in a "Gold" row but the alias is "gold-black", 
                // and we're looking for "black-gold" filenames, skip it
                if ((first === 'gold' && second === 'black') || 
                    (first === 'black' && second === 'gold')) {
                  // This is a problematic cross-match - skip it
                  return false;
                }
              }
            }
            
            return matchesAliasInFilename(alias, obj.original);
          });
          if (foundAlias && !multipackPattern.test(obj.normalized)) {
            matchedAlias = foundAlias;
            return true;
          }
          return false;
        });
        if (matchObj) {
          suggestedFilename = matchObj.original;
          matchUrl = matchObj.url || "";
          debugSwatchMatch += ` | Matched: ${matchedAlias}`;
        }
      } else if (packCode) {
        // Multipack: match color/alias AND pack size
        const packAliases = getPackSizeAliases(packCode);
        let matchedAlias = "";
        const matchObj = normalizedFilenameMap.find(obj => {
          // First try exact matches, but with filtering
          const exactAlias = swatchRow.aliases.find(alias => {
            const normAlias = normalizeColor(alias);
            const normObj = obj.normalized;
            
            // Skip problematic single-word aliases that cause cross-matches
            if (normAlias === 'black' && normObj.includes('black-gold')) return false;
            if (normAlias === 'black' && normObj.includes('black-silver')) return false;
            if (normAlias === 'gold' && normObj.includes('black-gold')) return false;
            if (normAlias === 'silver' && normObj.includes('black-silver')) return false;
            
            return (normObj === normAlias || normObj.includes(normAlias)) &&
                   packAliases.some(p => normObj.includes(p));
          });
          if (exactAlias) {
            matchedAlias = exactAlias;
            return true;
          }
          
          // Then try partial matches, but filter out problematic partial aliases
          const foundAlias = swatchRow.aliases.find(alias => {
            // Skip aliases that start with dash (like "-gold") as they cause false matches
            if (alias.startsWith('-') || alias.endsWith('-')) return false;
            
            // Skip problematic aliases that cause cross-matches
            const normAlias = normalizeColor(alias);
            const normObj = obj.normalized;
            
            // Skip problematic single-word aliases that cause cross-matches
            if (normAlias === 'black' && normObj.includes('black-gold')) return false;
            if (normAlias === 'black' && normObj.includes('black-silver')) return false;
            if (normAlias === 'gold' && normObj.includes('black-gold')) return false;
            if (normAlias === 'silver' && normObj.includes('black-silver')) return false;
            
            // Check if this alias would match the wrong color combination
            // For example, if we're in the "Brush Gold" row, we don't want "gold-black" 
            // to match "black-gold" filenames
            if (normAlias.includes('-')) {
              const parts = normAlias.split('-');
              // If this is a two-part color like "gold-black", check if it's in the wrong order
              // for the current swatch row
              if (parts.length === 2) {
                const [first, second] = parts;
                // If we're in a "Gold" row but the alias is "gold-black", 
                // and we're looking for "black-gold" filenames, skip it
                if ((first === 'gold' && second === 'black') || 
                    (first === 'black' && second === 'gold')) {
                  // This is a problematic cross-match - skip it
                  return false;
                }
              }
            }
            
            return matchesAliasInFilename(alias, obj.original) &&
                   packAliases.some(p => obj.normalized.includes(p));
          });
          if (foundAlias) {
            matchedAlias = foundAlias;
            return true;
          }
          return false;
        });
        if (matchObj) {
          suggestedFilename = matchObj.original;
          matchUrl = matchObj.url || "";
          debugSwatchMatch += ` | Matched: ${matchedAlias}`;
        } else {
          // Fallback: color-only match within this swatch row
          const fallbackObj = normalizedFilenameMap.find(obj => {
            // First try exact matches, but with filtering
            const exactAlias = swatchRow.aliases.find(alias => {
              const normAlias = normalizeColor(alias);
              const normObj = obj.normalized;
              
              // Skip problematic single-word aliases that cause cross-matches
              if (normAlias === 'black' && normObj.includes('black-gold')) return false;
              if (normAlias === 'black' && normObj.includes('black-silver')) return false;
              if (normAlias === 'gold' && normObj.includes('black-gold')) return false;
              if (normAlias === 'silver' && normObj.includes('black-silver')) return false;
              
              return normObj === normAlias || normObj.includes(normAlias);
            });
            if (exactAlias) {
              matchedAlias = exactAlias;
              return true;
            }
            
            // Then try partial matches, but filter out problematic partial aliases
            const foundAlias = swatchRow.aliases.find(alias => {
              // Skip aliases that start with dash (like "-gold") as they cause false matches
              if (alias.startsWith('-') || alias.endsWith('-')) return false;
              
              // Skip problematic aliases that cause cross-matches
              const normAlias = normalizeColor(alias);
              
              // Check if this alias would match the wrong color combination
              // For example, if we're in the "Brush Gold" row, we don't want "gold-black" 
              // to match "black-gold" filenames
              if (normAlias.includes('-')) {
                const parts = normAlias.split('-');
                // If this is a two-part color like "gold-black", check if it's in the wrong order
                // for the current swatch row
                if (parts.length === 2) {
                  const [first, second] = parts;
                  // If we're in a "Gold" row but the alias is "gold-black", 
                  // and we're looking for "black-gold" filenames, skip it
                  if ((first === 'gold' && second === 'black') || 
                      (first === 'black' && second === 'gold')) {
                    // This is a problematic cross-match - skip it
                    return false;
                  }
                }
              }
              
              return matchesAliasInFilename(alias, obj.original);
            });
            if (foundAlias) {
              matchedAlias = foundAlias;
              return true;
            }
            return false;
          });
          if (fallbackObj) {
            suggestedFilename = fallbackObj.original;
            matchUrl = fallbackObj.url || "";
            debugSwatchMatch += ` | Fallback: ${matchedAlias}`;
          }
        }
      }
    }

    output.push([
      title,
      suggestedFilename,
      matchUrl,
      dimensionUrl,
      ...Array.from({ length: 7 }, (_, i) => lifestyleUrls[i] || ""),
      debugSwatchMatch,
      `${colorField} → ${expandedColor} → ${color}`
    ]);
  }

  // Add debug column header
  output[0].push("Swatch Row Matched");
  output[0].push("Color Processing");

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
