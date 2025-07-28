function generateOnePackMatchPreview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Sellbrite CSV Export");
  const linkSheet = ss.getSheetByName("Image Link Generator");

  if (!flatSheet || !linkSheet) {
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

  const data = flatSheet.getRange(4, 1, flatSheet.getLastRow() - 3, 48).getValues();
  const preview = [];

  const headers = ["Product Title", "Suggested Filename", "Main Image URL", "Dimensions URL"];
  for (let i = 1; i <= 7; i++) headers.push(`Lifestyle ${i}`);
  preview.push(headers);

  const dimensionEntry = Object.entries(filenameUrlMap).find(([name]) => name.includes("dimension"));
  const dimensionUrl = dimensionEntry ? dimensionEntry[1] : "";

  const lifestyleUrls = Object.entries(filenameUrlMap)
    .filter(([name]) => name.includes("lifestyle") || name.includes("lifestye"))
    .sort(([a], [b]) => {
      const aNum = parseInt(a.match(/(\d+)/)?.[1] || "999");
      const bNum = parseInt(b.match(/(\d+)/)?.[1] || "999");
      return aNum - bNum;
    })
    .map(([, url]) => url);

  const onePackAliases = ["1pk", "1-pack", "1 pack", "single"];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const title = row[2]?.toLowerCase() || ""; // Column C — 'name'

    if (!onePackAliases.some(alias => title.includes(alias))) continue;

    const color = normalizeColor(row[31] || ""); // Column AE — Color

    let bestFilename = "";
    for (const [fname, url] of Object.entries(filenameUrlMap)) {
      if (
        fname.includes(color) &&
        !fname.match(/\b(2pk|3pk|4pk|5pk|10pk|2-pack|3-pack|5-pack|10-pack|packs?)\b/i)
      ) {
        bestFilename = fname;
        break;
      }
    }

    const normalizedBest = normalizeColor(bestFilename?.replace(/\.(jpg|png)$/i, "") || "");
    const bestUrl = filenameUrlMap[normalizedBest] || "";

    const lifestyleSlice = [...lifestyleUrls.slice(0, 7)];
    while (lifestyleSlice.length < 7) lifestyleSlice.push("");

    preview.push([
      row[2], // Product Title
      bestFilename || "",
      bestUrl || "",
      dimensionUrl,
      ...lifestyleSlice
    ]);
  }

  let previewSheet = ss.getSheetByName("1-Pack Match Preview");
  if (previewSheet) ss.deleteSheet(previewSheet);
  previewSheet = ss.insertSheet("1-Pack Match Preview");

  if (preview.length > 1) {
    previewSheet.getRange(1, 1, preview.length, headers.length).setValues(preview);

    const multipackPattern = /\b(2pk|3pk|4pk|5pk|10pk|2-pack|3-pack|5-pack|10-pack|2packs|3packs|5packs|10packs)\b/i;
const validList = filenames
  .map(([f]) => f)
  .filter(name => name && !multipackPattern.test(name));

    const rule = SpreadsheetApp.newDataValidation().requireValueInList(validList, true).build();
    previewSheet.getRange(2, 2, preview.length - 1, 1).setDataValidation(rule);

    SpreadsheetApp.getUi().alert("✅ 1-Pack Match Preview generated.");
  } else {
    SpreadsheetApp.getUi().alert("⚠️ No 1-Pack rows found to preview.");
  }
}
