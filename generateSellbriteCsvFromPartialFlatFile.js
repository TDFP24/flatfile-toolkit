function generateSellbriteCsvFromPartialFlatFile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const flatSheet = ss.getSheetByName("Partial Flat File");
  const attrSheet = ss.getSheetByName("Attributes");


  if (!flatSheet || !attrSheet) {
    SpreadsheetApp.getUi().alert("❌ Required sheets not found.");
    return;
  }


  const technicalHeaders = [ /* same header list as before, unchanged */
    "sku", "parent_sku", "name", "description", "brand", "condition", "condition_note", "price", "notes",
    "msrp", "category_name", "store_product_url", "manufacturer", "manufacturer_model_number",
    "upc", "ean", "isbn", "gtin", "gcid", "asin", "epid",
    "package_height", "package_length", "package_width", "package_weight",
    "feature_1", "feature_2", "feature_3", "feature_4", "feature_5",
    "variation_1", "variation_2", "variation_3", "variation_4", "variation_5",
    "product_image_1", "product_image_2", "product_image_3", "product_image_4",
    "product_image_5", "product_image_6", "product_image_7", "product_image_8",
    "product_image_9", "product_image_10", "product_image_11", "product_image_12",
    "delete"
  ];


  const columnMap = {
    sku: 2, parent_sku: 27, name: 10, description: 7, brand: 4,
    condition: 42, condition_note: 43, price: 13, manufacturer: 9,
    feature_1: 32, feature_2: 33, feature_3: 34, feature_4: 35, feature_5: 36,
    package_length: 45, package_width: 47,
    product_image_1: 14, product_image_2: 15, product_image_3: 16, product_image_4: 17,
    product_image_5: 18, product_image_6: 19, product_image_7: 20, product_image_8: 21
  };


  let exportSheet = ss.getSheetByName("Sellbrite CSV Export");
  if (exportSheet) ss.deleteSheet(exportSheet);
  exportSheet = ss.insertSheet("Sellbrite CSV Export");


  // Headers
  exportSheet.getRange(1, 1).setValue("SELLBRITE PRODUCT CSV TEMPLATE (Do NOT remove the first 3 rows).");
  exportSheet.getRange(2, 1, 1, technicalHeaders.length).setValues([technicalHeaders]);
  exportSheet.getRange(3, 1, 1, technicalHeaders.length).setValues([technicalHeaders]);


  const startRow = 4;
  const totalRows = flatSheet.getLastRow() - 3;
  const flatData = flatSheet.getRange(startRow, 1, totalRows, flatSheet.getLastColumn()).getValues();
  const attrData = attrSheet.getDataRange().getValues();


  const attrHeaders = attrData[0].map(h => h.toString().toLowerCase().trim());
  const exportRows = [];


  for (let i = 0; i < flatData.length; i++) {
    const row = flatData[i];
    const sellbriteRow = Array(technicalHeaders.length).fill("");


    // Basic mapping
    for (const [field, colIndex] of Object.entries(columnMap)) {
      const targetIndex = technicalHeaders.indexOf(field);
      if (targetIndex !== -1) sellbriteRow[targetIndex] = row[colIndex - 1] || "";
    }


    // SKU duplication
    sellbriteRow[technicalHeaders.indexOf("manufacturer_model_number")] = row[1] || "";


    // Static fields
    sellbriteRow[technicalHeaders.indexOf("condition")] = "new";


    // Variation logic
    sellbriteRow[30] = row[38] || "";  // Size
    sellbriteRow[31] = row[37] || "";  // Color


    // ID logic
    const id = row[4];
    const idType = (row[5] || "").toLowerCase();
    if (idType === "upc") sellbriteRow[technicalHeaders.indexOf("upc")] = id;
    else if (idType === "gtin") sellbriteRow[technicalHeaders.indexOf("gtin")] = id;


    // -------- PACKAGE ATTRIBUTES --------
    const title = (row[9] || "").toLowerCase().replace(/[^\w\s]/g, "").trim();
    const shape = attrData.map(r => r[1].toString().toLowerCase().trim());
    const sizeCol = attrData.map(r => r[2].toString().toLowerCase().replace(/[^\w\s]/g, "").trim());


    const matchedShape = shape.findIndex((s, idx) => title.includes(s) && sizeCol[idx] && title.includes(sizeCol[idx]));


    if (matchedShape !== -1) {
      const packMatch = title.match(/(\d+)\s*pack/);
      const packSize = packMatch ? packMatch[1] : "1";


      const headerSuffix = `(${packSize} pack)`;
      const dims = ["package_height", "package_length", "package_width", "package_weight"];


      dims.forEach((dim, j) => {
        const headerIndex = attrHeaders.findIndex(h => h.includes(dim) && h.includes(headerSuffix));
        const destIndex = technicalHeaders.indexOf(dim);
        if (headerIndex !== -1 && destIndex !== -1) {
          sellbriteRow[destIndex] = attrData[matchedShape][headerIndex] || "";
        }
      });
    }


    exportRows.push(sellbriteRow);
  }


  exportSheet.getRange(4, 1, exportRows.length, technicalHeaders.length).setValues(exportRows);


  // Write variation headers
  exportSheet.getRange(4, 31).setValue("Size");
  exportSheet.getRange(4, 32).setValue("Color");


  SpreadsheetApp.getUi().alert(`✅ Sellbrite export created with ${exportRows.length} rows.`);
}
