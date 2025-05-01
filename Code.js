function batchGenerateAndArchive() {
  Logger.log("🚀 Starting batch archive load...");

  loadOntracToArchive(); // updated name

  Logger.log("✅ Batch archive load complete.");
}


function generateBatchID() {
  const now = new Date();
  const formatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  return formatted;
}

function testBatchID() {
  const batchID = generateBatchID();
  const msg = `🧪 Batch ID: ${batchID}`;

  Logger.log(msg);
}

function listOntracTabs() {
  const sourceSheetId = "1hLx5vKUsnKvCaa5Xu-UgGWG_2zzX0V6em1jmumnx7Ko";
  const ss = SpreadsheetApp.openById(sourceSheetId);
  const allSheets = ss.getSheets();
  
  const ontracTabs = allSheets
    .map(sheet => sheet.getName())
    .filter(name => name.startsWith("Ontrac_"));

  Logger.log(`📋 Found ${ontracTabs.length} Ontrac tab(s):`);

  ontracTabs.forEach(name => {
    const { source, region } = extractMetadata(name);
    Logger.log(`- ${name} → Source: ${source}, Region: ${region}`);
  });
}

function extractMetadata(sheetName) {
  const parts = sheetName.split('_');
  const source = parts[0] || "UnknownSource";
  const region = parts[1] || "UnknownRegion";
  return { source, region };
}

function testBatchMetadata() {
  const sheetName = "Ontrac_Keystone"; // Simulated tab
  const { source, region } = extractMetadata(sheetName);
  const batchID = `${generateBatchID()}_${region}_${source}`;
  const msg = `🧪 Batch ID with Metadata: ${batchID}`;

  Logger.log(msg);
}

function previewOntracDataRows() {
  const sourceSheetId = "1hLx5vKUsnKvCaa5Xu-UgGWG_2zzX0V6em1jmumnx7Ko";
  const ss = SpreadsheetApp.openById(sourceSheetId);
  const allSheets = ss.getSheets();

  const ontracTabs = allSheets
    .map(sheet => sheet.getName())
    .filter(name => name.startsWith("Ontrac_"));

  Logger.log(`🔍 Previewing Ontrac row counts:`);

  ontracTabs.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const { region } = extractMetadata(name);

    const totalRows = sheet.getLastRow();
    const dataStartRow = 3;
    const dataRows = Math.max(0, totalRows - (dataStartRow - 1));

    Logger.log(`📄 ${region} → ${dataRows} data row(s)`);
  });
}

function auditOntracHeaders() {
  const sourceSheetId = "1hLx5vKUsnKvCaa5Xu-UgGWG_2zzX0V6em1jmumnx7Ko";
  const ss = SpreadsheetApp.openById(sourceSheetId);
  const ontracTabs = ss.getSheets()
    .map(s => s.getName())
    .filter(name => name.startsWith("Ontrac_"));

  ontracTabs.forEach(name => {
    const sheet = ss.getSheetByName(name);
    const { region } = extractMetadata(name);
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

    Logger.log(`🔎 Headers for ${region}:`);
    headers.forEach((header, index) => {
      Logger.log(`  Col ${index + 1}: ${header}`);
    });
  });
}

function loadOntracToArchive() {
  const batchID = generateBatchID();
  const sourceSheetId = "1hLx5vKUsnKvCaa5Xu-UgGWG_2zzX0V6em1jmumnx7Ko";
  const targetSheetId = "19zuCKxg_4Akn9CpubsAKk1ejdXqSo0id0K9s9nxAStw";

  const sourceSS = SpreadsheetApp.openById(sourceSheetId);
  const targetSS = SpreadsheetApp.openById(targetSheetId);
  const archiveSheet = targetSS.getSheetByName("Archive_Test");

  if (!archiveSheet) {
    throw new Error("❌ 'Archive_Test' sheet not found in target spreadsheet.");
  }

  const ontracTabs = sourceSS.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => name.startsWith("Ontrac_"));

  let totalAdded = 0;
  let allRows = [];

  ontracTabs.forEach(tabName => {
    const sheet = sourceSS.getSheetByName(tabName);
    const { source, region } = extractMetadata(tabName);
    const lastRow = sheet.getLastRow();
    const numRows = Math.max(0, lastRow - 2); // Skip 2 header rows

    if (numRows === 0) return;

    const techIDs = sheet.getRange(3, 1, numRows, 1).getValues(); // Col A
    const data = sheet.getRange(3, 2, numRows, 33).getValues();   // Col B to AH

    for (let i = 0; i < numRows; i++) {
      const techID = techIDs[i][0];
      if (!techID) continue;

      const uniqueKey = `${batchID}_${region}_${techID}`;
      const row = [batchID, region, source, techID, uniqueKey, ...data[i]];
      allRows.push(row);
    }

    Logger.log(`✅ Prepared ${allRows.length} rows from ${region}`);
  });

  if (allRows.length > 0) {
    const startRow = archiveSheet.getLastRow() + 1;
    archiveSheet.insertRowsAfter(startRow, allRows.length);
    archiveSheet.getRange(startRow + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
    Logger.log(`🧪 Archive Load Complete: ${allRows.length} total rows written.`);
  } else {
    Logger.log("⚠️ No rows to write.");
  }
}
