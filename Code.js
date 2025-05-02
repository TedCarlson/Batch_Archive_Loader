function batchGenerateAndArchive() {
  Logger.log("🚀 Starting batch archive load...");

  const batchID = generateBatchID();
  const defaultFM = getFiscalMonth(new Date());
  const fiscalMonth = promptForFiscalMonth(defaultFM);

  loadOntracToArchive(batchID, fiscalMonth);
  deduplicateArchive();

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

function getFiscalMonth(date) {
  const d = new Date(date);
  if (isNaN(d.getTime())) return "Invalid Date";


  const year = d.getFullYear();
  let month = d.getMonth();
  const day = d.getDate();


  if (day >= 22) month += 1;
  if (month > 11) {
    month = 0;
    return `${year + 1}-01`;  // Use backticks here
  }


  const monthStr = ('0' + (month + 1)).slice(-2);
  return `${year}-${monthStr}`;  // And here
  
}

function testFiscalMonth() {
  const now = new Date();
  const fiscalMonth = getFiscalMonth(now);
  Logger.log(`🧪 Current date: ${now.toDateString()}`);
  Logger.log(`🗓 Fiscal month: ${fiscalMonth}`);
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

function deduplicateArchive() {
  const targetSheetId = "19zuCKxg_4Akn9CpubsAKk1ejdXqSo0id0K9s9nxAStw";
  const ss = SpreadsheetApp.openById(targetSheetId);
  const sheet = ss.getSheetByName("Archive_Test");

  if (!sheet) throw new Error("❌ 'Archive_Test' not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("⚠️ Archive_Test has no data rows.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);
  const keyMap = new Map();

  rows.forEach(row => {
    const batchID = row[0];
    const fiscalMonth = row[1];
    const techID = row[4];

    // ✅ Safer footer check (only skip if techID is a string AND contains "total")
    if (typeof techID === 'string' && techID.toLowerCase().includes("total")) return;

    // 🧠 Dedupe key: fiscalMonth + techID
    const dedupeKey = `${fiscalMonth}_${techID}`;
    const existing = keyMap.get(dedupeKey);

    if (!existing || batchID > existing[0]) {
      keyMap.set(dedupeKey, row);
    }
  });

  const deduped = Array.from(keyMap.values());
  deduped.unshift(headers);

  sheet.clearContents();
  sheet.getRange(1, 1, deduped.length, deduped[0].length).setValues(deduped);

  const removed = rows.length - keyMap.size;
  Logger.log(`🧹 Dedup complete by FiscalMonth + TechID. Removed ${removed} row(s), kept ${keyMap.size}.`);
}


function loadOntracToArchive(batchID, fiscalMonth) {
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

  let allRows = [];

  ontracTabs.forEach(tabName => {
    const sheet = sourceSS.getSheetByName(tabName);
    const { source, region } = extractMetadata(tabName);
    const lastRow = sheet.getLastRow();
    const numRows = Math.max(0, lastRow - 2);

    if (numRows === 0) return;

    const techIDs = sheet.getRange(3, 1, numRows, 1).getValues(); // Col A
    const data = sheet.getRange(3, 2, numRows, 33).getValues();   // Col B–AH

    for (let i = 0; i < numRows; i++) {
      const techID = techIDs[i][0];
      if (!techID) continue;

      const uniqueKey = `${batchID}_${region}_${techID}`;
      const row = [batchID, fiscalMonth, region, source, techID, uniqueKey, ...data[i]];
      allRows.push(row);
    }

    Logger.log(`✅ Prepared ${allRows.length} rows from ${region}`);
  });

  if (allRows.length > 0) {
  const startRow = archiveSheet.getLastRow();
  archiveSheet.insertRowsAfter(startRow, allRows.length);
  archiveSheet.getRange(startRow + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
  Logger.log(`🧪 Archive Load Complete: ${allRows.length} total rows written.`); 
  } else {
    Logger.log("⚠️ No rows to write.");
  }
}

//Helper functions follow

function injectArchiveHeaders() {
  const targetSheetId = "19zuCKxg_4Akn9CpubsAKk1ejdXqSo0id0K9s9nxAStw";
  const sheet = SpreadsheetApp.openById(targetSheetId).getSheetByName("Archive_Test");

  if (!sheet) throw new Error("❌ 'Archive_Test' not found.");

  const headers = [
    "BatchID", "Region", "Source", "TechID", "UniqueKey",
    "TechName", "Supervisor", "Total Jobs", "Installs", "TCs", "SROs",
    "TUResult", "TUEligibleJobs", "ToolUsage",
    "Promoters", "Detractors", "tNPS Surveys", "tNPS Rate",
    "FTRFailJobs", "Total FTR/Contact Jobs", "FTR%",
    "48Hr Contact Orders", "48Hr Contact Rate%",
    "PHT Jobs", "PHT Pure Pass", "PHT Fails", "PHT RTM", "PHT Pass%", "PHT Pure Pass%",
    "TotalAppts", "TotalMetAppts", "MetRate",
    "Rework Count", "Rework Rate%", "SOI Count", "SOI Rate%", "Repeat Count", "Repeat Rate%"
  ];

  // Overwrite row 1
  sheet.insertRows(1);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  Logger.log("🧾 Header row injected at the top of Archive_Test.");
}

function promptForFiscalMonth(defaultFM) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Enter Fiscal Month (YYYY-MM)',
    `Default: ${defaultFM}. Leave blank or click Cancel to use it.`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    Logger.log(`ℹ️ User cancelled input. Using default: ${defaultFM}`);
    return defaultFM;
  }

  const input = response.getResponseText().trim();
  const isValid = /^\d{4}-\d{2}$/.test(input);

  if (isValid) {
    Logger.log(`🗓 Fiscal Month manually set to: ${input}`);
    return input;
  } else {
    Logger.log(`⚠️ Invalid or blank input. Using default: ${defaultFM}`);
    return defaultFM;
  }
}

function previewAndRunLoader() {
  const batchID = generateBatchID();
  const defaultFM = getFiscalMonth(new Date());
  const fiscalMonth = promptForFiscalMonth(defaultFM);

  Logger.log(`🧪 Batch ID: ${batchID}`);
  Logger.log(`🗓 Fiscal Month: ${fiscalMonth}`);

  loadOntracToArchive(batchID, fiscalMonth);
  deduplicateArchive();

  Logger.log("✅ Full loader run with manual input complete.");
}