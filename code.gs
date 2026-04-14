/**
 * MULTI-REGION + MULTI-CUSTOMER-TYPE KML ZONE EXPORT
 */

const CONFIG_SPREADSHEET_ID = "1hSsBUWbOetshDSDcuXUwPJk5S4s_9jzdKsYkSNI7Yts";
const MASTER_PROGRESS_SHEET = "Master_Progress";

const GLOBAL_SETTINGS = {
  BATCH_SIZE: 2,
  MAX_EXECUTION_TIME_MS: 300000,
  KML_ICON_URL: "http://maps.google.com/mapfiles/kml/shapes/man.png",
  ADMIN_EMAIL: "cadastralteam.grm@gmail.com",
  HEADERS: {
    CUSTOMER: "Account_Number",
    LAT: "Latitude",
    LON: "Longitude",
    ZONE: "Zone",
    BILL_STATUS: "Bill_Delivery_Status",
    MONITOR_STATUS: "Monitoring_Status",
    GPS_AVAILABLE: "GPS_Available?",
  }
};

/**
 * Load configuration dynamically from main dashboard "Config" sheet
 */
function loadConfigFromDashboard() {
  console.log("Loading configuration from Config sheet...");
  const ss = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Config");
  if (!sheet) throw new Error('Missing "Config" sheet');

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const regionIdx = headers.indexOf("Region");
  const typeIdx = headers.indexOf("Type");
  const ssIdIdx = headers.indexOf("SpreadsheetId");
  const dashIdIdx = headers.indexOf("DashboardId");
  const folderIdIdx = headers.indexOf("FolderId");
  const colorColIdx = headers.indexOf("ColorColumn");

  if ([regionIdx, typeIdx, ssIdIdx, dashIdIdx, folderIdIdx, colorColIdx].includes(-1)) {
    throw new Error("Config sheet missing required columns");
  }

  const configs = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[ssIdIdx]) continue;
    configs.push({
      region: row[regionIdx],
      type: row[typeIdx],
      spreadsheetId: row[ssIdIdx],
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${row[ssIdIdx]}/edit`,
      csvUrl: `https://docs.google.com/spreadsheets/d/${row[ssIdIdx]}/gviz/tq?tqx=out:csv`,
      dashboardId: row[dashIdIdx],
      folderId: row[folderIdIdx],
      colorColumn: row[colorColIdx]
    });
  }
  console.log(`Loaded ${configs.length} configurations.`);
  return configs;
}

/**
 * Ensure MASTER_PROGRESS exists & is synced with config
 */
function ensureMasterProgress(configs) {
  console.log("Ensuring Master Progress sheet exists...");
  const ss = SpreadsheetApp.openById(CONFIG_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_PROGRESS_SHEET);
  if (!sheet) {
    console.log("Master Progress sheet not found. Creating...");
    sheet = ss.insertSheet(MASTER_PROGRESS_SHEET);
    sheet.appendRow(["Region", "Type", "Status", "Last Updated"]);
  }

  const mpValues = sheet.getDataRange().getValues();
  const existingMap = {};
  for (let i = 1; i < mpValues.length; i++) {
    existingMap[`${mpValues[i][0]}__${mpValues[i][1]}`] = i + 1;
  }

  configs.forEach(cfg => {
    const key = `${cfg.region}__${cfg.type}`;
    if (!existingMap[key]) {
      console.log(`Adding missing entry to Master Progress: ${cfg.region} - ${cfg.type}`);
      sheet.appendRow([cfg.region, cfg.type, "Pending", ""]);
    }
  });
  return sheet;
}

/**
 * Get next pending or in-progress region/type
 */
function getNextRegionType(masterSheet) {
  const values = masterSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const status = (values[i][2] || "").toString().trim().toUpperCase();
    if (status === "PENDING" || status === "INPROGRESS" || status === "ERROR") {
      console.log(`Next task found: ${values[i][0]} - ${values[i][1]} (Status: ${status})`);
      return { rowIndex: i + 1, region: values[i][0], type: values[i][1], status };
    }
  }
  console.log("No pending or in-progress tasks found.");
  return null;
}

/**
 * Reset MASTER_PROGRESS when all done 
 */
function resetMasterProgress(masterSheet) {
  console.log("Checking if all tasks are complete for reset...");
  const values = masterSheet.getDataRange().getValues();
  let allDone = true;
  for (let i = 1; i < values.length; i++) {
    if ((values[i][2] || "").toString().trim().toUpperCase() !== "DONE") {
      allDone = false;
      break;
    }
  }
  if (allDone) {
    console.log("All tasks complete. Resetting to Pending...");
    for (let i = 1; i < values.length; i++) {
      masterSheet.getRange(i + 1, 3).setValue("Pending");
      masterSheet.getRange(i + 1, 4).setValue("");
    }
    SpreadsheetApp.flush();
  }
}

/**
 * Main runner
 */
function doExportAllZone() {
  console.log("=== Starting Export Run ===");
  const startTime = Date.now(); // ⬅ Track start time
  const configs = loadConfigFromDashboard();
  const masterSheet = ensureMasterProgress(configs);

  resetMasterProgress(masterSheet);

  let nextTask;
  while ((nextTask = getNextRegionType(masterSheet))) {

    if (Date.now() - startTime > GLOBAL_SETTINGS.MAX_EXECUTION_TIME_MS) {
      console.log("[SAFE EXIT] Approaching execution time limit, will resume next run.");
      break;
    }

    console.log(`Processing ${nextTask.region} - ${nextTask.type}...`);
    const cfg = configs.find(c => c.region === nextTask.region && c.type === nextTask.type);
    if (!cfg) {
      console.log(`Config missing for ${nextTask.region} - ${nextTask.type}`);
      masterSheet.getRange(nextTask.rowIndex, 3).setValue("Error");
      continue;
    }

    masterSheet.getRange(nextTask.rowIndex, 3).setValue("InProgress");
    SpreadsheetApp.flush();

    const success = processRegionType(cfg, startTime); // ⬅ PATCH: Pass startTime down

    if (success) {
      console.log(`Completed ${cfg.region} - ${cfg.type}`);
      masterSheet.getRange(nextTask.rowIndex, 3).setValue("Done");
      masterSheet.getRange(nextTask.rowIndex, 4).setValue(new Date());
      masterSheet.getRange(nextTask.rowIndex, 5).setValue(0);

    } else if (success === false || success === null) {
      console.log(`Keeping ${cfg.region} - ${cfg.type} as InProgress for next run`);
      
    } else {
      console.log(`Error processing ${cfg.region} - ${cfg.type}`);
      masterSheet.getRange(nextTask.rowIndex, 3).setValue("Error");
      let errorCount = Number(masterSheet.getRange(nextTask.rowIndex, 5).getValue()) || 0;
      errorCount++;
      masterSheet.getRange(nextTask.rowIndex, 5).setValue(errorCount);

      if (errorCount >= 4) {
        MailApp.sendEmail({
          to: GLOBAL_SETTINGS.ADMIN_EMAIL,
          subject: `Export Failure: ${cfg.region} - ${cfg.type}`,
          body: `Region ${cfg.region} - ${cfg.type} failed ${errorCount} times.\nCheck logs for details.`
        });
      }
      SpreadsheetApp.flush();
    }
  }
  console.log("=== Export Run Finished ===");
}

/**
 * Process one Region + Customer Type
 */
function processRegionType(cfg, startTime) {
  console.log(`--- Starting process for ${cfg.region} - ${cfg.type} ---`);
  const PROGRESS_SHEET_NAME = `Export_Progress_${cfg.region}_${cfg.type}`;
  const dashboardSS = SpreadsheetApp.openById(cfg.dashboardId);

  const data = fetchCSVData(cfg.csvUrl);
  console.log(`Fetched ${data.length - 1} rows of data for ${cfg.region} - ${cfg.type}`);

  const headers = data[0];
  const columnIndexes = getColumnIndexes(headers);

  let progressSheet = dashboardSS.getSheetByName(PROGRESS_SHEET_NAME);
  if (!progressSheet) {
    console.log(`Creating progress sheet: ${PROGRESS_SHEET_NAME}`);
    progressSheet = dashboardSS.insertSheet(PROGRESS_SHEET_NAME);
  }

  let pendingBatches = getPendingBatchesFromSheet(progressSheet, 999);
  console.log(`Found ${pendingBatches.length} pending batches initially.`);
  if (pendingBatches.length === 0) {
    ensureBatchListExists(progressSheet, data, columnIndexes);
    pendingBatches = getPendingBatchesFromSheet(progressSheet, 999);
    console.log(`Batch list created. Pending batches: ${pendingBatches.length}`);
  }
  if (pendingBatches.length === 0) {
    console.log("No batches to process. Marking as complete.");
    return true;
  }

  const groupedData = groupDataByZone(data, columnIndexes); // ⬅ Zone only
  const folder = DriveApp.getFolderById(cfg.folderId);

  while ((Date.now() - startTime) < (GLOBAL_SETTINGS.MAX_EXECUTION_TIME_MS - 2000)) {
    const chunk = pendingBatches.splice(0, GLOBAL_SETTINGS.BATCH_SIZE);
    if (!chunk.length) break;

    console.log(`Processing batch chunk: ${chunk.length} items for ${cfg.region} - ${cfg.type}`);
    const result = generateKMLByZone(groupedData, folder, progressSheet, startTime, chunk, cfg, headers, columnIndexes);
    if (!result.moreTime) {
      console.log("Time limit reached. Will resume in next run.");
      return null;
    }
  }
  const complete = pendingBatches.length === 0;
  console.log(`Process complete for ${cfg.region} - ${cfg.type}: ${complete}`);
  return complete;
}

/* --- Remaining helper functions are unchanged --- */

function fetchCSVData(url) {
  const csv = UrlFetchApp.fetch(url).getContentText();
  return Utilities.parseCsv(csv);
}

function getColumnIndexes(headers) {
  const map = {};
  for (const key in GLOBAL_SETTINGS.HEADERS) {
    map[key.toLowerCase()] = headers.indexOf(GLOBAL_SETTINGS.HEADERS[key]);
  }
  return map;
}

function groupDataByZone(data, columnIndexes) {
  const grouped = {};
  for (let i = 1; i < data.length; i++) {
    const parsed = parseZoneRow(data[i], columnIndexes);
    if (!parsed) continue;
    (grouped[parsed.key] = grouped[parsed.key] || []).push(data[i]);
  }
  return grouped;
}

function parseZoneRow(row, columnIndexes) {
  if ((row[columnIndexes.gps_available] || "").toString().trim().toUpperCase() !== "Y") return null;
  const rawZone = (row[columnIndexes.zone] || "").trim();
  if (!rawZone || rawZone === "-") return null;
  const zoneNum = parseInt(rawZone, 10);
  if (isNaN(zoneNum)) return null;
  return { zoneNum, key: `${zoneNum}` };
}

function ensureBatchListExists(sheet, data, columnIndexes) {
  sheet.clear().appendRow(["Zone", "Status"]);
  const zoneSet = new Set();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const parsed = parseZoneRow(data[i], columnIndexes);
    if (parsed && !zoneSet.has(parsed.key)) {
      zoneSet.add(parsed.key);
      rows.push([parsed.zoneNum.toString().padStart(3, "0"), "Pending"]);
    }
  }
  rows.sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function getPendingBatchesFromSheet(sheet, maxCount) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxZone = headers.indexOf("Zone");
  const idxStatus = headers.indexOf("Status");
  return values.slice(1)
    .map((row, i) => ({
      rowIndex: i + 2,
      zone: row[idxZone].toString().trim(),
      status: (row[idxStatus] || "").toString().trim().toUpperCase()
    }))
    .filter(r => r.status === "PENDING")
    .slice(0, maxCount);
}

function generateKMLByZone(groupedData, parentFolder, progressSheet, startTime, pendingBatches, cfg, headers, columnIndexes) {
  const SAFE_TIME = GLOBAL_SETTINGS.MAX_EXECUTION_TIME_MS - 2000;
  for (const entry of pendingBatches) {
    if (Date.now() - startTime >= SAFE_TIME) {
      console.log("⏸️ Time limit nearly reached before processing batch. Leaving batch pending...");
      return { moreTime: false };
    }

    const key = `${parseInt(entry.zone, 10)}`;
    const filteredRows = groupedData[key] || [];
    if (!filteredRows.length) {
      progressSheet.getRange(entry.rowIndex, 2).setValue("Done");
      SpreadsheetApp.flush();
      continue;
    }

    const zoneCode = entry.zone.toString().padStart(3, "0");
    const fileName = `Zone ${zoneCode} Customers.kml`;

    const kml = buildKMLMulti(filteredRows, headers, columnIndexes, GLOBAL_SETTINGS.KML_ICON_URL, `Zone ${zoneCode}`, cfg);

    const existingFile = getFileByName(parentFolder, fileName);
    if (existingFile) existingFile.setContent(kml);
    else parentFolder.createFile(fileName, kml, "application/vnd.google-earth.kml+xml");

    progressSheet.getRange(entry.rowIndex, 2).setValue("Done");
    SpreadsheetApp.flush();
  }
  return { moreTime: true };
}

function buildKMLMulti(rows, headers, columnIndexes, iconUrl, zoneLabel, cfg) {
  const colorColumnIndex = headers.indexOf(cfg.colorColumn);
  let kml = `<?xml version="1.0" encoding="UTF-8"?>\n` +
            `<kml xmlns="http://www.opengis.net/kml/2.2"><Document>\n` +
            `<name>${cfg.region} ${cfg.type} - ${zoneLabel}</name>\n`;
  const SAFE_TIME = GLOBAL_SETTINGS.MAX_EXECUTION_TIME_MS - 2000;
  const startTime = Date.now();

  rows.forEach((row, idx) => {
    if (idx > 0 && idx % 25 === 0) {
      if (Date.now() - startTime >= SAFE_TIME) {
        console.log(`⏸️ Time check triggered after ${idx} rows. Stopping KML generation early.`);
        return kml + "</Document>\n</kml>";
      }
    }

    const name = String(row[columnIndexes.customer]).trim();
    const lat = row[columnIndexes.lat];
    const lon = row[columnIndexes.lon];
    const status = row[colorColumnIndex];
    const color = status && status.toString().trim().toUpperCase() === "Y" ? "ff00ff00" : "ff000080";

    let description = "<table border='1'>";
    headers.forEach((header, j) => {
      description += `<tr><td><b>${header}</b></td><td>${row[j] || "N/A"}</td></tr>`;
    });
    description += "</table>";

    kml += `<Placemark><name>${name}</name><description><![CDATA[${description}]]></description>
      <Style><IconStyle><color>${color}</color><Icon><href>${iconUrl}</href></Icon><scale>1.2</scale></IconStyle></Style>
      <Point><coordinates>${lon},${lat}</coordinates></Point></Placemark>\n`;
  });

  return `${kml}</Document>\n</kml>`;
}

function findOrCreateSubFolder(parentFolder, subFolderName) {
  const folders = parentFolder.getFoldersByName(subFolderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(subFolderName);
}

function getFileByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}
