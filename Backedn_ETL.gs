/**
 * ETL PIPELINE & ADMIN FUNCTIONS
 */

function getHeadersFromCSVFolder(folderInput) {
  try {
    let folderId = folderInput.trim();
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let targetFile = null;
    while (files.hasNext()) {
      let f = files.next();
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) { targetFile = f; break;
      }
    }
    if (!targetFile) throw new Error("Could not find a valid .csv file.");
    const csvData = Utilities.parseCsv(targetFile.getBlob().getDataAsString());
    if (csvData.length === 0) throw new Error("The CSV file appears to be empty.");
    return csvData[0];
  } catch (e) { throw new Error("Drive Read Error: " + e.toString());
  }
}

function getHeadersFromHierarchySheet(sheetInput, tabName) {
  try {
    let sheetId = sheetInput.trim();
    if (sheetId.includes('/d/')) sheetId = sheetId.split('/d/')[1].split('/')[0];
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(tabName.trim());
    if (!sheet) throw new Error(`Tab "${tabName}" not found.`);
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } catch (e) { throw new Error("Hierarchy Read Error: " + e.toString());
  }
}

function getCampaignDimensionsFromCSV(payload) {
  try {
    let folderId = payload.csvFolderId.trim();
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const cMap = payload.csvMap;
    const dims = { seg1: new Set(), seg2: new Set(), seg3: new Set(), seg4: new Set(), seg5: new Set(), classCode: new Set() };
    while (files.hasNext()) {
      let f = files.next();
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) {
        let csvData = Utilities.parseCsv(f.getBlob().getDataAsString());
        for (let i = 1; i < csvData.length; i++) {
          let row = csvData[i];
          if(cMap.map_seg1 !== "" && row[cMap.map_seg1]) dims.seg1.add(String(row[cMap.map_seg1]).trim());
          if(cMap.map_seg2 !== "" && row[cMap.map_seg2]) dims.seg2.add(String(row[cMap.map_seg2]).trim());
          if(cMap.map_seg3 !== "" && row[cMap.map_seg3]) dims.seg3.add(String(row[cMap.map_seg3]).trim());
          if(cMap.map_seg4 !== "" && row[cMap.map_seg4]) dims.seg4.add(String(row[cMap.map_seg4]).trim());
          if(cMap.map_seg5 !== "" && row[cMap.map_seg5]) dims.seg5.add(String(row[cMap.map_seg5]).trim());
          if(cMap.map_class !== "" && row[cMap.map_class]) dims.classCode.add(String(row[cMap.map_class]).trim());
        }
      }
    }
    return { seg1: [...dims.seg1].sort(), seg2: [...dims.seg2].sort(), seg3: [...dims.seg3].sort(), seg4: [...dims.seg4].sort(), seg5: [...dims.seg5].sort(), classCode: [...dims.classCode].sort() };
  } catch(e) { throw new Error("Dimension Load Error: " + e.toString());
  }
}

function runETLPipeline(payload) {
  try {
    let sheetId = payload.hierSheetId.trim();
    if (sheetId.includes('/d/')) sheetId = sheetId.split('/d/')[1].split('/')[0];
    const hierSs = SpreadsheetApp.openById(sheetId);
    const hierSheet = hierSs.getSheetByName(payload.hierTab.trim());
    const hierData = hierSheet.getDataRange().getValues();
    
    const hMap = payload.hierMap;
    const hierarchyDict = {};
    for(let i = 1; i < hierData.length; i++) {
      let row = hierData[i];
      let aeKey = String(row[hMap.hmap_ae]).toLowerCase().trim();
      if(aeKey) {
        hierarchyDict[aeKey] = {
          sc: hMap.hmap_sc !== "" ?
          row[hMap.hmap_sc] : "", manager: hMap.hmap_mgr !== "" ? row[hMap.hmap_mgr] : "",
          region: hMap.hmap_reg !== "" ?
          row[hMap.hmap_reg] : "", email: hMap.hmap_email !== "" ? row[hMap.hmap_email] : ""
        };
      }
    }

    let folderId = payload.csvFolderId.trim();
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    const splitGroups = {};
    const cMap = payload.csvMap;
    let headersArr = [];
    for(let i=0; i<40; i++) headersArr.push("Col_"+i);
    headersArr[0] = "PRICE_LIST"; headersArr[2] = "ITEM_NUM"; headersArr[3] = "DESCRIPTION"; headersArr[5] = "CURRENT_PRICE"; 
    headersArr[6] = "L12M_SALES";
    headersArr[7] = "L12M_GP"; 
    headersArr[8] = "PREVIOUS_PRICE";
    headersArr[9] = "SEGMENT_1"; headersArr[10] = "SEGMENT_2"; headersArr[11] = "SEGMENT_3"; headersArr[12] = "SEGMENT_4"; headersArr[13] = "SEGMENT_5";
    headersArr[20] = "VOLUME"; headersArr[23] = "ACCOUNT_NUM"; headersArr[24] = "AE_NAME"; headersArr[25] = "SC_NAME"; headersArr[26] = "ACCOUNT_NAME"; headersArr[27] = "CLASS_CODE";
    headersArr[28] = "SUGGESTED_PRICE"; headersArr[29] = "EST_IMPACT"; headersArr[37] = "MANAGER"; headersArr[38] = "REGION"; headersArr[39] = "AE_EMAIL"; 

    let totalRowsProcessed = 0;
    while (files.hasNext()) {
      let f = files.next();
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) {
        let csvData = Utilities.parseCsv(f.getBlob().getDataAsString());
        if(csvData.length < 2) continue; 
        
        for(let i=1; i<csvData.length; i++) {
          let cRow = csvData[i];
          let outRow = new Array(40).fill("");
          
          if(cMap.map_pl !== "") outRow[0] = cRow[cMap.map_pl];
          if(cMap.map_item !== "") outRow[2] = cRow[cMap.map_item];
          if(cMap.map_desc !== "") outRow[3] = cRow[cMap.map_desc];
          
          let currentPrice = cMap.map_price !== "" ? parseFloat(cRow[cMap.map_price]) || 0 : 0;
          outRow[5] = currentPrice;
          
          let prevPrice = cMap.map_prev_price !== "" ? parseFloat(cRow[cMap.map_prev_price]) || 0 : 0;
          outRow[8] = prevPrice;
          
          let l12mSales = cMap.map_l12m_sales !== "" ? parseFloat(cRow[cMap.map_l12m_sales]) || 0 : 0; outRow[6] = l12mSales;
          let l12mGP = cMap.map_l12m_gp !== "" ? parseFloat(cRow[cMap.map_l12m_gp]) || 0 : 0; outRow[7] = l12mGP;
          if(cMap.map_seg1 !== "") outRow[9] = cRow[cMap.map_seg1];
          if(cMap.map_seg2 !== "") outRow[10] = cRow[cMap.map_seg2];
          if(cMap.map_seg3 !== "") outRow[11] = cRow[cMap.map_seg3];
          if(cMap.map_seg4 !== "") outRow[12] = cRow[cMap.map_seg4];
          if(cMap.map_seg5 !== "") outRow[13] = cRow[cMap.map_seg5];
          
          let volume = cMap.map_vol !== "" ?
          parseFloat(cRow[cMap.map_vol]) || 0 : 0; outRow[20] = volume;
          if(cMap.map_acct_num !== "") outRow[23] = cRow[cMap.map_acct_num];
          let aeName = cMap.map_ae !== "" ? String(cRow[cMap.map_ae]).trim() : ""; outRow[24] = aeName;
          if(cMap.map_acct_name !== "") outRow[26] = cRow[cMap.map_acct_name];
          if(cMap.map_class !== "") outRow[27] = cRow[cMap.map_class];

          let dictMatch = hierarchyDict[aeName.toLowerCase()];
          if(dictMatch) {
            outRow[25] = dictMatch.sc; outRow[37] = dictMatch.manager;
            outRow[38] = dictMatch.region; outRow[39] = dictMatch.email;
          } else {
            if(cMap.map_sc !== "") outRow[25] = cRow[cMap.map_sc];
          }

          let newSuggPrice = currentPrice;
          for(let r=0; r < payload.campaignRules.length; r++) {
            let rule = payload.campaignRules[r];
            let match = true;
            if(rule.seg1 !== "ALL" && String(outRow[9]).trim() !== rule.seg1) match = false;
            if(rule.seg2 !== "ALL" && String(outRow[10]).trim() !== rule.seg2) match = false;
            if(rule.seg3 !== "ALL" && String(outRow[11]).trim() !== rule.seg3) match = false;
            if(rule.seg4 !== "ALL" && String(outRow[12]).trim() !== rule.seg4) match = false;
            if(rule.seg5 !== "ALL" && String(outRow[13]).trim() !== rule.seg5) match = false;
            if(rule.classCode !== "ALL" && String(outRow[27]).trim() !== rule.classCode) match = false;
            if(match) { newSuggPrice = currentPrice * (1 + (rule.pct / 100)); break; }
          }
          
          // Force Suggested Price to round up to exactly 2 decimal places
          newSuggPrice = Math.ceil(newSuggPrice * 100) / 100;
          
          outRow[28] = newSuggPrice.toFixed(2);
          outRow[29] = ((newSuggPrice - currentPrice) * volume).toFixed(2); 

          let splitKey = outRow[24] ? String(outRow[24]).trim() : "Unassigned";
          if(!splitGroups[splitKey]) splitGroups[splitKey] = [];
          splitGroups[splitKey].push(outRow);
          totalRowsProcessed++;
        }
      }
    }

    if(totalRowsProcessed === 0) throw new Error("No data was processed. Ensure valid CSV files exist.");
    
    const timestamp = Utilities.formatDate(new Date(), "America/New_York", "yyyyMMdd_HHmm");
    let cName = (payload.campaignName && payload.campaignName.trim() !== "") ? payload.campaignName.replace(/[^a-zA-Z0-9 _-]/g, '_') : "BulkPack";
    
    const outFolder = DriveApp.createFolder(`${cName}_AE_Shards_${timestamp}`);
    const results = [];
    
    const masterSs = SpreadsheetApp.getActiveSpreadsheet();
    let dirSheet = masterSs.getSheetByName("Directory");
    if (!dirSheet) dirSheet = masterSs.insertSheet("Directory");
    dirSheet.clear();
    dirSheet.appendRow(["AE_NAME", "FILE_ID", "MANAGER", "SC_NAME", "PORTFOLIO_STATUS", "TOTAL_ITEMS", "PENDING", "APPROVED", "REJECTED", "EST_IMPACT", "REALIZED_IMPACT"]);
    dirSheet.getRange("A1:K1").setFontWeight("bold").setBackground(COLORS.PRIMARY).setFontColor(COLORS.WHITE);
    for (let key in splitGroups) {
      let safeKey = key.replace(/[^a-zA-Z0-9_-]/g, '_');
      let newSs = SpreadsheetApp.create(`${cName}_AE_${safeKey}`);
      let fileId = newSs.getId();
      DriveApp.getFileById(fileId).moveTo(outFolder);
      
      let targetSheet = newSs.getSheets()[0];
      targetSheet.setName("Sheet1");
      
      let finalData = [headersArr].concat(splitGroups[key]);
      targetSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
      let sampleRow = splitGroups[key][0];
      dirSheet.appendRow([key, fileId, sampleRow[37], sampleRow[25], "Not Submitted", 0, 0, 0, 0, 0, 0]);
      results.push({ key: key, count: splitGroups[key].length, url: newSs.getUrl() });
    }

    return { folderUrl: outFolder.getUrl(), processed: totalRowsProcessed, files: results };
  } catch(e) { throw new Error(e.toString()); }
}

function generateUploadTab() {
  const masterSs = SpreadsheetApp.getActiveSpreadsheet();
  const dirSheet = masterSs.getSheetByName("Directory");
  if(!dirSheet) return SpreadsheetApp.getUi().alert("No Directory found. ETL must be run first.");
  
  let uploadSheet = masterSs.getSheetByName("To be loaded");
  if (!uploadSheet) uploadSheet = masterSs.insertSheet("To be loaded");
  else uploadSheet.clear();

  let headers = new Array(20).fill("");
  headers[2] = "TYPE"; headers[3] = "ITEM_NUM";
  headers[6] = "START_DATE"; headers[8] = "UNIT_PRICE"; headers[9] = "UOM"; headers[10] = "ACTIVE"; headers[11] = "PRICE_LIST_NAME"; headers[12] = "PRECEDENCE";
  headers[17] = "CYCLE_DAYS"; headers[19] = "CYL_GROUP";

  uploadSheet.getRange(1, 1, 1, 20).setValues([headers]).setBackground(COLORS.PRIMARY).setFontColor(COLORS.WHITE).setFontWeight("bold");

  const dirData = dirSheet.getDataRange().getValues();
  const outputRows = [];
  let globalDefault = PropertiesService.getScriptProperties().getProperty('DEFAULT_START_DATE') || "2026-02-25";
  let gdObj = new Date(globalDefault + "T12:00:00");
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  let fallbackDateStr = `${gdObj.getDate()}-${months[gdObj.getMonth()]}-${gdObj.getFullYear()}`;

  for(let d = 1; d < dirData.length; d++) {
    let fileId = dirData[d][1];
    if(!fileId) continue;
    try {
      let sheet = SpreadsheetApp.openById(fileId).getSheetByName("Sheet1");
      if(!sheet) continue;
      let data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length > 33 && (row[33] === "APPROVED" || row[33] === "MODIFIED")) { 
          let newRow = new Array(20).fill("");
          newRow[2] = "Item Number"; 
          const storedDate = (row.length > 32) ? row[32] : ""; let finalDate = fallbackDateStr;
          if (storedDate instanceof Date) { finalDate = `${storedDate.getDate()}-${months[storedDate.getMonth()]}-${storedDate.getFullYear()}`; } 
          else if (storedDate) { finalDate = storedDate;
          }
          
          newRow[6] = finalDate;
          newRow[10] = "Yes"; newRow[3] = row[2]; newRow[8] = (row.length > 31) ? row[31] : ""; newRow[9] = row[4];
          newRow[11] = row[0]; newRow[12] = row[8]; newRow[17] = row[11]; newRow[19] = row[9];
          outputRows.push(newRow);
        }
      }
    } catch(e) { Logger.log("Skipping invalid shard ID: " + fileId);
    }
  }

  if (outputRows.length > 0) {
    uploadSheet.getRange(2, 1, outputRows.length, 20).setValues(outputRows);
    SpreadsheetApp.getUi().alert(`Success: Compiled ${outputRows.length} lines across all shards.`);
  } else { SpreadsheetApp.getUi().alert("No lines marked as 'APPROVED' found across shards.");
  }
}

function debugDriveFetch() {
  const testFolderId = "1HOisUPdotYFPi1iRDgat4k_jRxeea6v-";
  try {
    Logger.log("1. Starting test with input: " + testFolderId);
    let fId = testFolderId.trim();
    if (fId.includes('folders/')) fId = fId.split('folders/')[1].split('?')[0].split('/')[0];
    Logger.log("2. Parsed Folder ID: " + fId);
    
    const folder = DriveApp.getFolderById(fId);
    Logger.log("3. Successfully connected to folder: " + folder.getName());
    
    const files = folder.getFiles();
    let targetFile = null;
    while (files.hasNext()) {
      let f = files.next();
      Logger.log("4. Scanning file: " + f.getName() + " (MimeType: " + f.getMimeType() + ")");
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) { 
        targetFile = f;
        Logger.log("5. TARGET FILE ACQUIRED: " + f.getName());
        break; 
      }
    }
    
    if (!targetFile) {
      Logger.log("❌ ERROR: Looked through files but found no valid .csv");
      return;
    }
    
    const csvText = targetFile.getBlob().getDataAsString();
    Logger.log("6. Read blob data. String length: " + csvText.length);
    
    const csvData = Utilities.parseCsv(csvText);
    Logger.log("✅ SUCCESS! Headers found: " + JSON.stringify(csvData[0]));
  } catch (e) {
    Logger.log("❌ FATAL CRASH: " + e.toString());
  }
}
/**
 * THE LOADER: Screens all Google Sheets in a folder and extracts final pricing data.
 * @param {String} folderId - The ID or URL of the folder containing the shard files.
 * @returns {Object} - An object containing the URL of the new file and the number of files processed.
 */
function generateLoaderCSVFromFolder(folderId) {
  try {
    if (!folderId) throw new Error("Folder ID is required.");
    
    // Clean up the folder ID just in case the user pastes a full URL
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    let csvData = [];
    let fileCount = 0;
    
    // Add the CSV Header row
    csvData.push(["Price List Name", "Item Number", "Final Price", "Final Start Date", "Status", "Manager Approval Status"].join(","));
    
    while (files.hasNext()) {
      let file = files.next();
      let ss = SpreadsheetApp.open(file);
      let sheet = ss.getSheetByName("Sheet1"); 
      if (!sheet) continue; 
      
      let data = sheet.getDataRange().getValues();
      
      // Loop through rows (skipping the header)
      for (let i = 1; i < data.length; i++) {
        let row = data[i];
        while (row.length < 40) row.push("");
        
        let priceList = String(row[0] || "").trim();
        let itemNum = String(row[2] || "").trim();
        let finalPrice = row[31];
        let startDate = row[32];
        let status = String(row[33] || "").trim();
        let mgrStatus = String(row[36] || "").trim();
        
        if (priceList && itemNum) {
          // Format Date safely (YYYY-MM-DD)
          if (startDate instanceof Date) {
            let y = startDate.getFullYear();
            let m = String(startDate.getMonth() + 1).padStart(2, '0');
            let d = String(startDate.getDate()).padStart(2, '0');
            startDate = `${y}-${m}-${d}`;
          } else {
            startDate = String(startDate || "").trim();
          }
          
          csvData.push([`"${priceList}"`, `"${itemNum}"`, finalPrice, `"${startDate}"`, `"${status}"`, `"${mgrStatus}"`].join(","));
        }
      }
      fileCount++;
    }
    
    // Create the physical CSV file in the same Google Drive Folder
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const fileName = "Price_to_Load_" + timestamp + ".csv";
    const newFile = folder.createFile(fileName, csvData.join("\n"), MimeType.CSV);
    
    return { url: newFile.getUrl(), count: fileCount };
    
  } catch (e) {
    throw new Error("Loader Error: " + e.message);
  }
}
