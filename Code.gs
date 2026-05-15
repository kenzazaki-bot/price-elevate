/**
 * Air Liquide Canada - Bulk Pack Review Tool
 * HUB AND SPOKE ARCHITECTURE - FINAL
 */

const COLORS = { PRIMARY: "#006272", ACCENT: "#ff8200", SECONDARY: "#4ec3e0", NEUTRAL: "#d0d0ce", NEGATIVE: "#a73321", POSITIVE: "#58b947", WHITE: "#ffffff" };

function doGet() {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('Bulk Pack Review - Eastern Region')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function toJSONSafe(val, type) {
  if (val === null || val === undefined) return type === 'number' ? 0 : "";
  if (val instanceof Date) {
    const year = val.getFullYear(); const month = String(val.getMonth() + 1).padStart(2, '0'); const day = String(val.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  if (type === 'number') {
    if (typeof val === 'number') return (isNaN(val) || !isFinite(val)) ? 0 : val;
    const parsed = parseFloat(val);
    return (isNaN(parsed) || !isFinite(parsed)) ? 0 : parsed;
  }
  return String(val).trim();
}

function ensureColumns(sheet, requiredCols) {
  const currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
}

function getAeFileId(aeName) {
  const dirSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Directory");
  if (!dirSheet) throw new Error("Directory tab not found. Admin must run the ETL pipeline first.");
  const data = dirSheet.getDataRange().getValues();
  const target = String(aeName).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) return data[i][1];
  }
  throw new Error("No data file found for AE: " + aeName + ". Ensure they exist in the raw CSVs.");
}

function getAllUsers() {
  try {
    const dirSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Directory");
    if (!dirSheet) return {AE: [], SC: []};
    const data = dirSheet.getDataRange().getValues();
    const aeSet = new Set(); const scSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) aeSet.add(String(data[i][0]).trim());
      if (data[i][3]) scSet.add(String(data[i][3]).trim());
    }
    return { AE: [...aeSet].sort(), SC: [...scSet].sort() };
  } catch (e) {
    return {AE: [], SC: []};
  }
}

function getItemsForUser(selectedUser, role) {
  try {
    const fileId = getAeFileId(selectedUser);
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("Sheet1");
    if (!sheet) return { items: [], portfolioStatus: "Not Submitted" };

    ensureColumns(sheet, 40); 
    const data = sheet.getDataRange().getValues(); 
    if (data.length < 2) return { items: [], portfolioStatus: "Not Submitted" };

    const targetUser = toJSONSafe(selectedUser, 'string').toLowerCase();
    const groups = {};
    let portfolioStatus = "Not Submitted";

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      while (row.length < 40) row.push(""); 

      let rowUser = role === 'SC' ? toJSONSafe(row[25], 'string').toLowerCase() : toJSONSafe(row[24], 'string').toLowerCase();
      if (rowUser !== targetUser) continue;
      
      let akVal = toJSONSafe(row[36], 'string');
      if (akVal === "Pending Approval" || akVal === "Approved") portfolioStatus = akVal;
      
      const priceList = toJSONSafe(row[0], 'string'); const itemNum = toJSONSafe(row[2], 'string');
      const uniqueKey = priceList + "||" + itemNum;
      const acctNum = toJSONSafe(row[23], 'string'); const acctName = toJSONSafe(row[26], 'string'); const classCode = toJSONSafe(row[27], 'string');
      let impactStr = acctNum ? `<b>${acctNum}</b>${acctName ? ` - ${acctName}` : ''}${classCode ? ` , ${classCode}` : ''}` : "";

      if (!groups[uniqueKey]) {
        groups[uniqueKey] = {
          key: uniqueKey, rowIds: [], priceList: priceList, itemNum: itemNum, desc: toJSONSafe(row[3], 'string'),
          salesPerson: toJSONSafe(row[24], 'string'), currPrice: toJSONSafe(row[5], 'number'), suggPrice: toJSONSafe(row[28], 'number'),
          estImpact: 0, qty: 0, volDisplay: toJSONSafe(row[20], 'string'),
          finalPrice: (row[31] !== "" && row[31] !== null) ? toJSONSafe(row[31], 'number') : "",
          status: toJSONSafe(row[33], 'string'), comment: toJSONSafe(row[34], 'string'), startDate: toJSONSafe(row[32], 'string'), impacts: []
        };
      }

      groups[uniqueKey].rowIds.push(i + 1);
      groups[uniqueKey].estImpact += toJSONSafe(row[29], 'number'); 
      groups[uniqueKey].qty += toJSONSafe(row[20], 'number');
      if (impactStr && !groups[uniqueKey].impacts.includes(impactStr)) groups[uniqueKey].impacts.push(impactStr);
      if (row[33] !== "" && groups[uniqueKey].status === "") groups[uniqueKey].status = toJSONSafe(row[33], 'string');
      if (row[34] !== "" && groups[uniqueKey].comment === "") groups[uniqueKey].comment = toJSONSafe(row[34], 'string');
      if (row[32] !== "" && groups[uniqueKey].startDate === "") groups[uniqueKey].startDate = toJSONSafe(row[32], 'string');
    }
    
    const results = Object.values(groups);
    results.forEach(group => {
      if (group.rowIds.length > 0) group.estImpact = group.estImpact / group.rowIds.length;
      group.estImpact = toJSONSafe(group.estImpact, 'number'); group.currPrice = toJSONSafe(group.currPrice, 'number'); group.suggPrice = toJSONSafe(group.suggPrice, 'number'); group.qty = toJSONSafe(group.qty, 'number');
    });

    results.sort((a, b) => b.estImpact !== a.estImpact ? b.estImpact - a.estImpact : a.priceList.localeCompare(b.priceList));
    return { items: results, portfolioStatus: portfolioStatus };
  } catch (error) { throw new Error("Server Error: " + error.toString()); }
}

function getManagerSummary() {
  try {
    const dirSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Directory");
    if (!dirSheet) return [];
    const dirData = dirSheet.getDataRange().getValues();
    if (dirData.length < 2) return [];

    const managerAgg = {};
    
    for (let d = 1; d < dirData.length; d++) {
      let aeName = dirData[d][0];
      let fileId = dirData[d][1];
      if(!fileId) continue;

      try {
        let ss = SpreadsheetApp.openById(fileId);
        let sheet = ss.getSheetByName("Sheet1");
        if(!sheet) continue;
        
        let data = sheet.getDataRange().getValues();
        let aeStats = { aeName: aeName, total: 0, pending: 0, approved: 0, rejected: 0, estImpact: 0, realized: 0, portfolioStatus: "Not Submitted" };
        const groups = {};

        for(let i = 1; i < data.length; i++) {
          const row = data[i];
          while (row.length < 40) row.push(""); 
          
          const priceList = toJSONSafe(row[0], 'string'); const itemNum = toJSONSafe(row[2], 'string');
          const uniqueKey = priceList + "||" + itemNum;
          
          let akVal = toJSONSafe(row[36], 'string');
          if (akVal === "Pending Approval" || akVal === "Approved") aeStats.portfolioStatus = akVal;

          if (!groups[uniqueKey]) {
            groups[uniqueKey] = { rowIds: [], currPrice: toJSONSafe(row[5], 'number'), estImpact: 0, qty: 0, finalPrice: row[31] !== "" ? toJSONSafe(row[31], 'number') : "", status: toJSONSafe(row[33], 'string') || "PENDING" };
          }
          groups[uniqueKey].rowIds.push(i + 1);
          groups[uniqueKey].estImpact += toJSONSafe(row[29], 'number');
          groups[uniqueKey].qty += toJSONSafe(row[20], 'number');
          if (row[33] !== "") groups[uniqueKey].status = toJSONSafe(row[33], 'string');
          if (row[31] !== "") groups[uniqueKey].finalPrice = toJSONSafe(row[31], 'number');
        }

        Object.values(groups).forEach(g => {
          if (g.rowIds.length > 0) g.estImpact = g.estImpact / g.rowIds.length;
          aeStats.total++;
          aeStats.estImpact += g.estImpact;
          if (g.status === "APPROVED") {
            aeStats.approved++;
            if (g.finalPrice !== "") aeStats.realized += (g.finalPrice - g.currPrice) * g.qty;
          } else if (g.status === "MODIFIED") { aeStats.approved++; if (g.finalPrice !== "") aeStats.realized += (g.finalPrice - g.currPrice) * g.qty; }
          else if (g.status === "REJECTED") { aeStats.rejected++; } 
          else aeStats.pending++;
        });

        managerAgg[aeName] = aeStats;
      } catch(e) { Logger.log("Could not process shard for: " + aeName); }
    }

    return Object.values(managerAgg).sort((a,b) => b.estImpact - a.estImpact);
  } catch (error) { throw new Error(error.toString()); }
}

function submitPortfolio(aeName) {
  try {
    const fileId = getAeFileId(aeName);
    const sheet = SpreadsheetApp.openById(fileId).getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();
    const targetUser = String(aeName).toLowerCase();
    const outValues = [];
    for (let i = 1; i < data.length; i++) {
      while (data[i].length < 40) data[i].push("");
      outValues.push([String(data[i][24]).toLowerCase() === targetUser ? "Pending Approval" : data[i][36]]);
    }
    if (outValues.length > 0) sheet.getRange(2, 37, outValues.length, 1).setValues(outValues);
    return true;
  } catch (e) { throw new Error(e.toString()); }
}

function submitReview(updates, aeName) {
  const fileId = getAeFileId(aeName);
  const sheet = SpreadsheetApp.openById(fileId).getSheetByName("Sheet1");
  updates.forEach(u => {
    const ids = Array.isArray(u.rowIds) ? u.rowIds : [u.rowIds];
    ids.forEach(id => {
      if (u.decision) sheet.getRange(id, 34).setValue(u.decision);
      if (u.finalPrice !== undefined && u.finalPrice !== "") { sheet.getRange(id, 32).setValue(u.finalPrice); } 
      else if (u.decision === "REJECTED") { sheet.getRange(id, 32).setValue(u.currPrice); }
      if (u.startDate) sheet.getRange(id, 33).setValue(u.startDate);
      if (u.comment) sheet.getRange(id, 35).setValue(u.comment);
    });
  });
  return true;
}

function getHeadersFromCSVFolder(folderInput) {
  try {
    let folderId = folderInput.trim();
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let targetFile = null;
    while (files.hasNext()) {
      let f = files.next();
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) { targetFile = f; break; }
    }
    if (!targetFile) throw new Error("Could not find a valid .csv file.");
    const csvData = Utilities.parseCsv(targetFile.getBlob().getDataAsString());
    if (csvData.length === 0) throw new Error("The CSV file appears to be empty.");
    return csvData[0]; 
  } catch (e) { throw new Error("Drive Read Error: " + e.toString()); }
}

function getHeadersFromHierarchySheet(sheetInput, tabName) {
  try {
    let sheetId = sheetInput.trim();
    if (sheetId.includes('/d/')) sheetId = sheetId.split('/d/')[1].split('/')[0];
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(tabName.trim());
    if (!sheet) throw new Error(`Tab "${tabName}" not found.`);
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } catch (e) { throw new Error("Hierarchy Read Error: " + e.toString()); }
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
  } catch(e) { throw new Error("Dimension Load Error: " + e.toString()); }
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
          sc: hMap.hmap_sc !== "" ? row[hMap.hmap_sc] : "", manager: hMap.hmap_mgr !== "" ? row[hMap.hmap_mgr] : "",
          region: hMap.hmap_reg !== "" ? row[hMap.hmap_reg] : "", email: hMap.hmap_email !== "" ? row[hMap.hmap_email] : ""
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
    headersArr[0] = "PRICE_LIST"; headersArr[2] = "ITEM_NUM"; headersArr[3] = "DESCRIPTION"; headersArr[5] = "CURRENT_PRICE"; headersArr[9] = "SEGMENT_1"; headersArr[10] = "SEGMENT_2"; headersArr[11] = "SEGMENT_3"; headersArr[12] = "SEGMENT_4"; headersArr[13] = "SEGMENT_5"; headersArr[20] = "VOLUME"; headersArr[23] = "ACCOUNT_NUM"; headersArr[24] = "AE_NAME"; headersArr[25] = "SC_NAME"; headersArr[26] = "ACCOUNT_NAME"; headersArr[27] = "CLASS_CODE"; headersArr[28] = "SUGGESTED_PRICE"; headersArr[29] = "EST_IMPACT"; headersArr[37] = "MANAGER"; headersArr[38] = "REGION"; headersArr[39] = "AE_EMAIL"; 

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
          let currentPrice = cMap.map_price !== "" ? parseFloat(cRow[cMap.map_price]) || 0 : 0; outRow[5] = currentPrice;
          
          if(cMap.map_seg1 !== "") outRow[9] = cRow[cMap.map_seg1];
          if(cMap.map_seg2 !== "") outRow[10] = cRow[cMap.map_seg2];
          if(cMap.map_seg3 !== "") outRow[11] = cRow[cMap.map_seg3];
          if(cMap.map_seg4 !== "") outRow[12] = cRow[cMap.map_seg4];
          if(cMap.map_seg5 !== "") outRow[13] = cRow[cMap.map_seg5];
          
          let volume = cMap.map_vol !== "" ? parseFloat(cRow[cMap.map_vol]) || 0 : 0; outRow[20] = volume;
          if(cMap.map_acct_num !== "") outRow[23] = cRow[cMap.map_acct_num];
          
          let aeName = cMap.map_ae !== "" ? String(cRow[cMap.map_ae]).trim() : ""; outRow[24] = aeName;
          if(cMap.map_acct_name !== "") outRow[26] = cRow[cMap.map_acct_name];
          if(cMap.map_class !== "") outRow[27] = cRow[cMap.map_class];

          let dictMatch = hierarchyDict[aeName.toLowerCase()];
          if(dictMatch) {
            outRow[25] = dictMatch.sc; outRow[37] = dictMatch.manager; outRow[38] = dictMatch.region; outRow[39] = dictMatch.email;
          } else {
            if(cMap.map_sc !== "") outRow[25] = cRow[cMap.map_sc];
          }

          let newSuggPrice = currentPrice; 
          for(let r=0; r < payload.campaignRules.length; r++) {
            let rule = payload.campaignRules[r]; let match = true;
            if(rule.seg1 !== "ALL" && String(outRow[9]).trim() !== rule.seg1) match = false;
            if(rule.seg2 !== "ALL" && String(outRow[10]).trim() !== rule.seg2) match = false;
            if(rule.seg3 !== "ALL" && String(outRow[11]).trim() !== rule.seg3) match = false;
            if(rule.seg4 !== "ALL" && String(outRow[12]).trim() !== rule.seg4) match = false;
            if(rule.seg5 !== "ALL" && String(outRow[13]).trim() !== rule.seg5) match = false;
            if(rule.classCode !== "ALL" && String(outRow[27]).trim() !== rule.classCode) match = false;
            
            if(match) { newSuggPrice = currentPrice * (1 + (rule.pct / 100)); break; }
          }
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
    const outFolder = DriveApp.createFolder(`BulkPack_AE_Shards_${timestamp}`);
    const results = [];
    
    const masterSs = SpreadsheetApp.getActiveSpreadsheet();
    let dirSheet = masterSs.getSheetByName("Directory");
    if (!dirSheet) dirSheet = masterSs.insertSheet("Directory");
    dirSheet.clear();
    dirSheet.appendRow(["AE_NAME", "FILE_ID", "MANAGER", "SC_NAME"]);
    dirSheet.getRange("A1:D1").setFontWeight("bold").setBackground(COLORS.PRIMARY).setFontColor(COLORS.WHITE);

    for (let key in splitGroups) {
      let safeKey = key.replace(/[^a-zA-Z0-9_-]/g, '_');
      let newSs = SpreadsheetApp.create(`ReviewData_AE_${safeKey}`);
      let fileId = newSs.getId();
      DriveApp.getFileById(fileId).moveTo(outFolder);
      
      let targetSheet = newSs.getSheets()[0];
      targetSheet.setName("Sheet1");
      
      let finalData = [headersArr].concat(splitGroups[key]);
      targetSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
      
      let sampleRow = splitGroups[key][0];
      dirSheet.appendRow([key, fileId, sampleRow[37], sampleRow[25]]);
      results.push({ key: key, count: splitGroups[key].length, url: newSs.getUrl() });
    }

    return { folderUrl: outFolder.getUrl(), processed: totalRowsProcessed, files: results };

  } catch(e) { throw new Error(e.toString()); }
}

function getApprovedAccounts(selectedUser, role) {
  try {
    const fileId = getAeFileId(selectedUser);
    const sheet = SpreadsheetApp.openById(fileId).getSheetByName("Sheet1");
    const data = sheet.getDataRange().getValues();
    const targetUser = toJSONSafe(selectedUser, 'string').toLowerCase();
    const accts = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i]; while (row.length < 40) row.push(""); 
      let rowUser = role === 'SC' ? toJSONSafe(row[25], 'string').toLowerCase() : toJSONSafe(row[24], 'string').toLowerCase();
      if (rowUser !== targetUser) continue;
      if (toJSONSafe(row[33], 'string') !== "APPROVED") continue;
      
      const acctNum = toJSONSafe(row[23], 'string'); const acctName = toJSONSafe(row[26], 'string');
      if (acctNum) accts[acctNum] = acctName || "Unknown Customer";
    }
    return Object.keys(accts).map(k => ({ num: k, name: accts[k] })).sort((a, b) => a.name.localeCompare(b.name));
  } catch(e) { return []; }
}

function onOpen() { SpreadsheetApp.getUi().createMenu('ALC Pricing Tools').addItem('Generate "To be loaded" Tab', 'generateUploadTab').addToUi(); }

function getAppLogo() {
  const fileId = "1G2ZX-8kS-uXOimRuXpUj80X3svyKEeUK"; 
  try {
    const file = DriveApp.getFileById(fileId); const blob = file.getBlob(); const mimeType = blob.getContentType(); const base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + mimeType + ";base64," + base64;
  } catch (e) { return "https://upload.wikimedia.org/wikipedia/en/thumb/e/e0/Air_Liquide_Logo.svg/320px-Air_Liquide_Logo.svg.png"; }
}

function generateUploadTab() {
  const masterSs = SpreadsheetApp.getActiveSpreadsheet();
  const dirSheet = masterSs.getSheetByName("Directory");
  if(!dirSheet) return SpreadsheetApp.getUi().alert("No Directory found. ETL must be run first.");
  
  let uploadSheet = masterSs.getSheetByName("To be loaded");
  if (!uploadSheet) uploadSheet = masterSs.insertSheet("To be loaded");
  else uploadSheet.clear();

  let headers = new Array(20).fill("");
  headers[2] = "TYPE"; headers[3] = "ITEM_NUM"; headers[6] = "START_DATE"; headers[8] = "UNIT_PRICE"; headers[9] = "UOM"; headers[10] = "ACTIVE"; headers[11] = "PRICE_LIST_NAME"; headers[12] = "PRECEDENCE"; headers[17] = "CYCLE_DAYS"; headers[19] = "CYL_GROUP";

  uploadSheet.getRange(1, 1, 1, 20).setValues([headers]).setBackground(COLORS.PRIMARY).setFontColor(COLORS.WHITE).setFontWeight("bold");

  const dirData = dirSheet.getDataRange().getValues();
  const outputRows = [];

  for(let d = 1; d < dirData.length; d++) {
    let fileId = dirData[d][1]; if(!fileId) continue;
    try {
      let sheet = SpreadsheetApp.openById(fileId).getSheetByName("Sheet1");
      if(!sheet) continue;
      let data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length > 33 && (row[33] === "APPROVED" || row[33] === "MODIFIED")) { 
          let newRow = new Array(20).fill("");
          newRow[2] = "Item Number"; 
          const storedDate = (row.length > 32) ? row[32] : ""; let finalDate = "26-Feb-2026"; 
          if (storedDate instanceof Date) { const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]; finalDate = `${storedDate.getDate()}-${months[storedDate.getMonth()]}-${storedDate.getFullYear()}`; } 
          else if (storedDate) { finalDate = storedDate; }
          
          newRow[6] = finalDate; newRow[10] = "Yes"; newRow[3] = row[2]; newRow[8] = (row.length > 31) ? row[31] : ""; newRow[9] = row[4]; newRow[11] = row[0]; newRow[12] = row[8]; newRow[17] = row[11]; newRow[19] = row[9];
          outputRows.push(newRow);
        }
      }
    } catch(e) { Logger.log("Skipping invalid shard ID: " + fileId); }
  }

  if (outputRows.length > 0) {
    uploadSheet.getRange(2, 1, outputRows.length, 20).setValues(outputRows);
    SpreadsheetApp.getUi().alert(`Success: Compiled ${outputRows.length} lines across all shards.`);
  } else { SpreadsheetApp.getUi().alert("No lines marked as 'APPROVED' found across shards."); }
}
