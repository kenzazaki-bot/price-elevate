/**
 * Air Liquide Canada - Bulk Pack Review Tool
 * FINAL VERSION + ADMIN VIEW + DYNAMIC ETL PIPELINE (FIXED FLOW)
 */

const COLORS = {
  PRIMARY: "#006272",
  ACCENT: "#ff8200",
  SECONDARY: "#4ec3e0",
  NEUTRAL: "#d0d0ce",
  NEGATIVE: "#a73321",
  POSITIVE: "#58b947",
  WHITE: "#ffffff"
};

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
    const year = val.getFullYear();
    const month = String(val.getMonth() + 1).padStart(2, '0');
    const day = String(val.getDate()).padStart(2, '0');
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
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
}

function getAllUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    if (!sheet) return {AE: [], SC: []};
    ensureColumns(sheet, 40); 
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {AE: [], SC: []};

    const aeSet = new Set();
    const scSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const aeName = data[i].length > 24 ? toJSONSafe(data[i][24], 'string') : ""; 
      const scName = data[i].length > 25 ? toJSONSafe(data[i][25], 'string') : ""; 
      if (aeName !== "") aeSet.add(aeName);
      if (scName !== "") scSet.add(scName);
    }
    return { AE: [...aeSet].sort(), SC: [...scSet].sort() };
  } catch (e) {
    return {AE: [], SC: []};
  }
}

function getItemsForUser(selectedUser, role) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
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
      
      const priceList = toJSONSafe(row[0], 'string');
      const itemNum = toJSONSafe(row[2], 'string');
      const uniqueKey = priceList + "||" + itemNum;
      
      const acctNum = toJSONSafe(row[23], 'string');
      const acctName = toJSONSafe(row[26], 'string');
      const classCode = toJSONSafe(row[27], 'string');
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
      group.estImpact = toJSONSafe(group.estImpact, 'number');
      group.currPrice = toJSONSafe(group.currPrice, 'number');
      group.suggPrice = toJSONSafe(group.suggPrice, 'number');
      group.qty = toJSONSafe(group.qty, 'number');
    });

    results.sort((a, b) => b.estImpact !== a.estImpact ? b.estImpact - a.estImpact : a.priceList.localeCompare(b.priceList));
    return { items: results, portfolioStatus: portfolioStatus };
  } catch (error) {
    throw new Error("Server Error: " + error.toString());
  }
}

function getManagerSummary() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    if (!sheet) return [];
    ensureColumns(sheet, 40); 
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const groups = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      while (row.length < 40) row.push(""); 
      
      const priceList = toJSONSafe(row[0], 'string');
      const itemNum = toJSONSafe(row[2], 'string');
      const aeName = toJSONSafe(row[24], 'string') || "Unassigned";
      const uniqueKey = aeName + "||" + priceList + "||" + itemNum;

      if (!groups[uniqueKey]) {
        groups[uniqueKey] = {
          aeName: aeName, rowIds: [], currPrice: toJSONSafe(row[5], 'number'), estImpact: 0, qty: 0,
          finalPrice: row[31] !== "" ? toJSONSafe(row[31], 'number') : "", status: toJSONSafe(row[33], 'string') || "PENDING", portfolioStatus: "Not Submitted"
        };
      }
      
      groups[uniqueKey].rowIds.push(i + 1);
      groups[uniqueKey].estImpact += toJSONSafe(row[29], 'number');
      groups[uniqueKey].qty += toJSONSafe(row[20], 'number');
      if (row[33] !== "") groups[uniqueKey].status = toJSONSafe(row[33], 'string');
      if (row[31] !== "") groups[uniqueKey].finalPrice = toJSONSafe(row[31], 'number');
      let akVal = toJSONSafe(row[36], 'string');
      if (akVal === "Pending Approval" || akVal === "Approved") groups[uniqueKey].portfolioStatus = akVal;
    }

    const managerAgg = {};
    Object.values(groups).forEach(g => {
      if (g.rowIds.length > 0) g.estImpact = g.estImpact / g.rowIds.length;
      if (!managerAgg[g.aeName]) managerAgg[g.aeName] = { aeName: g.aeName, total: 0, pending: 0, approved: 0, rejected: 0, estImpact: 0, realized: 0, portfolioStatus: "Not Submitted" };
      managerAgg[g.aeName].total++;
      managerAgg[g.aeName].estImpact += g.estImpact;
      if (g.status === "APPROVED") {
        managerAgg[g.aeName].approved++;
        if (g.finalPrice !== "") managerAgg[g.aeName].realized += (g.finalPrice - g.currPrice) * g.qty;
      } else if (g.status === "REJECTED") {
        managerAgg[g.aeName].rejected++;
      } else managerAgg[g.aeName].pending++;
      
      if (g.portfolioStatus === "Pending Approval" || g.portfolioStatus === "Approved") managerAgg[g.aeName].portfolioStatus = g.portfolioStatus;
    });

    return Object.values(managerAgg).sort((a,b) => b.estImpact - a.estImpact);
  } catch (error) { throw new Error(error.toString()); }
}

function submitPortfolio(aeName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    ensureColumns(sheet, 40);
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

function approvePortfolio(aeName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    ensureColumns(sheet, 40);
    const data = sheet.getDataRange().getValues();
    const targetUser = String(aeName).toLowerCase();
    const outValues = [];
    for (let i = 1; i < data.length; i++) {
      while (data[i].length < 40) data[i].push("");
      outValues.push([String(data[i][24]).toLowerCase() === targetUser ? "Approved" : data[i][36]]);
    }
    if (outValues.length > 0) sheet.getRange(2, 37, outValues.length, 1).setValues(outValues);
    return true;
  } catch (e) { throw new Error(e.toString()); }
}

/**
 * ADMIN DATA INGESTION & CSV MAPPING
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
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) {
        targetFile = f; break;
      }
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

/**
 * DYNAMIC RULE DROPDOWN SCANNER
 * Scans the actual mapped CSVs to populate rule targeting options.
 */
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
    
    return {
      seg1: [...dims.seg1].sort(), seg2: [...dims.seg2].sort(), seg3: [...dims.seg3].sort(),
      seg4: [...dims.seg4].sort(), seg5: [...dims.seg5].sort(), classCode: [...dims.classCode].sort()
    };
  } catch(e) { 
    throw new Error("Dimension Load Error: " + e.toString()); 
  }
}

/**
 * THE CORE ETL PIPELINE: Extract, Transform (Map & Math), Shard, Load
 */
function runETLPipeline(payload) {
  try {
    // 1. EXTRACT & BUILD HIERARCHY DICTIONARY
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
          sc: hMap.hmap_sc !== "" ? row[hMap.hmap_sc] : "",
          manager: hMap.hmap_mgr !== "" ? row[hMap.hmap_mgr] : "",
          region: hMap.hmap_reg !== "" ? row[hMap.hmap_reg] : "",
          email: hMap.hmap_email !== "" ? row[hMap.hmap_email] : ""
        };
      }
    }

    // 2. EXTRACT RAW CSVs
    let folderId = payload.csvFolderId.trim();
    if (folderId.includes('folders/')) folderId = folderId.split('folders/')[1].split('?')[0].split('/')[0];
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    const splitGroups = {};
    const cMap = payload.csvMap;
    let headersArr = []; 
    
    for(let i=0; i<40; i++) headersArr.push("Col_"+i);
    headersArr[0] = "PRICE_LIST"; headersArr[2] = "ITEM_NUM"; headersArr[3] = "DESCRIPTION";
    headersArr[5] = "CURRENT_PRICE"; headersArr[9] = "SEGMENT_1"; headersArr[10] = "SEGMENT_2";
    headersArr[11] = "SEGMENT_3"; headersArr[12] = "SEGMENT_4"; headersArr[13] = "SEGMENT_5";
    headersArr[20] = "VOLUME"; headersArr[23] = "ACCOUNT_NUM"; headersArr[24] = "AE_NAME";
    headersArr[25] = "SC_NAME"; headersArr[26] = "ACCOUNT_NAME"; headersArr[27] = "CLASS_CODE";
    headersArr[28] = "SUGGESTED_PRICE"; headersArr[29] = "EST_IMPACT"; 
    headersArr[37] = "MANAGER"; headersArr[38] = "REGION"; headersArr[39] = "AE_EMAIL"; 

    let totalRowsProcessed = 0;

    // Loop all CSVs in folder
    while (files.hasNext()) {
      let f = files.next();
      if (f.getName().toLowerCase().endsWith('.csv') || f.getMimeType() === MimeType.CSV || f.getMimeType() === MimeType.PLAIN_TEXT) {
        
        let csvData = Utilities.parseCsv(f.getBlob().getDataAsString());
        if(csvData.length < 2) continue; 
        
        // 3. TRANSFORM (Process Rows)
        for(let i=1; i<csvData.length; i++) {
          let cRow = csvData[i];
          let outRow = new Array(40).fill("");
          
          if(cMap.map_pl !== "") outRow[0] = cRow[cMap.map_pl];
          if(cMap.map_item !== "") outRow[2] = cRow[cMap.map_item];
          if(cMap.map_desc !== "") outRow[3] = cRow[cMap.map_desc];
          
          let currentPrice = cMap.map_price !== "" ? parseFloat(cRow[cMap.map_price]) || 0 : 0;
          outRow[5] = currentPrice;
          
          if(cMap.map_seg1 !== "") outRow[9] = cRow[cMap.map_seg1];
          if(cMap.map_seg2 !== "") outRow[10] = cRow[cMap.map_seg2];
          if(cMap.map_seg3 !== "") outRow[11] = cRow[cMap.map_seg3];
          if(cMap.map_seg4 !== "") outRow[12] = cRow[cMap.map_seg4];
          if(cMap.map_seg5 !== "") outRow[13] = cRow[cMap.map_seg5];
          
          let volume = cMap.map_vol !== "" ? parseFloat(cRow[cMap.map_vol]) || 0 : 0;
          outRow[20] = volume;
          
          if(cMap.map_acct_num !== "") outRow[23] = cRow[cMap.map_acct_num];
          
          let aeName = cMap.map_ae !== "" ? String(cRow[cMap.map_ae]).trim() : "";
          outRow[24] = aeName;
          
          if(cMap.map_acct_name !== "") outRow[26] = cRow[cMap.map_acct_name];
          if(cMap.map_class !== "") outRow[27] = cRow[cMap.map_class];

          // APPLY ENRICHMENT FROM HIERARCHY DICTIONARY
          let dictMatch = hierarchyDict[aeName.toLowerCase()];
          if(dictMatch) {
            outRow[25] = dictMatch.sc;
            outRow[37] = dictMatch.manager;
            outRow[38] = dictMatch.region;
            outRow[39] = dictMatch.email;
          } else {
            // Fallback to CSV SC if no dictionary match
            if(cMap.map_sc !== "") outRow[25] = cRow[cMap.map_sc];
          }

          // APPLY CAMPAIGN MATH
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
            
            if(match) {
              newSuggPrice = currentPrice * (1 + (rule.pct / 100));
              break; 
            }
          }
          
          outRow[28] = newSuggPrice.toFixed(2); 
          outRow[29] = ((newSuggPrice - currentPrice) * volume).toFixed(2); 

          // SHARD INTO BUCKETS (Split by Segment 2 - Index 10)
          let splitKey = outRow[10] ? String(outRow[10]).trim() : "Unassigned";
          if(!splitGroups[splitKey]) splitGroups[splitKey] = [];
          splitGroups[splitKey].push(outRow);
          totalRowsProcessed++;
        }
      }
    }

    if(totalRowsProcessed === 0) throw new Error("No data was processed. Ensure valid CSV files exist.");

    // 4. LOAD (Generate Files)
    const timestamp = Utilities.formatDate(new Date(), "America/New_York", "yyyyMMdd_HHmm");
    const outFolder = DriveApp.createFolder(`BulkPack_Distributed_${timestamp}`);
    const results = [];

    for (let key in splitGroups) {
      let safeKey = key.replace(/[^a-zA-Z0-9_-]/g, '_');
      let newSs = SpreadsheetApp.create(`ReviewData_${safeKey}`);
      let fileId = newSs.getId();
      DriveApp.getFileById(fileId).moveTo(outFolder);
      
      let targetSheet = newSs.getSheets()[0];
      targetSheet.setName("Sheet1");
      
      let finalData = [headersArr].concat(splitGroups[key]);
      targetSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
      
      results.push({ key: key, count: splitGroups[key].length, url: newSs.getUrl() });
    }

    return { folderUrl: outFolder.getUrl(), processed: totalRowsProcessed, files: results };

  } catch(e) {
    throw new Error(e.toString());
  }
}

// ... (KEEP ALL EXISTING PDF AND UPLOAD TAB FUNCTIONS EXACTLY AS THEY WERE) ...
function getApprovedAccounts(selectedUser, role) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    if (!sheet) return [];
    
    ensureColumns(sheet, 40);
    const data = sheet.getDataRange().getValues();
    const targetUser = toJSONSafe(selectedUser, 'string').toLowerCase();
    const accts = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      while (row.length < 40) row.push(""); 
      
      let rowUser = role === 'SC' ? toJSONSafe(row[25], 'string').toLowerCase() : toJSONSafe(row[24], 'string').toLowerCase();
      
      if (rowUser !== targetUser) continue;
      if (toJSONSafe(row[33], 'string') !== "APPROVED") continue;
      
      const acctNum = toJSONSafe(row[23], 'string');
      const acctName = toJSONSafe(row[26], 'string');
      if (acctNum) {
        accts[acctNum] = acctName || "Unknown Customer";
      }
    }
    return Object.keys(accts).map(k => ({ num: k, name: accts[k] })).sort((a, b) => a.name.localeCompare(b.name));
  } catch(e) {
    Logger.log("Error in getApprovedAccounts: " + e.toString());
    return [];
  }
}

function generatePDF(selectedUser, role, acctNum) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    ensureColumns(sheet, 40); 
    const data = sheet.getDataRange().getValues();
    const targetUser = toJSONSafe(selectedUser, 'string').toLowerCase();
    
    let customerName = "Customer";
    let startDateStr = "[Date]";
    let itemsHtml = "";
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      while (row.length < 40) row.push(""); 

      let rowUser = role === 'SC' ? toJSONSafe(row[25], 'string').toLowerCase() : toJSONSafe(row[24], 'string').toLowerCase();
      if (rowUser !== targetUser) continue;
      if (toJSONSafe(row[33], 'string') !== "APPROVED") continue;
      if (toJSONSafe(row[23], 'string') !== acctNum) continue;
      
      let rName = toJSONSafe(row[26], 'string');
      if (rName) customerName = rName;
      
      let storedDate = row[32];
      if (startDateStr === "[Date]" && storedDate) {
        if (storedDate instanceof Date) {
           const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
           startDateStr = `${months[storedDate.getMonth()]} ${storedDate.getDate()}, ${storedDate.getFullYear()}`;
        } else {
           startDateStr = storedDate;
        }
      }

      let product = toJSONSafe(row[3], 'string') || toJSONSafe(row[2], 'string');
      let currentPrice = toJSONSafe(row[5], 'number');
      let newPrice = toJSONSafe(row[31], 'number');
      let pctIncrease = currentPrice > 0 ? (((newPrice - currentPrice) / currentPrice) * 100).toFixed(2) : 0;

      itemsHtml += `
      <tr>
          <td style="padding:8px; border-bottom:1px solid #eee;">${product}</td>
          <td style="padding:8px; border-bottom:1px solid #eee;">$${currentPrice.toFixed(2)}</td>
          <td style="padding:8px; border-bottom:1px solid #eee; font-weight:bold;">$${newPrice.toFixed(2)}</td>
          <td style="padding:8px; border-bottom:1px solid #eee;">${pctIncrease}%</td>
      </tr>`;
    }

    const today = new Date();
    const monthsFull = ["January","February","March","April","May","June","July","August","September","October","November","December"];
    const dateString = `${monthsFull[today.getMonth()]} ${String(today.getDate()).padStart(2, '0')}, ${today.getFullYear()}`;

    let html = `
    <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333; padding: 40px;">
      <p style="text-transform: uppercase;">${dateString}</p>
      <p style="margin-top:20px; margin-bottom:30px;">
        <b>${customerName}</b><br>
        [Billing Address]<br>
        [Billing City, Province, Postal Code]
      </p>
      <p style="font-weight: bold; font-size:16px;">Subject: Price Adjustments Notification</p>
      <p>Dear Customer,</p>
      <p>We value your continued partnership with Air Liquide Canada and are writing to inform you of an upcoming adjustment to your account pricing.</p>
      <p>We continuously strive to optimize our operations to provide you with high-quality service and a reliable supply at competitive rates. While we have made every effort to absorb recent market fluctuations internally, we are currently facing exceptional and unanticipated cost pressures across our supply chain. Specifically, we are experiencing unprecedented volatility and sustained increases in fuel and energy costs, which have impacted our distribution and facility operations. Furthermore, global supply constraints continue to affect helium sourcing and raw material acquisition.</p>
      <p>As a result of these factors, a price adjustment will be applied to your account effective <b>${startDateStr}</b>. You can expect to see the specific changes reflected in your upcoming billing cycle.</p>
      <p>The increase will be defined as the structure below:</p>
      <table style="width:100%; border-collapse: collapse; margin-top: 15px; text-align: left; font-size: 13px;">
          <thead>
              <tr>
                  <th style="padding:8px; border-bottom:2px solid #006272; color:#006272;">Product</th>
                  <th style="padding:8px; border-bottom:2px solid #006272; color:#006272;">Current Price</th>
                  <th style="padding:8px; border-bottom:2px solid #006272; color:#006272;">New Price</th>
                  <th style="padding:8px; border-bottom:2px solid #006272; color:#006272;">% Increase</th>
              </tr>
          </thead>
          <tbody>
              ${itemsHtml}
          </tbody>
      </table>
      <p style="margin-top: 40px;">Sincerely,</p>
      <p><b>Air Liquide Canada</b></p>
    </div>`;

    const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
    const base64 = Utilities.base64Encode(blob.getBytes());
    return { filename: `${customerName.replace(/[^a-zA-Z0-9]/g, '_')}_Price_Increase.pdf`, base64: base64 };
  } catch (error) {
    Logger.log("Error generating PDF: " + error.toString());
    throw new Error(error.toString());
  }
}

function submitReview(updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  ensureColumns(sheet, 40); 

  updates.forEach(u => {
    const ids = Array.isArray(u.rowIds) ? u.rowIds : [u.rowIds];
    ids.forEach(id => {
      if (u.decision) sheet.getRange(id, 34).setValue(u.decision);
      if (u.finalPrice !== undefined && u.finalPrice !== "") {
         sheet.getRange(id, 32).setValue(u.finalPrice);
      } else if (u.decision === "REJECTED") {
         sheet.getRange(id, 32).clearContent();
      }
      if (u.startDate) sheet.getRange(id, 33).setValue(u.startDate);
      if (u.comment) sheet.getRange(id, 35).setValue(u.comment);
    });
  });
  return true;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ALC Pricing Tools').addItem('Generate "To be loaded" Tab', 'generateUploadTab').addToUi();
}

function getAppLogo() {
  const fileId = "1G2ZX-8kS-uXOimRuXpUj80X3svyKEeUK"; 
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType(); 
    const base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + mimeType + ";base64," + base64;
  } catch (e) {
    return "https://upload.wikimedia.org/wikipedia/en/thumb/e/e0/Air_Liquide_Logo.svg/320px-Air_Liquide_Logo.svg.png";
  }
}

function generateUploadTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Sheet1");
  let uploadSheet = ss.getSheetByName("To be loaded");
  
  if (!uploadSheet) uploadSheet = ss.insertSheet("To be loaded");
  else uploadSheet.clear();

  let headers = new Array(20).fill("");
  headers[2] = "TYPE"; headers[3] = "ITEM_NUM"; headers[6] = "START_DATE";
  headers[8] = "UNIT_PRICE"; headers[9] = "UOM"; headers[10] = "ACTIVE";
  headers[11] = "PRICE_LIST_NAME"; headers[12] = "PRECEDENCE";
  headers[17] = "CYCLE_DAYS"; headers[19] = "CYL_GROUP";

  uploadSheet.getRange(1, 1, 1, 20).setValues([headers])
    .setBackground(COLORS.PRIMARY).setFontColor(COLORS.WHITE).setFontWeight("bold");

  const data = sourceSheet.getDataRange().getValues();
  const outputRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length > 33 && row[33] === "APPROVED") { 
      let newRow = new Array(20).fill("");
      newRow[2] = "Item Number"; 
      
      const storedDate = (row.length > 32) ? row[32] : "";
      let finalDate = "26-Feb-2026"; 
      
      if (storedDate instanceof Date) {
         const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
         finalDate = `${storedDate.getDate()}-${months[storedDate.getMonth()]}-${storedDate.getFullYear()}`;
      } else if (storedDate) {
         finalDate = storedDate;
      }

      newRow[6] = finalDate; 
      newRow[10] = "Yes";
      newRow[3] = row[2]; 
      newRow[8] = (row.length > 31) ? row[31] : ""; 
      newRow[9] = row[4];
      newRow[11] = row[0]; newRow[12] = row[8]; newRow[17] = row[11]; newRow[19] = row[9];
      outputRows.push(newRow);
    }
  }

  if (outputRows.length > 0) {
    uploadSheet.getRange(2, 1, outputRows.length, 20).setValues(outputRows);
    SpreadsheetApp.getUi().alert(`Success: Generated ${outputRows.length} lines.`);
  } else {
    SpreadsheetApp.getUi().alert("No lines marked as 'APPROVED' found.");
  }
}
