/**
 * GLOBAL CONFIGURATION & UTILITIES
 */

const COLORS = { PRIMARY: "#006272", ACCENT: "#ff8200", SECONDARY: "#4ec3e0", NEUTRAL: "#d0d0ce", NEGATIVE: "#a73321", POSITIVE: "#58b947", WHITE: "#ffffff" };

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

function saveDefaultStartDate(dateStr) {
  PropertiesService.getScriptProperties().setProperty('DEFAULT_START_DATE', dateStr);
  return true;
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
