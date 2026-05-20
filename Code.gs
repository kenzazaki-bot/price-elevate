/**
 * Air Liquide Canada - Bulk Pack Review Tool
 * ENTRY POINT & ROUTING
 */

function doGet(e) {
  const page = (e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : 'user';
  let template = HtmlService.createTemplateFromFile('WebApp');
  template.pageContext = page;
  return template
    .evaluate()
    .setTitle('')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// NEW: Helper function required to inject CSS and JS files into the main WebApp HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() { 
  SpreadsheetApp.getUi().createMenu('ALC Pricing Tools').addItem('Generate "To be loaded" Tab', 'generateUploadTab').addToUi(); 
}

function getAppLogo() {
  const fileId = "1G2ZX-8kS-uXOimRuXpUj80X3svyKEeUK"; 
  try {
    const file = DriveApp.getFileById(fileId); const blob = file.getBlob(); const mimeType = blob.getContentType(); const base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + mimeType + ";base64," + base64;
  } catch (e) { return "https://upload.wikimedia.org/wikipedia/en/thumb/e/e0/Air_Liquide_Logo.svg/320px-Air_Liquide_Logo.svg.png"; }
}

function forceDriveAuth() {
  DriveApp.getFiles();
}
