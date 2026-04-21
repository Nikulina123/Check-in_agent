/**
 * Webiz Inventory Agent – Google Apps Script
 *
 * HOW TO DEPLOY (one-time setup by IT admin):
 *  1. Open your Google Sheet → Extensions → Apps Script
 *  2. Paste this entire file, replacing the default code
 *  3. Click Deploy → New deployment
 *     Type: Web app
 *     Execute as: Me
 *     Who has access: Anyone          ← required so devices can POST
 *  4. Click Deploy → copy the Web App URL
 *  5. Paste that URL into WebizInventory_Windows.ps1 AND WebizInventory_macOS.sh AND
 *     WebizInventory_Linux.sh as APPS_SCRIPT_URL
 *  6. To update: make changes → Deploy → Manage deployments → edit existing → New version
 */

const SHEET_NAME = "Inventory";
const HEADERS    = [
  "Timestamp", "First Name", "Last Name", "Email", "Project",
  "Hostname", "IP Address", "Brand", "Model", "Serial Number",
  "CPU", "RAM", "Storage", "OS"
];

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Create header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
      hdr.setFontWeight("bold")
         .setBackground("#1A2B5A")
         .setFontColor("#FFFFFF")
         .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.timestamp    || new Date().toISOString(),
      data.first_name   || "",
      data.last_name    || "",
      data.email        || "",
      data.project      || "",
      data.hostname     || "",
      data.ip_address   || "",
      data.brand        || "",
      data.model        || "",
      data.serial_number|| "",
      data.cpu          || "",
      data.ram          || "",
      data.storage      || "",
      data.os           || "",
    ]);

    // Auto-resize columns for readability
    sheet.autoResizeColumns(1, HEADERS.length);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Simple health-check (open URL in browser to verify deployment is alive)
function doGet(e) {
  return ContentService.createTextOutput("Webiz Inventory API – OK");
}
