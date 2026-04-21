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

// Column indices (0-based) used for look-up and comparison
const COL = {
  TIMESTAMP:     0,
  FIRST_NAME:    1,
  LAST_NAME:     2,
  EMAIL:         3,
  PROJECT:       4,
  HOSTNAME:      5,
  IP_ADDRESS:    6,
  BRAND:         7,
  MODEL:         8,
  SERIAL_NUMBER: 9,
  CPU:           10,
  RAM:           11,
  STORAGE:       12,
  OS:            13,
};

/**
 * Returns the current timestamp formatted in Tbilisi time (UTC+4).
 * Example: "2026-04-21T14:30:00+04:00"
 */
function getTbilisiTimestamp() {
  const now    = new Date();
  const offset = 4 * 60 * 60 * 1000; // UTC+4 in ms
  const local  = new Date(now.getTime() + offset);
  // toISOString gives UTC — replace Z with +04:00 after shifting the time
  return local.toISOString().replace(/\.\d{3}Z$/, "+04:00");
}

/**
 * Returns the most recent row (as a 0-based array) for the same device,
 * excluding the row we just appended.
 *
 * Match priority:
 *   1. Serial Number (most reliable device identifier)
 *   2. Hostname (fallback when serial is absent)
 *
 * NOTE: a different First Name / Last Name / Email / Project on the same
 * device is intentionally NOT part of the look-up key — it is compared
 * afterwards to detect an ownership change.
 */
function findPreviousEntry(sheet, serialNumber, hostname, currentLastRow) {
  const searchRows = currentLastRow - 2; // rows between header row and the new row
  if (searchRows <= 0) return null;

  const data = sheet.getRange(2, 1, searchRows, HEADERS.length).getValues();

  // Walk backwards — most-recent previous entry wins
  for (let i = data.length - 1; i >= 0; i--) {
    const sn   = String(data[i][COL.SERIAL_NUMBER]).trim();
    const host = String(data[i][COL.HOSTNAME]).trim();

    // Serial number match takes priority
    if (serialNumber && sn && sn === serialNumber) return data[i];
    // Hostname fallback (only when serial is missing on either side)
    if (!serialNumber && hostname && host === hostname) return data[i];
  }
  return null;
}

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

    const newRowValues = [
      data.timestamp    || getTbilisiTimestamp(),
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
    ];

    sheet.appendRow(newRowValues);
    const newRowNum = sheet.getLastRow(); // 1-based sheet row of the row we just added

    // ── Change-detection note ────────────────────────────────────────────────
    const prevEntry = findPreviousEntry(
      sheet,
      String(data.serial_number || "").trim(),
      String(data.hostname      || "").trim(),
      newRowNum
    );

    let noteText;
    if (!prevEntry) {
      // No history for this device — first ever check-in, no note needed
      noteText = null;
    } else {
      // Compare every field except Timestamp (index 0)
      const fieldsToCompare = HEADERS.slice(1); // "First Name" … "OS"
      const allMatch = fieldsToCompare.every((_, i) => {
        const col = i + 1; // skip Timestamp
        return String(newRowValues[col]).trim() === String(prevEntry[col]).trim();
      });

      const noteTime = getTbilisiTimestamp();
      if (allMatch) {
        noteText = `Collected nothing changed\n${noteTime}`;
      } else {
        const prevFirst = String(prevEntry[COL.FIRST_NAME]).trim();
        const prevLast  = String(prevEntry[COL.LAST_NAME]).trim();
        noteText = `Previous Checkin device owner was: ${prevFirst} ${prevLast}\n${noteTime}`;
      }
    }

    if (noteText) {
      // Attach note to the Timestamp cell of the new row
      sheet.getRange(newRowNum, 1).setNote(noteText);
    }

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
