// ═══════════════════════════════════════════════════════════
// Marathon Registration Tracker — Google Apps Script Backend
// Deploy as a Web App (Execute as: Me, Access: Anyone)
// ═══════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
// Replace with your spreadsheet ID from the URL:
// https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit

const SHEET_NAME = "Sheet1";
// Replace with the exact name of your sheet tab (bottom of screen)

const PAID_COLUMN_INDEX = 12;
// Column L = 12 (1-indexed). Change if your "Paid" column is different.

const ADMIN_PASSWORD = "YOUR_SECRET_PASSWORD_HERE";
// Set a password to protect the mark-as-paid action.
// Users will need to enter this in the app to unlock marking.
// ──────────────────────────────────────────────────────────


function doGet(e) {
  const action = e.parameter.action;

  if (action === "read") {
    return handleRead();
  }

  return jsonResponse({ error: "Unknown action" }, 400);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "markPaid") {
      return handleMarkPaid(body);
    }

    return jsonResponse({ error: "Unknown action" }, 400);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}


// ── READ: Return all rows as JSON ──────────────────────────
function handleRead() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      return jsonResponse({ rows: [], headers: [] });
    }

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1).map((row, i) => {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = row[j]; });
      obj.__rowIndex = i + 2; // 1-indexed, skip header row
      return obj;
    });

    return jsonResponse({ headers, rows });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}


// ── WRITE: Mark a specific row as Paid ────────────────────
function handleMarkPaid(body) {
  const { rowIndex, password } = body;

  // Validate password
  if (password !== ADMIN_PASSWORD) {
    return jsonResponse({ error: "Incorrect password" }, 403);
  }

  // Validate row index
  if (!rowIndex || typeof rowIndex !== "number" || rowIndex < 2) {
    return jsonResponse({ error: "Invalid row index" }, 400);
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // Write "Yes" to the Paid column (column L = 12)
    sheet.getRange(rowIndex, PAID_COLUMN_INDEX).setValue("Yes");
    SpreadsheetApp.flush();

    return jsonResponse({ success: true, rowIndex });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}


// ── Helper: return JSON response with CORS headers ─────────
function jsonResponse(data, statusCode) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
