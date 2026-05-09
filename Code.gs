// ═══════════════════════════════════════════════════════════
// Marathon Registration Tracker — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────
const SPREADSHEET_ID    = "YOUR_SPREADSHEET_ID_HERE";
const SHEET_NAME        = "Sheet1";
const PAID_COLUMN_INDEX = 12;        // Column L = 12
const ADMIN_PASSWORD    = "YOUR_SECRET_PASSWORD_HERE";
// ──────────────────────────────────────────────────────────


// ── GET: handles read + CORS preflight ────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "";

  if (action === "read") {
    const result = handleRead();
    return buildCorsResponse(result);
  }

  // Unknown action — still return CORS-safe response
  return buildCorsResponse({ error: "Unknown action" });
}


// ── POST: handles markPaid ────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || "";

    if (action === "markPaid") {
      const result = handleMarkPaid(body);
      return buildCorsResponse(result);
    }

    return buildCorsResponse({ error: "Unknown action" });
  } catch (err) {
    return buildCorsResponse({ error: err.message });
  }
}


// ── READ ──────────────────────────────────────────────────
function handleRead() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data  = sheet.getDataRange().getValues();

    if (data.length === 0) return { rows: [], headers: [] };

    const headers = data[0].map(h => String(h).trim());
    const rows    = data.slice(1).map((row, i) => {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = row[j]; });
      obj.__rowIndex = i + 2; // 1-indexed, skipping header row
      return obj;
    });

    return { headers, rows };
  } catch (err) {
    return { error: err.message };
  }
}


// ── MARK PAID ─────────────────────────────────────────────
function handleMarkPaid(body) {
  const { rowIndex, password } = body;

  if (password !== ADMIN_PASSWORD) {
    return { error: "Incorrect password" };
  }

  if (!rowIndex || typeof rowIndex !== "number" || rowIndex < 2) {
    return { error: "Invalid row index" };
  }

  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    sheet.getRange(rowIndex, PAID_COLUMN_INDEX).setValue("Yes");
    SpreadsheetApp.flush();
    return { success: true, rowIndex };
  } catch (err) {
    return { error: err.message };
  }
}


// ── CORS-safe JSON response ───────────────────────────────
// Apps Script does not support setting response headers directly,
// so we use HtmlService to wrap our JSON — this bypasses the
// CORS block that ContentService triggers in browsers.
function buildCorsResponse(data) {
  const json = JSON.stringify(data);
  const html = HtmlService.createHtmlOutput(json);
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}
