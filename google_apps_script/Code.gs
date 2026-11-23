/**
 * Aptech Monthly BSR — Google Sheets bridge.
 *
 * Deploy as a web app. Set Script Properties:
 *   BSR_SYNC_TOKEN       - a long random secret
 *   BSR_SPREADSHEET_ID   - the one spreadsheet this bridge may update
 *
 * The desktop app sends generated worksheet values. This bridge clears and
 * writes values only; existing formatting, merges, dimensions and validations
 * are retained.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return json_({ success: false, error: "Another BSR sync is already running." });
  }

  try {
    var payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    var properties = PropertiesService.getScriptProperties();
    var expectedToken = properties.getProperty("BSR_SYNC_TOKEN");
    if (!expectedToken || payload.token !== expectedToken) {
      return json_({ success: false, error: "Unauthorized. Check the sync secret." });
    }
    if (payload.action !== "replaceSheetValues") {
      return json_({ success: false, error: "Unsupported sync action." });
    }

    var expectedSpreadsheetId = properties.getProperty("BSR_SPREADSHEET_ID");
    if (!expectedSpreadsheetId || expectedSpreadsheetId !== String(payload.spreadsheetId || "")) {
      return json_({ success: false, error: "Spreadsheet ID does not match the configured BSR spreadsheet." });
    }

    var spreadsheet = SpreadsheetApp.openById(expectedSpreadsheetId);
    var sheetName = String(payload.sheetName || "").trim();
    // Prefer the matching month tab. The configured gid remains a fallback
    // for institutes whose Google Sheet tab names do not match the BSR month.
    var sheet = sheetName ? spreadsheet.getSheetByName(sheetName) : null;
    if (!sheet && payload.sheetId !== undefined && payload.sheetId !== null) {
      sheet = spreadsheet.getSheetById(Number(payload.sheetId));
    }
    if (!sheet) {
      return json_({
        success: false,
        error: "No matching month tab was found and the configured worksheet ID is unavailable."
      });
    }

    var rowCount = Number(payload.rowCount);
    var columnCount = Number(payload.columnCount);
    if (!Number.isInteger(rowCount) || rowCount < 1 ||
        !Number.isInteger(columnCount) || columnCount < 1) {
      return json_({ success: false, error: "Invalid worksheet dimensions in the sync payload." });
    }
    if (rowCount > sheet.getMaxRows() || columnCount > sheet.getMaxColumns()) {
      return json_({
        success: false,
        error: "The target tab is smaller than the generated BSR. Expand its row/column grid before syncing."
      });
    }

    var cells = payload.cells || [];
    if (!Array.isArray(cells)) {
      return json_({ success: false, error: "Invalid cell list in the sync payload." });
    }

    // Build one rectangular value matrix and write it in one Sheets API call.
    // This avoids thousands of individual setValue calls for a full BSR.
    var values = [];
    for (var r = 0; r < rowCount; r++) {
      values.push(new Array(columnCount).fill(""));
    }

    cells.forEach(function (item) {
      var row = Number(item.row);
      var col = Number(item.col);
      if (!Number.isInteger(row) || !Number.isInteger(col) ||
          row < 1 || row > rowCount || col < 1 || col > columnCount) {
        throw new Error("A cell coordinate is outside the generated worksheet bounds.");
      }

      var value = item.value;
      if (value && typeof value === "object" && value.type === "date") {
        value = new Date(value.value);
        if (isNaN(value.getTime())) {
          throw new Error("The workbook contains an invalid date value.");
        }
      }
      values[row - 1][col - 1] = value;
    });

    var clearRows = Math.max(rowCount, sheet.getLastRow(), 1);
    var clearColumns = Math.max(columnCount, sheet.getLastColumn(), 1);

    // clearContent() preserves formatting, merged cells, dimensions and validations.
    sheet.getRange(1, 1, clearRows, clearColumns).clearContent();
    sheet.getRange(1, 1, rowCount, columnCount).setValues(values);

    SpreadsheetApp.flush();
    return json_({
      success: true,
      cellsWritten: cells.length,
      message: "Updated " + sheet.getName() + " without changing its formatting."
    });
  } catch (err) {
    return json_({ success: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({
    success: true,
    message: "Aptech BSR Google Sheets bridge is ready."
  });
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
