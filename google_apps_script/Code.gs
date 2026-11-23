/**
 * Aptech Monthly BSR — Google Sheets bridge.
 *
 * Deploy as a web app. Set Script Property BSR_SYNC_TOKEN to a long random
 * secret. The desktop app sends the generated worksheet values; this script
 * clears values only and writes into the existing tab, preserving formatting,
 * merged cells, widths, row heights and other sheet presentation.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return json_({ success: false, error: "Another BSR sync is already running." });
  }

  try {
    var payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    var expected = PropertiesService.getScriptProperties().getProperty("BSR_SYNC_TOKEN");
    if (!expected || payload.token !== expected) {
      return json_({ success: false, error: "Unauthorized. Check the sync secret." });
    }
    if (payload.action !== "replaceSheetValues") {
      return json_({ success: false, error: "Unsupported sync action." });
    }

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet || spreadsheet.getId() !== String(payload.spreadsheetId)) {
      return json_({ success: false, error: "Spreadsheet ID does not match this Apps Script project." });
    }

    var sheet = spreadsheet.getSheetById(Number(payload.sheetId));
    if (!sheet) {
      return json_({ success: false, error: "Target worksheet ID was not found." });
    }

    var rowCount = Math.max(1, Number(payload.rowCount) || 1);
    var columnCount = Math.max(1, Number(payload.columnCount) || 1);
    var clearRows = Math.max(rowCount, sheet.getLastRow(), 1);
    var clearColumns = Math.max(columnCount, sheet.getLastColumn(), 1);

    // clearContent() does not remove formatting, merges, dimensions or validations.
    sheet.getRange(1, 1, clearRows, clearColumns).clearContent();

    var cells = payload.cells || [];
    cells.forEach(function (item) {
      var value = item.value;
      if (value && typeof value === "object" && value.type === "date") {
        value = new Date(value.value);
      }
      // Values are written only to populated anchor cells. Merged-cell
      // placeholders are omitted by the Python exporter.
      sheet.getRange(Number(item.row), Number(item.col)).setValue(value);
    });

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
