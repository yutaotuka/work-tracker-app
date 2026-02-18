const TARGET_SHEET_ID = "1ggWSLbaj5vFMmkcJP4EWAUxQusQ12m8jpWmta0-lmDg";
const SHEET_NAME = "app_state";
const MAX_HISTORY_ROWS = 3000;
const DATA_START_ROW = 2;
const META_LATEST_ROW_LABEL = "meta_latest_row";
const META_LATEST_ROW_CELL = "E1";
const META_LATEST_ROW_LABEL_CELL = "D1";

function doGet(e) {
  const action = getParam(e, "action");
  if (action !== "load" && action !== "latest") {
    return jsonOut({ ok: false, message: "Invalid action" });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetId = getParam(e, "sheetId") || TARGET_SHEET_ID;
    const sheet = getOrCreateSheet(sheetId);
    const row = getLatestDataRow(sheet);
    if (!row) {
      if (action === "latest") {
        return jsonOut({ ok: true, hasData: false, savedAt: 0 });
      }
      return jsonOut({ ok: true, data: null, savedAt: 0, message: "No saved data" });
    }

    const savedAt = Number(sheet.getRange(row, 2).getValue()) || 0;
    if (action === "latest") {
      return jsonOut({ ok: true, hasData: true, savedAt });
    }

    const value = sheet.getRange(row, 3).getValue();
    if (!value) {
      return jsonOut({ ok: true, data: null, savedAt, message: "No saved data" });
    }

    const parsed = JSON.parse(value);
    return jsonOut({ ok: true, data: parsed, savedAt });
  } catch (err) {
    return jsonOut({ ok: false, message: String(err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Ignore unlock errors.
    }
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const body = JSON.parse(e.postData.contents || "{}");
    if (body.action !== "save") {
      return jsonOut({ ok: false, message: "Invalid action" });
    }

    const sheetId = body.sheetId || TARGET_SHEET_ID;
    const sheet = getOrCreateSheet(sheetId);
    const savedAt = body.savedAt || Date.now();
    const json = JSON.stringify(body.data || {});

    const writeRow = getNextWriteRow(sheet);
    sheet.getRange(writeRow, 1, 1, 3).setValues([[new Date(savedAt), savedAt, json]]);
    setLatestDataRow(sheet, writeRow);
    return jsonOut({ ok: true, row: writeRow });
  } catch (err) {
    return jsonOut({ ok: false, message: String(err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Ignore unlock errors.
    }
  }
}

function getOrCreateSheet(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([["saved_at_human", "saved_at_ms", "data_json"]]);
    sheet.getRange(META_LATEST_ROW_LABEL_CELL).setValue(META_LATEST_ROW_LABEL);
    sheet.getRange(META_LATEST_ROW_CELL).setValue(0);
  }
  return sheet;
}

function getLatestDataRow(sheet) {
  const metaRow = Number(sheet.getRange(META_LATEST_ROW_CELL).getValue());
  if (isValidDataRow(metaRow)) {
    const stamp = Number(sheet.getRange(metaRow, 2).getValue()) || 0;
    if (stamp > 0) return metaRow;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 0;
  return lastRow;
}

function getNextWriteRow(sheet) {
  const latestRow = getLatestDataRow(sheet);
  if (!latestRow) return DATA_START_ROW;
  const offset = (latestRow - DATA_START_ROW + 1) % MAX_HISTORY_ROWS;
  return DATA_START_ROW + offset;
}

function setLatestDataRow(sheet, row) {
  sheet.getRange(META_LATEST_ROW_LABEL_CELL).setValue(META_LATEST_ROW_LABEL);
  sheet.getRange(META_LATEST_ROW_CELL).setValue(row);
}

function isValidDataRow(row) {
  return Number.isFinite(row) && row >= DATA_START_ROW && row < DATA_START_ROW + MAX_HISTORY_ROWS;
}

function getParam(e, key) {
  return (e && e.parameter && e.parameter[key]) || "";
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
