const TARGET_SHEET_ID = "1ggWSLbaj5vFMmkcJP4EWAUxQusQ12m8jpWmta0-lmDg";
const SHEET_NAME = "app_state";

function doGet(e) {
  const action = getParam(e, "action");
  if (action !== "load") {
    return jsonOut({ ok: false, message: "Invalid action" });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheetId = getParam(e, "sheetId") || TARGET_SHEET_ID;
    const sheet = getOrCreateSheet(sheetId);
    const row = sheet.getLastRow();
    if (row < 2) {
      return jsonOut({ ok: true, data: null, message: "No saved data" });
    }

    const value = sheet.getRange(row, 3).getValue();
    if (!value) {
      return jsonOut({ ok: true, data: null, message: "No saved data" });
    }

    const parsed = JSON.parse(value);
    return jsonOut({ ok: true, data: parsed });
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

    sheet.appendRow([new Date(savedAt), savedAt, json]);
    return jsonOut({ ok: true });
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
  }
  return sheet;
}

function getParam(e, key) {
  return (e && e.parameter && e.parameter[key]) || "";
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
