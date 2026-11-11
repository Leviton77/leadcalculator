function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// Thread ID lives in column B of the Leads Log sheet.
function findRowByThreadId_(sheet, threadId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === threadId) return i + 1; // 1-based sheet row
  }
  return -1;
}

// Runs fn only if no other run of this script is in progress (e.g. the
// 15-minute trigger and a manual "Run now" at the same moment), so the
// same lead can't be logged twice. Waits up to 30s, then skips.
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn('Leads Calculator: another run is in progress, skipping this one.');
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('Another run is in progress. Try again in a minute.');
    } catch (e) {
      // No UI (time-based trigger) — the log line above is enough.
    }
    return;
  }
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}
