// Run this once from the Apps Script editor (or via the "Leads Calculator"
// menu in the sheet) after the first push. It is safe to run again any time
// — it only creates things that are missing.
function setup() {
  const leadsSheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  lockRowHeight_(leadsSheet);

  const periodsSheet = getOrCreateSheet_(SHEET_PERIODS, [
    'Period ID', 'City', 'Period Start', 'Period End',
    'Total Leads', 'Rejected Leads', 'Net Leads',
    'Cost / Lead (w/ tax)', 'Total Cost', 'Invoice Total', 'Match?'
  ]);
  lockRowHeight_(periodsSheet);

  const configSheet = getOrCreateSheet_(SHEET_CONFIG, [
    'City', 'Cost / Lead (before tax)', 'Tax Rate', 'Cost / Lead (with tax)'
  ]);
  if (configSheet.getLastRow() === 1) {
    // Fill in column B (before-tax cost) once you have real numbers.
    // Tax rates default to QC (Gatineau) and ON (Ottawa) — adjust if needed.
    configSheet.appendRow(['Gatineau', '', 0.14975, '=IF(B2="","",B2*(1+C2))']);
    configSheet.appendRow(['Ottawa', '', 0.13, '=IF(B3="","",B3*(1+C3))']);
  }

  getOrCreateLabel_(LABEL_LEAD_PROCESSED);
  getOrCreateLabel_(LABEL_RETOUR_PROCESSED);
  getOrCreateLabel_(LABEL_RENEWAL_PROCESSED);
  getOrCreateLabel_(OTTAWA_LABEL);

  ensureTrigger_();

  SpreadsheetApp.getActiveSpreadsheet().toast('Setup complete. Leads Calculator is ready.');
}

// Prevents long cell text (e.g. the Retour Note) from auto-expanding row
// height — clips overflow instead, and pins every row (current + future,
// up to the sheet's max) to the standard 21px height.
function lockRowHeight_(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxRows, maxCols).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.setRowHeights(1, maxRows, 21);
}

function ensureTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === 'runAll'
  );
  if (!exists) {
    ScriptApp.newTrigger('runAll').timeBased().everyMinutes(TRIGGER_INTERVAL_MINUTES).create();
  }
}
