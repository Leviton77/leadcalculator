// Undoes all renewal rollups so the next "Run now" recomputes Renewal
// Periods from scratch. Useful after fixing a bug (like the RETOUR
// detection issue) that affected numbers already rolled up. Does NOT touch
// Leads Log Status (Pending/Rejected) — only the rollup bookkeeping.
function resetRollups() {
  withScriptLock_(resetRollupsLocked_);
}

function resetRollupsLocked_() {
  const periodsSheet = getOrCreateSheet_(SHEET_PERIODS, [
    'Period ID', 'City', 'Period Start', 'Period End',
    'Total Leads', 'Rejected Leads', 'Net Leads',
    'Cost / Lead (w/ tax)', 'Total Cost', 'Invoice Total', 'Match?'
  ]);
  if (periodsSheet.getLastRow() > 1) {
    periodsSheet.getRange(2, 1, periodsSheet.getLastRow() - 1, periodsSheet.getLastColumn()).clearContent();
  }

  const leadsSheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  if (leadsSheet.getLastRow() > 1) {
    leadsSheet.getRange(2, 8, leadsSheet.getLastRow() - 1, 1).clearContent();
  }

  const props = PropertiesService.getScriptProperties();
  props.getKeys()
    .filter(k => k.indexOf('lastRenewalDate_') === 0)
    .forEach(k => props.deleteProperty(k));

  const renewalLabel = getOrCreateLabel_(LABEL_RENEWAL_PROCESSED);
  GmailApp.search(`label:${LABEL_RENEWAL_PROCESSED}`, 0, 100)
    .forEach(thread => thread.removeLabel(renewalLabel));

  SpreadsheetApp.getActiveSpreadsheet().toast('Rollups reset. Click "Run now" to recompute.');
}

// TESTING ONLY — wipes all logged data (Leads Log + Renewal Periods) and
// un-labels every Gmail thread the script has touched, so the next "Run
// now" rediscovers everything from scratch as if it were the first run.
// Safe while there's no real cost data riding on the sheet yet; do not use
// once this is tracking real invoices, since it forgets which leads were
// already reconciled.
function resetEverything() {
  withScriptLock_(resetEverythingLocked_);
}

function resetEverythingLocked_() {
  const leadsSheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  if (leadsSheet.getLastRow() > 1) {
    leadsSheet.getRange(2, 1, leadsSheet.getLastRow() - 1, leadsSheet.getLastColumn()).clearContent();
  }

  const periodsSheet = getOrCreateSheet_(SHEET_PERIODS, [
    'Period ID', 'City', 'Period Start', 'Period End',
    'Total Leads', 'Rejected Leads', 'Net Leads',
    'Cost / Lead (w/ tax)', 'Total Cost', 'Invoice Total', 'Match?'
  ]);
  if (periodsSheet.getLastRow() > 1) {
    periodsSheet.getRange(2, 1, periodsSheet.getLastRow() - 1, periodsSheet.getLastColumn()).clearContent();
  }

  const props = PropertiesService.getScriptProperties();
  props.getKeys()
    .filter(k => k.indexOf('lastRenewalDate_') === 0)
    .forEach(k => props.deleteProperty(k));

  [LABEL_LEAD_PROCESSED, LABEL_RETOUR_PROCESSED, LABEL_RENEWAL_PROCESSED].forEach(labelName => {
    const label = getOrCreateLabel_(labelName);
    GmailApp.search(`label:${labelName}`, 0, 100)
      .forEach(thread => thread.removeLabel(label));
  });

  lockRowHeight_(leadsSheet);
  lockRowHeight_(periodsSheet);

  SpreadsheetApp.getActiveSpreadsheet().toast('Full reset done. Click "Run now" to rebuild from scratch.');
}
