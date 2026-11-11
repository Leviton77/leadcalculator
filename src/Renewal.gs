// Watches for the renewal email. Each time one arrives, closes out the
// current period (per city) and rolls the counts into Renewal Periods.
function detectRenewalAndRollup() {
  const label = getOrCreateLabel_(LABEL_RENEWAL_PROCESSED);
  const query = `from:${RENEWAL_SENDER} subject:"${RENEWAL_SUBJECT}" -label:${LABEL_RENEWAL_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 50).reverse(); // process oldest first

  threads.forEach(thread => {
    const msg = thread.getMessages()[0];
    const parsed = parseInvoiceTotal_(msg.getPlainBody());
    rollupPeriod_(msg.getDate(), parsed === null ? DEFAULT_INVOICE_TOTAL : parsed);
    thread.addLabel(label);
  });
}

function rollupPeriod_(renewalDate, invoiceTotal) {
  const leadsSheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  const periodsSheet = getOrCreateSheet_(SHEET_PERIODS, [
    'Period ID', 'City', 'Period Start', 'Period End',
    'Total Leads', 'Rejected Leads', 'Net Leads',
    'Cost / Lead (w/ tax)', 'Total Cost', 'Invoice Total', 'Match?'
  ]);
  const configSheet = getOrCreateSheet_(SHEET_CONFIG, [
    'City', 'Cost / Lead (before tax)', 'Tax Rate', 'Cost / Lead (with tax)'
  ]);

  const configRows = configSheet.getDataRange().getValues().slice(1);
  const leadsData = leadsSheet.getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();

  configRows.forEach(configRow => {
    const city = configRow[0];
    if (!city) return;
    const costPerLead = configRow[3]; // computed with-tax cost, column D

    const lastKey = 'lastRenewalDate_' + city;
    const lastDateStr = props.getProperty(lastKey);
    const periodStart = lastDateStr ? new Date(lastDateStr) : new Date(0);

    let total = 0;
    let rejected = 0;
    const rowsToStamp = [];

    for (let i = 1; i < leadsData.length; i++) {
      const rowCity = leadsData[i][3];
      const status = leadsData[i][4];
      const alreadyRolledUp = leadsData[i][7];
      if (rowCity !== city || alreadyRolledUp) continue;

      const receivedDate = new Date(leadsData[i][0]);
      if (receivedDate > periodStart && receivedDate <= renewalDate) {
        total++;
        if (status === 'Rejected') rejected++;
        rowsToStamp.push(i + 1);
      }
    }

    // Nothing ever happened for this city and no prior period either —
    // skip so we don't create empty rows for a city that hasn't started.
    if (total === 0 && !lastDateStr) return;

    const net = total - rejected;
    const totalCost = (typeof costPerLead === 'number') ? net * costPerLead : '';
    const periodId = city + '_' + Utilities.formatDate(
      renewalDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    // Match? is filled in below, once every city for this renewal is written.
    periodsSheet.appendRow([
      periodId, city, periodStart, renewalDate,
      total, rejected, net,
      costPerLead, totalCost, invoiceTotal, ''
    ]);

    rowsToStamp.forEach(r => leadsSheet.getRange(r, 8).setValue(periodId));
    props.setProperty(lastKey, renewalDate.toISOString());
  });

  refreshRenewalMatch_(periodsSheet, renewalDate);
}

// One renewal email = one invoice covering every city rolled up by it, so
// the check compares the SUM of those cities' Total Cost against the
// invoice, and writes the same YES / CHECK on each of their rows. Left
// blank while any city is missing its cost in Config.
function refreshRenewalMatch_(periodsSheet, periodEnd) {
  const data = periodsSheet.getDataRange().getValues();
  const rows = [];
  let sum = 0;
  let complete = true;
  let invoice = '';

  for (let i = 1; i < data.length; i++) {
    if (!sameInstant_(data[i][3], periodEnd)) continue;
    rows.push(i + 1);
    if (typeof data[i][8] === 'number') sum += data[i][8];
    else complete = false;
    invoice = data[i][9];
  }

  const match = (!complete || typeof invoice !== 'number')
    ? ''
    : (Math.abs(sum - invoice) < 0.01 ? 'YES' : 'CHECK');
  rows.forEach(r => periodsSheet.getRange(r, 11).setValue(match));
}

// Recomputes a period's counts and cost from the Leads Log. Used when a
// RETOUR arrives after the lead was already rolled into this period.
function recountPeriod_(periodId) {
  const leadsSheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  const periodsSheet = getOrCreateSheet_(SHEET_PERIODS, [
    'Period ID', 'City', 'Period Start', 'Period End',
    'Total Leads', 'Rejected Leads', 'Net Leads',
    'Cost / Lead (w/ tax)', 'Total Cost', 'Invoice Total', 'Match?'
  ]);

  let total = 0;
  let rejected = 0;
  leadsSheet.getDataRange().getValues().slice(1).forEach(row => {
    if (row[7] !== periodId) return;
    total++;
    if (row[4] === 'Rejected') rejected++;
  });

  const periods = periodsSheet.getDataRange().getValues();
  for (let i = 1; i < periods.length; i++) {
    if (periods[i][0] !== periodId) continue;
    const costPerLead = periods[i][7];
    const net = total - rejected;
    const totalCost = (typeof costPerLead === 'number') ? net * costPerLead : '';
    periodsSheet.getRange(i + 1, 5, 1, 3).setValues([[total, rejected, net]]);
    periodsSheet.getRange(i + 1, 9).setValue(totalCost);
    refreshRenewalMatch_(periodsSheet, periods[i][3]);
    return;
  }
}

// Pulls the invoice amount out of the renewal email, e.g. "Total : 574,88 $"
// or "Total: $574.88". Takes the LAST "total" amount so a grand total wins
// over a "Sous-total" above it. Returns null if nothing looks like an
// amount, in which case DEFAULT_INVOICE_TOTAL is used instead.
function parseInvoiceTotal_(body) {
  if (!body) return null;
  const re = /total[^0-9\n]{0,40}?(\d{1,3}(?:[  ,.]\d{3})*[.,]\d{2})(?!\d)/gi;
  let m;
  let last = null;
  while ((m = re.exec(body)) !== null) last = m[1];
  if (last === null) return null;

  const cleaned = last.replace(/[  ]/g, '');
  const intPart = cleaned.slice(0, -3).replace(/[.,]/g, '');
  const amount = Number(intPart + '.' + cleaned.slice(-2));
  return isFinite(amount) && amount > 0 ? amount : null;
}

function sameInstant_(a, b) {
  if (!a || !b || typeof a.getTime !== 'function' || typeof b.getTime !== 'function') return false;
  // Sheets can drop sub-second precision on stored dates.
  return Math.abs(a.getTime() - b.getTime()) < 1000;
}
