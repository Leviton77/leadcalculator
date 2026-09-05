// Watches for the renewal email. Each time one arrives, closes out the
// current period (per city) and rolls the counts into Renewal Periods.
function detectRenewalAndRollup() {
  const label = getOrCreateLabel_(LABEL_RENEWAL_PROCESSED);
  const query = `from:${RENEWAL_SENDER} subject:"${RENEWAL_SUBJECT}" -label:${LABEL_RENEWAL_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 50).reverse(); // process oldest first

  threads.forEach(thread => {
    const msg = thread.getMessages()[0];
    rollupPeriod_(msg.getDate());
    thread.addLabel(label);
  });
}

function rollupPeriod_(renewalDate) {
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
    const matchResult = totalCost === ''
      ? ''
      : (Math.abs(totalCost - DEFAULT_INVOICE_TOTAL) < 0.01 ? 'YES' : 'CHECK');

    periodsSheet.appendRow([
      periodId, city, periodStart, renewalDate,
      total, rejected, net,
      costPerLead, totalCost, DEFAULT_INVOICE_TOTAL, matchResult
    ]);

    rowsToStamp.forEach(r => leadsSheet.getRange(r, 8).setValue(periodId));
    props.setProperty(lastKey, renewalDate.toISOString());
  });
}
