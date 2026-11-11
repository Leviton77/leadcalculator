// Logs any new lead email as a row in the Leads Log sheet, then labels the
// thread so it's never counted twice.
function ingestLeads() {
  const sheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  const processedLabel = getOrCreateLabel_(LABEL_LEAD_PROCESSED);

  const query = `from:${LEAD_SENDER} to:${MY_EMAIL} -label:${LABEL_LEAD_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 100);

  // Belt and braces against double-logging: skip threads already in the log
  // even if their processed label didn't stick.
  const loggedThreadIds = {};
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    if (row[1]) loggedThreadIds[row[1]] = true;
  });

  threads.forEach(thread => {
    if (loggedThreadIds[thread.getId()]) {
      thread.addLabel(processedLabel);
      return;
    }
    const firstMsg = thread.getMessages()[0];
    const isOttawa = thread.getLabels().some(l => l.getName() === OTTAWA_LABEL);
    const city = isOttawa ? 'Ottawa' : DEFAULT_CITY;

    sheet.appendRow([
      firstMsg.getDate(),
      thread.getId(),
      firstMsg.getId(),
      city,
      'Pending',
      '', '', ''
    ]);

    thread.addLabel(processedLabel);
  });
}

// Re-reads the Ottawa label for every lead not yet rolled into a renewal
// period, so the label can be added (or removed) any time before the next
// renewal instead of only in the 15 minutes before the lead is logged.
// Leads already rolled up are left alone.
function refreshCities() {
  const sheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);

  const ottawaThreadIds = {};
  const pageSize = 100;
  for (let start = 0; start < 2000; start += pageSize) {
    const page = GmailApp.search(`label:${OTTAWA_LABEL}`, start, pageSize);
    page.forEach(t => { ottawaThreadIds[t.getId()] = true; });
    if (page.length < pageSize) break;
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const city = data[i][3];
    if (data[i][7]) continue; // already rolled up
    if (city !== 'Ottawa' && city !== DEFAULT_CITY) continue; // hand-edited, leave it
    const wanted = ottawaThreadIds[data[i][1]] ? 'Ottawa' : DEFAULT_CITY;
    if (city !== wanted) sheet.getRange(i + 1, 4).setValue(wanted);
  }
}
