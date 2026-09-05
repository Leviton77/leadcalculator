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

  threads.forEach(thread => {
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
