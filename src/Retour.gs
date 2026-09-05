// Checks every still-Pending lead's own thread for a reply from you
// containing "RETOUR", and marks it Rejected. Iterates Leads Log directly
// (rather than a combined Gmail search) because Gmail's "from:X keyword"
// search requires a single message to match both terms — it won't match a
// thread where the lead came from X and your RETOUR reply is a separate
// message from you.
function detectRetours() {
  const sheet = getOrCreateSheet_(SHEET_LEADS, [
    'Date Received', 'Thread ID', 'Message ID', 'City', 'Status',
    'Retour Date', 'Retour Note', 'Renewal Period ID'
  ]);
  const retourLabel = getOrCreateLabel_(LABEL_RETOUR_PROCESSED);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][4] !== 'Pending') continue;

    const threadId = data[i][1];
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) continue;

    const retourMsg = thread.getMessages().find(m =>
      m.getFrom().toLowerCase().indexOf(MY_EMAIL.toLowerCase()) !== -1 &&
      m.getPlainBody().toUpperCase().indexOf(RETOUR_KEYWORD) !== -1
    );
    if (!retourMsg) continue;

    const row = i + 1;
    sheet.getRange(row, 5).setValue('Rejected');
    sheet.getRange(row, 6).setValue(retourMsg.getDate());
    sheet.getRange(row, 7).setValue(retourMsg.getPlainBody().substring(0, 300));

    thread.addLabel(retourLabel);
  }
}
