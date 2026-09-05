function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Leads Calculator')
    .addItem('Run now', 'runAll')
    .addItem('Setup (first time / re-check triggers)', 'setup')
    .addItem('Reset rollups (recompute from scratch)', 'resetRollups')
    .addItem('Full reset (testing only — wipes all data)', 'resetEverything')
    .addToUi();
}

function runAll() {
  ingestLeads();
  detectRetours();
  detectRenewalAndRollup();
}
