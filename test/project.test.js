const test = require('node:test');
const assert = require('node:assert');
const { loadProject, makeThread, makeMessage } = require('./fakes');

const LEAD = 'leads@soumissionsdemenageurs.ca';
const ME = 'Tom <tom@tommoving.ca>';
const ADMIN = 'admin@soumissionsdemenageurs.ca';
const SUBJECT = 'Renouvellement automatique bloc de soumissions';

function setUp() {
  const p = loadProject();
  p.ctx.setup();
  // Config: Gatineau 10 + 14.975% tax, Ottawa 20 + 13%.
  const cfg = p.sheet('Config');
  cfg.getRange(2, 2).setValue(10); cfg.getRange(2, 4).setValue(10 * 1.14975);
  cfg.getRange(3, 2).setValue(20); cfg.getRange(3, 4).setValue(20 * 1.13);
  return p;
}

const lead = (id, date, extra = [], labels = []) =>
  makeThread(id, [makeMessage({ from: LEAD, date: new Date(date) }), ...extra], labels);
const retour = date => makeMessage({ from: ME, date: new Date(date), body: 'RETOUR mauvais numéro' });
const renewal = (id, date, body = '') =>
  makeThread(id, [makeMessage({ from: ADMIN, date: new Date(date), subject: SUBJECT, body })]);

// Rows are built inside the vm context; copy into this realm for deepStrictEqual.
const periods = p => p.sheet('Renewal Periods').rows.slice(1).map(r => [...r]);
const leads = p => p.sheet('Leads Log').rows.slice(1).map(r => [...r]);

test('late RETOUR updates the period it was already counted in', () => {
  const p = setUp();
  const t1 = lead('t1', '2026-10-01');
  p.threads.push(t1, lead('t2', '2026-10-02'), renewal('r1', '2026-10-10'));
  p.ctx.runAll();
  assert.deepStrictEqual(periods(p)[0].slice(4, 7), [2, 0, 2]);

  t1.messages.push(retour('2026-10-12'));
  p.ctx.runAll();
  const row = periods(p)[0];
  assert.deepStrictEqual(row.slice(4, 7), [2, 1, 1]);
  assert.ok(Math.abs(row[8] - 10 * 1.14975) < 1e-9);
});

test('invoice check compares the combined cost of all cities to one invoice', () => {
  const p = setUp();
  p.threads.push(
    lead('g1', '2026-10-01'),
    lead('o1', '2026-10-02', [], ['LeadsCalc/City-Ottawa']),
    renewal('r1', '2026-10-10', 'Sous-total : 30,00 $\nTPS...\nTotal : 34,10 $'),
  );
  p.ctx.runAll();
  const rows = periods(p);
  assert.strictEqual(rows.length, 2);
  rows.forEach(r => {
    assert.strictEqual(r[9], 34.10);   // amount read from the email
    assert.strictEqual(r[10], 'YES');  // 11.4975 + 22.60 = 34.0975, within a cent
  });
});

test('invoice mismatch flags every city row of that renewal', () => {
  const p = setUp();
  p.threads.push(
    lead('g1', '2026-10-01'),
    lead('o1', '2026-10-02', [], ['LeadsCalc/City-Ottawa']),
    renewal('r1', '2026-10-10', 'Total: $574.88'),
  );
  p.ctx.runAll();
  assert.deepStrictEqual(periods(p).map(r => r[10]), ['CHECK', 'CHECK']);
});

test('falls back to the default invoice when the email has no amount', () => {
  const p = setUp();
  p.threads.push(lead('g1', '2026-10-01'), renewal('r1', '2026-10-10', 'Votre bloc a été renouvelé.'));
  p.ctx.runAll();
  assert.strictEqual(periods(p)[0][9], 574.88);
});

test('parses common invoice amount formats', () => {
  const { ctx } = loadProject();
  assert.strictEqual(ctx.parseInvoiceTotal_('Total : 574,88 $'), 574.88);
  assert.strictEqual(ctx.parseInvoiceTotal_('Total (taxes incluses): 1 234,56 $'), 1234.56);
  assert.strictEqual(ctx.parseInvoiceTotal_('TOTAL $1,234.56'), 1234.56);
  assert.strictEqual(ctx.parseInvoiceTotal_('Sous-total 500.00\nTotal 574.88'), 574.88);
  assert.strictEqual(ctx.parseInvoiceTotal_('Total de 20 soumissions'), null);
  assert.strictEqual(ctx.parseInvoiceTotal_(''), null);
});

test('Ottawa label added after the lead was logged still counts it as Ottawa', () => {
  const p = setUp();
  const t = lead('o1', '2026-10-01');
  p.threads.push(t);
  p.ctx.runAll();
  assert.strictEqual(leads(p)[0][3], 'Gatineau');

  t.labels.add('LeadsCalc/City-Ottawa');
  p.ctx.runAll();
  assert.strictEqual(leads(p)[0][3], 'Ottawa');

  p.threads.push(renewal('r1', '2026-10-10'));
  p.ctx.runAll();
  assert.deepStrictEqual(periods(p).map(r => [r[1], r[4]]), [['Ottawa', 1]]);

  // After rollup the city is frozen, even if the label is removed.
  t.labels.delete('LeadsCalc/City-Ottawa');
  p.ctx.runAll();
  assert.strictEqual(leads(p)[0][3], 'Ottawa');
});

test('a run is skipped while another holds the lock', () => {
  const p = setUp();
  p.threads.push(lead('g1', '2026-10-01'));
  p.holdLock();
  p.ctx.runAll();
  assert.strictEqual(leads(p).length, 0);
});

test('a thread already in the log is not logged again', () => {
  const p = setUp();
  const t = lead('g1', '2026-10-01');
  p.threads.push(t);
  p.ctx.runAll();
  t.labels.delete('LeadsCalc/Lead-Processed'); // e.g. label lost or removed by hand
  p.ctx.runAll();
  assert.strictEqual(leads(p).length, 1);
});
