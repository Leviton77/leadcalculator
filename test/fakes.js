// Minimal in-memory stand-ins for the Apps Script services the project uses,
// so the .gs files can run under plain Node (`npm test`).
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function makeSheet(name) {
  const rows = [];
  const cell = (r, c) => (rows[r - 1] || [])[c - 1];
  const set = (r, c, v) => {
    while (rows.length < r) rows.push([]);
    rows[r - 1][c - 1] = v;
  };
  const width = () => rows.reduce((w, r) => Math.max(w, r.length), 0);
  const range = (r, c, nr = 1, nc = 1) => ({
    getValues: () => Array.from({ length: nr }, (_, i) =>
      Array.from({ length: nc }, (_, j) => { const v = cell(r + i, c + j); return v === undefined ? '' : v; })),
    setValue: v => set(r, c, v),
    setValues: vals => vals.forEach((row, i) => row.forEach((v, j) => set(r + i, c + j, v))),
    clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) set(r + i, c + j, ''); },
    setWrapStrategy: () => {},
  });
  return {
    name, rows,
    getLastRow: () => rows.length,
    getLastColumn: () => width(),
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
    appendRow: r => rows.push(r.slice()),
    setFrozenRows: () => {},
    setRowHeights: () => {},
    getRange: range,
    getDataRange: () => range(1, 1, Math.max(rows.length, 1), Math.max(width(), 1)),
  };
}

function makeThread(id, messages, labels = []) {
  const t = {
    id, messages, labels: new Set(labels),
    getId: () => id,
    getMessages: () => messages,
    getLabels: () => [...t.labels].map(n => ({ getName: () => n })),
    addLabel: l => t.labels.add(l.getName()),
    removeLabel: l => t.labels.delete(l.getName()),
  };
  return t;
}

function makeMessage({ from, date, body = '', subject = '' }) {
  return {
    getFrom: () => from, getDate: () => date, getPlainBody: () => body,
    getSubject: () => subject, getId: () => 'm_' + Math.random().toString(36).slice(2),
  };
}

// Understands just the query shapes the project builds.
function matches(thread, query) {
  const first = thread.messages[0];
  const terms = query.match(/-?\w+:("[^"]*"|\S+)/g) || [];
  return terms.every(term => {
    const neg = term.startsWith('-');
    const [key, ...rest] = term.replace(/^-/, '').split(':');
    const val = rest.join(':').replace(/"/g, '');
    let ok;
    if (key === 'from') ok = first.getFrom().includes(val);
    else if (key === 'to') ok = true;
    else if (key === 'subject') ok = first.getSubject().includes(val);
    else if (key === 'label') ok = thread.labels.has(val);
    else throw new Error('unsupported query term ' + term);
    return neg ? !ok : ok;
  });
}

function loadProject() {
  const sheets = {};
  const threads = [];
  const props = {};
  let lockHeld = false;

  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = makeSheet(n)),
    toast: () => {},
  };

  const ctx = {
    console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, WrapStrategy: { CLIP: 'CLIP' } },
    GmailApp: {
      search: (q, start, max) => threads.filter(t => matches(t, q)).slice(start, start + max),
      getThreadById: id => threads.find(t => t.id === id) || null,
      getUserLabelByName: n => ({ getName: () => n }),
      createLabel: n => ({ getName: () => n }),
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (k in props ? props[k] : null),
        setProperty: (k, v) => { props[k] = v; },
        deleteProperty: k => { delete props[k]; },
        getKeys: () => Object.keys(props),
      }),
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => (lockHeld ? false : (lockHeld = true)),
        releaseLock: () => { lockHeld = false; },
      }),
    },
    Utilities: { formatDate: d => d.toISOString().slice(0, 10) },
    Session: { getScriptTimeZone: () => 'UTC' },
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => {} }) }) }) },
  };
  vm.createContext(ctx);

  const srcDir = path.join(__dirname, '..', 'src');
  const code = fs.readdirSync(srcDir).filter(f => f.endsWith('.gs')).sort()
    .map(f => fs.readFileSync(path.join(srcDir, f), 'utf8')).join('\n');
  vm.runInContext(code, ctx);

  return {
    ctx, sheets, threads, props,
    holdLock: () => { lockHeld = true; },
    sheet: n => sheets[n],
  };
}

module.exports = { loadProject, makeThread, makeMessage };
