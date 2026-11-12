// Treasury Cash Dashboard - server
// Reads every customer workbook (Treasury_Cash_Dashboard_<Customer>.xlsx) from DATA_DIR
// and serves them to the browser. Files are re-read automatically when they change on disk.
const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
let cfg = {};
try { cfg = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8')); } catch (e) {}
const DATA_DIR = path.resolve(ROOT, process.env.DATA_DIR || cfg.dataDir || '.');
const PORT = process.env.PORT || cfg.port || 3000;

const cache = new Map(); // file -> {mtimeMs, data}

function plain(c) {
  let v = c && c.value !== undefined ? c.value : c;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return isNaN(v) ? null : v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if ('result' in v) return plain(v.result);
    if ('richText' in v) return v.richText.map(t => t.text).join('');
    if ('error' in v) return null;
    if ('text' in v) return v.text;
    return null;
  }
  return v;
}
const num = v => (typeof v === 'number' && isFinite(v) ? v : 0);
const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();

function table(ws) {
  const head = [];
  ws.getRow(1).eachCell({ includeEmpty: false }, (c, i) => head[i] = String(plain(c)).trim());
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const first = plain(row.getCell(1));
    if (first === null || first === '') continue;
    const o = {};
    head.forEach((h, i) => { if (h) o[h] = plain(row.getCell(i)); });
    rows.push(o);
  }
  return rows;
}

const WATERFALL = {
  opening: 'opening cash balance',
  recvAdv: 'received: 50% advance',
  recvBal: 'received: 50% balance',
  expAdv: 'expected: 50% advance',
  expBal: 'expected: 50% balance',
  receipts: 'total customer receipts',
  paidAdvActual: 'paid to advance suppliers',
  paidCredActual: 'paid to credit suppliers',
  paidActual: 'total supplier payments already made',
  paidAdvFc: 'advance paid (forecast',
  paidCredFc: 'credit paid (forecast',
  payments: 'total supplier payments (actual + forecast)',
  closing: 'closing cash balance',
  deferred: 'total payables deferred',
  net: 'net cash position (',
};

async function parseWorkbook(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const wsR = wb.getWorksheet('Fact_Receipts');
  const wsP = wb.getWorksheet('Fact_Payables');
  const wsD = wb.getWorksheet('Dashboard');
  if (!wsR || !wsP) throw new Error('missing Fact_Receipts / Fact_Payables sheet');
  const warnings = [];

  const receipts = table(wsR).map(r => {
    const sales = num(r['Sales Value (USD)']);
    const tr = r['Tranche Amount (USD)'];
    const tranche = typeof tr === 'number' ? tr : sales / 2;
    return {
      id: r['Receipt ID'], po: r['PO No'], style: r['Style'], qty: num(r['Quantity']), sales,
      shipment: r['Shipment Date'], tranche: r['Tranche'], scheduled: r['Scheduled Date'],
      amount: tranche, received: num(r['Received Amount (USD)']), receivedDate: r['Received Date'],
      status: r['Status'] || '', pending: num(r['Pending Amount (USD)']),
      daysOverdue: num(r['Days Overdue']), cashMonth: r['Cash Month'], remarks: r['Remarks'],
      customer: r['Customer'] || null,
    };
  });
  const payables = table(wsP).map(r => ({
    id: r['Record ID'], cls: r['Payable Class'], po: r['PO No'], supplier: r['Supplier'],
    term: r['Payment Term'], delivery: r['Delivery Start Date'], due: r['Due Date'],
    value: num(r['PO Value (USD)']), paid: num(r['Paid (USD)']), paidDate: r['Paid Date'],
    outstanding: num(r['Outstanding (USD)']), status: r['Status'] || '',
    daysOverdue: num(r['Days Overdue']), cashMonth: r['Cash Month'],
    include: r['Include in Plan'], request: r['Request Status'], remarks: r['Remarks'],
  }));

  // Working-capital waterfall = the engine on the workbook's Dashboard sheet
  let months = [], phases = [], waterfall = {};
  if (wsD) {
    let hdr = 0;
    for (let r = 1; r <= 40; r++) if (norm(plain(wsD.getCell(r, 2))) === 'row labels') { hdr = r; break; }
    if (hdr) {
      for (let c = 3; c <= 30; c++) {
        const v = plain(wsD.getCell(hdr, c));
        if (typeof v === 'string' && /^\d{4}-\d{2}/.test(v)) months.push({ col: c, date: v.slice(0, 7) });
        else break;
      }
      phases = months.map(m => plain(wsD.getCell(hdr + 1, m.col)));
      for (let r = hdr + 2; r <= hdr + 40; r++) {
        const label = norm(plain(wsD.getCell(r, 2)));
        for (const [k, pat] of Object.entries(WATERFALL)) {
          if (!(k in waterfall) && label.includes(pat))
            waterfall[k] = months.map(m => num(plain(wsD.getCell(r, m.col))));
        }
      }
      if (!waterfall.receipts) warnings.push('Dashboard sheet has no calculated values - open and save the workbook in Excel so values are stored.');
    }
  } else warnings.push('Dashboard sheet not found - working-capital charts unavailable.');

  const asOf = wb.getWorksheet('Assumptions') ? plain(wb.getWorksheet('Assumptions').getCell('C4')) : null;
  const eur = wb.getWorksheet('Assumptions') ? plain(wb.getWorksheet('Assumptions').getCell('C12')) : null;
  return { receipts, payables, months: months.map(m => m.date), phases, waterfall, asOf, eurUsd: eur, warnings };
}

function customerName(file) {
  const base = path.basename(file, path.extname(file));
  const m = base.match(/Treasury[_ ]Cash[_ ]Dashboard[_ -]+(.+)$/i);
  return (m ? m[1] : base).replace(/_/g, ' ').trim();
}

async function load() {
  let files = [];
  try { files = fs.readdirSync(DATA_DIR).filter(f => /\.xlsx$/i.test(f) && !f.startsWith('~$')); } catch (e) {
    return { dataDir: DATA_DIR, customers: [], skipped: [{ file: DATA_DIR, reason: 'folder not readable: ' + e.message }] };
  }
  const customers = [], skipped = [];
  for (const f of files.sort()) {
    const full = path.join(DATA_DIR, f);
    const st = fs.statSync(full);
    let hit = cache.get(full);
    if (!hit || hit.mtimeMs !== st.mtimeMs) {
      try { hit = { mtimeMs: st.mtimeMs, data: await parseWorkbook(full), modified: st.mtime.toISOString() }; }
      catch (e) { hit = { mtimeMs: st.mtimeMs, error: e.message }; }
      cache.set(full, hit);
    }
    if (hit.error) { skipped.push({ file: f, reason: hit.error }); continue; }
    customers.push({ name: customerName(f), file: f, modified: hit.modified, ...hit.data });
  }
  return { dataDir: DATA_DIR, loadedAt: new Date().toISOString(), customers, skipped };
}

const app = express();
app.get('/api/data', async (req, res) => {
  try { res.json(await load()); } catch (e) { res.status(500).json({ error: e.message }); }
});
app.use('/vendor/chart.js', express.static(path.join(ROOT, 'node_modules/chart.js/dist')));
app.use(express.static(path.join(ROOT, 'public')));
if (require.main === module) {
  app.listen(PORT, () => console.log(`Treasury dashboard: http://localhost:${PORT}\nReading workbooks from: ${DATA_DIR}`));
}
module.exports = { load };
