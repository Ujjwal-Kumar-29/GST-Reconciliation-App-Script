/**
 * GST RECONCILIATION - Purchase Register vs GSTR-2B
 * Creates a Google Sheet named "GST reconciliation" with 3 tabs:
 *   purchase_register | GSTR2B | result
 * fills dummy data, and reconciles the two.
 *
 * HOW TO RUN:
 *  1. Go to script.google.com > New project, paste this whole code, Save.
 *  2. Select  setupGSTReconciliation  in the function dropdown > Run.
 *  3. Allow permissions when asked. Open the link shown in View > Logs
 *     (or find "GST reconciliation" in your Google Drive).
 *  4. Later, use the menu  GST Tools > Run Reconciliation  inside that sheet
 *     (menu works if this code is also pasted in Extensions > Apps Script of that sheet).
 */

const SHEET_PR  = 'purchase_register';
const SHEET_2B  = 'GSTR2B';
const SHEET_RES = 'result';
const FILE_NAME = 'GST reconciliation';
const TOLERANCE = 1; // rupees; differences up to this are treated as matched

const HEADERS = ['GSTIN', 'Supplier', 'Invoice No', 'Date', 'Taxable Value', 'IGST', 'CGST', 'SGST'];

// Indian number system (lakh / crore style grouping)
const INR_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';

/* ------------------------------------------------------------------ */
/*  MENU                                                               */
/* ------------------------------------------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GST Tools')
    .addItem('Set up sheet + dummy data', 'setupGSTReconciliation')
    .addItem('Run Reconciliation', 'reconcile')
    .addToUi();
}

/* ------------------------------------------------------------------ */
/*  SETUP: create sheet, tabs, dummy data                              */
/* ------------------------------------------------------------------ */
function setupGSTReconciliation() {
  let ss = SpreadsheetApp.getActiveSpreadsheet(); // null when run from a standalone script
  if (!ss) {
    ss = SpreadsheetApp.create(FILE_NAME);
    PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  } else {
    ss.rename(FILE_NAME);
  }

  const pr  = getOrCreateSheet_(ss, SHEET_PR);
  const b2  = getOrCreateSheet_(ss, SHEET_2B);
  const res = getOrCreateSheet_(ss, SHEET_RES);

  // remove the default empty "Sheet1" if it is still there
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(def);

  writeInputSheet_(pr, getPurchaseRegisterData_());
  writeInputSheet_(b2, getGSTR2BData_());

  // order tabs
  ss.setActiveSheet(pr);  ss.moveActiveSheet(1);
  ss.setActiveSheet(b2);  ss.moveActiveSheet(2);
  ss.setActiveSheet(res); ss.moveActiveSheet(3);

  reconcile();
  ss.setActiveSheet(res);

  Logger.log('Sheet ready: ' + ss.getUrl());
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeInputSheet_(sheet, rows) {
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@'); // invoice no as text
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.getRange(2, 4, rows.length, 1).setNumberFormat('dd-mm-yyyy');
  sheet.getRange(2, 5, rows.length, 4).setNumberFormat(INR_FMT);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

/* ------------------------------------------------------------------ */
/*  DUMMY DATA (buyer assumed in Maharashtra, state code 27)           */
/*  Cols: GSTIN, Supplier, Invoice No, Date, Taxable, IGST, CGST, SGST */
/* ------------------------------------------------------------------ */
function getPurchaseRegisterData_() {
  return [
    // Perfect match
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV/002',      new Date(2026, 3, 12),  50000, 0,     4500, 4500],
    ['33AABCT1332L1ZY', 'Chennai Packaging',    'CP-88',        new Date(2026, 4, 21),  60000, 10800, 0,    0],
    ['29AABCU9603R1ZX', 'Bright Electronics',   'BE-2026-0045', new Date(2026, 3, 18), 200000, 36000, 0,    0],
    // Invoice format differs (books INV/001 vs 2B INV-1, etc.)
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV/001',      new Date(2026, 3, 5),  100000, 0,     9000, 9000],
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV/003',      new Date(2026, 5, 3),   75000, 0,     6750, 6750],
    ['29AABCU9603R1ZX', 'Bright Electronics',   'BE-2026-0051', new Date(2026, 4, 2),   80000, 14400, 0,    0],
    // In books but missing in 2B
    ['07AAACR5055K1Z8', 'Delhi Stationers',     'DS/0007',      new Date(2026, 4, 9),   25000, 4500,  0,    0],
    // Tax amount differs (books IGST 27,000 vs 2B IGST 24,300)
    ['24AAACC1206D1ZM', 'Gujarat Chemicals',    'GC/101',       new Date(2026, 4, 15), 150000, 27000, 0,    0]
  ];
}

function getGSTR2BData_() {
  return [
    // Perfect match
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV/002',      new Date(2026, 3, 12),  50000, 0,     4500, 4500],
    ['33AABCT1332L1ZY', 'Chennai Packaging',    'CP-88',        new Date(2026, 4, 21),  60000, 10800, 0,    0],
    ['29AABCU9603R1ZX', 'Bright Electronics',   'BE-2026-0045', new Date(2026, 3, 18), 200000, 36000, 0,    0],
    // Different invoice formats
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV-1',        new Date(2026, 3, 5),  100000, 0,     9000, 9000],
    ['27AAAPL1234C1ZV', 'Sharma Traders',       'INV-3',        new Date(2026, 5, 3),   75000, 0,     6750, 6750],
    ['29AABCU9603R1ZX', 'Bright Electronics',   'BE/2026/0051', new Date(2026, 4, 2),   80000, 14400, 0,    0],
    // Tax amount differs
    ['24AAACC1206D1ZM', 'Gujarat Chemicals',    'GC/101',       new Date(2026, 4, 15), 150000, 24300, 0,    0],
    // In 2B but not booked
    ['27AABCM5678F1ZP', 'Mehta Logistics',      'ML/445',       new Date(2026, 5, 10),  40000, 0,     3600, 3600]
  ];
}

/* ------------------------------------------------------------------ */
/*  RECONCILIATION                                                     */
/* ------------------------------------------------------------------ */
function reconcile() {
  const ss = getSpreadsheet_();
  if (!ss) { Logger.log('Spreadsheet not found. Run setupGSTReconciliation first.'); return; }

  const prSheet = ss.getSheetByName(SHEET_PR);
  const b2Sheet = ss.getSheetByName(SHEET_2B);
  const resSheet = ss.getSheetByName(SHEET_RES);
  if (!prSheet || !b2Sheet || !resSheet) { Logger.log('One or more tabs are missing.'); return; }

  const pr = prSheet.getDataRange().getValues().slice(1).filter(r => r[0] !== '');
  const b2 = b2Sheet.getDataRange().getValues().slice(1).filter(r => r[0] !== '');

  // Invoice numbers like INV/001, INV-1, inv 001 all become the same key
  const normInv = s => {
    let t = String(s).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean)
      .map(p => /^\d+$/.test(p) ? String(parseInt(p, 10)) : p).join('');
    return t.replace(/([A-Z])0+(?=\d)/g, '$1');
  };
  const key = r => String(r[0]).trim().toUpperCase() + '|' + normInv(r[2]);
  const tax = r => (Number(r[5]) || 0) + (Number(r[6]) || 0) + (Number(r[7]) || 0);
  const round2 = n => Math.round(n * 100) / 100;

  const map = {};
  b2.forEach(r => { map[key(r)] = { row: r, used: false }; });

  const header = ['GSTIN', 'Supplier', 'Invoice No (Books)', 'Invoice No (2B)',
                  'Tax in Books', 'Tax in 2B', 'Difference', 'Status'];
  const out = [];
  const stat = { matched: 0, mismatch: 0, booksOnly: 0, b2Only: 0, itcRisk: 0, itcNotBooked: 0 };

  pr.forEach(r => {
    const m = map[key(r)];
    if (!m) {
      out.push([r[0], r[1], r[2], '-', tax(r), 0, tax(r), 'In books only - ITC at risk']);
      stat.booksOnly++; stat.itcRisk += tax(r);
      return;
    }
    m.used = true;
    const diff = round2(tax(r) - tax(m.row));
    if (Math.abs(diff) <= TOLERANCE) {
      const fmtNote = String(r[2]) !== String(m.row[2]) ? 'Matched (invoice format differs)' : 'Matched';
      out.push([r[0], r[1], r[2], m.row[2], tax(r), tax(m.row), diff, fmtNote]);
      stat.matched++;
    } else {
      out.push([r[0], r[1], r[2], m.row[2], tax(r), tax(m.row), diff, 'Tax amount mismatch']);
      stat.mismatch++;
      if (diff > 0) stat.itcRisk += diff;
    }
  });

  Object.keys(map).forEach(k => {
    const m = map[k];
    if (m.used) return;
    const r = m.row;
    out.push([r[0], r[1], '-', r[2], 0, tax(r), -tax(r), 'In 2B only - not booked']);
    stat.b2Only++; stat.itcNotBooked += tax(r);
  });

  // ----- write result -----
  resSheet.clear();
  resSheet.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  if (out.length) {
    resSheet.getRange(2, 1, out.length, header.length).setValues(out);
    resSheet.getRange(2, 5, out.length, 3).setNumberFormat(INR_FMT);

    const colors = out.map(row => {
      const s = row[7];
      const c = s.indexOf('Matched') === 0 ? '#d9ead3'
              : s.indexOf('mismatch') > -1 ? '#fff2cc'
              : s.indexOf('books only') > -1 ? '#f4cccc'
              : '#cfe2f3';
      return new Array(header.length).fill(c);
    });
    resSheet.getRange(2, 1, out.length, header.length).setBackgrounds(colors);
  }

  // ----- summary block -----
  const summary = [
    ['SUMMARY', ''],
    ['Matched invoices', stat.matched],
    ['Tax amount mismatches', stat.mismatch],
    ['In books only', stat.booksOnly],
    ['In 2B only (not booked)', stat.b2Only],
    ['ITC at risk (Rs.)', round2(stat.itcRisk)],
    ['ITC available in 2B but not booked (Rs.)', round2(stat.itcNotBooked)]
  ];
  resSheet.getRange(1, 10, summary.length, 2).setValues(summary);
  resSheet.getRange(1, 10, 1, 2).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  resSheet.getRange(6, 11, 2, 1).setNumberFormat(INR_FMT);
  resSheet.getRange(2, 10, summary.length - 1, 1).setFontWeight('bold');

  resSheet.setFrozenRows(1);
  resSheet.autoResizeColumns(1, 11);
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : null;
}
