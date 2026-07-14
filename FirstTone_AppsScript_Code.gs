/**
 * First Tone App — Google Apps Script backend
 * Deploy as Web App (Execute as: Me, Who has access: Anyone with the link)
 *
 * Setup:
 *  1. Create a new Google Sheet.
 *  2. Extensions > Apps Script, paste this whole file in as Code.gs.
 *  3. Run setupSpreadsheet() once (top toolbar > select function > Run) to create sheets + seed data.
 *  4. Deploy > New deployment > Web app > Execute as Me / Access: Anyone.
 *  5. Copy the Web app URL into APPS_SCRIPT_URL in FirstTone_App.html and set USE_LOCAL = false.
 *
 * Updating an existing deployment (schema changed — new columns added):
 *  1. Paste the new code over the old Code.gs (Ctrl+A, paste, Ctrl+S).
 *  2. Run migrateAddCategoryColumn() once, then migrateAddItemNoAndTagsColumns() once.
 *     Both are safe to re-run (no-op if the column already exists) and do NOT wipe data.
 *     Do NOT re-run setupSpreadsheet() on a live sheet — it clears existing data.
 *  3. Deploy > Manage deployments > pencil icon on the active deployment > Version: New version > Deploy.
 *     (This keeps the same Web app URL — no need to update FirstTone_App.html.)
 */

const SHEET_NAMES = {
  SETTINGS: 'Settings',
  STAFF: 'Staff',
  TEACHERS: 'Teachers',
  STUDENTS: 'Students',
  ITEMS: 'Items',
  INVOICES: 'Invoices',
  INVOICE_ITEMS: 'InvoiceItems',
  PAYMENTS: 'Payments',
  WRITEOFFS: 'Writeoffs'
};

const SHEET_HEADERS = {
  Settings: ['staffPIN', 'adminPIN', 'lowStockThreshold', 'supplierName', 'supplierWA', 'supplierNotes'],
  Staff: ['id', 'name'],
  Teachers: ['id', 'name', 'notes'],
  Students: ['id', 'name', 'teacherId', 'notes'],
  Items: ['barcode', 'name', 'type', 'category', 'itemNo', 'tags', 'price', 'cost', 'qty', 'alertOn', 'createdAt'],
  Invoices: ['id', 'no', 'date', 'buyerType', 'buyerId', 'teacherId', 'staffId', 'total', 'discount', 'paid', 'status'],
  InvoiceItems: ['invoiceId', 'barcode', 'name', 'originalPrice', 'discounted'],
  Payments: ['invoiceId', 'date', 'amount', 'method'],
  Writeoffs: ['date', 'invoiceNo', 'studentName', 'amount', 'reason']
};

// ── ENTRY POINTS ─────────────────────────────────────────────────────────────

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'getAll';
  try {
    if (action === 'getAll') return jsonOut({ success: true, data: getAllData() });
    return jsonOut({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};
    let result;
    switch (action) {
      case 'getAll': result = getAllData(); break;
      case 'saveInvoice': result = saveInvoice(payload); break;
      case 'updatePayment': result = updatePayment(payload); break;
      case 'saveItem': result = saveItem(payload); break;
      case 'bulkImportItems': result = bulkImportItems(payload); break;
      case 'updateQty': result = updateQty(payload); break;
      case 'savePerson': result = savePerson(payload); break;
      case 'writeoff': result = writeoffInvoice(payload); break;
      case 'cancelInvoice': result = cancelInvoiceBackend(payload); break;
      case 'saveSettings': result = saveSettings(payload); break;
      default: return jsonOut({ success: false, error: 'Unknown action: ' + action });
    }
    return jsonOut({ success: true, data: result });
  } catch (err) {
    return jsonOut({ success: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ───────────────────────────────────────────────────────────────────

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name) { return ss().getSheetByName(name); }

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheetToObjects(name) {
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => { const o = {}; headers.forEach((h, i) => o[h] = r[i]); return o; });
}

function appendRow(sh, rowArray) { sh.appendRow(rowArray); }

function findRowIndexByKey(sh, keyCol, keyVal) {
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][keyCol]) === String(keyVal)) return r + 1; // 1-based sheet row
  }
  return -1;
}

function updateInvoiceFields(id, fields) {
  const sh = sheet(SHEET_NAMES.INVOICES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(id)) {
      Object.keys(fields).forEach(k => {
        const c = headers.indexOf(k);
        if (c > -1) sh.getRange(r + 1, c + 1).setValue(fields[k]);
      });
      return true;
    }
  }
  return false;
}

function adjustItemQtys(deltaByBarcode) {
  const sh = sheet(SHEET_NAMES.ITEMS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const barcodeCol = headers.indexOf('barcode');
  const qtyCol = headers.indexOf('qty');
  for (let r = 1; r < data.length; r++) {
    const bc = String(data[r][barcodeCol]);
    if (Object.prototype.hasOwnProperty.call(deltaByBarcode, bc)) {
      const cur = data[r][qtyCol];
      if (cur !== '' && cur !== null && cur !== undefined) {
        const newQty = Math.max(0, Number(cur) + deltaByBarcode[bc]);
        sh.getRange(r + 1, qtyCol + 1).setValue(newQty);
      }
    }
  }
}

// ── READ: getAll ──────────────────────────────────────────────────────────────

function getAllData() {
  const settingsRows = sheetToObjects(SHEET_NAMES.SETTINGS);
  const s = settingsRows[0] || {};
  const settings = {
    staffPIN: s.staffPIN !== undefined && s.staffPIN !== '' ? String(s.staffPIN) : '1234',
    adminPIN: s.adminPIN !== undefined && s.adminPIN !== '' ? String(s.adminPIN) : '9999',
    lowStockThreshold: Number(s.lowStockThreshold || 5),
    supplier: { name: s.supplierName || '', wa: s.supplierWA || '', notes: s.supplierNotes || '' }
  };

  const staff = sheetToObjects(SHEET_NAMES.STAFF).map(r => ({ id: String(r.id), name: r.name }));
  const teachers = sheetToObjects(SHEET_NAMES.TEACHERS).map(r => ({ id: String(r.id), name: r.name, notes: r.notes || '' }));
  const students = sheetToObjects(SHEET_NAMES.STUDENTS).map(r => ({
    id: String(r.id), name: r.name, teacherId: r.teacherId ? String(r.teacherId) : '', notes: r.notes || ''
  }));

  const itemRows = sheetToObjects(SHEET_NAMES.ITEMS);
  const items = {};
  itemRows.forEach(r => {
    items[String(r.barcode)] = {
      name: r.name,
      type: r.type,
      category: r.category || '',
      itemNo: r.itemNo || '',
      tags: r.tags ? String(r.tags).split(',').map(t => t.trim()).filter(Boolean) : [],
      price: Number(r.price) || 0,
      cost: Number(r.cost) || 0,
      qty: (r.qty === '' || r.qty === null || r.qty === undefined) ? null : Number(r.qty),
      alertOn: r.alertOn === true || r.alertOn === 'TRUE',
      createdAt: Number(r.createdAt) || Date.now()
    };
  });

  const invoiceRows = sheetToObjects(SHEET_NAMES.INVOICES);
  const invItemRows = sheetToObjects(SHEET_NAMES.INVOICE_ITEMS);
  const paymentRows = sheetToObjects(SHEET_NAMES.PAYMENTS);
  const invoices = invoiceRows.map(r => {
    const id = String(r.id);
    return {
      id,
      no: r.no,
      date: Number(r.date),
      buyerType: r.buyerType,
      buyerId: String(r.buyerId),
      teacherId: r.teacherId ? String(r.teacherId) : '',
      staffId: String(r.staffId),
      total: Number(r.total) || 0,
      discount: Number(r.discount) || 0,
      paid: Number(r.paid) || 0,
      status: r.status,
      items: invItemRows.filter(ii => String(ii.invoiceId) === id).map(ii => ({
        barcode: String(ii.barcode), name: ii.name,
        originalPrice: Number(ii.originalPrice) || 0, discounted: Number(ii.discounted) || 0
      })),
      payments: paymentRows.filter(p => String(p.invoiceId) === id).map(p => ({
        date: Number(p.date), amount: Number(p.amount) || 0, method: p.method
      }))
    };
  });

  const writeoffs = sheetToObjects(SHEET_NAMES.WRITEOFFS).map(r => ({
    date: Number(r.date), invoiceNo: r.invoiceNo, studentName: r.studentName,
    amount: Number(r.amount) || 0, reason: r.reason || ''
  }));

  return { settings, staff, teachers, students, items, invoices, writeoffs };
}

// ── WRITE ACTIONS ─────────────────────────────────────────────────────────────

function saveInvoice(inv) {
  return withLock(() => {
    appendRow(sheet(SHEET_NAMES.INVOICES), [
      inv.id, inv.no, inv.date, inv.buyerType, inv.buyerId,
      inv.teacherId || '', inv.staffId, inv.total, inv.discount || 0, inv.paid || 0, inv.status
    ]);
    const itemSh = sheet(SHEET_NAMES.INVOICE_ITEMS);
    (inv.items || []).forEach(it => appendRow(itemSh, [inv.id, it.barcode, it.name, it.originalPrice, it.discounted]));
    if (inv.payments && inv.payments.length) {
      const paySh = sheet(SHEET_NAMES.PAYMENTS);
      inv.payments.forEach(p => appendRow(paySh, [inv.id, p.date, p.amount, p.method]));
    }
    const delta = {};
    (inv.items || []).forEach(it => { delta[it.barcode] = (delta[it.barcode] || 0) - 1; });
    adjustItemQtys(delta);
    return { id: inv.id };
  });
}

function updatePayment(p) {
  return withLock(() => {
    appendRow(sheet(SHEET_NAMES.PAYMENTS), [p.invoiceId, p.date, p.amount, p.method]);
    updateInvoiceFields(p.invoiceId, { paid: p.newPaid, status: p.newStatus });
    return { ok: true };
  });
}

function cancelInvoiceBackend(payload) {
  return withLock(() => {
    const id = payload.id;
    updateInvoiceFields(id, { status: 'cancelled' });
    const invItems = sheetToObjects(SHEET_NAMES.INVOICE_ITEMS).filter(ii => String(ii.invoiceId) === String(id));
    const delta = {};
    invItems.forEach(ii => { delta[String(ii.barcode)] = (delta[String(ii.barcode)] || 0) + 1; });
    adjustItemQtys(delta);
    return { ok: true };
  });
}

function writeoffInvoice(payload) {
  return withLock(() => {
    appendRow(sheet(SHEET_NAMES.WRITEOFFS), [
      payload.date, payload.invoiceNo, payload.studentName, payload.amount, payload.reason || 'Bad debt'
    ]);
    updateInvoiceFields(payload.invoiceId, { status: 'writeoff' });
    return { ok: true };
  });
}

function saveItem(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.ITEMS);
    const headers = SHEET_HEADERS.Items;
    const barcodeCol = headers.indexOf('barcode');

    if (payload.deleted) {
      const row = findRowIndexByKey(sh, barcodeCol, payload.barcode);
      if (row > -1) sh.deleteRow(row);
      return { ok: true };
    }

    if (payload.oldBarcode && payload.oldBarcode !== payload.barcode) {
      const oldRow = findRowIndexByKey(sh, barcodeCol, payload.oldBarcode);
      if (oldRow > -1) sh.deleteRow(oldRow);
    }

    const rowArr = [
      payload.barcode, payload.name, payload.type, payload.category || '', payload.itemNo || '',
      (payload.tags || []).join(','), payload.price, payload.cost || 0,
      (payload.qty === null || payload.qty === undefined) ? '' : payload.qty,
      !!payload.alertOn, payload.createdAt || Date.now()
    ];
    const foundRow = findRowIndexByKey(sh, barcodeCol, payload.barcode);
    if (foundRow > -1) sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    return { ok: true };
  });
}

// Writes many item rows in one batched operation — much faster than calling
// saveItem() once per item. Skips barcodes that already exist (safe to re-run
// a partial/interrupted import without creating duplicates).
function bulkImportItems(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.ITEMS);
    const existing = new Set(sheetToObjects(SHEET_NAMES.ITEMS).map(r => String(r.barcode)));
    const rows = (payload.items || [])
      .filter(it => !existing.has(String(it.barcode)))
      .map(it => [
        it.barcode, it.name, it.type || 'book', it.category || '', it.itemNo || '',
        (it.tags || []).join(','), it.price || 0, it.cost || 0,
        (it.qty === null || it.qty === undefined) ? '' : it.qty,
        it.alertOn !== false, it.createdAt || Date.now()
      ]);
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { inserted: rows.length, skipped: (payload.items || []).length - rows.length };
  });
}

function updateQty(payload) {
  return withLock(() => {
    const updates = payload.items ? payload.items : [{ barcode: payload.barcode, qty: payload.qty }];
    const sh = sheet(SHEET_NAMES.ITEMS);
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const barcodeCol = headers.indexOf('barcode');
    const qtyCol = headers.indexOf('qty');
    const map = {};
    updates.forEach(u => { map[String(u.barcode)] = u.qty; });
    for (let r = 1; r < data.length; r++) {
      const bc = String(data[r][barcodeCol]);
      if (Object.prototype.hasOwnProperty.call(map, bc)) {
        const v = map[bc];
        sh.getRange(r + 1, qtyCol + 1).setValue((v === null || v === undefined) ? '' : v);
      }
    }
    return { ok: true };
  });
}

function savePerson(payload) {
  return withLock(() => {
    const sheetName = payload.type === 'teacher' ? SHEET_NAMES.TEACHERS
      : payload.type === 'student' ? SHEET_NAMES.STUDENTS : SHEET_NAMES.STAFF;
    const sh = sheet(sheetName);
    const headers = SHEET_HEADERS[sheetName];
    const idCol = headers.indexOf('id');

    if (payload.deleted) {
      const row = findRowIndexByKey(sh, idCol, payload.id);
      if (row > -1) sh.deleteRow(row);
      return { ok: true };
    }

    const id = payload.id || Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    let rowArr;
    if (payload.type === 'teacher') rowArr = [id, payload.name, payload.notes || ''];
    else if (payload.type === 'student') rowArr = [id, payload.name, payload.teacherId || '', payload.notes || ''];
    else rowArr = [id, payload.name];

    const foundRow = findRowIndexByKey(sh, idCol, id);
    if (foundRow > -1) sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    return { ok: true, id };
  });
}

function saveSettings(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.SETTINGS);
    const headers = SHEET_HEADERS.Settings;
    if (sh.getLastRow() < 2) sh.appendRow(headers.map(() => ''));
    const rowIdx = 2;
    Object.keys(payload).forEach(k => {
      const c = headers.indexOf(k);
      if (c > -1) sh.getRange(rowIdx, c + 1).setValue(payload[k]);
    });
    return { ok: true };
  });
}

// ── ONE-TIME SETUP ────────────────────────────────────────────────────────────

function setupSpreadsheet() {
  const spreadsheet = ss();

  Object.keys(SHEET_HEADERS).forEach(name => {
    let sh = spreadsheet.getSheetByName(name);
    if (!sh) sh = spreadsheet.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
    sh.setFrozenRows(1);
  });

  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1) spreadsheet.deleteSheet(defaultSheet);

  // Seed default data (mirrors the app's built-in defaults)
  sheet('Settings').getRange(2, 1, 1, 6).setValues([['1234', '9999', 5, '', '', '']]);
  sheet('Staff').appendRow(['s1', 'Front Desk']);
  sheet('Teachers').appendRow(['t1', 'Teacher A']);
  sheet('Teachers').appendRow(['t2', 'Teacher B']);
  sheet('Students').appendRow(['stu1', 'Ahmad Bin Ali', 't1', 'Mon 3pm']);
  sheet('Students').appendRow(['stu2', 'Siti Binti Hassan', 't1', '']);
  sheet('Students').appendRow(['stu3', 'Wei Chen', 't2', 'Sat 10am']);
  sheet('Items').appendRow(['9789670362175', 'Theory of Music Made Easy Grade 1', 'book', 'Theory', '', 'Theory', 12, 0, '', true, Date.now()]);
  sheet('Items').appendRow(['9789670362182', 'Theory of Music Made Easy Grade 2', 'book', 'Theory', '', 'Theory', 12, 0, '', true, Date.now()]);
  sheet('Items').appendRow(['INST-001', 'Ukulele (Standard)', 'stock', 'Ukulele', '', '', 200, 0, '', true, Date.now()]);
  sheet('Items').appendRow(['ACC-001', 'Guitar Capo', 'stock', 'Guitar', '', '', 24, 0, '', true, Date.now()]);

  SpreadsheetApp.flush();
  Logger.log('Setup complete — sheets created and seeded.');
}

// ── MIGRATIONS ────────────────────────────────────────────────────────────────
// Run once after updating this file on an already-deployed sheet, to add new
// columns without wiping existing data. Safe to re-run — it's a no-op if the
// column already exists.

function migrateAddCategoryColumn() {
  const sh = sheet(SHEET_NAMES.ITEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('category') > -1) {
    Logger.log('category column already exists — nothing to do.');
    return;
  }
  const typeCol = headers.indexOf('type'); // 0-based
  const insertAt = typeCol + 2; // 1-based column right after 'type'
  sh.insertColumnAfter(typeCol + 1);
  sh.getRange(1, insertAt).setValue('category');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — category column added to Items.');
}

function migrateAddItemNoAndTagsColumns() {
  const sh = sheet(SHEET_NAMES.ITEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  ['itemNo', 'tags'].forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) {
      Logger.log(colName + ' column already exists — skipping.');
      return;
    }
    const categoryCol = freshHeaders.indexOf('category'); // 0-based
    const afterCol = categoryCol > -1 ? categoryCol : freshHeaders.length - 1;
    const insertAt = afterCol + 2; // 1-based column right after 'category' (or at the end)
    sh.insertColumnAfter(afterCol + 1);
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — itemNo/tags columns added to Items.');
}
