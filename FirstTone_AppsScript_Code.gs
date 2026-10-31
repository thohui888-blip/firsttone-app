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
 *  2. Run any new migrateXxx() functions listed in this version's release notes once each
 *     (e.g. migrateAddCategoryColumn, migrateAddItemNoAndTagsColumns, migrateAddInvoiceItemTypeMonth,
 *     migrateAddStudentMonthlyFee, migrateAddStudentCourseFields, migrateAddSettingsHolidays,
 *     migrateAddAttendanceSheet, migrateAddStudentLessonDay, migrateAddStudentMonthStatus,
 *     migrateAddStudentLessonTime, migrateAddStudentDurationFields, migrateAddAttendanceSlotFields,
 *     migrateAddAttendanceSlotTeacherId, migrateFixLessonTimeFormat, migrateFixSlotTimeFormat,
 *     migrateFixInvoiceItemMonthFormat, migrateFixAttendanceDateFormat, migrateAddBillersSheet,
 *     migrateAddSettingsDirectorPhone, migrateAddExpenseBillerAndSentColumns,
 *     migrateFixCommissionVoucherMonthKeyFormat, migrateAddAuditLogSheet, migrateAddExpenseFieldsV2, migrateAddBankReconSheets).
 *     All are safe to re-run (no-op if the column already exists) and do NOT wipe data.
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
  WRITEOFFS: 'Writeoffs',
  ATTENDANCE: 'Attendance',
  EXPENSES: 'Expenses',
  COMMISSION_VOUCHERS: 'CommissionVouchers',
  BILLERS: 'Billers',
  AUDIT_LOG: 'AuditLog',
  BANK_TXNS: 'BankTxns',
  PAYER_ALIASES: 'PayerAliases'
};

const SHEET_HEADERS = {
  Settings: ['staffPIN', 'adminPIN', 'lowStockThreshold', 'supplierName', 'supplierWA', 'supplierNotes', 'holidays', 'directorPhone'],
  Staff: ['id', 'name'],
  Teachers: ['id', 'name', 'notes', 'courseCode', 'commissionRate', 'bankName', 'bankAccount', 'icNumber', 'wa', 'nickname'],
  Students: ['id', 'name', 'teacherId', 'notes', 'monthlyFee', 'ageGroup', 'instrument', 'grade', 'icNumber', 'feeOverride', 'examRecords', 'lessonDay', 'monthStatus', 'lessonTime', 'lessonDuration', 'durationOverride', 'parentWa', 'distanceKm', 'address', 'createdAt'],
  Items: ['barcode', 'name', 'type', 'category', 'itemNo', 'tags', 'price', 'cost', 'qty', 'alertOn', 'createdAt'],
  Invoices: ['id', 'no', 'date', 'buyerType', 'buyerId', 'teacherId', 'staffId', 'total', 'discount', 'paid', 'status'],
  InvoiceItems: ['invoiceId', 'barcode', 'name', 'originalPrice', 'discounted', 'type', 'month'],
  Payments: ['invoiceId', 'date', 'amount', 'method'],
  Writeoffs: ['date', 'invoiceNo', 'studentName', 'amount', 'reason'],
  Attendance: ['id', 'studentId', 'date', 'status', 'leaveBy', 'makeupDate', 'note', 'slotTime', 'slotDuration', 'slotTeacherId', 'parentNotified', 'makeupForDates'],
  Expenses: ['id', 'date', 'category', 'description', 'amount', 'method', 'receiptUrl', 'voucherNo', 'billerId', 'sentToDirector', 'payee', 'supplierInvNo', 'refNo', 'claimant', 'claimStatus', 'reimbursedDate', 'attachments'],
  CommissionVouchers: ['id', 'teacherId', 'monthKey', 'voucherNo', 'issuedAt', 'adjustments', 'excludedInvoiceIds', 'expenseId'],
  Billers: ['id', 'name', 'style', 'category', 'refNo', 'payeeName', 'bankName', 'bankAccount'],
  AuditLog: ['ts', 'entity', 'entityId', 'field', 'oldValue', 'newValue', 'by', 'reason'],
  BankTxns: ['id', 'key', 'date', 'text', 'detail', 'type', 'amount', 'balance', 'month', 'status', 'matches', 'note', 'auto'],
  PayerAliases: ['id', 'name', 'studentIds', 'teacherId']
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
      case 'saveAttendance': result = saveAttendance(payload); break;
      case 'writeoff': result = writeoffInvoice(payload); break;
      case 'cancelInvoice': result = cancelInvoiceBackend(payload); break;
      case 'saveSettings': result = saveSettings(payload); break;
      case 'saveExpense': result = saveExpense(payload); break;
      case 'saveCommissionVoucher': result = saveCommissionVoucher(payload); break;
      case 'saveBiller': result = saveBiller(payload); break;
      case 'editInvoice': result = editInvoice(payload); break;
      case 'saveBankTxns': result = saveBankTxns(payload); break;
      case 'saveAlias': result = saveAlias(payload); break;
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

// Google Sheets auto-detects things like "16:45" as a Time value, so Apps Script
// hands it back as a JS Date instead of the string we wrote. Coerce it back to "HH:MM".
function timeCellToString(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return String(v.getHours()).padStart(2, '0') + ':' + String(v.getMinutes()).padStart(2, '0');
  }
  return String(v);
}

// Same problem as timeCellToString but for "YYYY-MM" month strings — Sheets
// auto-detects "2026-07" as a date, so Apps Script hands back a Date object.
function monthCellToString(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0');
  }
  return String(v);
}

// Same problem again but for full "YYYY-MM-DD" dates (Attendance.date, Attendance.makeupDate).
function dateCellToString(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
  }
  return String(v);
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
    supplier: { name: s.supplierName || '', wa: s.supplierWA || '', notes: s.supplierNotes || '' },
    holidays: (function () { try { return s.holidays ? JSON.parse(s.holidays) : []; } catch (e) { return []; } })(),
    directorPhone: s.directorPhone || ''
  };

  const staff = sheetToObjects(SHEET_NAMES.STAFF).map(r => ({ id: String(r.id), name: r.name }));
  const teachers = sheetToObjects(SHEET_NAMES.TEACHERS).map(r => ({
    id: String(r.id), name: r.name, notes: r.notes || '',
    courseCode: r.courseCode || '', commissionRate: Number(r.commissionRate) || 0,
    bankName: r.bankName || '', bankAccount: r.bankAccount || '', icNumber: r.icNumber || '',
    wa: r.wa || '', nickname: r.nickname || ''
  }));
  const students = sheetToObjects(SHEET_NAMES.STUDENTS).map(r => ({
    id: String(r.id), name: r.name, teacherId: r.teacherId ? String(r.teacherId) : '', notes: r.notes || '',
    monthlyFee: Number(r.monthlyFee) || 0,
    ageGroup: r.ageGroup || 'child',
    instrument: r.instrument || '',
    grade: r.grade || '',
    icNumber: r.icNumber || '',
    feeOverride: r.feeOverride === true || r.feeOverride === 'TRUE',
    examRecords: (function () { try { return r.examRecords ? JSON.parse(r.examRecords) : []; } catch (e) { return []; } })(),
    lessonDay: r.lessonDay || '',
    monthStatus: (function () { try { return r.monthStatus ? JSON.parse(r.monthStatus) : {}; } catch (e) { return {}; } })(),
    lessonTime: timeCellToString(r.lessonTime),
    lessonDuration: Number(r.lessonDuration) || 0,
    durationOverride: r.durationOverride === true || r.durationOverride === 'TRUE',
    parentWa: r.parentWa || '',
    distanceKm: Number(r.distanceKm) || 0,
    address: r.address || '',
    createdAt: Number(r.createdAt) || 0
  }));

  const itemRows = sheetToObjects(SHEET_NAMES.ITEMS);
  const items = {};
  itemRows.forEach(r => {
    items[String(r.barcode)] = {
      name: String(r.name || ''),
      type: r.type,
      category: String(r.category || ''),
      itemNo: r.itemNo === '' || r.itemNo === null || r.itemNo === undefined ? '' : String(r.itemNo),
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
        originalPrice: Number(ii.originalPrice) || 0, discounted: Number(ii.discounted) || 0,
        type: ii.type || 'book', month: monthCellToString(ii.month)
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

  const attendance = sheetToObjects(SHEET_NAMES.ATTENDANCE).map(r => ({
    id: String(r.id), studentId: String(r.studentId), date: dateCellToString(r.date), status: r.status,
    leaveBy: r.leaveBy || '', makeupDate: dateCellToString(r.makeupDate), note: r.note || '',
    slotTime: timeCellToString(r.slotTime), slotDuration: Number(r.slotDuration) || 0,
    slotTeacherId: r.slotTeacherId ? String(r.slotTeacherId) : '',
    parentNotified: r.parentNotified === true || r.parentNotified === 'TRUE',
    makeupForDates: (function () { try { return r.makeupForDates ? JSON.parse(r.makeupForDates) : []; } catch (e) { return []; } })()
  }));

  const expenses = sheetToObjects(SHEET_NAMES.EXPENSES).map(r => ({
    id: String(r.id), date: Number(r.date) || 0, category: r.category || '', description: r.description || '',
    amount: Number(r.amount) || 0, method: r.method || 'cash', receiptUrl: r.receiptUrl || '', voucherNo: r.voucherNo || '',
    billerId: r.billerId || '', sentToDirector: r.sentToDirector === true || r.sentToDirector === 'TRUE',
    payee: r.payee || '', supplierInvNo: r.supplierInvNo === undefined ? '' : String(r.supplierInvNo), refNo: r.refNo === undefined ? '' : String(r.refNo),
    claimant: r.claimant || '', claimStatus: r.claimStatus || 'none', reimbursedDate: Number(r.reimbursedDate) || 0,
    attachments: (function () {
      let a = [];
      try { a = r.attachments ? JSON.parse(r.attachments) : []; } catch (e) { a = []; }
      // Receipts saved before multi-attachment support only have receiptUrl.
      if (!a.length && r.receiptUrl) a = [{ url: String(r.receiptUrl), name: 'Receipt' }];
      return a;
    })()
  }));

  const commissionVouchers = sheetToObjects(SHEET_NAMES.COMMISSION_VOUCHERS).map(r => ({
    id: String(r.id), teacherId: String(r.teacherId), monthKey: monthCellToString(r.monthKey),
    voucherNo: r.voucherNo || '', issuedAt: Number(r.issuedAt) || 0,
    adjustments: (function () { try { return r.adjustments ? JSON.parse(r.adjustments) : []; } catch (e) { return []; } })(),
    excludedInvoiceIds: (function () { try { return r.excludedInvoiceIds ? JSON.parse(r.excludedInvoiceIds) : []; } catch (e) { return []; } })(),
    expenseId: r.expenseId || ''
  }));

  const billers = sheetToObjects(SHEET_NAMES.BILLERS).map(r => ({
    id: String(r.id), name: r.name || '', style: r.style || 'utility', category: r.category || '',
    refNo: r.refNo || '', payeeName: r.payeeName || '', bankName: r.bankName || '', bankAccount: r.bankAccount || ''
  }));

  const auditLog = (function () {
    if (!sheet(SHEET_NAMES.AUDIT_LOG)) return [];
    return sheetToObjects(SHEET_NAMES.AUDIT_LOG).map(r => ({
      ts: Number(r.ts) || 0, entity: r.entity || '', entityId: String(r.entityId || ''), field: r.field || '',
      oldValue: String(r.oldValue === undefined ? '' : r.oldValue), newValue: String(r.newValue === undefined ? '' : r.newValue),
      by: r.by || '', reason: r.reason || ''
    }));
  })();

  const bankTxns = (function () {
    if (!sheet(SHEET_NAMES.BANK_TXNS)) return [];
    return sheetToObjects(SHEET_NAMES.BANK_TXNS).map(r => ({
      id: String(r.id), key: String(r.key || ''), date: Number(r.date) || 0, text: String(r.text || ''), detail: String(r.detail || ''),
      type: String(r.type || ''), amount: Number(r.amount) || 0, balance: r.balance === '' || r.balance === undefined ? null : Number(r.balance),
      month: monthCellToString(r.month), status: r.status || 'unmatched',
      matches: (function () { try { return r.matches ? JSON.parse(r.matches) : []; } catch (e) { return []; } })(),
      note: String(r.note || ''), auto: r.auto === true || r.auto === 'TRUE'
    }));
  })();
  const payerAliases = (function () {
    if (!sheet(SHEET_NAMES.PAYER_ALIASES)) return [];
    return sheetToObjects(SHEET_NAMES.PAYER_ALIASES).map(r => ({
      id: String(r.id), name: String(r.name || ''), teacherId: r.teacherId ? String(r.teacherId) : '',
      studentIds: (function () { try { return r.studentIds ? JSON.parse(r.studentIds) : []; } catch (e) { return []; } })()
    }));
  })();

  return { settings, staff, teachers, students, items, invoices, writeoffs, attendance, expenses, commissionVouchers, billers, auditLog, bankTxns, payerAliases };
}

// ── WRITE ACTIONS ─────────────────────────────────────────────────────────────

function saveInvoice(inv) {
  return withLock(() => {
    // Idempotent: a retried request for an invoice that is already saved must not add it again.
    if (findRowIndexByKey(sheet(SHEET_NAMES.INVOICES), 0, inv.id) > -1) return { id: inv.id, duplicate: true };
    appendRow(sheet(SHEET_NAMES.INVOICES), [
      inv.id, inv.no, inv.date, inv.buyerType, inv.buyerId,
      inv.teacherId || '', inv.staffId, inv.total, inv.discount || 0, inv.paid || 0, inv.status
    ]);
    const itemSh = sheet(SHEET_NAMES.INVOICE_ITEMS);
    const itemHeaders = SHEET_HEADERS.InvoiceItems;
    const monthCol = itemHeaders.indexOf('month');
    (inv.items || []).forEach(it => {
      appendRow(itemSh, [inv.id, it.barcode, it.name, it.originalPrice, it.discounted, it.type || 'book', it.month || '']);
      // Force plain-text format so Sheets doesn't auto-convert "2026-07" into a Date on this or a future save.
      if (monthCol > -1) itemSh.getRange(itemSh.getLastRow(), monthCol + 1).setNumberFormat('@').setValue(it.month || '');
    });
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
    // Idempotent: skip if this exact payment is already on the invoice (a retried request).
    const already = sheetToObjects(SHEET_NAMES.PAYMENTS).some(r => String(r.invoiceId) === String(p.invoiceId) && Number(r.date) === Number(p.date) && Number(r.amount) === Number(p.amount) && String(r.method) === String(p.method));
    if (already) return { ok: true, duplicate: true };
    appendRow(sheet(SHEET_NAMES.PAYMENTS), [p.invoiceId, p.date, p.amount, p.method]);
    updateInvoiceFields(p.invoiceId, { paid: p.newPaid, status: p.newStatus });
    return { ok: true };
  });
}

// Appends audit rows. Each change is {field, oldValue, newValue}. Skips silently if the
// AuditLog sheet hasn't been created yet (migrateAddAuditLogSheet not run) so older
// deployments keep working.
function writeAuditRows(entity, entityId, changes, by, reason) {
  const sh = sheet(SHEET_NAMES.AUDIT_LOG);
  if (!sh) return;
  const now = Date.now();
  (changes || []).forEach(c => sh.appendRow([now, entity, entityId, c.field, c.oldValue, c.newValue, by || '', reason || '']));
}

function cancelInvoiceBackend(payload) {
  return withLock(() => {
    const id = payload.id;
    // Idempotent: cancelling twice must not return the stock twice.
    const cur = sheetToObjects(SHEET_NAMES.INVOICES).find(r => String(r.id) === String(id));
    if (cur && cur.status === 'cancelled') return { ok: true, duplicate: true };
    updateInvoiceFields(id, { status: 'cancelled' });
    if (payload.audit) writeAuditRows('invoice', id, [{ field: 'status', oldValue: payload.oldStatus || '', newValue: 'cancelled' }], payload.by, payload.reason);
    const invItems = sheetToObjects(SHEET_NAMES.INVOICE_ITEMS).filter(ii => String(ii.invoiceId) === String(id));
    const delta = {};
    invItems.forEach(ii => { delta[String(ii.barcode)] = (delta[String(ii.barcode)] || 0) + 1; });
    adjustItemQtys(delta);
    return { ok: true };
  });
}

// Bank statement lines (parsed in the browser from the monthly PDF) and how each was matched
// to invoice payments / expenses. Upsert by id, in one batch per import/confirm so a whole
// statement is a single request. Rows are looked up once, not per row.
function saveBankTxns(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.BANK_TXNS);
    const headers = SHEET_HEADERS.BankTxns;
    const idCol = headers.indexOf('id');
    const monthCol = headers.indexOf('month'), keyCol = headers.indexOf('key');
    const rows = payload.rows || [];
    const lastRow = sh.getLastRow();
    const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, headers.length).getValues() : [];
    const rowOf = {};
    existing.forEach((r, i) => { rowOf[String(r[idCol])] = i + 2; });
    const toRow = t => [t.id, t.key || '', t.date || 0, t.text || '', t.detail || '', t.type || '', t.amount || 0, t.balance === null || t.balance === undefined ? '' : t.balance,
      t.month || '', t.status || 'unmatched', JSON.stringify(t.matches || []), t.note || '', t.auto === true];
    const fresh = [];
    let touched = 0;
    rows.forEach(t => {
      const at = rowOf[String(t.id)];
      if (t.deleted) { if (at) { sh.getRange(at, 1, 1, headers.length).clearContent(); } return; }
      if (at) { sh.getRange(at, 1, 1, headers.length).setValues([toRow(t)]); sh.getRange(at, monthCol + 1).setNumberFormat('@').setValue(t.month || ''); touched++; }
      else fresh.push(t);
    });
    if (fresh.length) {
      const start = sh.getLastRow() + 1;
      // Plain-text first so "2026-09" and long reference keys never get auto-converted.
      sh.getRange(start, monthCol + 1, fresh.length, 1).setNumberFormat('@');
      sh.getRange(start, keyCol + 1, fresh.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, fresh.length, headers.length).setValues(fresh.map(toRow));
    }
    return { ok: true, updated: touched, added: fresh.length };
  });
}

// "HENG CHIA TYNG" paid for Isabel + Emma last time → remember it so the next transfer from
// that payer is matched automatically.
function saveAlias(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.PAYER_ALIASES);
    const headers = SHEET_HEADERS.PayerAliases;
    const idCol = headers.indexOf('id');
    if (payload.deleted) {
      const row = findRowIndexByKey(sh, idCol, payload.id);
      if (row > -1) sh.deleteRow(row);
      return { ok: true };
    }
    const rowArr = [payload.id, payload.name || '', JSON.stringify(payload.studentIds || []), payload.teacherId || ''];
    const found = findRowIndexByKey(sh, idCol, payload.id);
    if (found > -1) sh.getRange(found, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    return { ok: true, id: payload.id };
  });
}

// In-place correction of an already-saved invoice: payment method/date and a fee item's
// month (+ its name, which embeds the month). Amounts, students and item add/remove are
// deliberately NOT editable here — those go through cancel + reissue.
// An invoice can have several Payments / InvoiceItems rows, so rows are located by their
// position among that invoice's rows, then verified against the old values the client
// saw; if the position doesn't line up (e.g. a sync landed out of order) it falls back to
// matching on the old values, and refuses rather than editing the wrong row.
function editInvoice(payload) {
  return withLock(() => {
    const id = String(payload.invoiceId);
    const changes = [];

    const paySh = sheet(SHEET_NAMES.PAYMENTS);
    const payHeaders = SHEET_HEADERS.Payments;
    const pInv = payHeaders.indexOf('invoiceId'), pDate = payHeaders.indexOf('date'), pAmt = payHeaders.indexOf('amount'), pMethod = payHeaders.indexOf('method');
    const payData = paySh.getLastRow() > 1 ? paySh.getRange(2, 1, paySh.getLastRow() - 1, payHeaders.length).getValues() : [];
    const payRows = [];
    payData.forEach((r, i) => { if (String(r[pInv]) === id) payRows.push({ row: i + 2, vals: r }); });
    (payload.payments || []).forEach(p => {
      let target = payRows[p.idx];
      if (!target || Math.abs(Number(target.vals[pAmt]) - Number(p.oldAmount)) > 0.001) {
        target = payRows.find(x => Math.abs(Number(x.vals[pAmt]) - Number(p.oldAmount)) <= 0.001 && Number(x.vals[pDate]) === Number(p.oldDate));
      }
      if (!target) throw new Error('Payment row not found for invoice ' + id);
      if (p.method !== undefined && p.method !== String(target.vals[pMethod])) {
        changes.push({ field: 'payment#' + (p.idx + 1) + ' method', oldValue: target.vals[pMethod], newValue: p.method });
        paySh.getRange(target.row, pMethod + 1).setValue(p.method);
      }
      if (p.date !== undefined && Number(p.date) !== Number(target.vals[pDate])) {
        changes.push({ field: 'payment#' + (p.idx + 1) + ' date', oldValue: target.vals[pDate], newValue: p.date });
        paySh.getRange(target.row, pDate + 1).setValue(p.date);
      }
    });

    const itemSh = sheet(SHEET_NAMES.INVOICE_ITEMS);
    const itHeaders = SHEET_HEADERS.InvoiceItems;
    const iInv = itHeaders.indexOf('invoiceId'), iName = itHeaders.indexOf('name'), iMonth = itHeaders.indexOf('month');
    const itData = itemSh.getLastRow() > 1 ? itemSh.getRange(2, 1, itemSh.getLastRow() - 1, itHeaders.length).getValues() : [];
    const itRows = [];
    itData.forEach((r, i) => { if (String(r[iInv]) === id) itRows.push({ row: i + 2, vals: r }); });
    (payload.items || []).forEach(it => {
      let target = itRows[it.idx];
      if (!target || String(target.vals[iName]) !== String(it.oldName)) {
        target = itRows.find(x => String(x.vals[iName]) === String(it.oldName));
      }
      if (!target) throw new Error('Invoice item row not found for invoice ' + id);
      const oldMonth = monthCellToString(target.vals[iMonth]);
      if (it.month !== undefined && it.month !== oldMonth) {
        changes.push({ field: 'item#' + (it.idx + 1) + ' month', oldValue: oldMonth, newValue: it.month });
        itemSh.getRange(target.row, iMonth + 1).setNumberFormat('@').setValue(it.month);
      }
      // The name just embeds the month, so it follows the month change without its own audit row.
      if (it.name !== undefined && it.name !== String(target.vals[iName])) {
        itemSh.getRange(target.row, iName + 1).setValue(it.name);
      }
    });

    writeAuditRows('invoice', id, changes, payload.by, payload.reason);
    return { ok: true, changed: changes.length };
  });
}

function writeoffInvoice(payload) {
  return withLock(() => {
    const cur = sheetToObjects(SHEET_NAMES.INVOICES).find(r => String(r.id) === String(payload.invoiceId));
    if (cur && cur.status === 'writeoff') return { ok: true, duplicate: true };
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
    if (payload.type === 'teacher') rowArr = [id, payload.name, payload.notes || '',
      payload.courseCode || '', payload.commissionRate || 0, payload.bankName || '', payload.bankAccount || '', payload.icNumber || '', payload.wa || '', payload.nickname || ''];
    else if (payload.type === 'student') rowArr = [id, payload.name, payload.teacherId || '', payload.notes || '', payload.monthlyFee || 0,
      payload.ageGroup || 'child', payload.instrument || '', payload.grade || '', payload.icNumber || '', payload.feeOverride ? true : false,
      JSON.stringify(payload.examRecords || []), payload.lessonDay || '', JSON.stringify(payload.monthStatus || {}), payload.lessonTime || '',
      payload.lessonDuration || 0, payload.durationOverride ? true : false, payload.parentWa || '', payload.distanceKm || 0, payload.address || '',
      payload.createdAt || ''];
    else rowArr = [id, payload.name];

    const foundRow = findRowIndexByKey(sh, idCol, id);
    let targetRow;
    if (foundRow > -1) { sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]); targetRow = foundRow; }
    else { sh.appendRow(rowArr); targetRow = sh.getLastRow(); }
    // Force plain-text format so Sheets doesn't auto-convert "16:45" into a Time value on this or the next save.
    if (payload.type === 'student') {
      const lessonTimeCol = headers.indexOf('lessonTime');
      if (lessonTimeCol > -1) sh.getRange(targetRow, lessonTimeCol + 1).setNumberFormat('@').setValue(payload.lessonTime || '');
    }
    return { ok: true, id };
  });
}

function getOrCreateReceiptsFolder() {
  const folderName = 'First Tempo Expense Receipts';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// Kept for older cached clients that still send a single receiptDataUri.
function uploadReceiptToDrive(dataUri, id) {
  return uploadExpenseAttachment({ dataUri: dataUri, name: 'receipt' }, id, Date.now(), 1).url;
}

// Files go to "First Tempo Expense Receipts/YYYY-MM/<PVno>_<n>_<name>.<ext>" so the accountant
// can browse a month at a time and every file name already says which voucher it belongs to.
function getMonthReceiptsFolder(dateTs) {
  const root = getOrCreateReceiptsFolder();
  const monthName = Utilities.formatDate(new Date(Number(dateTs) || Date.now()), Session.getScriptTimeZone(), 'yyyy-MM');
  const sub = root.getFoldersByName(monthName);
  return sub.hasNext() ? sub.next() : root.createFolder(monthName);
}

function uploadExpenseAttachment(att, ref, dateTs, n) {
  const match = String(att.dataUri || '').match(/^data:(.+);base64,(.*)$/);
  if (!match) return { url: '', name: att.name || '' };
  const mimeType = match[1];
  const ext = mimeType === 'application/pdf' ? '.pdf' : mimeType === 'image/png' ? '.png' : mimeType === 'image/jpeg' ? '.jpg' : '';
  const base = String(att.name || 'receipt').replace(/\.[A-Za-z0-9]{1,5}$/, '').replace(/[^A-Za-z0-9_\- ]/g, '').trim() || 'receipt';
  const fileName = (ref || 'expense') + '_' + n + '_' + base + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, fileName);
  const file = getMonthReceiptsFolder(dateTs).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), name: fileName };
}

function saveExpense(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.EXPENSES);
    const headers = SHEET_HEADERS.Expenses;
    const idCol = headers.indexOf('id');

    if (payload.deleted) {
      const row = findRowIndexByKey(sh, idCol, payload.id);
      if (row > -1) sh.deleteRow(row);
      return { ok: true };
    }

    const id = payload.id || Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    const foundRow = findRowIndexByKey(sh, idCol, id);
    const existing = foundRow > -1 ? sh.getRange(foundRow, 1, 1, headers.length).getValues()[0] : null;
    // A field the client didn't send keeps whatever the sheet already has, so a partial payload
    // (e.g. one that only re-sends the older columns) can never blank the newer ones.
    const pick = (f, dflt) => payload[f] !== undefined ? payload[f] : (existing ? existing[headers.indexOf(f)] : dflt);

    let attachments = [];
    if (payload.attachments !== undefined) attachments = payload.attachments || [];
    else if (existing) { try { attachments = JSON.parse(existing[headers.indexOf('attachments')] || '[]'); } catch (e) { attachments = []; } }
    const dateTs = pick('date', Date.now());
    const voucherNo = pick('voucherNo', '');
    const newOnes = (payload.newAttachments || []).slice();
    if (payload.receiptDataUri) newOnes.push({ dataUri: payload.receiptDataUri, name: 'receipt' });
    newOnes.forEach(att => {
      const up = uploadExpenseAttachment(att, voucherNo || id, dateTs, attachments.length + 1);
      if (up.url) attachments.push({ url: up.url, name: up.name });
    });
    const receiptUrl = attachments.length ? attachments[0].url : (payload.receiptUrl !== undefined ? payload.receiptUrl : (existing ? existing[headers.indexOf('receiptUrl')] : ''));

    const sent = pick('sentToDirector', false);
    const rowArr = [id, dateTs, pick('category', ''), pick('description', ''), pick('amount', 0), pick('method', 'cash'), receiptUrl, voucherNo,
      pick('billerId', ''), sent === true || sent === 'true' || sent === 'TRUE',
      pick('payee', ''), pick('supplierInvNo', ''), pick('refNo', ''), pick('claimant', ''), pick('claimStatus', 'none'), pick('reimbursedDate', ''),
      JSON.stringify(attachments)];

    if (foundRow > -1) sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    // Reference numbers like "00123456" must stay text.
    const rowNo = foundRow > -1 ? foundRow : sh.getLastRow();
    ['supplierInvNo', 'refNo'].forEach(f => sh.getRange(rowNo, headers.indexOf(f) + 1).setNumberFormat('@').setValue(String(pick(f, ''))));
    return { ok: true, id, receiptUrl, attachments };
  });
}

// Recurring bill payees (utilities, landlord, etc.) with fixed account/bank details,
// used to prefill Cash Flow expenses and to build the combined WhatsApp payment
// request sent to the director — see buildPaymentRequestMessage in the frontend.
function saveBiller(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.BILLERS);
    const headers = SHEET_HEADERS.Billers;
    const idCol = headers.indexOf('id');

    if (payload.deleted) {
      const row = findRowIndexByKey(sh, idCol, payload.id);
      if (row > -1) sh.deleteRow(row);
      return { ok: true };
    }

    const id = payload.id || Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    const rowArr = [id, payload.name || '', payload.style || 'utility', payload.category || '', payload.refNo || '', payload.payeeName || '', payload.bankName || '', payload.bankAccount || ''];

    const foundRow = findRowIndexByKey(sh, idCol, id);
    if (foundRow > -1) sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    return { ok: true, id };
  });
}

// One row per teacher+month. The voucher number is issued once and never changes,
// but adjustments/exclusions get updated in place on every re-save (see
// saveCommissionVoucherState in the frontend), so reopening the same teacher+month
// restores exactly what was there instead of resetting.
function saveCommissionVoucher(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.COMMISSION_VOUCHERS);
    const headers = SHEET_HEADERS.CommissionVouchers;
    const idCol = headers.indexOf('id');
    const teacherIdCol = headers.indexOf('teacherId');
    const monthKeyCol = headers.indexOf('monthKey');
    const voucherNoCol = headers.indexOf('voucherNo');
    // The frontend's local cache can be stale right after a fire-and-forget sync
    // (e.g. two saves close together, or a reload before the first one lands) —
    // if that happens it may send a freshly-generated id/voucherNo believing this is
    // a new voucher. Enforce one row per teacher+month here regardless: if a row for
    // this teacher+month already exists (under any id), keep ITS id and voucherNo —
    // never let two rows, or two PV numbers, exist for the same voucher.
    let targetRow = -1, rowId = payload.id, rowVoucherNo = payload.voucherNo;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
      for (let i = 0; i < data.length; i++) {
        // monthCellToString, not String() — Sheets can silently auto-convert a
        // "2026-08"-looking cell to a real Date, and String(Date) never equals
        // the "2026-08" string we're comparing against, which breaks this
        // dedup match and lets a duplicate row slip through.
        if (String(data[i][teacherIdCol]) === String(payload.teacherId) && monthCellToString(data[i][monthKeyCol]) === monthCellToString(payload.monthKey)) {
          targetRow = i + 2;
          rowId = data[i][idCol];
          rowVoucherNo = data[i][voucherNoCol];
          break;
        }
      }
    }
    const rowArr = [rowId, payload.teacherId, payload.monthKey, rowVoucherNo, payload.issuedAt || Date.now(),
      JSON.stringify(payload.adjustments || []), JSON.stringify(payload.excludedInvoiceIds || []), payload.expenseId || ''];
    if (targetRow > -1) sh.getRange(targetRow, 1, 1, rowArr.length).setValues([rowArr]);
    else sh.appendRow(rowArr);
    // Force plain-text so this cell can never be auto-converted to a Date again.
    if (monthKeyCol > -1) sh.getRange(targetRow > -1 ? targetRow : sh.getLastRow(), monthKeyCol + 1).setNumberFormat('@').setValue(payload.monthKey);
    return { ok: true, id: rowId, voucherNo: rowVoucherNo };
  });
}

function saveAttendance(payload) {
  return withLock(() => {
    const sh = sheet(SHEET_NAMES.ATTENDANCE);
    const headers = SHEET_HEADERS.Attendance;
    const idCol = headers.indexOf('id');
    const rowArr = [payload.id, payload.studentId, payload.date, payload.status,
      payload.leaveBy || '', payload.makeupDate || '', payload.note || '',
      payload.slotTime || '', payload.slotDuration || 0, payload.slotTeacherId || '',
      payload.parentNotified ? true : false, JSON.stringify(payload.makeupForDates || [])];
    const foundRow = findRowIndexByKey(sh, idCol, payload.id);
    let targetRow;
    if (foundRow > -1) { sh.getRange(foundRow, 1, 1, rowArr.length).setValues([rowArr]); targetRow = foundRow; }
    else { sh.appendRow(rowArr); targetRow = sh.getLastRow(); }
    const slotTimeCol = headers.indexOf('slotTime');
    if (slotTimeCol > -1) sh.getRange(targetRow, slotTimeCol + 1).setNumberFormat('@').setValue(payload.slotTime || '');
    // Force plain-text on 'date' and 'makeupDate' too — same auto-conversion risk as slotTime.
    const dateCol = headers.indexOf('date');
    if (dateCol > -1) sh.getRange(targetRow, dateCol + 1).setNumberFormat('@').setValue(payload.date || '');
    const makeupDateCol = headers.indexOf('makeupDate');
    if (makeupDateCol > -1) sh.getRange(targetRow, makeupDateCol + 1).setNumberFormat('@').setValue(payload.makeupDate || '');
    return { ok: true };
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
  let afterColName = 'category';
  ['itemNo', 'tags'].forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) {
      Logger.log(colName + ' column already exists — skipping.');
      afterColName = colName; // keep chain order correct even when skipping
      return;
    }
    const afterIdx = freshHeaders.indexOf(afterColName); // 0-based
    const afterCol = afterIdx > -1 ? afterIdx : freshHeaders.length - 1;
    const insertAt = afterCol + 2; // 1-based column right after afterColName (or at the end)
    sh.insertColumnAfter(afterCol + 1);
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
    afterColName = colName; // next column chains after this one, preserving order
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — itemNo/tags columns added to Items.');
}

// One-off repair for deployments that already ran the buggy version of
// migrateAddItemNoAndTagsColumns() above, which inserted the columns in
// reversed order (tags before itemNo) while the rest of the code always
// wrote values assuming itemNo comes first. The data in those two columns
// is fine — it just needs correct header labels. Safe to re-run.
function fixItemNoTagsColumnOrder() {
  const sh = sheet(SHEET_NAMES.ITEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const tagsCol = headers.indexOf('tags');
  const itemNoCol = headers.indexOf('itemNo');
  if (tagsCol > -1 && itemNoCol > -1 && tagsCol < itemNoCol) {
    sh.getRange(1, tagsCol + 1).setValue('itemNo');
    sh.getRange(1, itemNoCol + 1).setValue('tags');
    SpreadsheetApp.flush();
    Logger.log('Fixed — itemNo/tags headers were swapped, now corrected.');
  } else {
    Logger.log('Columns already in correct order — nothing to do.');
  }
}

function migrateAddInvoiceItemTypeMonth() {
  const sh = sheet(SHEET_NAMES.INVOICE_ITEMS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let nextCol = headers.length;
  ['type', 'month'].forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    nextCol = sh.getLastColumn() + 1;
    sh.getRange(1, nextCol).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, nextCol, lastRow - 1, 1).setValue(colName === 'type' ? 'book' : '');
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — type/month columns added to InvoiceItems.');
}

function migrateAddStudentMonthlyFee() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('monthlyFee') > -1) {
    Logger.log('monthlyFee column already exists — nothing to do.');
    return;
  }
  const insertAt = headers.length + 1; // append at the end
  sh.getRange(1, insertAt).setValue('monthlyFee');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(0);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — monthlyFee column added to Students.');
}

function migrateAddStudentCourseFields() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const defaults = { ageGroup: 'child', instrument: '', grade: '', icNumber: '', feeOverride: true, examRecords: '[]' };
  Object.keys(defaults).forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    // existing students get feeOverride=true so their current manual monthlyFee is never silently recalculated
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(defaults[colName]);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — ageGroup/instrument/grade/icNumber/feeOverride/examRecords columns added to Students.');
}

function migrateAddSettingsHolidays() {
  const sh = sheet(SHEET_NAMES.SETTINGS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('holidays') > -1) { Logger.log('holidays column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('holidays');
  if (sh.getLastRow() > 1) sh.getRange(2, insertAt).setValue('[]');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — holidays column added to Settings.');
}

function migrateAddStudentMonthStatus() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('monthStatus') > -1) { Logger.log('monthStatus column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('monthStatus');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('{}');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — monthStatus column added to Students.');
}

function migrateAddStudentDurationFields() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const defaults = { lessonDuration: 0, durationOverride: false };
  Object.keys(defaults).forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(defaults[colName]);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — lessonDuration/durationOverride columns added to Students.');
}

function migrateAddStudentLessonTime() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('lessonTime') > -1) { Logger.log('lessonTime column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('lessonTime');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — lessonTime column added to Students.');
}

function migrateAddAttendanceSlotFields() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  const defaults = { slotTime: '', slotDuration: 0 };
  Object.keys(defaults).forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(defaults[colName]);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — slotTime/slotDuration columns added to Attendance.');
}

function migrateFixLessonTimeFormat() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf('lessonTime');
  if (col === -1) { Logger.log('lessonTime column not found — nothing to do.'); return; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No student rows to fix.'); return; }
  const range = sh.getRange(2, col + 1, lastRow - 1, 1);
  const fixed = range.getValues().map(row => [timeCellToString(row[0])]);
  range.setNumberFormat('@').setValues(fixed);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — lessonTime column normalized to plain-text HH:mm and locked to Plain Text format.');
}

function migrateFixAttendanceDateFormat() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  if (!sh) { Logger.log('Attendance sheet not found — nothing to do.'); return; }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No attendance rows to fix.'); return; }
  ['date', 'makeupDate'].forEach(colName => {
    const col = headers.indexOf(colName);
    if (col === -1) { Logger.log(colName + ' column not found — skipping.'); return; }
    const range = sh.getRange(2, col + 1, lastRow - 1, 1);
    const fixed = range.getValues().map(row => [dateCellToString(row[0])]);
    range.setNumberFormat('@').setValues(fixed);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — Attendance date/makeupDate columns normalized to plain-text YYYY-MM-DD and locked to Plain Text format.');
}

function migrateFixInvoiceItemMonthFormat() {
  const sh = sheet(SHEET_NAMES.INVOICE_ITEMS);
  if (!sh) { Logger.log('InvoiceItems sheet not found — nothing to do.'); return; }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf('month');
  if (col === -1) { Logger.log('month column not found — nothing to do.'); return; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No invoice item rows to fix.'); return; }
  const range = sh.getRange(2, col + 1, lastRow - 1, 1);
  const fixed = range.getValues().map(row => [monthCellToString(row[0])]);
  range.setNumberFormat('@').setValues(fixed);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — InvoiceItems month column normalized to plain-text YYYY-MM and locked to Plain Text format.');
}

function migrateFixSlotTimeFormat() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  if (!sh) { Logger.log('Attendance sheet not found — nothing to do.'); return; }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf('slotTime');
  if (col === -1) { Logger.log('slotTime column not found — nothing to do.'); return; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No attendance rows to fix.'); return; }
  const range = sh.getRange(2, col + 1, lastRow - 1, 1);
  const fixed = range.getValues().map(row => [timeCellToString(row[0])]);
  range.setNumberFormat('@').setValues(fixed);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — slotTime column normalized to plain-text HH:mm and locked to Plain Text format.');
}

function migrateAddAttendanceSlotTeacherId() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('slotTeacherId') > -1) { Logger.log('slotTeacherId column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('slotTeacherId');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — slotTeacherId column added to Attendance.');
}

function migrateAddAttendanceParentNotified() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('parentNotified') > -1) { Logger.log('parentNotified column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('parentNotified');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(false);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — parentNotified column added to Attendance.');
}

function migrateAddAttendanceMakeupForDates() {
  const sh = sheet(SHEET_NAMES.ATTENDANCE);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('makeupForDates') > -1) { Logger.log('makeupForDates column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('makeupForDates');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('[]');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — makeupForDates column added to Attendance.');
}

function migrateAddAttendanceSheet() {
  const spreadsheet = ss();
  let sh = spreadsheet.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (sh) { Logger.log('Attendance sheet already exists — nothing to do.'); return; }
  sh = spreadsheet.insertSheet(SHEET_NAMES.ATTENDANCE);
  sh.getRange(1, 1, 1, SHEET_HEADERS.Attendance.length).setValues([SHEET_HEADERS.Attendance]);
  sh.setFrozenRows(1);
  Logger.log('Migration complete — Attendance sheet created.');
}

function migrateAddExpensesSheet() {
  const spreadsheet = ss();
  let sh = spreadsheet.getSheetByName(SHEET_NAMES.EXPENSES);
  if (sh) { Logger.log('Expenses sheet already exists — nothing to do.'); return; }
  sh = spreadsheet.insertSheet(SHEET_NAMES.EXPENSES);
  sh.getRange(1, 1, 1, SHEET_HEADERS.Expenses.length).setValues([SHEET_HEADERS.Expenses]);
  sh.setFrozenRows(1);
  Logger.log('Migration complete — Expenses sheet created.');
}

function migrateAddExpenseVoucherNo() {
  const sh = sheet(SHEET_NAMES.EXPENSES);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('voucherNo') > -1) { Logger.log('voucherNo column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('voucherNo');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — voucherNo column added to Expenses.');
}

function migrateAddAuditLogSheet() {
  const spreadsheet = ss();
  let sh = spreadsheet.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  if (sh) { Logger.log('AuditLog sheet already exists — nothing to do.'); return; }
  sh = spreadsheet.insertSheet(SHEET_NAMES.AUDIT_LOG);
  sh.getRange(1, 1, 1, SHEET_HEADERS.AuditLog.length).setValues([SHEET_HEADERS.AuditLog]);
  sh.setFrozenRows(1);
  // oldValue/newValue can hold month strings like "2026-08" — keep them as plain text.
  sh.getRange(2, 5, sh.getMaxRows() - 1, 2).setNumberFormat('@');
  Logger.log('Migration complete — AuditLog sheet created.');
}

function migrateAddExpenseFieldsV2() {
  const sh = sheet(SHEET_NAMES.EXPENSES);
  ['payee', 'supplierInvNo', 'refNo', 'claimant', 'claimStatus', 'reimbursedDate', 'attachments'].forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — nothing to do.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const def = colName === 'claimStatus' ? 'none' : colName === 'attachments' ? '[]' : '';
      sh.getRange(2, insertAt, lastRow - 1, 1).setValue(def);
    }
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — payee/supplierInvNo/refNo/claimant/claimStatus/reimbursedDate/attachments added to Expenses.');
}

function migrateAddBankReconSheets() {
  const spreadsheet = ss();
  [[SHEET_NAMES.BANK_TXNS, SHEET_HEADERS.BankTxns], [SHEET_NAMES.PAYER_ALIASES, SHEET_HEADERS.PayerAliases]].forEach(pair => {
    let sh = spreadsheet.getSheetByName(pair[0]);
    if (sh) { Logger.log(pair[0] + ' sheet already exists — nothing to do.'); return; }
    sh = spreadsheet.insertSheet(pair[0]);
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]);
    sh.setFrozenRows(1);
    Logger.log('Migration complete — ' + pair[0] + ' sheet created.');
  });
  const bt = spreadsheet.getSheetByName(SHEET_NAMES.BANK_TXNS);
  const h = SHEET_HEADERS.BankTxns;
  // month ("2026-09") and key must stay plain text or Sheets turns them into dates.
  bt.getRange(2, h.indexOf('month') + 1, bt.getMaxRows() - 1, 1).setNumberFormat('@');
  bt.getRange(2, h.indexOf('key') + 1, bt.getMaxRows() - 1, 1).setNumberFormat('@');
}

function migrateAddBillersSheet() {
  const spreadsheet = ss();
  let sh = spreadsheet.getSheetByName(SHEET_NAMES.BILLERS);
  if (sh) { Logger.log('Billers sheet already exists — nothing to do.'); return; }
  sh = spreadsheet.insertSheet(SHEET_NAMES.BILLERS);
  sh.getRange(1, 1, 1, SHEET_HEADERS.Billers.length).setValues([SHEET_HEADERS.Billers]);
  sh.setFrozenRows(1);
  Logger.log('Migration complete — Billers sheet created.');
}

function migrateAddSettingsDirectorPhone() {
  const sh = sheet(SHEET_NAMES.SETTINGS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('directorPhone') > -1) { Logger.log('directorPhone column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('directorPhone');
  if (sh.getLastRow() > 1) sh.getRange(2, insertAt).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — directorPhone column added to Settings.');
}

// One-time repair for existing CommissionVouchers rows whose monthKey cell got
// silently auto-converted from "2026-08" text into a real Date by Sheets (the
// column was never forced to plain-text format on write until this version).
// A corrupted monthKey breaks every teacherId+monthKey lookup — both the
// dedup check in saveCommissionVoucher and the frontend's restore-on-reopen —
// which looks like the voucher/exclusions randomly "resetting". Safe to re-run.
function migrateFixCommissionVoucherMonthKeyFormat() {
  const sh = sheet(SHEET_NAMES.COMMISSION_VOUCHERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const monthKeyCol = headers.indexOf('monthKey');
  if (monthKeyCol === -1) { Logger.log('monthKey column not found — nothing to do.'); return; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('No CommissionVouchers rows — nothing to do.'); return; }
  const values = sh.getRange(2, monthKeyCol + 1, lastRow - 1, 1).getValues();
  const fixed = values.map(row => [monthCellToString(row[0])]);
  sh.getRange(2, monthKeyCol + 1, fixed.length, 1).setNumberFormat('@').setValues(fixed);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — monthKey column in CommissionVouchers repaired to plain text for ' + fixed.length + ' row(s).');
}

function migrateAddExpenseBillerAndSentColumns() {
  const sh = sheet(SHEET_NAMES.EXPENSES);
  ['billerId', 'sentToDirector'].forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — nothing to do.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(colName === 'sentToDirector' ? false : '');
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — billerId and sentToDirector columns added to Expenses.');
}

function migrateAddCommissionVouchersSheet() {
  const spreadsheet = ss();
  let sh = spreadsheet.getSheetByName(SHEET_NAMES.COMMISSION_VOUCHERS);
  if (sh) { Logger.log('CommissionVouchers sheet already exists — nothing to do.'); return; }
  sh = spreadsheet.insertSheet(SHEET_NAMES.COMMISSION_VOUCHERS);
  sh.getRange(1, 1, 1, SHEET_HEADERS.CommissionVouchers.length).setValues([SHEET_HEADERS.CommissionVouchers]);
  sh.setFrozenRows(1);
  Logger.log('Migration complete — CommissionVouchers sheet created.');
}

function migrateAddCommissionVoucherPersistedState() {
  const sh = sheet(SHEET_NAMES.COMMISSION_VOUCHERS);
  const defaults = { adjustments: '[]', excludedInvoiceIds: '[]', expenseId: '' };
  Object.keys(defaults).forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(defaults[colName]);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — adjustments/excludedInvoiceIds/expenseId columns added to CommissionVouchers.');
}

function migrateAddStudentLessonDay() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('lessonDay') > -1) { Logger.log('lessonDay column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('lessonDay');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — lessonDay column added to Students.');
}

function migrateAddStudentParentWa() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('parentWa') > -1) { Logger.log('parentWa column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('parentWa');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — parentWa column added to Students.');
}

function migrateAddStudentDistanceKm() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('distanceKm') > -1) { Logger.log('distanceKm column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('distanceKm');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(0);
  SpreadsheetApp.flush();
  Logger.log('Migration complete — distanceKm column added to Students.');
}

function migrateAddTeacherCommissionFields() {
  const sh = sheet(SHEET_NAMES.TEACHERS);
  const defaults = { courseCode: '', commissionRate: 60, bankName: '', bankAccount: '', icNumber: '' };
  Object.keys(defaults).forEach(colName => {
    const freshHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (freshHeaders.indexOf(colName) > -1) { Logger.log(colName + ' column already exists — skipping.'); return; }
    const insertAt = sh.getLastColumn() + 1;
    sh.getRange(1, insertAt).setValue(colName);
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue(defaults[colName]);
  });
  SpreadsheetApp.flush();
  Logger.log('Migration complete — courseCode/commissionRate/bankName/bankAccount/icNumber columns added to Teachers.');
}

function migrateAddTeacherWa() {
  const sh = sheet(SHEET_NAMES.TEACHERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('wa') > -1) { Logger.log('wa column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('wa');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — wa column added to Teachers.');
}

function migrateAddTeacherNickname() {
  const sh = sheet(SHEET_NAMES.TEACHERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('nickname') > -1) { Logger.log('nickname column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('nickname');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — nickname column added to Teachers.');
}

function migrateAddStudentAddress() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('address') > -1) { Logger.log('address column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('address');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — address column added to Students.');
}

// New column only — existing students have no recorded join date, so their growth
// trend starts from whenever this migration runs, not retroactively.
function migrateAddStudentCreatedAt() {
  const sh = sheet(SHEET_NAMES.STUDENTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('createdAt') > -1) { Logger.log('createdAt column already exists — nothing to do.'); return; }
  const insertAt = sh.getLastColumn() + 1;
  sh.getRange(1, insertAt).setValue('createdAt');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  SpreadsheetApp.flush();
  Logger.log('Migration complete — createdAt column added to Students.');
}

// ── ONE-TIME CLEAN-UP: invoices that were saved twice ───────────────────────────────────────
// A request whose reply was lost got re-sent and the old saveInvoice appended the invoice a second
// time (Invoices, InvoiceItems and Payments each got a second copy). For every invoice id that has
// more than one Invoices row, keep one row and, if the items / payments are exact repeats of
// themselves, keep a single copy of those too. Anything that is not a clean repeat is left alone
// and reported in the log. Safe to run more than once.
function migrateRemoveDuplicateInvoiceRows() {
  withLock(() => {
    const invSh = sheet(SHEET_NAMES.INVOICES);
    const invData = invSh.getDataRange().getValues();
    const seen = {}, copies = {};
    for (let r = 1; r < invData.length; r++) {
      const id = String(invData[r][0]);
      if (!id) continue;
      copies[id] = (copies[id] || 0) + 1;
    }
    const dupIds = Object.keys(copies).filter(id => copies[id] > 1);
    if (!dupIds.length) { Logger.log('No duplicate invoices found — nothing to do.'); return; }

    const dropFromRepeats = (sh, idCol, id, c, label) => {
      const data = sh.getDataRange().getValues();
      const rowsOf = [];
      for (let r = 1; r < data.length; r++) if (String(data[r][idCol]) === id) rowsOf.push(r);
      if (!rowsOf.length) return 0;
      if (rowsOf.length % c !== 0) { Logger.log(label + ' for ' + id + ': ' + rowsOf.length + ' rows is not a multiple of ' + c + ' — left alone, check by hand.'); return 0; }
      const keep = rowsOf.length / c;
      for (let k = keep; k < rowsOf.length; k++) {
        if (JSON.stringify(data[rowsOf[k]]) !== JSON.stringify(data[rowsOf[k % keep]])) { Logger.log(label + ' for ' + id + ' are not exact repeats — left alone, check by hand.'); return 0; }
      }
      const del = rowsOf.slice(keep).sort((a, b) => b - a);
      del.forEach(r => sh.deleteRow(r + 1));
      return del.length;
    };

    dupIds.forEach(id => {
      const c = copies[id];
      const nItems = dropFromRepeats(sheet(SHEET_NAMES.INVOICE_ITEMS), 0, id, c, 'Items');
      const nPay = dropFromRepeats(sheet(SHEET_NAMES.PAYMENTS), 0, id, c, 'Payments');
      // Remove the extra Invoices rows (bottom first so row numbers stay valid).
      const data = invSh.getDataRange().getValues();
      const rows = [];
      for (let r = 1; r < data.length; r++) if (String(data[r][0]) === id) rows.push(r);
      rows.slice(1).sort((a, b) => b - a).forEach(r => invSh.deleteRow(r + 1));
      Logger.log('Invoice ' + id + ': removed ' + (rows.length - 1) + ' duplicate invoice row(s), ' + nItems + ' item row(s), ' + nPay + ' payment row(s).');
    });
    Logger.log('Done. Reload the app.');
  });
}
