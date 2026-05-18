// ══════════════════════════════════════════════════════════════════
//  AJ Medina POS — Google Apps Script Backend (code.gs)
//  Logs POS orders to a Google Sheet + User Authentication
// ══════════════════════════════════════════════════════════════════
//
//  SETUP GUIDE:
//  1. Open Google Sheets → Extensions → Apps Script
//  2. Paste this entire file as code.gs
//  3. Click Deploy → New deployment → Web App
//     - Execute as:  Me
//     - Who has access:  Anyone
//  4. Copy the Web App URL
//  5. Paste it into index.html where it says:
//     const GAS_URL = 'YOUR_GAS_WEB_APP_URL_HERE';
//
//  FIRST-TIME SETUP:
//  The admin account is auto-created on first run:
//    Username: admin
//    Password: admin123
//  Change it immediately after first login via the Admin panel.
//
// ══════════════════════════════════════════════════════════════════

// ── Sheet names ──────────────────────────────────────────────────
var SHEET_NAME_ORDERS    = 'Orders';
var SHEET_NAME_ITEMS     = 'Order Items';
var SHEET_NAME_INVENTORY = 'Inventory';
var SHEET_NAME_USERS     = 'Users';

// ── JSON helper ──────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
//  GET handler — read actions + GET-based auth/user routes
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  try {
    if (action === 'login')        return handleLogin(e);
    if (action === 'inventory')    return handleGetInventory(e);
    if (action === 'transactions') return handleGetTransactions(e);
    if (action === 'getUsers')     return handleGetUsers(e);
    if (action === 'createUser')   return handleCreateUser(e);
    if (action === 'updateUser')   return handleUpdateUser(e);
    if (action === 'deleteUser')   return handleDeleteUser(e);

    return jsonResponse({ status: 'ok', app: 'AJ Medina POS', time: new Date().toISOString() });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ════════════════════════════════════════════════════════════════════
//  POST handler — orders + inventory write actions
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData ? e.postData.contents : '{}');
    var action = body.action || '';

    // No action = legacy order save
    if (!action)                        return handleSaveOrder(body);
    if (action === 'addInventory')      return handleAddInventory(body);
    if (action === 'updateInventory')   return handleUpdateInventory(body);
    if (action === 'deleteInventory')   return handleDeleteInventory(body);

    return jsonResponse({ success: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ════════════════════════════════════════════════════════════════════
//  AUTH — Login  (GET ?action=login&username=…&password=…)
// ════════════════════════════════════════════════════════════════════
function handleLogin(e) {
  var username = ((e.parameter && e.parameter.username) || '').trim().toLowerCase();
  var password = ((e.parameter && e.parameter.password) || '').trim();

  if (!username || !password) {
    return jsonResponse({ success: false, error: 'Username and password required.' });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS,
    ['Username', 'Password', 'Role', 'Created At', 'Last Login']);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if ((row[0] || '').toString().trim().toLowerCase() === username &&
        (row[1] || '').toString().trim() === password) {
      // Update last login
      sheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      return jsonResponse({
        success:  true,
        username: row[0].toString().trim(),
        role:     (row[2] || 'user').toString().trim()
      });
    }
  }

  // Seed default admin on first run
  if (data.length <= 1 && username === 'admin' && password === 'admin123') {
    sheet.appendRow(['admin', 'admin123', 'admin', new Date().toISOString(), new Date().toISOString()]);
    return jsonResponse({ success: true, username: 'admin', role: 'admin' });
  }

  return jsonResponse({ success: false, error: 'Invalid username or password.' });
}

// ════════════════════════════════════════════════════════════════════
//  USERS CRUD  (all via GET URL params)
// ════════════════════════════════════════════════════════════════════
function handleGetUsers(e) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS,
    ['Username', 'Password', 'Role', 'Created At', 'Last Login']);
  var data  = sheet.getDataRange().getValues();
  var users = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    users.push({
      username:  r[0].toString(),
      role:      r[2] ? r[2].toString() : 'user',
      createdAt: r[3] ? r[3].toString() : '',
      lastLogin: r[4] ? r[4].toString() : '',
      rowIndex:  i - 1  // 0-based index for client
    });
  }

  return jsonResponse({ success: true, users: users });
}

function handleCreateUser(e) {
  var username = ((e.parameter && e.parameter.username) || '').trim();
  var password = ((e.parameter && e.parameter.password) || '').trim();
  var role     = ((e.parameter && e.parameter.role)     || 'user').trim();

  if (!username || !password) {
    return jsonResponse({ success: false, error: 'Username and password required.' });
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_USERS,
    ['Username', 'Password', 'Role', 'Created At', 'Last Login']);
  var data  = sheet.getDataRange().getValues();

  // Check for duplicate username
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === username.toLowerCase()) {
      return jsonResponse({ success: false, error: 'Username already exists.' });
    }
  }

  sheet.appendRow([username, password, role, new Date().toISOString(), '']);
  return jsonResponse({ success: true });
}

function handleUpdateUser(e) {
  var rowIndex = parseInt((e.parameter && e.parameter.rowIndex) || '0');
  var username = ((e.parameter && e.parameter.username) || '').trim();
  var password = ((e.parameter && e.parameter.password) || '').trim();
  var role     = ((e.parameter && e.parameter.role)     || 'user').trim();

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheet    = getOrCreateSheet(ss, SHEET_NAME_USERS,
    ['Username', 'Password', 'Role', 'Created At', 'Last Login']);
  var sheetRow = rowIndex + 2; // +1 header, +1 for 1-based

  if (sheetRow > sheet.getLastRow()) {
    return jsonResponse({ success: false, error: 'User row not found.' });
  }

  sheet.getRange(sheetRow, 1).setValue(username);
  if (password) sheet.getRange(sheetRow, 2).setValue(password);
  sheet.getRange(sheetRow, 3).setValue(role);

  return jsonResponse({ success: true });
}

function handleDeleteUser(e) {
  var rowIndex = parseInt((e.parameter && e.parameter.rowIndex) || '0');
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheet    = getOrCreateSheet(ss, SHEET_NAME_USERS,
    ['Username', 'Password', 'Role', 'Created At', 'Last Login']);
  var sheetRow = rowIndex + 2;

  if (sheetRow > sheet.getLastRow()) {
    return jsonResponse({ success: false, error: 'User row not found.' });
  }

  // Protect: never delete the last admin
  var role = sheet.getRange(sheetRow, 3).getValue();
  if (role === 'admin') {
    var last       = sheet.getLastRow();
    var adminCount = 0;
    if (last >= 2) {
      sheet.getRange(2, 3, last - 1, 1).getValues()
           .forEach(function(r) { if (r[0] === 'admin') adminCount++; });
    }
    if (adminCount <= 1) {
      return jsonResponse({ success: false, error: 'Cannot delete the last admin account.' });
    }
  }

  sheet.deleteRow(sheetRow);
  return jsonResponse({ success: true });
}

// ════════════════════════════════════════════════════════════════════
//  INVENTORY
// ════════════════════════════════════════════════════════════════════
function handleGetInventory(e) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY,
    ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS']);
  var data      = sheet.getDataRange().getValues();
  var inventory = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    inventory.push({
      item:       r[0].toString(),
      unit:       r[1] ? r[1].toString() : '',
      beginQty:   r[2] !== undefined ? r[2] : 0,
      withdrawal: r[3] !== undefined ? r[3] : 0,
      balance:    r[4] !== undefined ? r[4] : 0,
      remarks:    r[5] ? r[5].toString() : '',
      rowIndex:   i - 1  // 0-based
    });
  }

  return jsonResponse({ success: true, inventory: inventory });
}

function handleAddInventory(body) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY,
    ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS']);

  sheet.appendRow([
    body.item       || '',
    body.unit       || '',
    parseFloat(body.beginQty)   || 0,
    parseFloat(body.withdrawal) || 0,
    parseFloat(body.balance)    || 0,
    body.remarks    || ''
  ]);
  return jsonResponse({ success: true });
}

function handleUpdateInventory(body) {
  var rowIndex = parseInt(body.rowIndex !== undefined ? body.rowIndex : -1);
  if (rowIndex < 0) return jsonResponse({ success: false, error: 'rowIndex required.' });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY,
    ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS']);
  var sheetRow = rowIndex + 2;

  if (sheetRow > sheet.getLastRow()) {
    return jsonResponse({ success: false, error: 'Inventory row not found.' });
  }

  sheet.getRange(sheetRow, 1, 1, 6).setValues([[
    body.item       || '',
    body.unit       || '',
    parseFloat(body.beginQty)   || 0,
    parseFloat(body.withdrawal) || 0,
    parseFloat(body.balance)    || 0,
    body.remarks    || ''
  ]]);
  return jsonResponse({ success: true });
}

function handleDeleteInventory(body) {
  var rowIndex = parseInt(body.rowIndex !== undefined ? body.rowIndex : -1);
  if (rowIndex < 0) return jsonResponse({ success: false, error: 'rowIndex required.' });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_NAME_INVENTORY,
    ['ITEM', 'UNIT-M', 'BEGINNING QTY', 'TOTAL WITHDRAWAL', 'AVAILABLE BALANCE', 'REMARKS']);
  var sheetRow = rowIndex + 2;

  if (sheetRow > sheet.getLastRow()) {
    return jsonResponse({ success: false, error: 'Inventory row not found.' });
  }

  sheet.deleteRow(sheetRow);
  return jsonResponse({ success: true });
}

// ════════════════════════════════════════════════════════════════════
//  TRANSACTIONS — Save order (POST, no action field)
// ════════════════════════════════════════════════════════════════════
function handleSaveOrder(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Orders sheet ──
  var ordersSheet = getOrCreateSheet(ss, SHEET_NAME_ORDERS,
    ['Order #', 'Date & Time', 'Line Items', 'Total Qty', 'Grand Total (₱)']);

  ordersSheet.appendRow([
    body.orderNum || '',
    formatDateTime(body.timestamp),
    body.items ? body.items.length : 0,
    body.items ? body.items.reduce(function(s, i) { return s + i.qty; }, 0) : 0,
    parseFloat(body.total) || 0
  ]);

  // ── Order Items sheet ──
  var itemsSheet = getOrCreateSheet(ss, SHEET_NAME_ITEMS,
    ['Order #', 'Date & Time', 'Product', 'Qty', 'Unit Price (₱)', 'Subtotal (₱)']);

  if (Array.isArray(body.items) && body.items.length > 0) {
    body.items.forEach(function(item) {
      itemsSheet.appendRow([
        body.orderNum || '',
        formatDateTime(body.timestamp),
        item.name     || '',
        item.qty      || 0,
        parseFloat(item.price)    || 0,
        parseFloat(item.subtotal) || 0
      ]);
    });
  }

  try {
    ordersSheet.autoResizeColumns(1, 5);
    itemsSheet.autoResizeColumns(1, 6);
  } catch(e) { /* ignore */ }

  return jsonResponse({ success: true, orderNum: body.orderNum });
}

// ════════════════════════════════════════════════════════════════════
//  TRANSACTIONS — Fetch  (GET ?action=transactions)
// ════════════════════════════════════════════════════════════════════
function handleGetTransactions(e) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_ORDERS);

  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ success: true, transactions: [] });
  }

  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[2]) continue;
    rows.push({
      orderNum: r[0] !== undefined ? r[0].toString() : '',
      time:     r[1] ? r[1].toString() : '',
      items:    r[2] !== undefined ? r[2] : 0,
      total:    r[4] !== undefined ? r[4] : 0,
      paid:     0,
      change:   0
    });
  }

  rows.reverse(); // newest first
  return jsonResponse({ success: true, transactions: rows });
}

// ════════════════════════════════════════════════════════════════════
//  UTILITY — get or create a sheet with headers
// ════════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#1a1a2e')
         .setFontColor('#f5c842')
         .setFontWeight('bold')
         .setFontSize(11)
         .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ════════════════════════════════════════════════════════════════════
//  UTILITY — format date/time for the sheet
// ════════════════════════════════════════════════════════════════════
function formatDateTime(isoString) {
  if (!isoString) return new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });
  try {
    return Utilities.formatDate(new Date(isoString), 'Asia/Manila', 'yyyy-MM-dd HH:mm:ss');
  } catch(e) {
    return isoString;
  }
}

// ══════════════════════════════════════════════════════════════════
//  OPTIONAL: Daily Summary Email
// ══════════════════════════════════════════════════════════════════

/*
function sendDailySummary() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return;

  var today = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
  var data  = sheet.getDataRange().getValues();
  var todayOrders = data.slice(1).filter(function(row) {
    return String(row[1]).startsWith(today);
  });

  var totalRevenue = todayOrders.reduce(function(s, r) { return s + (r[4] || 0); }, 0);

  var body = 'AJ Medina POS Daily Summary\n\n' +
             'Date: '    + today + '\n' +
             'Orders: '  + todayOrders.length + '\n' +
             'Revenue: ₱' + totalRevenue.toFixed(2) + '\n\n' +
             'View full report: ' + ss.getUrl();

  MailApp.sendEmail({
    to:      Session.getActiveUser().getEmail(),
    subject: 'AJ Medina POS — Daily Summary ' + today,
    body:    body
  });
}
*/
