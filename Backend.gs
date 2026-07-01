// =============================================================
// Backend.gs — Complete FIFO Dashboard Backend (Single File)
// Paste this entire file as ONE .gs file in Google Apps Script.
//
// What's inside (in order):
//  • Core router, doGet, shared helpers, sheet utilities
//  • Sheet schema setup + seed data
//  • Authentication + session management
//  • Cache engine + audit log
//  • FIFO batch ranking
//  • Violation detection + compliance metrics
//  • Dashboard summary + financial KPIs
//  • Gatepass upload + inventory import + exports
// =============================================================

// =============================================================
// Code.gs — Router · Shared Helpers · Sheet Setup · Auth · Cache · Audit
// =============================================================

/**
 * FIFO Violation Monitoring Dashboard
 * Main Entry Point - Google Apps Script
 * Version: 1.0.0
 */

// ─── GLOBAL CONSTANTS ────────────────────────────────────────────────────────
const APP_VERSION = '1.0.0';
const BUILD_ID = 'fifo-scope-today-yesterday-mtd-20260617-v4';
const CACHE_TTL   = 21600; // 6 hours in seconds
const APP_TIMEZONE = 'Asia/Kolkata';

const SHEETS = {
  INVENTORY   : 'Inventory_Current',
  GATEPASS    : 'Gatepass_Data',
  VIOLATIONS  : 'FIFO_Violations',
  COMPLIANCE  : 'User_Compliance',
  USERS       : 'Users',
  SETTINGS    : 'Settings',
  AUDIT_LOG   : 'Audit_Log',
  CACHE       : 'Cache_Data'
};

const FACILITIES = ['SL Ambient', 'SL Mother Hub'];

// ─── WEB APP ENTRY POINT ─────────────────────────────────────────────────────
function doGet(e) {
  try {
    const html = HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('FIFO Control Tower')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return html;
  } catch (err) {
    logSystemError('doGet', err);
    return HtmlService.createHtmlOutput('<p>Error loading application. Please contact admin.</p>');
  }
}

/**
 * Required by Apps Script templating.
 * Lets index.html pull in css.html and js.html via <?!= include("css"); ?>
 * without Apps Script sanitizing the JS inside them.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── HTTP API ENDPOINT (for GitHub Pages / external hosting) ─────────────────
/**
 * doPost — Called when GitHub Pages frontend makes a fetch() POST request.
 * Accepts JSON body: { action, params, token }
 * Returns JSON response with CORS headers so any origin can call it.
 */
function doPost(e) {
  var output;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action  || '';
    var params = body.params  || {};
    var token  = body.token   || '';
    var result = apiCall(action, params, token);
    output = ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    output = ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'PARSE_ERROR', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}


// ─── MAIN API ROUTER ─────────────────────────────────────────────────────────
/**
 * Universal API dispatcher called from frontend via google.script.run
 * All frontend calls go through App.api() which calls this function.
 */
function apiCall(action, params, token) {
  try {
    // Auth is handled client-side (passwordless identity capture).
    // _user is passed in params for audit tracking — no server token needed.

    // Route to correct handler
    switch (action) {
      // Auth
      // No authentication — open access, identity captured client-side

      // Dashboard
      case 'getDashboardSummary':   return getDashboardSummary(params);
      case 'getBuildInfo':          return getBuildInfo(params);
      case 'debugDateScopes':       return debugDateScopes(params);

      // Violations
      case 'getViolations':         return getViolations(params);
      case 'getViolationDetail':    return getViolationDetail(params);

      // Prevention
      case 'getPreventionData':     return getPreventionData(params);

      // Financial
      // Financial module removed

      // User Compliance
      case 'getUserCompliance':     return getUserCompliance(params);

      // Gatepass
      case 'uploadGatepass':        return uploadGatepass(params);
      case 'getGatepassHistory':    return getGatepassHistory(params);

      // Inventory
      case 'getInventoryStatus':    return getInventoryStatus(params);
      case 'triggerInventoryImport': return triggerInventoryImport(params);
      case 'uploadInventory':        return uploadInventory(params);

      // Settings
      case 'getSettings':           return getSettings(params);
      case 'saveSettings':          return saveSettings(params);

      // Audit
      case 'getAuditLog':           return getAuditLog(params);
      case 'logAuditEvent':         return logAuditEventAPI(params);

      // Cache
      case 'invalidateCache':       return invalidateCacheManual(params);

      // Export helpers
      case 'exportViolations':      return exportViolationsReport(params);
      case 'exportPrevention':      return exportPreventionReport(params);
      // removed: case 'exportFinancial':       return exportFinancialReport(params);
      case 'exportCompliance':      return exportComplianceReport(params);
      case 'exportFacility':        return exportFacilitySummaryReport(params);

      default:
        return { success: false, error: 'UNKNOWN_ACTION', message: `Unknown action: ${action}` };
    }
  } catch (err) {
    logSystemError(`apiCall[${action}]`, err);
    return { success: false, error: 'SERVER_ERROR', message: err.message };
  }
}

// ─── SPREADSHEET HELPERS ─────────────────────────────────────────────────────
/**
 * ⚠️  REQUIRED SETUP: Paste your Google Spreadsheet ID below.
 * Find it in the spreadsheet URL:
 * https://docs.google.com/spreadsheets/d/<<PASTE_ID_HERE>>/edit
 */
const SPREADSHEET_ID = '1I8zak_WUPFmO9KWYiCPPO8Ey_2g3uK8_Zd9khBLw9Jc';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * Returns a sheet by name, throws if not found.
 */
function getSheet(name) {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet not found: ${name}`);
  return sheet;
}

/**
 * Returns all data from a sheet as array of objects using header row.
 * Skips empty rows.
 */
function getSheetData(name) {
  const sheet = getSheet(name);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows    = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Skip completely empty rows
    if (row.every(cell => cell === '' || cell === null || cell === undefined)) continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

/**
 * Clears all data rows (keeps header) and writes new rows.
 */
function replaceSheetData(name, headers, rows) {
  const sheet = getSheet(name);
  const lastRow = sheet.getLastRow();

  // Clear existing data rows
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }

  if (rows.length === 0) return;

  // Build 2D array
  const values = rows.map(row => headers.map(h => {
    const v = row[h];
    return (v === undefined || v === null) ? '' : v;
  }));

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

/**
 * Appends rows to a sheet.
 */
function appendSheetRows(name, headers, rows) {
  if (rows.length === 0) return;
  const sheet = getSheet(name);
  const values = rows.map(row => headers.map(h => {
    const v = row[h];
    return (v === undefined || v === null) ? '' : v;
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, headers.length).setValues(values);
}

// ─── DATE HELPERS ─────────────────────────────────────────────────────────────
function toDateObj(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(val);
  if (typeof val === 'string') {
    var s = val.trim();
    // Handle dd-mm-yyyy and dd/mm/yyyy (Uniware export format)
    var dmyDash  = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    var dmySlash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmyDash)  return new Date(+dmyDash[3],  +dmyDash[2]-1,  +dmyDash[1]);
    if (dmySlash) return new Date(+dmySlash[3], +dmySlash[2]-1, +dmySlash[1]);
    // Handle yyyy-mm-dd
    var ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (ymd) return new Date(+ymd[1], +ymd[2]-1, +ymd[3]);
    // Fallback: let JS parse (handles ISO, etc.)
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function formatDate(val) {
  const d = toDateObj(val);
  if (!d) return '';
  return Utilities.formatDate(d, APP_TIMEZONE, 'yyyy-MM-dd');
}

function formatDateTime(val) {
  const d = toDateObj(val);
  if (!d) return '';
  return Utilities.formatDate(d, APP_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function daysBetween(d1, d2) {
  const a = toDateObj(d1);
  const b = toDateObj(d2);
  if (!a || !b) return null;
  return Math.round((b - a) / 86400000);
}

function todayStart() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now;
}

// ─── NUMBER HELPERS ───────────────────────────────────────────────────────────
function getReportDateRange_(params) {
  params = params || {};
  var scope = safeStr(params.scope || '').toLowerCase();
  var from = safeStr(params.dateFrom || '');
  var to = safeStr(params.dateTo || '');
  var today = todayStart();

  if (scope === 'today') {
    from = formatDate(today);
    to = formatDate(today);
  } else if (scope === 'yesterday') {
    var yesterday = new Date(today.getTime());
    yesterday.setDate(yesterday.getDate() - 1);
    from = formatDate(yesterday);
    to = formatDate(yesterday);
  } else if (scope === 'mtd') {
    var first = new Date(today.getFullYear(), today.getMonth(), 1);
    from = formatDate(first);
    to = formatDate(today);
  }

  return { dateFrom: from, dateTo: to, scope: scope || 'today' };
}

function filterRowsByDateRange_(rows, dateFrom, dateTo) {
  dateFrom = normaliseDateKey_(dateFrom);
  dateTo = normaliseDateKey_(dateTo);
  if (!dateFrom && !dateTo) return rows;

  return rows.filter(function(r) {
    var keys = getComparableDateKeys_(r['Date']);
    for (var i = 0; i < keys.length; i++) {
      var d = keys[i];
      if (dateFrom && d < dateFrom) continue;
      if (dateTo && d > dateTo) continue;
      return true;
    }
    return false;
  });
}

function filterRowsByScope_(rows, range) {
  range = range || {};
  return filterRowsByDateRange_(rows, range.dateFrom || '', range.dateTo || '');
}

function getDistinctActualRowDates_(rows) {
  var found = {};
  rows.forEach(function(r) {
    var d = formatDate(r['Date']);
    if (d) found[d] = true;
  });
  return Object.keys(found).sort();
}

function normaliseDateKey_(val) {
  if (!val) return '';
  if (typeof val === 'string') {
    var s = val.trim();
    var ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (ymd) return ymd[1] + '-' + String(ymd[2]).padStart(2, '0') + '-' + String(ymd[3]).padStart(2, '0');
    var dmyDash = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (dmyDash) return dmyDash[3] + '-' + String(dmyDash[2]).padStart(2, '0') + '-' + String(dmyDash[1]).padStart(2, '0');
    var dmySlash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmySlash) return dmySlash[3] + '-' + String(dmySlash[2]).padStart(2, '0') + '-' + String(dmySlash[1]).padStart(2, '0');
  }
  return formatDate(val);
}

function getComparableDateKeys_(val) {
  var d = toDateObj(val);
  if (!d) return [];
  var keys = {};
  keys[formatDate(d)] = true;
  var shifted = new Date(d.getTime());
  shifted.setDate(shifted.getDate() + 1);
  keys[formatDate(shifted)] = true;
  return Object.keys(keys);
}

function safeNum(val, fallback = 0) {
  const n = Number(val);
  return isNaN(n) ? fallback : n;
}

function safeStr(val) {
  if (val === null || val === undefined) return '';
  return String(val).trim();
}

// ─── SYSTEM ERROR LOGGER ──────────────────────────────────────────────────────
/**
 * Public API for frontend to log audit events (no auth required).
 */
function logAuditEventAPI(params) {
  try {
    var type        = safeStr((params || {}).type        || 'USER_ACTION');
    var description = safeStr((params || {}).description || '');
    var user        = safeStr((params || {}).user        || 'anonymous');
    var facility    = safeStr((params || {}).facility    || '');
    writeAuditLog({ eventType: type, description: description, user: user, facility: facility });
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function logSystemError(context, err) {
  try {
    console.error(`[${context}] ${err.message}`, err.stack);
    writeAuditLog({
      eventType    : 'SYSTEM_ERROR',
      description  : `[${context}] ${err.message}`,
      user         : 'SYSTEM',
      facility     : '',
      metadata     : err.stack || ''
    });
  } catch (_) {
    // Silently fail if audit log itself errors
  }
}


// =============================================================
// SHEET SETUP
// =============================================================

/**
 * SheetSetup.gs
 * Creates and configures all required Google Sheets with headers,
 * formatting, and initial seed data (Settings, Users).
 */

// ─── SHEET SCHEMA DEFINITIONS ─────────────────────────────────────────────────
const SCHEMA = {
  Inventory_Current: {
    headers: [
      'Facility', 'Facility Code', 'Item Type SKU Code', 'Item Type Name',
      'Inventory Type', 'Shelf', 'Quantity', 'Batch Code', 'Vendor Batch Code',
      'Manufacturing Date', 'Expiry', 'MRP', 'Batch Status'
    ],
    frozen: 1,
    widths: [120, 100, 140, 200, 80, 80, 120, 140, 130, 130, 80]
  },

  Gatepass_Data: {
    headers: [
      'Upload ID', 'Upload Date', 'Gatepass Code', 'Gatepass Created By',
      'Gatepass Created At', 'Gatepass Status', 'To Party', 'Item Name', 'Item SkuCode',
      'Quantity', 'Shelf', 'From Party', 'Uniware Batch Code',
      'Vendor Batch No', 'Manufacturing date', 'Expiry Date', 'Facility'
    ],
    frozen: 1,
    widths: [120, 120, 130, 160, 130, 160, 180, 140, 80, 80, 120, 140, 140, 130, 130, 120]
  },

  FIFO_Violations: {
    headers: [
      'ID', 'Date', 'GP Number', 'User', 'User Email', 'SKU Code', 'Item Name',
      'Batch Code', 'Vendor Batch', 'Dispatched Qty', 'Shelf',
      'Older Batch Count', 'Older Inventory Qty', 'Severity',
      'Status', 'GP Status', 'Facility', 'To Party', 'Loss Opportunity',
      'Manufacturing Date', 'Expiry Date', 'FIFO Rank',
      'Skipped Batch Details', 'Upload ID'
    ],
    frozen: 1,
    widths: [80, 110, 120, 140, 180, 140, 180, 130, 140, 100, 80,
             120, 130, 90, 140, 120, 160, 120, 130, 130, 80, 200, 120]
  },

  User_Compliance: {
    headers: [
      'User Name', 'User Email', 'Facility',
      'Total GPs', 'Violations', 'Partial FIFO', 'Compliant',
      'Compliance %', 'Skipped Qty', 'Loss Opportunity',
      'Last Violation Date', 'Violations Last 30 Days', 'Offender Level',
      'Is Champion', 'Last Updated'
    ],
    frozen: 1,
    widths: [140, 200, 120, 90, 90, 100, 90, 100, 100, 120, 140, 150, 120, 100, 130]
  },

  Users: {
    headers: [
      'Name', 'Email', 'Password Hash', 'Role', 'Facility',
      'Active', 'Created At', 'Last Login'
    ],
    frozen: 1,
    widths: [140, 200, 200, 90, 130, 70, 130, 130]
  },

  Settings: {
    headers: ['Key', 'Value', 'Label', 'Description', 'Updated By', 'Updated At'],
    frozen: 1,
    widths: [180, 120, 180, 280, 140, 140]
  },

  Audit_Log: {
    headers: [
      'ID', 'Timestamp', 'Event Type', 'Description',
      'User', 'Facility', 'Metadata'
    ],
    frozen: 1,
    widths: [80, 140, 130, 300, 140, 120, 300]
  },

  Cache_Data: {
    headers: ['Cache Key', 'Value', 'Created At', 'Expires At'],
    frozen: 1,
    widths: [180, 400, 140, 140]
  }
};

// ─── MAIN SETUP FUNCTION ──────────────────────────────────────────────────────
/**
 * Run this ONCE after deployment to initialize all sheets.
 * Safe to re-run — will not destroy existing data.
 */
function setupAllSheets() {
  const ss = getSpreadsheet();
  const log = [];

  Object.entries(SCHEMA).forEach(([sheetName, config]) => {
    try {
      let sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        log.push(`Created: ${sheetName}`);
      } else {
        log.push(`Exists:  ${sheetName}`);
      }

      applySchema(sheet, config);
    } catch (err) {
      log.push(`ERROR on ${sheetName}: ${err.message}`);
    }
  });

  // Seed default data
  seedSettings();
  seedAdminUser();

  // Move sheets to logical order
  reorderSheets(ss);

  console.log('Setup complete:\n' + log.join('\n'));
  return { success: true, log };
}

/**
 * Apply headers, frozen rows, column widths to a sheet.
 * Only writes headers if row 1 is empty.
 */
function applySchema(sheet, config) {
  const { headers, frozen, widths } = config;

  // Always write correct headers (safe - just overwrites row 1)
  // This ensures schema changes are picked up without full data reset
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1e3a5f');
  headerRange.setFontColor('#ffffff');
  headerRange.setWrap(false);

  // Freeze rows
  if (frozen) sheet.setFrozenRows(frozen);

  // Set column widths
  if (widths && widths.length) {
    widths.forEach((w, idx) => {
      try { sheet.setColumnWidth(idx + 1, w); } catch (_) {}
    });
  }

  // Auto-resize sheet tab if too many blank rows
  try {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (sheet.getMaxRows() > lastRow + 100) {
      sheet.deleteRows(lastRow + 2, sheet.getMaxRows() - lastRow - 1);
    }
  } catch (_) {}
}

/**
 * Seed default Settings rows. Skips if already seeded.
 */
function seedSettings() {
  const sheet = getSheet(SHEETS.SETTINGS);
  const existing = sheet.getLastRow();
  if (existing > 1) return; // Already seeded

  const defaults = [
    ['min_inventory_threshold', '50',    'Min Inventory Threshold',    'Older inventory below this qty is ignored', 'SYSTEM', formatDateTime(new Date())],
    ['severity_minor_min',      '50',    'Severity Minor Min',         'Min skipped qty for Minor violation',       'SYSTEM', formatDateTime(new Date())],
    ['severity_minor_max',      '499',   'Severity Minor Max',         'Max skipped qty for Minor violation',       'SYSTEM', formatDateTime(new Date())],
    ['severity_major_min',      '500',   'Severity Major Min',         'Min skipped qty for Major violation',       'SYSTEM', formatDateTime(new Date())],
    ['severity_major_max',      '999',   'Severity Major Max',         'Max skipped qty for Major violation',       'SYSTEM', formatDateTime(new Date())],
    ['severity_critical_min',   '1000',  'Severity Critical Min',      'Min skipped qty for Critical violation',    'SYSTEM', formatDateTime(new Date())],
    ['repeat_offender_days',    '30',    'Repeat Offender Window Days','Rolling window for repeat offender calc',   'SYSTEM', formatDateTime(new Date())],
    ['repeat_offender_min',     '5',     'Repeat Offender Min Violations','Min violations to be a repeat offender','SYSTEM', formatDateTime(new Date())],
    ['champion_min_gps',        '50',    'Champion Min GPs',           'Min gatepasses to qualify as champion',     'SYSTEM', formatDateTime(new Date())],
    ['champion_min_compliance', '98',    'Champion Min Compliance %',  'Min compliance % to qualify as champion',   'SYSTEM', formatDateTime(new Date())],
    ['inventory_import_email',  'noreply@e.unicommerce.com', 'Inventory Import Email', 'Gmail sender address for auto inventory import', 'SYSTEM', formatDateTime(new Date())],
    ['inventory_import_subject', 'Export Job Complete - All facility Shelfwise Inventory', 'Inventory Import Subject', 'Email subject line for auto inventory import', 'SYSTEM', formatDateTime(new Date())],
    ['cache_ttl_hours',         '6',     'Cache TTL Hours',            'Cache duration in hours',                   'SYSTEM', formatDateTime(new Date())]
  ];

  sheet.getRange(2, 1, defaults.length, 6).setValues(defaults);
}

/**
 * Seed default admin user. Skips if users already exist.
 */
function seedAdminUser() {
  const sheet = getSheet(SHEETS.USERS);
  const existing = sheet.getLastRow();
  if (existing > 1) return;

  // Default admin: admin@demo.com / admin123
  // Password is SHA-256 hashed
  const hash = hashPassword('admin123');
  const now  = formatDateTime(new Date());

  sheet.getRange(2, 1, 1, 8).setValues([[
    'Administrator', 'admin@demo.com', hash, 'admin', 'SL Ambient', true, now, ''
  ]]);
}

/**
 * Reorder sheets in logical sequence.
 */
function reorderSheets(ss) {
  const order = [
    'Inventory_Current', 'Gatepass_Data', 'FIFO_Violations',
    'User_Compliance', 'Users', 'Settings', 'Audit_Log', 'Cache_Data'
  ];
  order.forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.setActiveSheet(sheet), ss.moveActiveSheet(idx + 1);
  });
}

// ─── SCHEMA VALIDATION ────────────────────────────────────────────────────────
/**
 * Validates that a file's columns include required fields.
 * Returns { valid, missing }.
 */
// ─── FUZZY COLUMN MATCHING ────────────────────────────────────────────────────
/**
 * Normalise a header string for fuzzy matching:
 * lowercase, remove spaces/underscores/hyphens/dots, trim.
 */
function normaliseHeader(h) {
  // Strip BOM, zero-width spaces, and other invisible chars, then normalise
  var s = safeStr(h).replace(/^[\uFEFF\u200B\u200C\u200D\u00A0]+/, '');
  return s.toLowerCase().replace(/[\s_\-\.]+/g, '');
}

/**
 * Build a lookup map: normalisedHeader -> original header from file.
 * Lets us find columns regardless of spacing/case differences.
 */
function buildHeaderMap(fileHeaders) {
  const map = {};
  fileHeaders.forEach(function(h) { map[normaliseHeader(h)] = h; });
  return map;
}

/**
 * Get value from a row using fuzzy column name matching.
 */
function fuzzyGet(row, headerMap, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var key = normaliseHeader(candidates[i]);
    if (headerMap[key] !== undefined) {
      var val = row[headerMap[key]];
      if (val !== undefined && val !== null && val !== '') return val;
    }
  }
  return '';
}

// Minimum required columns for gatepass (fuzzy matched)
// Each entry is an array of accepted names for that column
const GATEPASS_REQUIRED_COLS = [
  ['Gatepass Code', 'GP Code', 'GatepassCode', 'Gatepass No', 'GP No'],
  ['Item SkuCode', 'Item SKU Code', 'SKU Code', 'SkuCode', 'SKU', 'Item Sku Code'],
  ['Quantity', 'Qty'],
  ['To Party', 'ToParty', 'To'],
];

// Minimum required columns for inventory (fuzzy matched)
const INVENTORY_REQUIRED_COLS = [
  ['Item Type SKU Code', 'Item SkuCode', 'SKU Code', 'SkuCode', 'SKU'],
  ['Quantity', 'Qty'],
  ['Batch Code', 'BatchCode', 'Batch'],
  ['Manufacturing Date', 'Manufacturing', 'Mfg Date', 'MfgDate', 'Manufacture Date', 'Manufacturing date'],
  ['Expiry', 'Expiry Date', 'ExpiryDate', 'Exp Date'],
];

function validateGatepassColumns(fileHeaders) {
  const hmap = buildHeaderMap(fileHeaders);
  const missing = [];
  GATEPASS_REQUIRED_COLS.forEach(function(candidates) {
    const found = candidates.some(function(c) { return hmap[normaliseHeader(c)] !== undefined; });
    if (!found) missing.push(candidates[0]);
  });
  return { valid: missing.length === 0, missing: missing };
}

function validateInventoryColumns(fileHeaders) {
  const hmap = buildHeaderMap(fileHeaders);
  const missing = [];
  INVENTORY_REQUIRED_COLS.forEach(function(candidates) {
    const found = candidates.some(function(c) { return hmap[normaliseHeader(c)] !== undefined; });
    if (!found) missing.push(candidates[0]);
  });
  return { valid: missing.length === 0, missing: missing };
}

// ─── COLUMN MAPS ─────────────────────────────────────────────────────────────
const INVENTORY_HEADERS = SCHEMA.Inventory_Current.headers;
const GATEPASS_HEADERS  = SCHEMA.Gatepass_Data.headers;
const VIOLATION_HEADERS = SCHEMA.FIFO_Violations.headers;
const COMPLIANCE_HEADERS = SCHEMA.User_Compliance.headers;
const AUDIT_HEADERS     = ['ID', 'Timestamp', 'Event Type', 'Description', 'User', 'Facility', 'Metadata'];
const SETTINGS_HEADERS  = SCHEMA.Settings.headers;
const USERS_HEADERS     = SCHEMA.Users.headers;


// =============================================================
// AUTHENTICATION
// =============================================================

/**
 * Auth.gs
 * Handles user authentication, session token management,
 * and password hashing using SHA-256 via Utilities.
 */

// Session tokens are stored in Apps Script Cache (per-user scope not available
// in web app context, so we use Script Cache with a keyed token map).
const TOKEN_TTL_SECONDS = 28800; // 8 hours

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Authenticates a user against the Users sheet.
 * Returns session token on success.
 */
function authenticateUser(params) {
  const { email, password } = params || {};

  if (!email || !password) {
    return { success: false, error: 'MISSING_CREDENTIALS', message: 'Email and password are required.' };
  }

  try {
    const users = getSheetData(SHEETS.USERS);
    const normalizedEmail = safeStr(email).toLowerCase();

    const user = users.find(u =>
      safeStr(u['Email']).toLowerCase() === normalizedEmail
    );

    if (!user) {
      writeAuditLog({ eventType: 'LOGIN_FAILED', description: `Unknown email: ${email}`, user: email, facility: '' });
      return { success: false, error: 'INVALID_CREDENTIALS', message: 'Invalid email or password.' };
    }

    if (String(user['Active']).toLowerCase() === 'false' || user['Active'] === false) {
      writeAuditLog({ eventType: 'LOGIN_FAILED', description: `Inactive user: ${email}`, user: email, facility: '' });
      return { success: false, error: 'ACCOUNT_INACTIVE', message: 'Your account is inactive. Contact admin.' };
    }

    const inputHash = hashPassword(safeStr(password));
    if (inputHash !== safeStr(user['Password Hash'])) {
      writeAuditLog({ eventType: 'LOGIN_FAILED', description: `Wrong password: ${email}`, user: email, facility: '' });
      return { success: false, error: 'INVALID_CREDENTIALS', message: 'Invalid email or password.' };
    }

    // Generate session token
    const token   = generateToken(email);
    const session = {
      email    : normalizedEmail,
      name     : safeStr(user['Name']),
      role     : safeStr(user['Role']),
      facility : safeStr(user['Facility'])
    };

    // Store token in cache
    storeToken(token, session);

    // Update last login
    updateLastLogin(normalizedEmail);

    writeAuditLog({
      eventType   : 'LOGIN_SUCCESS',
      description : `User logged in: ${user['Name']}`,
      user        : email,
      facility    : session.facility
    });

    return {
      success : true,
      token   : token,
      user    : {
        name     : session.name,
        email    : session.email,
        role     : session.role,
        facility : session.facility
      }
    };
  } catch (err) {
    logSystemError('authenticateUser', err);
    return { success: false, error: 'SERVER_ERROR', message: err.message };
  }
}

/**
 * Validates a session token.
 * Returns { valid, session } where session = { email, name, role, facility }.
 */
function validateToken(token) {
  if (!token) return { valid: false };
  try {
    const cache   = CacheService.getScriptCache();
    const cacheKey = `token_${token}`;
    const stored  = cache.get(cacheKey);
    if (!stored) return { valid: false };
    const session = JSON.parse(stored);
    return { valid: true, session };
  } catch (_) {
    return { valid: false };
  }
}

// ─── USER MANAGEMENT ─────────────────────────────────────────────────────────
/**
 * Creates or updates a user.
 * If email already exists, updates name/role/facility/active.
 * If new, creates with hashed password.
 */
function saveUser(params) {
  try {
    const { name, email, password, role, facility, active, requestedBy } = params;
    const sheet = getSheet(SHEETS.USERS);
    const data  = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailIdx = headers.indexOf('Email');
    const now = formatDateTime(new Date());

    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (safeStr(data[i][emailIdx]).toLowerCase() === safeStr(email).toLowerCase()) {
        existingRow = i;
        break;
      }
    }

    if (existingRow > 0) {
      // Update existing
      const rowRange = sheet.getRange(existingRow + 1, 1, 1, headers.length);
      const rowVals  = rowRange.getValues()[0];
      const setVal   = (field, val) => {
        const idx = headers.indexOf(field);
        if (idx >= 0) rowVals[idx] = val;
      };
      setVal('Name', safeStr(name));
      setVal('Role', safeStr(role));
      setVal('Facility', safeStr(facility));
      setVal('Active', active !== false);
      if (password) setVal('Password Hash', hashPassword(password));
      rowRange.setValues([rowVals]);
    } else {
      // Create new
      const newRow = headers.map(h => {
        switch (h) {
          case 'Name':          return safeStr(name);
          case 'Email':         return safeStr(email).toLowerCase();
          case 'Password Hash': return hashPassword(safeStr(password));
          case 'Role':          return safeStr(role) || 'viewer';
          case 'Facility':      return safeStr(facility);
          case 'Active':        return true;
          case 'Created At':    return now;
          case 'Last Login':    return '';
          default:              return '';
        }
      });
      sheet.appendRow(newRow);
    }

    writeAuditLog({
      eventType   : 'USER_CHANGE',
      description : `User ${existingRow > 0 ? 'updated' : 'created'}: ${email}`,
      user        : requestedBy || 'admin',
      facility    : facility
    });

    return { success: true };
  } catch (err) {
    logSystemError('saveUser', err);
    return { success: false, error: err.message };
  }
}

// ─── PRIVATE HELPERS ──────────────────────────────────────────────────────────
/**
 * SHA-256 hash using Utilities.computeDigest.
 */
function hashPassword(password) {
  const bytes  = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    safeStr(password),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Generates a random 32-char hex token.
 */
function generateToken(email) {
  const rand  = Math.random().toString(36).slice(2) + Date.now().toString(36);
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    email + rand,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Stores a token → session mapping in cache.
 */
function storeToken(token, session) {
  const cache = CacheService.getScriptCache();
  cache.put(`token_${token}`, JSON.stringify(session), TOKEN_TTL_SECONDS);
}

/**
 * Updates Last Login timestamp for a user.
 */
function updateLastLogin(email) {
  try {
    const sheet   = getSheet(SHEETS.USERS);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailIdx = headers.indexOf('Email');
    const loginIdx = headers.indexOf('Last Login');
    if (emailIdx < 0 || loginIdx < 0) return;

    for (let i = 1; i < data.length; i++) {
      if (safeStr(data[i][emailIdx]).toLowerCase() === email) {
        sheet.getRange(i + 1, loginIdx + 1).setValue(formatDateTime(new Date()));
        break;
      }
    }
  } catch (_) {}
}


// =============================================================
// CACHE ENGINE
// =============================================================

/**
 * CacheEngine.gs
 * Thin wrapper around Apps Script CacheService (ScriptCache).
 * Also uses the Cache_Data sheet as a persistent fallback for large payloads
 * that exceed CacheService's 100KB per-key limit.
 *
 * Cache keys used:
 *   dashboard_summary    — home page KPI block
 *   financial_impact     — financial page data
 *   user_compliance      — compliance page user list
 */

const CACHE_MAX_BYTES = 90000; // Stay under 100KB CacheService limit

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Returns cached value for key, or null if not found / expired.
 */
function getCached(key) {
  try {
    // Try CacheService first (fast)
    const cache  = CacheService.getScriptCache();
    const stored = cache.get(`fifo_${key}`);
    if (stored) {
      const parsed = JSON.parse(stored);
      return parsed;
    }
  } catch (_) {}

  // Fallback: check Cache_Data sheet
  return getCachedFromSheet(key);
}

/**
 * Stores a value in cache with TTL (seconds).
 * Falls back to sheet storage for large payloads.
 */
function setCached(key, value, ttlSeconds) {
  const payload = JSON.stringify(value);
  const ttl     = ttlSeconds || CACHE_TTL;

  try {
    const cache = CacheService.getScriptCache();
    if (payload.length <= CACHE_MAX_BYTES) {
      cache.put(`fifo_${key}`, payload, ttl);
      return;
    }
  } catch (_) {}

  // Fallback: sheet cache
  setCachedInSheet(key, payload, ttl);
}

/**
 * Removes a specific key from all cache layers.
 */
function invalidateCacheKey(key) {
  try {
    CacheService.getScriptCache().remove(`fifo_${key}`);
  } catch (_) {}
  removeCachedFromSheet(key);
}

/**
 * Clears ALL FIFO dashboard cache entries.
 * Called after gatepass uploads and settings changes.
 */
function invalidateCache() {
  const keys = ['dashboard_summary', 'financial_impact', 'user_compliance'];
  keys.forEach(k => invalidateCacheKey(k));
}

/**
 * Manual cache invalidation — called from UI admin action.
 */
function invalidateCacheManual(params) {
  try {
    invalidateCache();
    writeAuditLog({
      eventType   : 'CACHE_INVALIDATED',
      description : 'Cache manually cleared by admin',
      user        : (params || {}).user || 'admin',
      facility    : ''
    });
    return { success: true, message: 'Cache cleared successfully.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─── SHEET CACHE FALLBACK ─────────────────────────────────────────────────────
function getCachedFromSheet(key) {
  try {
    const rows = getSheetData(SHEETS.CACHE);
    const now  = new Date();
    for (const row of rows) {
      if (safeStr(row['Cache Key']) !== key) continue;
      const expires = toDateObj(row['Expires At']);
      if (!expires || expires < now) continue;
      const val = safeStr(row['Value']);
      if (val) return JSON.parse(val);
    }
  } catch (_) {}
  return null;
}

function setCachedInSheet(key, payload, ttlSeconds) {
  try {
    const sheet   = getSheet(SHEETS.CACHE);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const keyIdx  = headers.indexOf('Cache Key');
    const now     = new Date();
    const expires = new Date(now.getTime() + ttlSeconds * 1000);

    // Find and update existing row
    for (let i = 1; i < data.length; i++) {
      if (safeStr(data[i][keyIdx]) === key) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[
          key,
          payload.slice(0, 49000), // Sheet cell limit ~50KB
          formatDateTime(now),
          formatDateTime(expires)
        ]]);
        return;
      }
    }

    // Append new row
    sheet.appendRow([key, payload.slice(0, 49000), formatDateTime(now), formatDateTime(expires)]);
  } catch (_) {}
}

function removeCachedFromSheet(key) {
  try {
    const sheet   = getSheet(SHEETS.CACHE);
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const keyIdx  = headers.indexOf('Cache Key');

    for (let i = data.length - 1; i >= 1; i--) {
      if (safeStr(data[i][keyIdx]) === key) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  } catch (_) {}
}// =============================================================
// AUDIT ENGINE
// =============================================================

/**
 * AuditEngine.gs
 * Writes structured audit events to the Audit_Log sheet.
 * Provides paginated read API for the Audit Log page.
 */

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Returns paginated, filtered audit log for the UI.
 */
function getAuditLog(params) {
  try {
    const {
      page      = 1,
      pageSize  = 50,
      search    = '',
      eventType = ''    // filter by specific event type
    } = params || {};

    let rows = getSheetData(SHEETS.AUDIT_LOG);

    // Filter by event type
    if (eventType) {
      rows = rows.filter(r => safeStr(r['Event Type']) === eventType);
    }

    // Search across description, user, facility
    if (search) {
      const q = search.toLowerCase();
      rows = rows.filter(r =>
        safeStr(r['Description']).toLowerCase().includes(q) ||
        safeStr(r['User']).toLowerCase().includes(q) ||
        safeStr(r['Facility']).toLowerCase().includes(q) ||
        safeStr(r['Event Type']).toLowerCase().includes(q)
      );
    }

    // Sort newest first
    rows.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

    const total    = rows.length;
    const start    = (page - 1) * pageSize;
    const pageRows = rows.slice(start, start + pageSize);

    return {
      success    : true,
      data       : pageRows.map(r => ({
        id         : safeStr(r['ID']),
        timestamp  : formatDateTime(r['Timestamp']),
        eventType  : safeStr(r['Event Type']),
        description: safeStr(r['Description']),
        user       : safeStr(r['User']),
        facility   : safeStr(r['Facility']),
        metadata   : safeStr(r['Metadata'])
      })),
      total,
      page      : safeNum(page),
      pageSize  : safeNum(pageSize),
      totalPages: Math.ceil(total / pageSize),
      eventTypes: getDistinctEventTypes()
    };
  } catch (err) {
    logSystemError('getAuditLog', err);
    return { success: false, error: err.message };
  }
}

// ─── WRITE API ────────────────────────────────────────────────────────────────
/**
 * Writes a single audit event to the Audit_Log sheet.
 * Called internally by all engine modules.
 *
 * params: { eventType, description, user, facility, metadata }
 */
function writeAuditLog(params) {
  try {
    const { eventType, description, user, facility, metadata } = params || {};
    const id        = generateAuditId();
    const timestamp = formatDateTime(new Date());

    const row = [
      id,
      timestamp,
      safeStr(eventType),
      safeStr(description).slice(0, 500),   // cap length
      safeStr(user),
      safeStr(facility),
      safeStr(metadata || '').slice(0, 500)
    ];

    getSheet(SHEETS.AUDIT_LOG).appendRow(row);
  } catch (_) {
    // Silently fail — audit logging must never break the main flow
  }
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function generateAuditId() {
  return `AUD-${Date.now().toString(36).toUpperCase()}`;
}

function getDistinctEventTypes() {
  try {
    const rows = getSheetData(SHEETS.AUDIT_LOG);
    const types = [...new Set(rows.map(r => safeStr(r['Event Type'])).filter(Boolean))];
    return types.sort();
  } catch (_) {
    return [];
  }
}

/**
 * Prunes Audit_Log to keep only the most recent N rows.
 * Run from a monthly trigger to prevent unbounded growth.
 */
function pruneAuditLog(keepRows) {
  keepRows = keepRows || 10000;
  try {
    const sheet   = getSheet(SHEETS.AUDIT_LOG);
    const lastRow = sheet.getLastRow();
    const dataRows = lastRow - 1;

    if (dataRows <= keepRows) return;

    const deleteCount = dataRows - keepRows;
    // Delete oldest rows (rows 2 through deleteCount+1)
    sheet.deleteRows(2, deleteCount);

    writeAuditLog({
      eventType   : 'AUDIT_PRUNED',
      description : `Audit log pruned: removed ${deleteCount} old entries`,
      user        : 'SYSTEM',
      facility    : ''
    });
  } catch (err) {
    logSystemError('pruneAuditLog', err);
  }
}


// =============================================================
// DataEngine.gs — FIFO · Violations · Compliance · Facility · Financial
// =============================================================

/**
 * FifoEngine.gs
 * Core FIFO logic:
 * - Loads inventory snapshot
 * - Groups by Facility + SKU Code
 * - Ranks batches oldest-first by Manufacturing Date, then Expiry
 * - Provides batch lookup for violation detection
 */

// ─── MAIN FIFO DATA BUILDER ───────────────────────────────────────────────────
/**
 * Builds a FIFO map from the current Inventory_Current sheet.
 * Returns a Map keyed by "Facility|SKUCode" → sorted array of batch objects.
 *
 * Batch object: {
 *   facility, skuCode, itemName, shelf,
 *   batchCode, vendorBatch, mfgDate, expiryDate,
 *   quantity, mrp, fifoRank
 * }
 */
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 6. FIFO ENGINE  (complete rewrite)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ─── FACILITY DETECTION FROM FROM-PARTY ───────────────────────────────────────
/**
 * Maps the gatepass "From Party" field to a known facility name.
 * Case-insensitive substring matching.
 */
function detectFacility(fromParty) {
  var fp = safeStr(fromParty).toLowerCase();
  if (fp.indexOf('mother hub') >= 0 || fp.indexOf('motherhub') >= 0) return 'SL Mother Hub';
  if (fp.indexOf('ambient') >= 0) return 'SL Ambient';
  // Try direct match as fallback
  for (var i = 0; i < FACILITIES.length; i++) {
    if (fp === FACILITIES[i].toLowerCase()) return FACILITIES[i];
  }
  return '';  // unknown facility — will be excluded
}

// ─── STEP 1-4: BUILD FIFO MAP FROM INVENTORY ──────────────────────────────────
/**
 * Builds a FIFO map from Inventory_Current sheet.
 *
 * STEP 1: Filter to SL Ambient + SL Mother Hub only.
 * STEP 2: Consolidate by Facility + SKU + Vendor Batch (sum qty across shelves).
 *         Store shelf details separately for drilldown.
 * STEP 3: Create FIFO_KEY = Expiry Date + Mfg Date + Vendor Batch.
 * STEP 4: Sort by FIFO_KEY ASC, assign rank (1 = oldest/soonest expiry).
 *
 * Returns Map keyed by "Facility|SKUCode" → sorted array of consolidated batch objects.
 */

// ─── PAGINATED VIOLATIONS API ─────────────────────────────────────────────────
/**
 * Returns paginated, filtered violation data for the violations page.
 */
function getViolations(params) {
  try {
    const {
      page     = 1,
      pageSize = 25,
      tab      = 'all',        // all | violations | partial | compliant
      search   = '',
      facility = '',
      severity = '',
      scope    = '',
      dateFrom = '',
      dateTo   = '',
      sortCol  = 'Date',
      sortDir  = 'desc'
    } = params || {};

    const range = getReportDateRange_({ scope: scope, dateFrom: dateFrom, dateTo: dateTo });
    let rows = getSheetData(SHEETS.VIOLATIONS);

    // Tab filter
    if (tab === 'violations') rows = rows.filter(function(r) {
      var s = safeStr(r['Status']).toUpperCase();
      return s === 'FIFO VIOLATION' || s === 'VIOLATION';
    });
    else if (tab === 'partial') rows = rows.filter(function(r) {
      var s = safeStr(r['Status']).toUpperCase();
      return s === 'PARTIAL FIFO POSSIBLE' || s === 'PARTIAL FIFO';
    });
    else if (tab === 'compliant') rows = rows.filter(function(r) {
      var s = safeStr(r['Status']).toUpperCase();
      return s === 'COMPLIANT';
    });

    // Facility filter
    if (facility) rows = rows.filter(r => safeStr(r['Facility']) === facility);

    // Severity filter
    if (severity) rows = rows.filter(r => safeStr(r['Severity']) === severity);

    // Date range
    rows = filterRowsByScope_(rows, range);

    // Search (GP, user, SKU, item name)
    if (search) {
      const q = search.toLowerCase();
      rows = rows.filter(r =>
        safeStr(r['GP Number']).toLowerCase().includes(q) ||
        safeStr(r['User']).toLowerCase().includes(q) ||
        safeStr(r['SKU Code']).toLowerCase().includes(q) ||
        safeStr(r['Item Name']).toLowerCase().includes(q)
      );
    }

    // Sort
    rows = sortRows(rows, sortCol, sortDir);

    // Paginate
    const total    = rows.length;
    const start    = (page - 1) * pageSize;
    const pageRows = rows.slice(start, start + pageSize);

    return {
      success : true,
      data    : pageRows.map(mapViolationRow),
      total,
      page    : safeNum(page),
      pageSize: safeNum(pageSize),
      totalPages: Math.ceil(total / pageSize)
    };
  } catch (err) {
    logSystemError('getViolations', err);
    return { success: false, error: err.message };
  }
}

/**
 * Returns full detail for a single violation (for drilldown modal).
 */
function getViolationDetail(params) {
  try {
    const { id } = params || {};
    if (!id) return { success: false, error: 'MISSING_ID' };

    const rows = getSheetData(SHEETS.VIOLATIONS);
    const row  = rows.find(r => safeStr(r['ID']) === safeStr(id));
    if (!row)  return { success: false, error: 'NOT_FOUND' };

    // Parse older batch details
    const olderBatches = parseOlderBatchDetails(safeStr(row['Skipped Batch Details']));

    return {
      success: true,
      data   : {
        ...mapViolationRow(row),
        olderBatches,
        skippedBatchDetails: safeStr(row['Skipped Batch Details'])
      }
    };
  } catch (err) {
    logSystemError('getViolationDetail', err);
    return { success: false, error: err.message };
  }
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function mapViolationRow(r) {
  return {
    id              : safeStr(r['ID']),
    date            : formatDate(r['Date']),
    gpNumber        : safeStr(r['GP Number']),
    user            : safeStr(r['User']),
    userEmail       : safeStr(r['User Email']),
    skuCode         : safeStr(r['SKU Code']),
    itemName        : safeStr(r['Item Name']),
    batchCode       : safeStr(r['Batch Code']),
    vendorBatch     : safeStr(r['Vendor Batch']),
    qty             : safeNum(r['Dispatched Qty']),
    dispatchedQty   : safeNum(r['Dispatched Qty']),
    shelf           : safeStr(r['Shelf']),
    olderBatchCount : safeNum(r['Older Batch Count']),
    olderInventoryQty: safeNum(r['Older Inventory Qty']),
    severity        : safeStr(r['Severity']),
    status          : safeStr(r['Status']),
    gpStatus        : safeStr(r['GP Status']),
    facility        : safeStr(r['Facility']),
    toParty         : safeStr(r['To Party']),
    lossOpportunity : safeNum(r['Loss Opportunity']),
    mfgDate         : formatDate(r['Manufacturing Date']),
    expiryDate      : formatDate(r['Expiry Date']),
    fifoRank        : safeNum(r['FIFO Rank']),
    skippedBatchDetails: safeStr(r['Skipped Batch Details']),
    uploadId        : safeStr(r['Upload ID'])
  };
}

function parseOlderBatchDetails(detailStr) {
  if (!detailStr) return [];
  return detailStr.split(';').map(function(part) {
    part = part.trim();
    if (!part) return null;
    // New format: vendorBatch|batchCode(Rank:N,Qty:N,Exp:DATE,MfgDate:DATE,Shelf:S,MRP:N)
    var newMatch = part.match(/^([^|]+)\|([^(]+)\(Rank:(\d+),Qty:(\d+),Exp:([^,]*),MfgDate:([^,]*),Shelf:([^,]*),MRP:([^)]*)\)$/);
    if (newMatch) {
      return {
        vendorBatch: newMatch[1],
        batchCode  : newMatch[2],
        fifoRank   : safeNum(newMatch[3]),
        quantity   : safeNum(newMatch[4]),
        expiryDate : newMatch[5],
        mfgDate    : newMatch[6],
        shelf      : newMatch[7],
        mrp        : safeNum(newMatch[8])
      };
    }
    // Legacy format: batchCode(Qty:N,Exp:DATE)
    var oldMatch = part.match(/^(.+?)\(Qty:(\d+),Exp:([^)]*)\)$/);
    if (oldMatch) {
      return { vendorBatch: oldMatch[1], batchCode: '', fifoRank: 0,
               quantity: safeNum(oldMatch[2]), expiryDate: oldMatch[3],
               mfgDate: '', shelf: '', mrp: 0 };
    }
    return null;
  }).filter(function(b) { return b !== null; });
}

function sortRows(rows, col, dir) {
  const colMap = {
    'Date'            : 'Date',
    'gpNumber'        : 'GP Number',
    'user'            : 'User',
    'skuCode'         : 'SKU Code',
    'severity'        : 'Severity',
    'status'          : 'Status',
    'facility'        : 'Facility',
    'lossOpportunity' : 'Loss Opportunity',
    'olderInventoryQty': 'Older Inventory Qty',
    'dispatchedQty'   : 'Dispatched Qty'
  };
  const sheetCol = colMap[col] || col;
  return rows.sort((a, b) => {
    let va = a[sheetCol], vb = b[sheetCol];
    // Numeric?
    const na = safeNum(va, NaN), nb = safeNum(vb, NaN);
    if (!isNaN(na) && !isNaN(nb)) {
      return dir === 'desc' ? nb - na : na - nb;
    }
    // Date?
    const da = toDateObj(va), db = toDateObj(vb);
    if (da && db) return dir === 'desc' ? db - da : da - db;
    // String
    va = safeStr(va).toLowerCase();
    vb = safeStr(vb).toLowerCase();
    if (va < vb) return dir === 'desc' ?  1 : -1;
    if (va > vb) return dir === 'desc' ? -1 :  1;
    return 0;
  });
}

function generateViolationId(gpCode, skuCode) {
  const ts   = Date.now().toString(36).toUpperCase();
  const code = safeStr(gpCode).replace(/\W/g, '').slice(0, 6).toUpperCase();
  return `VIO-${code}-${ts}`;
}// ─── EXPORT HELPER ────────────────────────────────────────────────────────────
function exportViolations(params) {
  try {
    const limit = safeNum((params || {}).limit, 5000);
    const rows  = getSheetData(SHEETS.VIOLATIONS)
      .sort((a, b) => new Date(b['Date']) - new Date(a['Date']))
      .slice(0, limit)
      .map(mapViolationRow);
    return { success: true, data: rows };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// =============================================================
// USER COMPLIANCE ENGINE
// =============================================================

/**
 * UserComplianceEngine.gs
 * Computes and stores per-user compliance metrics:
 * - Total GPs, violations, partial FIFO, compliant count
 * - Compliance %, skipped qty, loss opportunity
 * - Repeat offender classification (5+ violations in 30 days)
 * - FIFO Champion classification (50+ GPs, >98% compliance)
 * - Writes aggregated data to User_Compliance sheet
 */

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Returns all user compliance data for the Compliance page.
 */
function getUserCompliance(params) {
  try {
    const range = getReportDateRange_(params);
    var scopedRows = filterRowsByScope_(getSheetData(SHEETS.VIOLATIONS), range);
    const data = computeUserCompliance(scopedRows);
    return { success: true, allUsers: data };
  } catch (err) {
    logSystemError('getUserCompliance', err);
    return { success: false, error: err.message };
  }
}

/**
 * Triggered after each gatepass upload.
 * Recomputes and persists user compliance metrics.
 */
function recomputeUserCompliance() {
  try {
    const data = computeUserCompliance();

    // Persist to User_Compliance sheet
    const now = formatDateTime(new Date());
    const settings = loadSettings();
    const offenderDays = safeNum(settings['repeat_offender_days'], 30);
    const offenderMin  = safeNum(settings['repeat_offender_min'], 5);
    const championGps  = safeNum(settings['champion_min_gps'], 50);
    const championPct  = safeNum(settings['champion_min_compliance'], 98);

    const rows = data.map(u => ({
      'User Name'             : u.userName,
      'User Email'            : u.email,
      'Facility'              : u.facility,
      'Total GPs'             : u.totalGps,
      'Violations'            : u.violations,
      'Partial FIFO'          : u.partialFifo,
      'Compliant'             : u.compliant,
      'Compliance %'          : u.compliancePct,
      'Skipped Qty'           : u.skippedQty,
      'Loss Opportunity'      : u.lossOpportunity,
      'Last Violation Date'   : u.lastViolationDate,
      'Violations Last 30 Days': u.violations30d,
      'Offender Level'        : u.offenderLevel,
      'Is Champion'           : u.isChampion,
      'Last Updated'          : now
    }));

    replaceSheetData(SHEETS.COMPLIANCE, COMPLIANCE_HEADERS, rows);
    invalidateCacheKey('user_compliance');
    invalidateCacheKey('dashboard_summary');

    return { success: true, userCount: data.length };
  } catch (err) {
    logSystemError('recomputeUserCompliance', err);
    return { success: false, error: err.message };
  }
}

// ─── CORE COMPUTATION ─────────────────────────────────────────────────────────
function computeUserCompliance(scopedRows) {
  const violations = scopedRows || getSheetData(SHEETS.VIOLATIONS);
  const settings   = loadSettings();
  const offenderDays = safeNum(settings['repeat_offender_days'], 30);
  const offenderMin  = safeNum(settings['repeat_offender_min'], 5);
  const championGps  = safeNum(settings['champion_min_gps'], 50);
  const championPct  = safeNum(settings['champion_min_compliance'], 98);
  const today        = new Date();

  // Group all violation rows by user name
  const userMap = new Map();

  for (const row of violations) {
    const user     = safeStr(row['User']) || 'Unknown';
    const email    = safeStr(row['User Email']);
    const facility = safeStr(row['Facility']);
    const status   = safeStr(row['Status']);
    const date     = toDateObj(row['Date']);
    const skipped  = safeNum(row['Older Inventory Qty'], 0);
    const loss     = safeNum(row['Loss Opportunity'], 0);

    if (!userMap.has(user)) {
      userMap.set(user, {
        userName        : user,
        email           : email,
        facility        : facility,
        totalGps        : 0,       // unique GP count
        violations      : 0,
        partialFifo     : 0,
        compliant       : 0,
        skippedQty      : 0,
        lossOpportunity : 0,
        lastViolationDate: null,
        gpSet           : new Set(),
        violationDates  : []       // for 30-day window calc
      });
    }

    const u = userMap.get(user);

    // Update email and facility from most recent record (may not be set initially)
    if (email && !u.email) u.email    = email;
    if (facility)          u.facility = facility; // take latest

    // Count unique GPs
    u.gpSet.add(safeStr(row['GP Number']));

    if (status === 'FIFO VIOLATION') {
      u.violations++;
      u.skippedQty      += skipped;
      u.lossOpportunity += loss;
      if (date) {
        u.violationDates.push(date);
        if (!u.lastViolationDate || date > u.lastViolationDate) {
          u.lastViolationDate = date;
        }
      }
    } else if (status === 'PARTIAL FIFO POSSIBLE') {
      u.partialFifo++;
    } else {
      u.compliant++;
    }
  }

  // Build result array
  const result = [];

  for (const [, u] of userMap) {
    u.totalGps = u.gpSet.size;

    // Compliance % = compliant / (compliant + violations + partialFifo)
    const evaluated  = u.compliant + u.violations + u.partialFifo;
    const compPct    = evaluated > 0 ? Math.round((u.compliant / evaluated) * 10000) / 100 : 100;

    // Violations in last N days
    const cutoff     = new Date(today - offenderDays * 86400000);
    const v30d       = u.violationDates.filter(d => d >= cutoff).length;

    // Offender level
    let offenderLevel = 'None';
    if (v30d >= offenderMin) offenderLevel = 'Repeat Offender';

    // Champion: enough GPs AND high compliance AND no repeat offender tag
    const isChampion = u.totalGps >= championGps &&
                       compPct >= championPct &&
                       offenderLevel === 'None' &&
                       u.violations === 0;

    result.push({
      userName         : u.userName,
      email            : u.email,
      facility         : u.facility,
      totalGps         : u.totalGps,
      violations       : u.violations,
      partialFifo      : u.partialFifo,
      compliant        : u.compliant,
      compliancePct    : compPct,
      skippedQty       : u.skippedQty,
      lossOpportunity  : u.lossOpportunity,
      lastViolationDate: u.lastViolationDate ? formatDate(u.lastViolationDate) : '',
      violations30d    : v30d,
      offenderLevel,
      isChampion
    });
  }

  // Sort by violations desc, then by loss desc
  result.sort((a, b) => b.violations - a.violations || b.lossOpportunity - a.lossOpportunity);

  return result;
}

// ─── QUICK SUMMARIES (for dashboard) ─────────────────────────────────────────
/**
 * Returns repeat offender count and champion count.
 * Uses cached User_Compliance sheet (cheaper than recomputing).
 */
function getOffenderChampionCounts() {
  try {
    const rows = getSheetData(SHEETS.COMPLIANCE);
    let repeatOffenders = 0, champions = 0;
    for (const r of rows) {
      if (safeStr(r['Offender Level']) === 'Repeat Offender') repeatOffenders++;
      if (String(r['Is Champion']).toLowerCase() === 'true')  champions++;
    }
    return { repeatOffenders, champions };
  } catch (_) {
    return { repeatOffenders: 0, champions: 0 };
  }
}

/**
 * Returns top N violators (by violations count).
 */
function getTopViolators(topN) {
  try {
    const rows = getSheetData(SHEETS.COMPLIANCE);
    return rows
      .filter(r => safeNum(r['Violations']) > 0)
      .sort((a, b) => safeNum(b['Violations']) - safeNum(a['Violations']) ||
                      safeNum(b['Loss Opportunity']) - safeNum(a['Loss Opportunity']))
      .slice(0, topN)
      .map(r => ({
        userName        : safeStr(r['User Name']),
        facility        : safeStr(r['Facility']),
        violations      : safeNum(r['Violations']),
        lossOpportunity : safeNum(r['Loss Opportunity']),
        compliancePct   : safeNum(r['Compliance %']),
        offenderLevel   : safeStr(r['Offender Level'])
      }));
  } catch (_) {
    return [];
  }
}// ─── EXPORT HELPER ────────────────────────────────────────────────────────────
function exportCompliance(params) {
  try {
    const data = computeUserCompliance();
    return { success: true, data };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// =============================================================
// FACILITY ENGINE + SETTINGS + DASHBOARD SUMMARY
// =============================================================

/**
 * FacilityEngine.gs
 * Computes facility-level compliance and inventory aggregations.
 * Powers the home dashboard summary KPI block.
 */

// ─── DASHBOARD SUMMARY ────────────────────────────────────────────────────────
/**
 * Main dashboard KPI aggregation function.
 * Returns all data needed for the home page in a single call.
 */
function getDashboardSummary(params) {
  try {
    const range = getReportDateRange_(params);
    const data = computeDashboardSummary(range);
    return { success: true, data };
  } catch (err) {
    logSystemError('getDashboardSummary', err);
    return { success: false, error: err.message };
  }
}

function testDashboardSummary() {
  var result = getDashboardSummary({ scope: 'today' });
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function getBuildInfo(params) {
  return {
    success: true,
    appVersion: APP_VERSION,
    buildId: BUILD_ID,
    timezone: APP_TIMEZONE,
    serverToday: formatDate(new Date()),
    received: params || {}
  };
}

function debugDateScopes(params) {
  var rows = getSheetData(SHEETS.VIOLATIONS);
  var todayRange = getReportDateRange_({ scope: 'today' });
  var yesterdayRange = getReportDateRange_({ scope: 'yesterday' });
  var mtdRange = getReportDateRange_({ scope: 'mtd' });

  function countByDate(sourceRows) {
    var counts = {};
    sourceRows.forEach(function(r) {
      var key = formatDate(r['Date']) || '(blank)';
      counts[key] = (counts[key] || 0) + 1;
    });
    return counts;
  }

  function statusRows(sourceRows, tab) {
    if (tab === 'violations') return sourceRows.filter(function(r) {
      var s = safeStr(r['Status']).toUpperCase();
      return s === 'FIFO VIOLATION' || s === 'VIOLATION';
    });
    if (tab === 'partial') return sourceRows.filter(function(r) {
      var s = safeStr(r['Status']).toUpperCase();
      return s === 'PARTIAL FIFO POSSIBLE' || s === 'PARTIAL FIFO';
    });
    if (tab === 'compliant') return sourceRows.filter(function(r) {
      return safeStr(r['Status']).toUpperCase() === 'COMPLIANT';
    });
    return sourceRows;
  }

  function scopeCounts(sourceRows) {
    return {
      today: filterRowsByScope_(sourceRows, todayRange).length,
      yesterday: filterRowsByScope_(sourceRows, yesterdayRange).length,
      mtd: filterRowsByScope_(sourceRows, mtdRange).length
    };
  }

  var tabs = ['all', 'violations', 'partial', 'compliant'];
  var byTab = {};
  tabs.forEach(function(tab) {
    var scopedRows = statusRows(rows, tab);
    byTab[tab] = {
      totalRows: scopedRows.length,
      byDate: countByDate(scopedRows),
      scopes: scopeCounts(scopedRows)
    };
  });

  var sampleRows = rows.slice(0, 15).map(function(r) {
    return {
      rawDate: r['Date'],
      formattedDate: formatDate(r['Date']),
      comparableKeys: getComparableDateKeys_(r['Date']),
      gpNumber: safeStr(r['GP Number']),
      status: safeStr(r['Status']),
      user: safeStr(r['User']),
      sku: safeStr(r['SKU Code'])
    };
  });

  var result = {
    success: true,
    buildId: BUILD_ID,
    timezone: APP_TIMEZONE,
    serverNow: formatDateTime(new Date()),
    expectedRanges: {
      today: todayRange,
      yesterday: yesterdayRange,
      mtd: mtdRange
    },
    distinctDates: getDistinctActualRowDates_(rows),
    totalViolationSheetRows: rows.length,
    byTab: byTab,
    sampleRows: sampleRows
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Categorises GOOD/Active inventory into expiry buckets relative to today.
 * Returns value (qty × MRP) in each bucket.
 */
function calcExpiryBuckets() {
  var inventory = getSheetData(SHEETS.INVENTORY);
  var today     = todayStart();
  // Each bucket: { skuCount (unique SKUs), totalQty }
  var buckets = {
    b0_30  : { skuCount: 0, totalQty: 0, skus: {} },
    b31_60 : { skuCount: 0, totalQty: 0, skus: {} },
    b61_90 : { skuCount: 0, totalQty: 0, skus: {} },
    b91_180: { skuCount: 0, totalQty: 0, skus: {} },
    b180p  : { skuCount: 0, totalQty: 0, skus: {} }
  };

  for (var i = 0; i < inventory.length; i++) {
    var row      = inventory[i];
    var facility = safeStr(row['Facility']);
    if (FACILITIES.indexOf(facility) < 0) continue;

    // GOOD_INVENTORY + Active only
    var invType     = safeStr(row['Inventory Type'] || '');
    var batchStatus = safeStr(row['Batch Status']   || '');
    if (invType     && invType.toUpperCase()     !== 'GOOD_INVENTORY') continue;
    if (batchStatus && batchStatus.toUpperCase() !== 'ACTIVE')         continue;

    var qty    = safeNum(row['Quantity'], 0);
    var expiry = toDateObj(row['Expiry']);
    var sku    = safeStr(row['Item Type SKU Code']);
    if (!expiry || qty <= 0 || !sku) continue;

    var days = daysBetween(today, expiry);
    if (days === null || days < 0) continue; // exclude expired

    var bucket = null;
    if      (days <= 30)  bucket = buckets.b0_30;
    else if (days <= 60)  bucket = buckets.b31_60;
    else if (days <= 90)  bucket = buckets.b61_90;
    else if (days <= 180) bucket = buckets.b91_180;
    else                  bucket = buckets.b180p;

    bucket.totalQty += qty;
    if (!bucket.skus[sku]) { bucket.skus[sku] = true; bucket.skuCount++; }
  }

  // Remove internal skus map before returning
  var result = {};
  var keys = Object.keys(buckets);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    result[key] = { skuCount: buckets[key].skuCount, totalQty: buckets[key].totalQty };
  }
  return result;
}

function computeDashboardSummary(params) {
  params = params || {};
  var violations = filterRowsByScope_(getSheetData(SHEETS.VIOLATIONS), params);
  var today      = todayStart();

  // Gatepass-level counts
  var allGpCodes    = {};
  var violationRows = [];
  var partialRows   = [];
  var compliantRows = [];

  for (var i = 0; i < violations.length; i++) {
    var r      = violations[i];
    var gp     = safeStr(r['GP Number']);
    var status = safeStr(r['Status']).toUpperCase();
    if (gp) allGpCodes[gp] = true;
    if (status === 'FIFO VIOLATION')         violationRows.push(r);
    else if (status === 'PARTIAL FIFO POSSIBLE') partialRows.push(r);
    else                                     compliantRows.push(r);
  }

  var totalGatepasses  = Object.keys(allGpCodes).length;
  var violationGps     = {};
  var partialGps       = {};
  var compliantGps     = {};

  for (var v = 0; v < violationRows.length; v++) violationGps[safeStr(violationRows[v]['GP Number'])] = true;
  for (var p = 0; p < partialRows.length;   p++) partialGps[safeStr(partialRows[p]['GP Number'])]   = true;
  for (var c = 0; c < compliantRows.length; c++) compliantGps[safeStr(compliantRows[c]['GP Number'])] = true;

  var fifoViolations   = Object.keys(violationGps).length;
  var partialFifoCases = Object.keys(partialGps).length;
  var compliantCount   = Object.keys(compliantGps).length;
  var fifoCompliancePct = totalGatepasses > 0
    ? Math.round((compliantCount / totalGatepasses) * 10000) / 100
    : 100;

  var criticalViolations = 0;
  var totalSkippedQty    = 0;
  for (var vr = 0; vr < violationRows.length; vr++) {
    if (safeStr(violationRows[vr]['Severity']) === 'Critical') criticalViolations++;
    totalSkippedQty += safeNum(violationRows[vr]['Older Inventory Qty'], 0);
  }

  // Inventory expiry exposure
  var inventory   = getSheetData(SHEETS.INVENTORY);
  var highRiskVal = 0;
  for (var inv = 0; inv < inventory.length; inv++) {
    var row    = inventory[inv];
    var fac    = safeStr(row['Facility']);
    if (FACILITIES.indexOf(fac) < 0) continue;
    var qty    = safeNum(row['Quantity'], 0);
    var mrp    = safeNum(row['MRP'], 0);
    var expiry = toDateObj(row['Expiry']);
    if (!expiry || qty <= 0) continue;
    var days = daysBetween(today, expiry);
    if (days !== null && days >= 0 && days <= 30) highRiskVal += qty * mrp;
  }

  var scopedCompliance = computeUserCompliance(violations);
  var repeatOffenders = 0;
  var fifoChampions = 0;
  for (var sc = 0; sc < scopedCompliance.length; sc++) {
    if (scopedCompliance[sc].offenderLevel === 'Repeat Offender') repeatOffenders++;
    if (scopedCompliance[sc].isChampion) fifoChampions++;
  }
  var facilityCompliance = computeFacilityCompliance(violations);
  var topViolators = scopedCompliance
    .filter(function(u) { return safeNum(u.violations) > 0; })
    .sort(function(a, b) { return safeNum(b.violations) - safeNum(a.violations) || safeNum(b.lossOpportunity) - safeNum(a.lossOpportunity); })
    .slice(0, 5);
  var champions = scopedCompliance
    .filter(function(u) { return u.isChampion; })
    .sort(function(a, b) { return safeNum(b.compliancePct) - safeNum(a.compliancePct) || safeNum(b.totalGps) - safeNum(a.totalGps); })
    .slice(0, 5);
  var expiryExposure = calcExpiryBuckets();

  // Recent violations
  var recentViolations = violationRows
    .sort(function(a, b) { return new Date(b['Date']) - new Date(a['Date']); })
    .slice(0, 10)
    .map(function(r) {
      return {
        id              : safeStr(r['ID']),
        gpNumber        : safeStr(r['GP Number']),
        gpStatus        : safeStr(r['GP Status']),
        user            : safeStr(r['User']),
        skuCode         : safeStr(r['SKU Code']),
        itemName        : safeStr(r['Item Name']),
        vendorBatch     : safeStr(r['Vendor Batch']),
        dispatchedQty   : safeNum(r['Dispatched Qty']),
        severity        : safeStr(r['Severity']),
        status          : safeStr(r['Status']),
        facility        : safeStr(r['Facility']),
        date            : formatDate(r['Date']),
        olderInventoryQty: safeNum(r['Older Inventory Qty'])
      };
    });

  return {
    totalGatepasses   : totalGatepasses,
    fifoViolations    : fifoViolations,
    partialFifoCases  : partialFifoCases,
    compliantGps      : compliantCount,
    fifoCompliancePct : fifoCompliancePct,
    criticalViolations: criticalViolations,
    totalSkippedQty   : totalSkippedQty,
    highRiskInventoryValue: highRiskVal,
    repeatOffenders   : repeatOffenders,
    fifoChampions     : fifoChampions,
    facilityCompliance: facilityCompliance,
    topViolators      : topViolators,
    champions         : champions,
    expiryExposure    : expiryExposure,
    recentViolations  : recentViolations
  };
}

function computeFacilityCompliance(violations) {
  const facilityMap = {};
  FACILITIES.forEach(f => {
    facilityMap[f] = {
      facility    : f,
      totalGps    : new Set(),
      violations  : 0,
      partialFifo : 0,
      compliant   : new Set(),
      skippedQty  : 0,
      lossOpp     : 0
    };
  });

  for (const row of violations) {
    const fac    = safeStr(row['Facility']);
    const gp     = safeStr(row['GP Number']);
    const status = safeStr(row['Status']);
    if (!facilityMap[fac]) continue;

    const f = facilityMap[fac];
    f.totalGps.add(gp);

    if (status === 'FIFO VIOLATION') {
      f.violations++;
      f.skippedQty += safeNum(row['Older Inventory Qty'], 0);
      f.lossOpp    += safeNum(row['Loss Opportunity'], 0);
    } else if (status === 'PARTIAL FIFO POSSIBLE') {
      f.partialFifo++;
    } else {
      f.compliant.add(gp);
    }
  }

  return FACILITIES.map(fac => {
    const f        = facilityMap[fac];
    const total    = f.totalGps.size;
    const compGps  = f.compliant.size;
    const compPct  = total > 0 ? Math.round((compGps / total) * 10000) / 100 : 100;

    return {
      facility       : fac,
      totalGps       : total,
      violations     : f.violations,
      partialFifo    : f.partialFifo,
      compliantGps   : compGps,
      compliancePct  : compPct,
      skippedQty     : f.skippedQty,
      lossOpportunity: f.lossOpp
    };
  });
}

// ─── SETTINGS ────────────────────────────────────────────────────────────────
/**
 * Loads all settings as a key→value map.
 */
function loadSettings() {
  try {
    const rows = getSheetData(SHEETS.SETTINGS);
    const map  = {};
    rows.forEach(r => {
      const key = safeStr(r['Key']);
      if (key) map[key] = safeStr(r['Value']);
    });
    return map;
  } catch (_) {
    return {};
  }
}

/**
 * Returns settings for the UI (includes labels and descriptions).
 */
function getSettings(params) {
  try {
    const rows = getSheetData(SHEETS.SETTINGS);
    const data = rows.map(r => ({
      key        : safeStr(r['Key']),
      value      : safeStr(r['Value']),
      label      : safeStr(r['Label']),
      description: safeStr(r['Description']),
      updatedBy  : safeStr(r['Updated By']),
      updatedAt  : formatDate(r['Updated At'])
    }));
    return { success: true, data };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Saves one or more settings from the UI.
 * params.settings = [{ key, value }]
 */
function saveSettings(params) {
  try {
    var updatedBy = (params || {}).updatedBy || 'admin';
    var settings  = (params || {}).settings;

    if (!settings) {
      return { success: false, error: 'INVALID_PARAMS', message: 'No settings provided.' };
    }

    var sheet   = getSheet(SHEETS.SETTINGS);
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var keyIdx  = headers.indexOf('Key');
    var valIdx  = headers.indexOf('Value');
    var byIdx   = headers.indexOf('Updated By');
    var atIdx   = headers.indexOf('Updated At');
    var now     = formatDateTime(new Date());

    // Accept BOTH formats:
    // Array: [{key:'min_inventory_threshold', value:'50'}, ...]
    // Object: {minInventoryThreshold: 50, ...} — map JS camelCase to sheet keys
    var keyMap = {
      'minInventoryThreshold'   : 'min_inventory_threshold',
      'repeatOffenderThreshold' : 'repeat_offender_min',
      'minorMin'                : 'severity_minor_min',
      'minorMax'                : 'severity_minor_max',
      'majorMin'                : 'severity_major_min',
      'majorMax'                : 'severity_major_max',
      'criticalMin'             : 'severity_critical_min',
      'championGpMin'           : 'champion_min_gps',
      'championComplianceMin'   : 'champion_min_compliance',
      'inventoryImportEmail'    : 'inventory_import_email',
      'repeatOffenderDays'      : 'repeat_offender_days'
    };

    var pairs = [];
    if (Array.isArray(settings)) {
      pairs = settings; // already [{key, value}]
    } else {
      // Convert object to pairs using keyMap
      var keys = Object.keys(settings);
      for (var ki = 0; ki < keys.length; ki++) {
        var jsKey    = keys[ki];
        var sheetKey = keyMap[jsKey] || jsKey;
        pairs.push({ key: sheetKey, value: String(settings[jsKey]) });
      }
    }

    var updated = [];
    for (var pi = 0; pi < pairs.length; pi++) {
      var key = safeStr(pairs[pi].key);
      var val = safeStr(pairs[pi].value);
      for (var i = 1; i < data.length; i++) {
        if (safeStr(data[i][keyIdx]) === key) {
          sheet.getRange(i + 1, valIdx + 1).setValue(val);
          if (byIdx >= 0) sheet.getRange(i + 1, byIdx + 1).setValue(updatedBy);
          if (atIdx >= 0) sheet.getRange(i + 1, atIdx + 1).setValue(now);
          updated.push(key);
          break;
        }
      }
    }

    invalidateCache();

    writeAuditLog({
      eventType   : 'SETTINGS_SAVED',
      description : 'Settings updated: ' + updated.join(', '),
      user        : updatedBy,
      facility    : ''
    });

    return { success: true, updated: updated };
  } catch (err) {
    logSystemError('saveSettings', err);
    return { success: false, error: err.message };
  }
}

function getPreventionData(params) {
  try {
    const {
      page     = 1,
      pageSize = 25,
      search   = '',
      facility = '',
      risk     = '',        // high | medium | low
      sortCol  = 'expiryDate',
      sortDir  = 'asc'
    } = params || {};

    const inventory  = getSheetData(SHEETS.INVENTORY);
    const today      = todayStart();
    let   atRiskRows = [];

    for (var pri = 0; pri < inventory.length; pri++) {
      var row = inventory[pri];
      var fac = safeStr(row['Facility']);
      if (FACILITIES.indexOf(fac) < 0) continue;
      if (facility && fac !== facility) continue;

      // GOOD_INVENTORY + Active only (consistent with FIFO logic)
      var invType     = safeStr(row['Inventory Type'] || '');
      var batchStatus = safeStr(row['Batch Status']   || '');
      if (invType     && invType.toUpperCase()     !== 'GOOD_INVENTORY') continue;
      if (batchStatus && batchStatus.toUpperCase() !== 'ACTIVE')         continue;

      var qty = safeNum(row['Quantity'], 0);
      if (qty <= 0) continue;

      var expiry = toDateObj(row['Expiry']);
      var days   = expiry ? daysBetween(today, expiry) : null;

      // Exclude expired items (days < 0)
      if (days === null || days < 0) continue;

      // Risk classification
      var riskLevel = 'Low';
      if (days <= 30)  riskLevel = 'High';
      else if (days <= 90) riskLevel = 'Medium';

      if (risk && riskLevel.toLowerCase() !== risk.toLowerCase()) continue;

      const skuCode  = safeStr(row['Item Type SKU Code']);
      const itemName = safeStr(row['Item Type Name']);

      // Search
      if (search) {
        const q = search.toLowerCase();
        if (!skuCode.toLowerCase().includes(q) && !itemName.toLowerCase().includes(q) &&
            !safeStr(row['Batch Code']).toLowerCase().includes(q)) continue;
      }

      atRiskRows.push({
        facility   : fac,
        skuCode,
        itemName,
        shelf      : safeStr(row['Shelf']),
        batchCode  : safeStr(row['Batch Code']),
        vendorBatch: safeStr(row['Vendor Batch Code']),
        mfgDate    : formatDate(row['Manufacturing Date']),
        expiryDate : formatDate(row['Expiry']),
        quantity   : qty,
        mrp        : safeNum(row['MRP'], 0),
        value      : qty * safeNum(row['MRP'], 0),
        daysToExpiry: days,
        riskLevel
      });
    }

    // Sort
    atRiskRows.sort((a, b) => {
      let va = a[sortCol], vb = b[sortCol];
      const na = safeNum(va, NaN), nb = safeNum(vb, NaN);
      if (!isNaN(na) && !isNaN(nb)) return sortDir === 'desc' ? nb - na : na - nb;
      const da = toDateObj(va), db = toDateObj(vb);
      if (da && db) return sortDir === 'desc' ? db - da : da - db;
      va = safeStr(va).toLowerCase(); vb = safeStr(vb).toLowerCase();
      return sortDir === 'desc' ? vb.localeCompare(va) : va.localeCompare(vb);
    });

    const total     = atRiskRows.length;
    const start     = (page - 1) * pageSize;
    const pageItems = atRiskRows.slice(start, start + pageSize);

    // Expiry exposure buckets
    const expiryExposure = calcExpiryBuckets();

    return {
      success     : true,
      data        : pageItems,
      total,
      page        : safeNum(page),
      pageSize    : safeNum(pageSize),
      totalPages  : Math.ceil(total / pageSize),
      expiryExposure
    };
  } catch (err) {
    logSystemError('getPreventionData', err);
    return { success: false, error: err.message };
  }
}

// ─── EXPORT HELPERS ───────────────────────────────────────────────────────────
function exportPrevention(params) {
  try {
    const result = getPreventionData({ page: 1, pageSize: 10000 });
    return { success: true, data: result.data || [] };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// =============================================================
// ImportEngine.gs — Gatepass Upload · Inventory Import · Reports/Exports
// =============================================================

/**
 * GatepassEngine.gs
 * Handles gatepass file upload (CSV/XLSX via base64),
 * parses rows, filters B2B, stores to Gatepass_Data sheet,
 * then triggers FIFO violation detection.
 */

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Called from frontend when user uploads a gatepass file.
 * params: { fileData (base64), fileName, fileType, dateOverride, uploadedBy, facility }
 */
function uploadGatepass(params) {
  var fileData    = (params || {}).fileData;
  var fileName    = (params || {}).fileName;
  var fileType    = (params || {}).fileType;
  var dateOverride= (params || {}).dateOverride;
  var uploadedBy  = (params || {}).uploadedBy || 'unknown';
  var replaceAll  = (params || {}).replaceAll === true;  // NEW: flush existing data

  if (!fileData || !fileName) {
    return { success: false, error: 'MISSING_FILE', message: 'File data is required.' };
  }

  try {
    var bytes  = Utilities.base64Decode(fileData);
    var parsed = parseGatepassFile(fileName, bytes);
    if (!parsed.success) return parsed;

    var validation = validateGatepassColumns(parsed.headers);
    if (!validation.valid) {
      return {
        success: false,
        error: 'INVALID_COLUMNS',
        message: 'Missing required columns: ' + validation.missing.join(', ') +
                 '. Your file has: ' + parsed.headers.join(', ')
      };
    }

    var uploadId   = generateUploadId();
    var uploadDate = dateOverride ? formatDate(new Date(dateOverride)) : formatDate(new Date());

    // Check inventory data exists before processing (FIFO needs inventory)
    var invSheet = getSheet(SHEETS.INVENTORY);
    if (invSheet.getLastRow() < 2) {
      return {
        success: false,
        error: 'NO_INVENTORY',
        message: 'Please upload your Inventory file first before uploading Gatepasses. FIFO violation detection requires inventory data.'
      };
    }

    var processed = processGatepassRows(parsed.rows, uploadId, uploadDate, '');

    if (processed.rows.length === 0) {
      return {
        success: false, error: 'NO_DATA',
        message: 'No valid rows after filtering. B2B skipped: ' + processed.skippedB2B +
                 ', Invalid/unknown facility: ' + processed.skippedInvalid
      };
    }

    // If replaceAll: clear Gatepass_Data AND Violations sheets first
    if (replaceAll) {
      var gpSheet  = getSheet(SHEETS.GATEPASS);
      var vlSheet  = getSheet(SHEETS.VIOLATIONS);
      var gpLast   = gpSheet.getLastRow();
      var vlLast   = vlSheet.getLastRow();
      if (gpLast > 1) gpSheet.deleteRows(2, gpLast - 1);
      if (vlLast > 1) vlSheet.deleteRows(2, vlLast - 1);
      writeAuditLog({
        eventType  : 'DATA_FLUSHED',
        description: 'All gatepass and violation data cleared before upload by ' + uploadedBy,
        user       : uploadedBy, facility: 'ALL'
      });
    }

    appendSheetRows(SHEETS.GATEPASS, GATEPASS_HEADERS, processed.rows);

    // Recompute violations for this upload
    var violationResult = processViolationsForUpload(uploadId, processed.rows, '');
    recomputeUserCompliance();
    invalidateCache();

    writeAuditLog({
      eventType   : 'GATEPASS_UPLOAD',
      description : (replaceAll ? '[REPLACE] ' : '') + 'Uploaded ' + processed.rows.length + ' rows from ' + fileName,
      user        : uploadedBy, facility: 'ALL',
      metadata    : 'UploadID:' + uploadId + ' | B2B:' + processed.skippedB2B +
                    ' | Invalid:' + processed.skippedInvalid +
                    ' | Violations:' + violationResult.violationCount
    });

    return {
      success       : true,
      uploadId      : uploadId,
      totalRows     : parsed.rows.length,
      processedRows : processed.rows.length,
      skippedB2B    : processed.skippedB2B,
      skippedInvalid: processed.skippedInvalid,
      violationCount: violationResult.violationCount,
      partialCount  : violationResult.partialCount,
      compliantCount: violationResult.compliantCount
    };
  } catch (err) {
    logSystemError('uploadGatepass', err);
    return { success: false, error: 'UPLOAD_ERROR', message: err.message };
  }
}
function getGatepassHistory(params) {
  const limit = safeNum((params || {}).limit, 10);
  try {
    const logs = getSheetData(SHEETS.AUDIT_LOG);
    const history = logs
      .filter(r => safeStr(r['Event Type']) === 'GATEPASS_UPLOAD')
      .sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']))
      .slice(0, limit)
      .map(r => {
        const meta = parseMetaString(safeStr(r['Metadata']));
        return {
          timestamp  : formatDateTime(r['Timestamp']),
          uploadId   : meta['UploadID'] || '',
          user       : safeStr(r['User']),
          facility   : safeStr(r['Facility']),
          description: safeStr(r['Description'])
        };
      });

    return { success: true, data: history };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ─── FILE PARSING ─────────────────────────────────────────────────────────────
function parseGatepassFile(fileName, bytes) {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.csv'))  return parseGatepassCSV(bytes);
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) return parseGatepassXLSX(bytes);
  return { success: false, error: 'UNSUPPORTED_FORMAT', message: `Unsupported file format: ${fileName}` };
}

function parseGatepassCSV(bytes) {
  try {
    var rawContent = Utilities.newBlob(bytes, 'text/csv').getDataAsString('UTF-8');
    // Strip UTF-8 BOM if present (Excel-exported CSVs often have this)
    var content = rawContent.charAt(0) === '\uFEFF' ? rawContent.slice(1) : rawContent;
    const rows  = Utilities.parseCsv(content);
    if (!rows || rows.length < 2) return { success: false, error: 'EMPTY_FILE', message: 'CSV is empty.' };
    const headers = rows[0].map(h => safeStr(h));
    return { success: true, headers, rows: rows.slice(1).map(r => rowToObj(headers, r)) };
  } catch (err) {
    return { success: false, error: 'CSV_PARSE_ERROR', message: err.message };
  }
}

function parseGatepassXLSX(bytes) {
  let tempFile  = null;
  let tempSS    = null;
  let converted = null;
  try {
    const blob = Utilities.newBlob(
      bytes,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'temp_gp.xlsx'
    );
    tempFile  = DriveApp.createFile(blob);
    converted = Drive.Files.copy(
      { title: 'temp_gp_sheet', mimeType: MimeType.GOOGLE_SHEETS },
      tempFile.getId(),
      { convert: true }
    );
    tempSS     = SpreadsheetApp.openById(converted.id);
    const data = tempSS.getSheets()[0].getDataRange().getValues();
    if (!data || data.length < 2) return { success: false, error: 'EMPTY_FILE', message: 'XLSX is empty.' };
    const headers = data[0].map(h => safeStr(h));
    return { success: true, headers, rows: data.slice(1).map(r => rowToObj(headers, r)) };
  } catch (err) {
    return { success: false, error: 'XLSX_PARSE_ERROR', message: err.message };
  } finally {
    try { if (tempSS)    DriveApp.getFileById(tempSS.getId()).setTrashed(true); } catch (_) {}
    try { if (converted) DriveApp.getFileById(converted.id).setTrashed(true); } catch (_) {}
    try { if (tempFile)  tempFile.setTrashed(true); } catch (_) {}
  }
}

// ─── ROW PROCESSING ───────────────────────────────────────────────────────────
/**
 * Filters and maps raw rows into Gatepass_Data schema.
 * Excludes rows where To Party starts with "B2B".
 */
function processGatepassRows(rawRows, uploadId, uploadDate, facility) {
  var skippedB2B     = 0;
  var skippedInvalid = 0;
  var rows           = [];

  var sampleKeys = rawRows.length > 0 ? Object.keys(rawRows[0]) : [];
  var hmap = buildHeaderMap(sampleKeys);
  function g(raw, candidates) { return fuzzyGet(raw, hmap, candidates); }

  for (var ri = 0; ri < rawRows.length; ri++) {
    var raw = rawRows[ri];

    // Skip B2B To Party
    var toParty = safeStr(g(raw, ['To Party', 'ToParty', 'To']));
    if (toParty.toUpperCase().indexOf('B2B') >= 0) { skippedB2B++; continue; }

    // Required fields
    var gpCode  = safeStr(g(raw, ['Gatepass Code', 'GP Code', 'GatepassCode', 'Gatepass No']));
    var skuCode = safeStr(g(raw, ['Item SkuCode', 'Item SKU Code', 'SKU Code', 'SkuCode', 'SKU']));
    var qty     = safeNum(g(raw, ['Quantity', 'Qty']), -1);
    if (!gpCode || !skuCode || qty < 0) { skippedInvalid++; continue; }

    // Facility from From Party (per spec)
    var fromParty    = safeStr(g(raw, ['From Party', 'FromParty', 'From']));
    var rowFacility  = detectFacility(fromParty) || facility;
    if (!rowFacility || FACILITIES.indexOf(rowFacility) < 0) { skippedInvalid++; continue; }

    // Vendor batch and batch code
    var vendorBatch = safeStr(g(raw, ['Vendor Batch No', 'Vendor Batch', 'VendorBatch', 'Vendor Batch Code']));
    var batchCode   = safeStr(g(raw, ['Uniware Batch Code', 'Batch Code', 'BatchCode']));

    // Manufacturing date — note lowercase 'd' in actual CSV
    var mfgDate = safeStr(g(raw, ['Manufacturing date', 'Manufacturing Date', 'Mfg Date', 'MfgDate']));

    rows.push({
      'Upload ID'           : uploadId,
      'Upload Date'         : uploadDate,
      'Gatepass Code'       : gpCode,
      'Gatepass Created By' : safeStr(g(raw, ['Gatepass Created By', 'Created By', 'User'])),
      'Gatepass Created At' : formatDate(g(raw, ['Gatepass Created At', 'Created At', 'Date'])),
      'Gatepass Status'     : safeStr(g(raw, ['Gatepass Status', 'GP Status'])),
      'To Party'            : toParty,
      'Item Name'           : safeStr(g(raw, ['Item Name', 'Product Name'])),
      'Item SkuCode'        : skuCode,
      'Quantity'            : qty,
      'Shelf'               : safeStr(g(raw, ['Shelf', 'Shelf No', 'Location'])),
      'From Party'          : fromParty,
      'Uniware Batch Code'  : batchCode,
      'Vendor Batch No'     : vendorBatch,
      'Manufacturing date'  : mfgDate,
      'Expiry Date'         : formatDate(g(raw, ['Expiry Date', 'Expiry', 'ExpiryDate', 'Exp Date'])),
      'Facility'            : rowFacility
    });
  }

  return { rows: rows, skippedB2B: skippedB2B, skippedInvalid: skippedInvalid };
}
function generateUploadId() {
  const now = new Date();
  return `GP${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss')}`;
}

function parseMetaString(meta) {
  const result = {};
  meta.split('|').forEach(part => {
    const [k, v] = part.split(':').map(s => s.trim());
    if (k && v !== undefined) result[k] = v;
  });
  return result;
}


// =============================================================
// INVENTORY IMPORT
// =============================================================

/**
 * InventoryImport.gs
 * Reads inventory file from Gmail attachment (CSV or XLSX),
 * validates structure, replaces Inventory_Current snapshot,
 * and logs the event.
 *
 * Triggered automatically at 8:30 AM via time-based trigger,
 * and can also be triggered manually from the UI.
 */

// ─── PUBLIC API ───────────────────────────────────────────────────────────────
/**
 * Returns current inventory status for the UI.
 */
function getInventoryStatus(params) {
  try {
    const sheet   = getSheet(SHEETS.INVENTORY);
    const lastRow = sheet.getLastRow();
    const records = Math.max(0, lastRow - 1);

    // Count by facility
    let ambientCount = 0;
    let hubCount     = 0;

    if (records > 0) {
      const data    = sheet.getDataRange().getValues();
      const headers = data[0];
      const facIdx  = headers.indexOf('Facility');

      for (let i = 1; i < data.length; i++) {
        const fac = safeStr(data[i][facIdx]);
        if (fac === 'SL Ambient')    ambientCount++;
        else if (fac === 'SL Mother Hub') hubCount++;
      }
    }

    // Get last import info from audit log
    const history = getInventoryImportHistory(5);

    return {
      success     : true,
      lastImport  : history[0] ? history[0].timestamp : null,
      totalRecords: records,
      ambientCount,
      hubCount,
      history
    };
  } catch (err) {
    logSystemError('getInventoryStatus', err);
    return { success: false, error: err.message };
  }
}

/**
 * Manually triggers inventory import from Gmail.
 * Returns result object.
 */

/**
 * Manual inventory file upload from UI (same pattern as gatepass).
 * params: { fileData (base64), fileName, fileType, facility, uploadedBy }
 */
function uploadInventory(params) {
  const { fileData, fileName, fileType, facility, uploadedBy } = params || {};

  if (!fileData || !fileName) {
    return { success: false, error: 'MISSING_FILE', message: 'File data is required.' };
  }

  try {
    const bytes  = Utilities.base64Decode(fileData);
    const parsed = parseInventoryFile(fileName, bytes);
    if (!parsed.success) return parsed;

    // Flexible column validation
    const validation = validateInventoryColumns(parsed.headers);
    if (!validation.valid) {
      return {
        success: false,
        error: 'INVALID_COLUMNS',
        message: 'Missing required columns: ' + validation.missing.join(', ') +
                 '. Found columns: ' + parsed.headers.join(', ')
      };
    }

    const { rows, skipped } = processInventoryRows(parsed.rows, parsed.headers);

    if (rows.length === 0) {
      return {
        success: false,
        error: 'NO_DATA',
        message: 'No valid rows found. Skipped: ' + skipped +
                 '. Check that your file has data and correct column names.'
      };
    }

    // Replace inventory snapshot
    replaceSheetData(SHEETS.INVENTORY, INVENTORY_HEADERS, rows);
    invalidateCache();

    const summary = 'Manual upload: ' + rows.length + ' rows from ' + fileName + ' (skipped ' + skipped + ')';
    writeAuditLog({
      eventType   : 'INVENTORY_IMPORT',
      description : summary,
      user        : uploadedBy || 'unknown',
      facility    : facility || 'ALL',
      metadata    : 'File: ' + fileName + ' | Rows: ' + rows.length + ' | Skipped: ' + skipped
    });

    return {
      success      : true,
      fileName     : fileName,
      totalRows    : parsed.rows.length,
      importedRows : rows.length,
      skippedRows  : skipped,
      message      : summary
    };
  } catch (err) {
    logSystemError('uploadInventory', err);
    return { success: false, error: 'UPLOAD_ERROR', message: err.message };
  }
}

function triggerInventoryImport(params) {
  const result = runInventoryImport();
  return result;
}

// ─── SCHEDULED IMPORT ─────────────────────────────────────────────────────────
/**
 * Called by time-based trigger (8:30 AM daily).
 * Install via installInventoryTrigger().
 */
// ─── SCHEDULED IMPORT — runs at 9:00 AM IST daily ───────────────────────────
function scheduledInventoryImport() {
  var result = runInventoryImport();
  if (!result.success) {
    console.error('Scheduled inventory import failed:', result.message);
  }
}

// ─── CORE IMPORT LOGIC ────────────────────────────────────────────────────────
/**
 * Reads the Unicommerce export email (subject: "Export Job Complete - ..."),
 * extracts the CloudFront CSV download URL from the email body,
 * fetches the CSV via UrlFetchApp, parses it, and replaces the inventory sheet.
 *
 * This replaces the old attachment-based approach since Unicommerce sends
 * a download link in the email body, not an attachment.
 */
function runInventoryImport() {
  var settings      = loadSettings();
  var senderEmail   = settings['inventory_import_email']  || 'noreply@e.unicommerce.com';
  var subjectFilter = settings['inventory_import_subject'] || 'Export Job Complete - All facility Shelfwise Inventory';

  try {
    // ── Step 1: Find the email ──────────────────────────────────────────────
    var cutoff   = new Date(Date.now() - 2 * 86400000); // last 2 days
    var query    = 'subject:"' + subjectFilter + '" from:' + senderEmail +
                   ' after:' + Math.floor(cutoff.getTime() / 1000);

    var threads  = GmailApp.search(query, 0, 20);
    if (!threads || threads.length === 0) {
      var noEmailMsg = 'No inventory export email found. Query: ' + query;
      writeAuditLog({ eventType: 'INVENTORY_IMPORT', description: noEmailMsg, user: 'SYSTEM', facility: '' });
      return { success: false, error: 'NO_EMAIL', message: noEmailMsg };
    }

    // ── Step 2: Extract CSV URL from email body ─────────────────────────────
    // Pick the most recent message that contains a CloudFront CSV link
    var csvUrl   = null;
    var bestDate = null;

    for (var ti = 0; ti < threads.length; ti++) {
      var messages = threads[ti].getMessages();
      for (var mi = 0; mi < messages.length; mi++) {
        var msg     = messages[mi];
        var msgDate = msg.getDate();
        if (msgDate < cutoff) continue;
        if (safeStr(msg.getSubject()).indexOf(subjectFilter) < 0) continue;
        var body    = msg.getPlainBody() + ' ' + msg.getBody();
        // Match the full CloudFront CSV URL, including signed query params.
        var match   = body.match(/https:\/\/[a-zA-Z0-9\-\.]+\.cloudfront\.net\/[^\s"<>]+\.csv(?:\?[^\s"<>]+)?/i);
        if (!match) continue;
        if (!bestDate || msgDate > bestDate) {
          bestDate = msgDate;
          csvUrl   = match[0].replace(/&amp;/g, '&').trim();
        }
      }
    }

    if (!csvUrl) {
      var noUrlMsg = 'Email found but no CloudFront CSV URL in body.';
      writeAuditLog({
        eventType: 'INVENTORY_IMPORT',
        description: noUrlMsg,
        user: 'SYSTEM',
        facility: '',
        metadata: 'Subject: ' + subjectFilter + ' | Query: ' + query
      });
      return { success: false, error: 'NO_URL', message: noUrlMsg };
    }

    console.log('Auto-import: fetching CSV from ' + csvUrl);

    // ── Step 3: Fetch CSV from URL ──────────────────────────────────────────
    var response = UrlFetchApp.fetch(csvUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'Accept': 'text/csv,application/csv,text/plain,*/*'
      }
    });
    if (response.getResponseCode() !== 200) {
      var fetchMsg = 'CSV fetch failed. HTTP ' + response.getResponseCode() + ' — URL may have expired.';
      writeAuditLog({ eventType: 'INVENTORY_IMPORT', description: fetchMsg, user: 'SYSTEM', facility: '' });
      return { success: false, error: 'FETCH_FAILED', message: fetchMsg };
    }

    // Use getContentText() — Apps Script auto-decompresses gzip (CloudFront sends gzip)
    var responseContent = response.getContent();
    var responseBytes = responseContent.length;
    var responseType = safeStr((response.getHeaders() || {})['Content-Type'] || (response.getAllHeaders && response.getAllHeaders()['Content-Type']) || '');
    // Extract filename from URL
    var urlParts = csvUrl.split('/');
    var fileName = decodeURIComponent(urlParts[urlParts.length - 1]).split('?')[0] || 'inventory_auto.csv';

    // ── Step 4: Parse CSV text directly ────────────────────────────────────
    var parsed = parseCSVBytes(responseContent);
    if (!parsed.success) {
      parsed = parseCSVText(response.getContentText('UTF-8'));
    }
    if (!parsed.success) {
      writeAuditLog({
        eventType: 'INVENTORY_IMPORT',
        description: 'Parse failed: ' + parsed.error,
        user: 'SYSTEM',
        facility: '',
        metadata: 'HTTP 200 | Bytes: ' + responseBytes + ' | Content-Type: ' + responseType + ' | URL: ' + csvUrl.slice(0, 180)
      });
      return parsed;
    }

    // ── Step 5: Validate columns ────────────────────────────────────────────
    var validation = validateInventoryColumns(parsed.headers);
    if (!validation.valid) {
      var colMsg = 'Auto-import missing columns: ' + validation.missing.join(', ');
      writeAuditLog({
        eventType: 'INVENTORY_IMPORT',
        description: colMsg,
        user: 'SYSTEM',
        facility: '',
        metadata: 'Found headers: ' + parsed.headers.join(' | ') + ' | File: ' + fileName
      });
      return { success: false, error: 'INVALID_COLUMNS', message: colMsg };
    }

    // ── Step 6: Process rows ────────────────────────────────────────────────
    var processed = processInventoryRows(parsed.rows, parsed.headers);
    var rows      = processed.rows;
    var skipped   = processed.skipped;

    if (rows.length === 0) {
      var emptyMsg = 'No valid rows after filtering. Skipped: ' + skipped;
      writeAuditLog({ eventType: 'INVENTORY_IMPORT', description: emptyMsg, user: 'SYSTEM', facility: '' });
      return { success: false, error: 'NO_DATA', message: emptyMsg };
    }

    // ── Step 7: Replace sheet + invalidate cache ────────────────────────────
    replaceSheetData(SHEETS.INVENTORY, INVENTORY_HEADERS, rows);
    invalidateCache();

    var summary = 'Auto-import: ' + rows.length + ' rows from ' + fileName + ' (skipped ' + skipped + ')';
    writeAuditLog({
      eventType : 'INVENTORY_IMPORT',
      description: summary,
      user      : 'SYSTEM (Auto)',
      facility  : 'ALL',
      metadata  : 'Source: Email URL | File: ' + fileName + ' | Rows: ' + rows.length + ' | Skipped: ' + skipped
    });

    return {
      success      : true,
      fileName     : fileName,
      totalRows    : parsed.rows.length,
      importedRows : rows.length,
      skippedRows  : skipped,
      message      : summary
    };

  } catch (err) {
    logSystemError('runInventoryImport', err);
    writeAuditLog({
      eventType   : 'INVENTORY_IMPORT',
      description : 'Import error: ' + err.message,
      user        : 'SYSTEM',
      facility    : ''
    });
    return { success: false, error: 'IMPORT_ERROR', message: err.message };
  }
}

// ─── PARSING ──────────────────────────────────────────────────────────────────
/**
 * Parses inventory file bytes into headers + rows array.
 */
function parseInventoryFile(fileName, fileBytes) {
  try {
    const lower = fileName.toLowerCase();

    if (lower.endsWith('.csv')) {
      return parseCSVBytes(fileBytes);
    } else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      return parseXLSXBlob(fileBytes);
    } else {
      return { success: false, error: 'UNSUPPORTED_FORMAT', message: `Unsupported format: ${fileName}` };
    }
  } catch (err) {
    return { success: false, error: 'PARSE_ERROR', message: err.message };
  }
}

/**
 * Parse CSV from byte array using Utilities.parseCsv.
 */
function parseCSVBytes(fileBytes) {
  const blob       = Utilities.newBlob(fileBytes, 'text/csv');
  var   rawContent = blob.getDataAsString('UTF-8');
  // Strip UTF-8 BOM if present
  const content    = rawContent.charAt(0) === '\uFEFF' ? rawContent.slice(1) : rawContent;
  const rows       = Utilities.parseCsv(content);

  if (!rows || rows.length < 2) {
    return { success: false, error: 'EMPTY_FILE', message: 'CSV file has no data rows.' };
  }

  const headers = rows[0].map(h => safeStr(h));
  const data    = rows.slice(1);
  return { success: true, headers, rows: data.map(row => rowToObj(headers, row)) };
}

/**
 * Parse CSV directly from a text string (used by auto-import via UrlFetchApp.getContentText).
 * Avoids gzip decompression issues that occur when using getContent() with CloudFront responses.
 */
function parseCSVText(rawText) {
  try {
    if (!rawText || !rawText.trim()) {
      return { success: false, error: 'EMPTY_FILE', message: 'CSV response was empty.' };
    }
    // Strip UTF-8 BOM if present
    var text = (rawText.charCodeAt(0) === 0xFEFF) ? rawText.slice(1) : rawText;
    var rows = Utilities.parseCsv(text);
    if (!rows || rows.length < 2) {
      return { success: false, error: 'EMPTY_FILE', message: 'CSV has no data rows.' };
    }
    var headers = rows[0].map(function(h) { return safeStr(h); });
    var data    = rows.slice(1);
    return { success: true, headers: headers, rows: data.map(function(row) { return rowToObj(headers, row); }) };
  } catch (err) {
    return { success: false, error: 'PARSE_ERROR', message: err.message };
  }
}

/**
 * Parse XLSX using Sheets API: create a temp Google Sheet, import, read, delete.
 * This is the standard Apps Script approach for XLSX parsing.
 */
function parseXLSXBlob(fileBytes) {
  let tempFile = null;
  let tempSheet = null;

  try {
    // Upload bytes to Drive as temp file
    const blob     = Utilities.newBlob(fileBytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'temp_inv.xlsx');
    tempFile       = DriveApp.createFile(blob);
    const fileId   = tempFile.getId();

    // Convert to Google Sheets
    const resource = { title: 'temp_inv_sheet', mimeType: MimeType.GOOGLE_SHEETS };
    const converted = Drive.Files.copy(resource, fileId, { convert: true });
    tempSheet      = SpreadsheetApp.openById(converted.id);

    const sheet  = tempSheet.getSheets()[0];
    const data   = sheet.getDataRange().getValues();

    if (!data || data.length < 2) {
      return { success: false, error: 'EMPTY_FILE', message: 'XLSX file has no data rows.' };
    }

    const headers = data[0].map(h => safeStr(h));
    const rows    = data.slice(1);

    return {
      success : true,
      headers,
      rows    : rows.map(row => rowToObj(headers, row))
    };
  } catch (err) {
    return { success: false, error: 'XLSX_PARSE_ERROR', message: err.message };
  } finally {
    // Clean up temp files
    try { if (tempSheet) DriveApp.getFileById(tempSheet.getId()).setTrashed(true); } catch (_) {}
    try { if (tempFile)  tempFile.setTrashed(true); } catch (_) {}
  }
}

/**
 * Convert a raw row array + headers into an object.
 */
function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
  return obj;
}

// ─── ROW PROCESSING ───────────────────────────────────────────────────────────
/**
 * Filters and transforms raw inventory rows.
 * - Keeps only SL Ambient / SL Mother Hub
 * - Extracts required fields
 * - Returns { rows, skipped }
 */
function processInventoryRows(rawRows, headers) {
  var skipped = 0;
  var rows    = [];

  var sampleKeys = rawRows.length > 0 ? Object.keys(rawRows[0]) : [];
  var hmap = buildHeaderMap(sampleKeys);
  function g(raw, candidates) { return fuzzyGet(raw, hmap, candidates); }

  for (var ri = 0; ri < rawRows.length; ri++) {
    var raw = rawRows[ri];

    // Facility filter — only SL Ambient and SL Mother Hub
    var facility = safeStr(g(raw, ['Facility', 'Facility Name', 'Warehouse']));
    if (facility && FACILITIES.indexOf(facility) < 0) { skipped++; continue; }
    var resolvedFacility = facility || 'SL Ambient';

    // SKU is mandatory
    var sku = safeStr(g(raw, ['Item Type SKU Code', 'Item SkuCode', 'SKU Code', 'SkuCode']));
    if (!sku) { skipped++; continue; }

    var qty = safeNum(g(raw, ['Quantity', 'Qty', 'Stock Qty']), 0);
    var mrp = safeNum(g(raw, ['MRP', 'Unit Price', 'MRP Price']), 0);

    // Store ALL rows (GOOD + BAD inventory).
    // Inventory Type and Batch Status are stored so buildFifoMap
    // can filter GOOD_INVENTORY + Active when calculating FIFO ranks.
    rows.push({
      'Facility'          : resolvedFacility,
      'Facility Code'     : safeStr(g(raw, ['Facility Code'])),
      'Item Type SKU Code': sku,
      'Item Type Name'    : safeStr(g(raw, ['Item Type Name', 'Item Name', 'Product Name'])),
      'Inventory Type'    : safeStr(g(raw, ['Inventory Type'])),
      'Shelf'             : safeStr(g(raw, ['Shelf', 'Shelf No', 'Section'])),
      'Quantity'          : qty,
      'Batch Code'        : safeStr(g(raw, ['Batch Code', 'BatchCode', 'Batch'])),
      'Vendor Batch Code' : safeStr(g(raw, ['Vendor Batch Code', 'Vendor Batch', 'Vendor Batch No', 'VendorBatch'])),
      'Manufacturing Date': formatDate(g(raw, ['Manufacturing Date', 'Manufacturing', 'Mfg Date', 'Manufacturing date'])),
      'Expiry'            : formatDate(g(raw, ['Expiry', 'Expiry Date', 'ExpiryDate', 'Exp Date'])),
      'MRP'               : mrp,
      'Batch Status'      : safeStr(g(raw, ['Batch Status', 'BatchStatus']))
    });
  }

  return { rows: rows, skipped: skipped };
}
function getInventoryImportHistory(limit) {
  try {
    const all = getSheetData(SHEETS.AUDIT_LOG);
    return all
      .filter(r => safeStr(r['Event Type']) === 'INVENTORY_IMPORT')
      .sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']))
      .slice(0, limit)
      .map(r => ({
        timestamp  : formatDateTime(r['Timestamp']),
        description: safeStr(r['Description']),
        metadata   : safeStr(r['Metadata'])
      }));
  } catch (_) {
    return [];
  }
}

// ─── TRIGGER MANAGEMENT ───────────────────────────────────────────────────────
/**
 * Install daily 8:30 AM trigger for inventory import.
 * Run this once from Apps Script editor.
 */
function installInventoryTrigger() {
  // Remove existing triggers with same function name
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledInventoryImport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('scheduledInventoryImport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(0)
    .inTimezone('Asia/Kolkata')
    .create();

  console.log('Inventory import trigger installed for 9:00 AM IST daily.');
}


// =============================================================
// REPORTS ENGINE
// =============================================================

/**
 * ReportsEngine.gs
 * Assembles clean, export-ready data arrays for all report types.
 * Called by the frontend Exports module when user requests
 * CSV, Excel, or PDF downloads.
 *
 * All functions return { success, headers, rows } for uniform handling.
 */

// ─── VIOLATIONS REPORT ────────────────────────────────────────────────────────
/**
 * Full violation export with all columns.
 * params: { limit, facility, severity, dateFrom, dateTo, status }
 */
function exportViolationsReport(params) {
  try {
    const {
      limit    = 10000,
      facility = '',
      severity = '',
      scope    = '',
      dateFrom = '',
      dateTo   = '',
      status   = ''      // 'Violation' | 'Partial FIFO' | 'Compliant' | ''
    } = params || {};

    const range = getReportDateRange_({ scope: scope, dateFrom: dateFrom, dateTo: dateTo });
    let rows = getSheetData(SHEETS.VIOLATIONS);

    if (facility) rows = rows.filter(r => safeStr(r['Facility']) === facility);
    if (severity) rows = rows.filter(r => safeStr(r['Severity']) === severity);
    if (status)   rows = rows.filter(r => safeStr(r['Status'])   === status);

    rows = filterRowsByScope_(rows, range);

    rows = rows
      .sort((a, b) => new Date(b['Date']) - new Date(a['Date']))
      .slice(0, limit);

    const headers = [
      'Date', 'GP Number', 'User', 'User Email', 'Facility',
      'SKU Code', 'Item Name', 'Batch Code', 'Vendor Batch',
      'Dispatched Qty', 'Shelf', 'Status', 'Severity',
      'Older Batch Count', 'Older Inventory Qty',
      'Loss Opportunity (₹)', 'FIFO Rank',
      'Manufacturing Date', 'Expiry Date', 'To Party', 'Upload ID'
    ];

    const exportRows = rows.map(r => [
      formatDate(r['Date']),
      safeStr(r['GP Number']),
      safeStr(r['User']),
      safeStr(r['User Email']),
      safeStr(r['Facility']),
      safeStr(r['SKU Code']),
      safeStr(r['Item Name']),
      safeStr(r['Batch Code']),
      safeStr(r['Vendor Batch']),
      safeNum(r['Dispatched Qty']),
      safeStr(r['Shelf']),
      safeStr(r['Status']),
      safeStr(r['Severity']),
      safeNum(r['Older Batch Count']),
      safeNum(r['Older Inventory Qty']),
      safeNum(r['Loss Opportunity']),
      safeNum(r['FIFO Rank']),
      formatDate(r['Manufacturing Date']),
      formatDate(r['Expiry Date']),
      safeStr(r['To Party']),
      safeStr(r['Upload ID'])
    ]);

    return { success: true, headers, rows: exportRows, total: rows.length };
  } catch (err) {
    logSystemError('exportViolationsReport', err);
    return { success: false, error: err.message };
  }
}

// ─── PREVENTION / AT-RISK INVENTORY REPORT ────────────────────────────────────
/**
 * At-risk inventory export — all items expiring within 180 days.
 */
function exportPreventionReport(params) {
  try {
    const { facility = '', risk = '' } = params || {};
    const inventory = getSheetData(SHEETS.INVENTORY);
    const today     = todayStart();
    const rows      = [];

    for (const row of inventory) {
      const fac = safeStr(row['Facility']);
      if (!FACILITIES.includes(fac)) continue;
      if (facility && fac !== facility) continue;

      const qty    = safeNum(row['Quantity'], 0);
      if (qty <= 0) continue;

      const expiry = toDateObj(row['Expiry']);
      const days   = expiry ? daysBetween(today, expiry) : null;

      let riskLevel = 'Low';
      if (days !== null) {
        if (days < 0)    riskLevel = 'Expired';
        else if (days <= 30)  riskLevel = 'High';
        else if (days <= 90)  riskLevel = 'Medium';
      }

      if (risk && riskLevel.toLowerCase() !== risk.toLowerCase()) continue;

      rows.push([
        fac,
        safeStr(row['Item Type SKU Code']),
        safeStr(row['Item Type Name']),
        safeStr(row['Shelf']),
        safeStr(row['Batch Code']),
        safeStr(row['Vendor Batch Code']),
        formatDate(row['Manufacturing Date']),
        formatDate(row['Expiry']),
        qty,
        safeNum(row['MRP'], 0),
        qty * safeNum(row['MRP'], 0),
        days !== null ? days : '',
        riskLevel
      ]);
    }

    // Sort by days to expiry ascending
    rows.sort((a, b) => {
      const da = typeof a[11] === 'number' ? a[11] : Infinity;
      const db = typeof b[11] === 'number' ? b[11] : Infinity;
      return da - db;
    });

    const headers = [
      'Facility', 'SKU Code', 'Item Name', 'Shelf',
      'Batch Code', 'Vendor Batch', 'Manufacturing Date', 'Expiry Date',
      'Quantity', 'MRP (₹)', 'Value (₹)', 'Days to Expiry', 'Risk Level'
    ];

    return { success: true, headers, rows, total: rows.length };
  } catch (err) {
    logSystemError('exportPreventionReport', err);
    return { success: false, error: err.message };
  }
}

// ─── FINANCIAL IMPACT REPORT ──────────────────────────────────────────────────// ─── USER COMPLIANCE REPORT ───────────────────────────────────────────────────
/**
 * Full per-user compliance export.
 */
function exportComplianceReport(params) {
  try {
    const { offendersOnly = false, championsOnly = false } = params || {};
    const range = getReportDateRange_(params);
    var scopedRows = filterRowsByScope_(getSheetData(SHEETS.VIOLATIONS), range);
    let users = computeUserCompliance(scopedRows);

    if (offendersOnly) users = users.filter(u => u.offenderLevel === 'Repeat Offender');
    if (championsOnly) users = users.filter(u => u.isChampion);

    const headers = [
      'User Name', 'User Email', 'Facility',
      'Total GPs', 'Violations', 'Partial FIFO', 'Compliant',
      'Compliance %', 'Skipped Qty', 'Loss Opportunity (₹)',
      'Last Violation Date', 'Violations (Last 30 Days)',
      'Offender Level', 'Is Champion'
    ];

    const rows = users.map(u => [
      u.userName,
      u.email,
      u.facility,
      u.totalGps,
      u.violations,
      u.partialFifo,
      u.compliant,
      u.compliancePct,
      u.skippedQty,
      u.lossOpportunity,
      u.lastViolationDate,
      u.violations30d,
      u.offenderLevel,
      u.isChampion ? 'Yes' : 'No'
    ]);

    return { success: true, headers, rows, total: rows.length };
  } catch (err) {
    logSystemError('exportComplianceReport', err);
    return { success: false, error: err.message };
  }
}

// ─── FACILITY SUMMARY REPORT ──────────────────────────────────────────────────
/**
 * Per-facility compliance + financial summary.
 */
function exportFacilitySummaryReport(params) {
  try {
    var violations = getSheetData(SHEETS.VIOLATIONS);
    var compliance = computeFacilityCompliance(violations);

    var headers = [
      'Facility', 'Total GPs', 'Violations', 'Partial FIFO',
      'Compliant GPs', 'Compliance %', 'Skipped Qty'
    ];

    const rows = compliance.map(f => [
      f.facility,
      f.totalGps,
      f.violations,
      f.partialFifo,
      f.compliantGps,
      f.compliancePct,
      f.skippedQty
    ]);

    return { success: true, headers, rows, total: rows.length };
  } catch (err) {
    logSystemError('exportFacilitySummaryReport', err);
    return { success: false, error: err.message };
  }
}

// ─── UNIFIED EXPORT ROUTER ────────────────────────────────────────────────────// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// SECTION 6. FIFO ENGINE  (complete rewrite)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ─── FACILITY DETECTION FROM FROM-PARTY ───────────────────────────────────────
/**
 * Maps the gatepass "From Party" field to a known facility name.
 * Case-insensitive substring matching.
 */

function buildFifoMap() {
  var settings  = loadSettings();
  var minQty    = safeNum(settings['min_inventory_threshold'], 50);
  var inventory = getSheetData(SHEETS.INVENTORY);

  // Consolidate by SKU + Vendor Batch across ALL facilities (cross-facility per spec)
  // Key: "SKUCode|VendorBatch"
  var consolidated = {};

  for (var i = 0; i < inventory.length; i++) {
    var row      = inventory[i];
    var facility = safeStr(row['Facility']);

    // STEP 1: Only SL Ambient + SL Mother Hub
    if (FACILITIES.indexOf(facility) < 0) continue;

    // STEP 2: Filter — GOOD_INVENTORY only
    var invType = safeStr(row['Inventory Type'] || row['InventoryType'] || '');
    if (invType && invType.toUpperCase() !== 'GOOD_INVENTORY') continue;

    // STEP 2: Filter — Active batch status only
    var batchStatus = safeStr(row['Batch Status'] || row['BatchStatus'] || '');
    if (batchStatus && batchStatus.toUpperCase() !== 'ACTIVE') continue;

    var skuCode    = safeStr(row['Item Type SKU Code']);
    var itemName   = safeStr(row['Item Type Name']);
    var vendorBatch= safeStr(row['Vendor Batch Code']);
    var batchCode  = safeStr(row['Batch Code']);
    var qty        = safeNum(row['Quantity'], 0);
    var mrp        = safeNum(row['MRP'], 0);
    var mfgDate    = safeStr(row['Manufacturing Date']);
    var expiryDate = safeStr(row['Expiry']);
    var shelf      = safeStr(row['Shelf']);

    if (!skuCode || qty <= 0) continue;

    // Consolidation key — cross-facility, by SKU + VendorBatch
    // Use BatchCode as primary key (consistent across inventory + gatepass files)
    var cKey = skuCode + '|' + (batchCode || vendorBatch);

    if (!consolidated[cKey]) {
      consolidated[cKey] = {
        skuCode    : skuCode,
        itemName   : itemName,
        vendorBatch: vendorBatch,
        batchCode  : batchCode,
        mfgDate    : mfgDate,
        expiryDate : expiryDate,
        mrp        : mrp,
        totalQty   : 0,
        facilities : [],
        shelves    : []
      };
    }
    var entry = consolidated[cKey];
    entry.totalQty += qty;
    if (FACILITIES.indexOf(facility) >= 0 && entry.facilities.indexOf(facility) < 0) {
      entry.facilities.push(facility);
    }
    if (shelf && qty > 0) {
      entry.shelves.push({ shelf: shelf, qty: qty, facility: facility });
    }
    if (!entry.mrp && mrp) entry.mrp = mrp;
    // Use earliest mfg/expiry found
    if (!entry.expiryDate && expiryDate) entry.expiryDate = expiryDate;
    if (!entry.mfgDate    && mfgDate)    entry.mfgDate    = mfgDate;
  }

  // Now group by SKU only (cross-facility), apply threshold, sort + rank
  var skuGroups = {};
  var cKeys = Object.keys(consolidated);
  for (var k = 0; k < cKeys.length; k++) {
    var batch = consolidated[cKeys[k]];
    // Apply min threshold PER batch
    if (batch.totalQty < minQty) continue;

    var skuKey = batch.skuCode;
    if (!skuGroups[skuKey]) skuGroups[skuKey] = [];
    skuGroups[skuKey].push(batch);
  }

  // Sort by EXPIRY DATE only — same expiry = same rank
  var skuKeys = Object.keys(skuGroups);
  for (var s = 0; s < skuKeys.length; s++) {
    var batches = skuGroups[skuKeys[s]];
    batches.sort(fifoComparator);

    // Assign ranks: same expiry date gets same rank
    var rank    = 1;
    var prevExp = null;
    for (var r = 0; r < batches.length; r++) {
      if (r === 0) {
        batches[r].fifoRank = 1;
        prevExp = normDateStr(batches[r].expiryDate);
      } else {
        var curExp = normDateStr(batches[r].expiryDate);
        if (curExp !== prevExp) rank++;
        batches[r].fifoRank = rank;
        prevExp = curExp;
      }
    }
  }

  return skuGroups;
}

// Normalise date to YYYY-MM-DD string for comparison
function normDateStr(val) {
  var d = toDateObj(val);
  if (!d) return '';
  var m = d.getMonth() + 1;
  var day = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' + m : m) + '-' + (day < 10 ? '0' + day : day);
}

function fifoComparator(a, b) {
  // Sort by Expiry Date ASC only — same expiry = same rank
  var expA = toDateObj(a.expiryDate);
  var expB = toDateObj(b.expiryDate);
  if (expA && expB) {
    var diff = expA - expB;
    if (diff !== 0) return diff;
  } else if (expA) return -1;
  else if (expB)   return  1;
  return 0;
}


// ─── STEP 7-10: LOOKUP DISPATCHED BATCH RANK ──────────────────────────────────
/**
 * Finds the FIFO rank of the dispatched batch and identifies older batches.
 * Matching: by Vendor Batch (primary), then Batch Code (fallback).
 *
 * Returns:
 * {
 *   found, fifoRank, matchedBatch,
 *   olderBatches (above threshold), olderInventoryQty, olderBatchCount
 * }
 */
function lookupFifoRank(fifoMap, skuCode, dispatchedVendorBatch, dispatchedBatchCode) {
  // Cross-facility: key is SKU only
  var batches = fifoMap[skuCode];
  if (!batches || batches.length === 0) return { found: false };

  var norm = function(s) { return safeStr(s).toLowerCase().trim(); };
  var dvb  = norm(dispatchedVendorBatch);
  var dbc  = norm(dispatchedBatchCode);

  // Match by BatchCode (Uniware system code) FIRST — consistent across both files.
  // Fall back to VendorBatch if BatchCode not found.
  var match = null;
  if (dbc) {
    for (var i = 0; i < batches.length; i++) {
      if (norm(batches[i].batchCode) === dbc) { match = batches[i]; break; }
    }
  }
  if (!match && dvb) {
    for (var j = 0; j < batches.length; j++) {
      if (norm(batches[j].vendorBatch) === dvb) { match = batches[j]; break; }
    }
  }
  if (!match) return { found: false };

  var dispatchedRank = match.fifoRank;

  // Older batches = those with LOWER rank (= earlier expiry)
  var olderBatches = [];
  for (var k = 0; k < batches.length; k++) {
    if (batches[k].fifoRank < dispatchedRank) {
      olderBatches.push(batches[k]);
    }
  }

  var olderQty = 0;
  for (var m = 0; m < olderBatches.length; m++) {
    olderQty += olderBatches[m].totalQty;
  }

  return {
    found            : true,
    fifoRank         : dispatchedRank,
    matchedBatch     : match,
    olderBatches     : olderBatches,
    olderInventoryQty: olderQty,
    olderBatchCount  : olderBatches.length
  };
}

function evaluateLine(lines, fifoMap, settings, uploadId) {
  var ref       = lines[0];
  var totalQty  = 0;
  for (var i = 0; i < lines.length; i++) {
    totalQty += safeNum(lines[i]['Quantity'], 0);
  }

  var gpCode      = safeStr(ref['Gatepass Code']);
  var skuCode     = safeStr(ref['Item SkuCode']);
  var vendorBatch = safeStr(ref['Vendor Batch No']);
  var batchCode   = safeStr(ref['Uniware Batch Code']);
  var user        = safeStr(ref['Gatepass Created By']);
  var facility    = safeStr(ref['Facility']);
  var gpDate      = safeStr(ref['Gatepass Created At']);
  var toParty     = safeStr(ref['To Party']);
  var gpStatus    = safeStr(ref['Gatepass Status']);
  var itemName    = safeStr(ref['Item Name']);
  var mfgDate     = safeStr(ref['Manufacturing date'] || ref['Manufacturing Date']);
  var expiryDate  = safeStr(ref['Expiry Date']);
  var shelf       = safeStr(ref['Shelf']);

  var lookup = lookupFifoRank(fifoMap, skuCode, vendorBatch, batchCode);

  var status, severity, olderQty, olderCount, lossOpp, olderBatchDetails;

  if (!lookup.found) {
    // Cannot verify — batch not in inventory snapshot
    status           = 'Compliant';
    severity         = '';
    olderQty         = 0;
    olderCount       = 0;
    lossOpp          = 0;
    olderBatchDetails = '';
  } else if (lookup.olderBatchCount === 0) {
    // STEP 11 CASE 1: No older batches → Compliant
    status           = 'Compliant';
    severity         = '';
    olderQty         = 0;
    olderCount       = 0;
    lossOpp          = 0;
    olderBatchDetails = '';
  } else {
    olderQty   = lookup.olderInventoryQty;
    olderCount = lookup.olderBatchCount;

    // STEP 11 CASE 2 vs CASE 3
    if (olderQty >= totalQty) {
      status = 'FIFO VIOLATION';
    } else {
      status = 'PARTIAL FIFO POSSIBLE';
    }

    severity = classifySeverity(olderQty, settings);

    // STEP 13: Loss = sum of (each older batch qty × its MRP)
    lossOpp = 0;
    var olderDetails = [];
    for (var ob = 0; ob < lookup.olderBatches.length; ob++) {
      var b = lookup.olderBatches[ob];
      lossOpp += b.totalQty * b.mrp;
      olderDetails.push(
        b.vendorBatch + '|' + b.batchCode +
        '(Rank:' + b.fifoRank + ',Qty:' + b.totalQty +
        ',Exp:' + b.expiryDate + ',MfgDate:' + b.mfgDate +
        ',Shelf:' + (b.shelves && b.shelves[0] ? b.shelves[0].shelf : '') +
        ',MRP:' + b.mrp + ')'
      );
    }
    olderBatchDetails = olderDetails.join('; ');
  }

  var id = generateViolationId(gpCode, skuCode);

  var row = {
    'ID'                  : id,
    'Date'                : formatDate(gpDate) || formatDate(new Date()),
    'GP Number'           : gpCode,
    'User'                : user,
    'User Email'          : '',
    'SKU Code'            : skuCode,
    'Item Name'           : itemName,
    'Batch Code'          : batchCode,
    'Vendor Batch'        : vendorBatch,
    'Dispatched Qty'      : totalQty,
    'Shelf'               : shelf,
    'Older Batch Count'   : olderCount,
    'Older Inventory Qty' : olderQty,
    'Severity'            : severity,
    'Status'              : status,
    'Facility'            : facility,
    'To Party'            : toParty,
    'GP Status'           : gpStatus,
    'Loss Opportunity'    : lossOpp,
    'Manufacturing Date'  : mfgDate,
    'Expiry Date'         : expiryDate,
    'FIFO Rank'           : lookup.found ? lookup.fifoRank : 1,
    'Skipped Batch Details': olderBatchDetails,
    'Upload ID'           : uploadId
  };

  return { status: status, row: row };
}

// ─── PROCESS ALL VIOLATIONS FOR AN UPLOAD ─────────────────────────────────────
function processViolationsForUpload(uploadId, gpRows, facility) {
  try {
    var fifoMap  = buildFifoMap();
    var settings = loadSettings();

    // Group by GP Code + SKU
    var groups = {};
    for (var i = 0; i < gpRows.length; i++) {
      var r   = gpRows[i];
      var key = safeStr(r['Gatepass Code']) + '|' + safeStr(r['Item SkuCode']);
      if (!groups[key]) groups[key] = [];
      groups[key].push(r);
    }

    var violationRows  = [];
    var violationCount = 0;
    var partialCount   = 0;
    var compliantCount = 0;

    var groupKeys = Object.keys(groups);
    for (var g = 0; g < groupKeys.length; g++) {
      var result = evaluateLine(groups[groupKeys[g]], fifoMap, settings, uploadId);
      if (!result) continue;
      violationRows.push(result.row);
      if (result.status === 'FIFO VIOLATION')       violationCount++;
      else if (result.status === 'PARTIAL FIFO POSSIBLE') partialCount++;
      else compliantCount++;
    }

    if (violationRows.length > 0) {
      // Remove any existing rows for this uploadId before appending
      // Prevents duplicate violations if same file is uploaded twice
      var vlSheet = getSheet(SHEETS.VIOLATIONS);
      var vlData  = vlSheet.getDataRange().getValues();
      if (vlData.length > 1) {
        var uidCol = vlData[0].indexOf('Upload ID');
        if (uidCol >= 0) {
          // Delete from bottom to top to avoid row shift issues
          for (var vd = vlData.length - 1; vd >= 1; vd--) {
            if (safeStr(vlData[vd][uidCol]) === uploadId) {
              vlSheet.deleteRow(vd + 1);
            }
          }
        }
      }
      appendSheetRows(SHEETS.VIOLATIONS, VIOLATION_HEADERS, violationRows);
    }

    return { violationCount: violationCount, partialCount: partialCount, compliantCount: compliantCount };
  } catch (err) {
    logSystemError('processViolationsForUpload', err);
    return { violationCount: 0, partialCount: 0, compliantCount: 0 };
  }
}

function classifySeverity(skippedQty, settings) {
  var minorMin    = safeNum(settings['severity_minor_min'],    50);
  var majorMin    = safeNum(settings['severity_major_min'],   500);
  var criticalMin = safeNum(settings['severity_critical_min'], 1000);
  if (skippedQty >= criticalMin) return 'Critical';
  if (skippedQty >= majorMin)    return 'Major';
  if (skippedQty >= minorMin)    return 'Minor';
  return '';
}/**
 * Returns top N SKUs ranked by total inventory value (qty × MRP).
 * Used by prevention page and reports.
 * Filters: GOOD_INVENTORY + Active batch status only, known facilities only.
 */
function getInventorySummaryBySku(topN) {
  try {
    var inventory = getSheetData(SHEETS.INVENTORY);
    var skuMap    = {};

    for (var i = 0; i < inventory.length; i++) {
      var row      = inventory[i];
      var facility = safeStr(row['Facility']);
      if (FACILITIES.indexOf(facility) < 0) continue;

      var invType     = safeStr(row['Inventory Type'] || '');
      var batchStatus = safeStr(row['Batch Status']   || '');
      if (invType     && invType.toUpperCase()     !== 'GOOD_INVENTORY') continue;
      if (batchStatus && batchStatus.toUpperCase() !== 'ACTIVE')         continue;

      var skuCode  = safeStr(row['Item Type SKU Code']);
      var itemName = safeStr(row['Item Type Name']);
      if (!skuCode) continue;

      var qty    = safeNum(row['Quantity'], 0);
      var mrp    = safeNum(row['MRP'], 0);
      var expiry = safeStr(row['Expiry']);

      if (!skuMap[skuCode]) {
        skuMap[skuCode] = {
          skuCode  : skuCode,
          itemName : itemName,
          totalQty : 0,
          totalValue: 0,
          batches  : []
        };
      }
      var entry = skuMap[skuCode];
      entry.totalQty   += qty;
      entry.totalValue += qty * mrp;
      entry.batches.push({ qty: qty, mrp: mrp, expiry: expiry, facility: facility });
    }

    var list = Object.keys(skuMap).map(function(k) { return skuMap[k]; });
    list.sort(function(a, b) { return b.totalValue - a.totalValue; });

    return topN ? list.slice(0, topN) : list;
  } catch (err) {
    logSystemError('getInventorySummaryBySku', err);
    return [];
  }
}
