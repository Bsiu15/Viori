/**
 * Config.gs
 * ---------------------------------------------------------------------------
 * Central configuration for the Inventory Forecasting Calendar.
 * Contains SKU definitions, date range constants, and color palette.
 * All values that could change across SKUs are driven from the Settings tab;
 * this file only holds structural constants used by every other module.
 * ---------------------------------------------------------------------------
 */

// ── Date range ──────────────────────────────────────────────────────────────
/** First day of the forecast window.
 *  Reads from the "Forecast start date" cell in the Settings tab.
 *  Falls back to today if the named range hasn't been created yet.
 */
var START_DATE = (function() {
  var fallback = new Date();
  fallback.setHours(0, 0, 0, 0);
  try {
    var r = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('GLOBAL__FORECAST_START_DATE');
    if (r) {
      var v = r.getValue();
      if (v instanceof Date && !isNaN(v.getTime())) {
        v.setHours(0, 0, 0, 0);
        return v;
      }
    }
  } catch (e) { /* spreadsheet not available (e.g. during API testing) */ }
  return fallback;
})();
/** Last day of the forecast window.
 *  Reads from the "Forecast end date" cell in the Settings tab.
 *  Falls back to Apr 30, 2026 if the named range hasn't been created yet
 *  (e.g. first run before Settings tab is built).
 */
var END_DATE = (function() {
  var fallback = new Date(2026, 3, 30); // Apr 30, 2026
  try {
    var r = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('GLOBAL__FORECAST_END_DATE');
    if (r) {
      var v = r.getValue();
      if (v instanceof Date && !isNaN(v.getTime())) {
        v.setHours(23, 59, 59, 0);
        return v;
      }
    }
  } catch (e) { /* spreadsheet not available (e.g. during API testing) */ }
  return fallback;
})();

// ── SKU limits ──────────────────────────────────────────────────────────────
var MAX_SKUS = 15;
var REGISTRY_TAB_NAME = '_Registry';

// ── SKU master list (defaults — used on first run, then registry takes over) ─
var DEFAULT_SKUS = [
  {
    id:    'SB-HW-100W-FBA',
    name:  'Hidden Waterfall Shampoo Bar',
    tab:   'SKU - SB-HW-100W-FBA'
  },
  {
    id:    'SB-CY-100W-FBA',
    name:  'Citrus Yao Clarifying Shampoo Bar',
    tab:   'SKU - SB-CY-100W-FBA'
  },
  {
    id:    'SB-TG-100W-FBA',
    name:  'Terrace Garden Shampoo Bar',
    tab:   'SKU - SB-TG-100W-FBA'
  },
  {
    id:    'BUND-1HWS-1HWC-S-FBA',
    name:  'Hidden Waterfall Shampoo and Conditioner Bar Set',
    tab:   'SKU - BUND-1HWS-1HWC-S'
  },
  {
    id:    'BUND-1CYS-1CYC-S-FBA',
    name:  'Citrus Yao Shampoo and Conditioner Bar Set',
    tab:   'SKU - BUND-1CYS-1CYC-S'
  },
  {
    id:    'BUND-1TGS-1TGC-S-FBA',
    name:  'Terrace Garden Shampoo and Conditioner Bar Set',
    tab:   'SKU - BUND-1TGS-1TGC-S'
  }
];

/**
 * Reads the SKU list from the _Registry sheet. Falls back to DEFAULT_SKUS
 * if the registry doesn't exist yet (pre-setup).
 * @return {Object[]} array of {id, name, tab}
 */
function getSkus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REGISTRY_TAB_NAME);
  if (!sheet) return DEFAULT_SKUS;

  var data = sheet.getDataRange().getValues();
  var skus = [];
  for (var i = 1; i < data.length; i++) { // skip header row
    if (data[i][0] && data[i][0] !== '') {
      skus.push({
        id:   String(data[i][0]),
        name: String(data[i][1]),
        tab:  String(data[i][2])
      });
    }
  }
  return skus.length > 0 ? skus : DEFAULT_SKUS;
}

/**
 * Creates or rebuilds the _Registry sheet with the given SKU list.
 * @param {Object[]} skus  array of {id, name, tab}
 */
function initRegistry(skus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(REGISTRY_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_TAB_NAME);
  }
  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['SKU ID', 'SKU Name', 'Tab Name']])
       .setFontWeight('bold');
  for (var i = 0; i < skus.length; i++) {
    sheet.getRange(i + 2, 1, 1, 3).setValues([[skus[i].id, skus[i].name, skus[i].tab]]);
  }
  sheet.hideSheet();
}

// For backwards compatibility — modules that reference SKUS directly
var SKUS = DEFAULT_SKUS;

// ── Color palette (hex) ─────────────────────────────────────────────────────
var COLORS = {
  FBA:        '#C6EFCE', // Green  — fulfilling from FBA
  FBM:        '#FFEB9C', // Yellow — fulfilling from FBM (FBA zero)
  DTC:        '#FCD5B4', // Orange — fulfilling from DTC bridge
  OOS:        '#FFC7CE', // Red    — true out of stock
  PROCESSING: '#BDD7EE', // Blue   — shipment arriving / entering processing
  GRAY:       '#D9D9D9', // Gray   — inventory sitting in FBA processing
  HEADER:     '#4472C4', // Dark blue header background
  HEADER_FG:  '#FFFFFF', // White header text
  SECTION_BG: '#D6E4F0', // Light blue section header in Settings
  WHITE:      '#FFFFFF'
};

// ── Settings tab layout constants ───────────────────────────────────────────
var SETTINGS_TAB_NAME      = 'Settings';
var SUMMARY_TAB_NAME       = 'Summary';
var GORILLA_DATA_TAB_NAME  = 'Gorilla Data';

/**
 * Number of input rows per SKU block in the Settings tab.
 * Each SKU block has the following rows:
 *  1  SKU header (merged, colored)
 *  ── Starting Inventory (mirrors Seller Central) ──
 *  2  FBA (available for sale) — auto-calc or manual override
 *  3  Inbound: Working
 *  3  Inbound: Shipped
 *  4  Inbound: Receiving
 *  5  On-hand: Available
 *  6  On-hand: FC transfer
 *  7  Reserved: Customer order
 *  8  Reserved: FC processing
 *  9  Researching
 * 10  Unfulfillable: Warehouse damaged
 * 11  Unfulfillable: Defective
 * 12  Unfulfillable: Expired
 * 13  Unfulfillable: Customer damaged
 * 14  Unfulfillable: Carrier damaged
 * 15  Unfulfillable: Distributor damaged
 * 16  FBA check-in and processing delay (days)
 * 17  FBM on-hand
 *  ── Sales Velocity ──
 * 18  Daily sales velocity (units/day)
 * 19  Velocity override 1: start date
 * 20  Velocity override 1: end date
 * 21  Velocity override 1: units/day
 * 22  Velocity override 2: start date
 * 23  Velocity override 2: end date
 * 24  Velocity override 2: units/day
 * 25  Velocity override 3: start date
 * 26  Velocity override 3: end date
 * 27  Velocity override 3: units/day
 *  ── Conversion Rate ──
 * 28  Conversion rate (%)
 * 29  CR override 1: start date
 * 30  CR override 1: end date
 * 31  CR override 1: value (%)
 * 32  CR override 2: start date
 * 33  CR override 2: end date
 * 34  CR override 2: value (%)
 * 35  CR override 3: start date
 * 36  CR override 3: end date
 * 37  CR override 3: value (%)
 *  ── Shipments ──
 * 38  SPD shipment: units
 * 39  SPD shipment: send date
 * 40  SPD shipment: transit time (days)
 * 41  LTL shipment: units
 * 42  LTL shipment: send date
 * 43  LTL shipment: transit time (days)
 * 44  DTC bridge: units
 * 45  DTC bridge: start date
 * 46  DTC bridge: end date
 * 47  Ad-hoc shipment: units
 * 48  Ad-hoc shipment: send date
 * 49  Ad-hoc shipment: transit time (days)
 *  ── Financials ──
 * 50  Selling price ($)
 * 51  DPP margin (%)
 * 52  Past OOS days (already experienced)
 * 53  (blank spacer row)
 */
var ROWS_PER_SKU_BLOCK = 53;

/** Column A = labels, Column B = values in the Settings tab */
var SETTINGS_LABEL_COL = 1;
var SETTINGS_VALUE_COL = 2;

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Returns an array of {year, month, name} objects for every calendar month
 * that falls within the START_DATE → END_DATE forecast window.
 * month is 0-indexed (0=Jan, 1=Feb, …). name is the full "Month YYYY" label.
 * @return {Object[]}
 */
function forecastMonths() {
  var MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];
  var months = [];
  var y = START_DATE.getFullYear();
  var m = START_DATE.getMonth();
  var endY = END_DATE.getFullYear();
  var endM = END_DATE.getMonth();

  while (y < endY || (y === endY && m <= endM)) {
    months.push({ year: y, month: m, name: MONTH_NAMES[m] + ' ' + y });
    m++;
    if (m > 11) { m = 0; y++; }
  }
  return months;
}

/**
 * Same as forecastMonths() but with short 3-letter names (e.g. "Feb", "Mar").
 * @return {Object[]}
 */
function forecastMonthsShort() {
  var SHORT = ['Jan','Feb','Mar','Apr','May','Jun',
               'Jul','Aug','Sep','Oct','Nov','Dec'];
  var full = forecastMonths();
  var result = [];
  for (var i = 0; i < full.length; i++) {
    var yr = String(full[i].year).slice(-2); // "26", "27", etc.
    result.push({ year: full[i].year, month: full[i].month, name: SHORT[full[i].month] + ' \'' + yr });
  }
  return result;
}

/**
 * Returns the total number of days in the forecast window (inclusive).
 * @return {number}
 */
function forecastDayCount() {
  return Math.round((END_DATE - START_DATE) / 86400000) + 1;
}

/**
 * Returns an array of Date objects for every day in the forecast window.
 * @return {Date[]}
 */
function forecastDates() {
  var dates = [];
  var count = forecastDayCount();
  for (var i = 0; i < count; i++) {
    dates.push(new Date(START_DATE.getFullYear(), START_DATE.getMonth(), START_DATE.getDate() + i));
  }
  return dates;
}

/**
 * Formats a Date as "M/D/YYYY" for display.
 * @param {Date} d
 * @return {string}
 */
function fmtDate(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
}

/**
 * Formats a Date as "MMM D" for compact column headers.
 * @param {Date} d
 * @return {string}
 */
function fmtDateShort(d) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months[d.getMonth()] + ' ' + d.getDate();
}

/**
 * Compares two dates ignoring time component. Returns true if same calendar day.
 * @param {Date} a
 * @param {Date} b
 * @return {boolean}
 */
function sameDay(a, b) {
  if (!a || !b) return false;
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth()    === b.getMonth() &&
         a.getDate()     === b.getDate();
}

/**
 * Returns true if date d is within [start, end] inclusive.
 * @param {Date} d
 * @param {Date} start
 * @param {Date} end
 * @return {boolean}
 */
function dateInRange(d, start, end) {
  if (!start || !end) return false;
  var t = d.getTime();
  return t >= start.getTime() && t <= end.getTime();
}
