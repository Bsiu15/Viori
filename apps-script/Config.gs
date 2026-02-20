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
/** First day of the forecast window */
var START_DATE = new Date(2026, 1, 19); // Feb 19, 2026 (months are 0-indexed)
/** Last day of the forecast window */
var END_DATE   = new Date(2026, 3, 30); // Apr 30, 2026

// ── SKU master list (order matters — used for row ordering everywhere) ──────
var SKUS = [
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
var SETTINGS_TAB_NAME  = 'Settings';
var SUMMARY_TAB_NAME   = 'Summary';

/**
 * Number of input rows per SKU block in the Settings tab.
 * Each SKU block has the following rows:
 *  1  SKU header (merged, colored)
 *  2  FBA (available units)
 *  3  Inbound (processing/receiving units)
 *  4  FBA check-in and processing delay (days)
 *  5  On-hand (FBM on-hand units)
 *  6  Reserved
 *  7  Researching
 *  8  Unfulfillable
 *  9  Daily sales velocity (units/day)
 * 10  Velocity override 1: start date
 * 11  Velocity override 1: end date
 * 12  Velocity override 1: units/day
 * 13  Velocity override 2: start date
 * 14  Velocity override 2: end date
 * 15  Velocity override 2: units/day
 * 16  Velocity override 3: start date
 * 17  Velocity override 3: end date
 * 18  Velocity override 3: units/day
 * 19  Conversion rate (%)
 * 20  Conversion rate override 1: start date
 * 21  Conversion rate override 1: end date
 * 22  Conversion rate override 1: value (%)
 * 23  Conversion rate override 2: start date
 * 24  Conversion rate override 2: end date
 * 25  Conversion rate override 2: value (%)
 * 26  Conversion rate override 3: start date
 * 27  Conversion rate override 3: end date
 * 28  Conversion rate override 3: value (%)
 * 29  SPD shipment: units
 * 30  SPD shipment: send date
 * 31  SPD shipment: transit time (days)
 * 32  LTL shipment: units
 * 33  LTL shipment: send date
 * 34  LTL shipment: transit time (days)
 * 35  DTC bridge: units
 * 36  DTC bridge: start date
 * 37  DTC bridge: end date
 * 38  Ad-hoc shipment: units
 * 39  Ad-hoc shipment: send date
 * 40  Ad-hoc shipment: transit time (days)
 * 41  (blank spacer row)
 */
var ROWS_PER_SKU_BLOCK = 41;

/** Column A = labels, Column B = values in the Settings tab */
var SETTINGS_LABEL_COL = 1;
var SETTINGS_VALUE_COL = 2;

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Returns the total number of days in the forecast window (inclusive).
 * @return {number}
 */
function forecastDayCount() {
  return Math.round((END_DATE - START_DATE) / 86400000) + 1; // 71 days
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
