// ============================================================================
// VIORI INVENTORY FORECASTING CALENDAR — COMBINED APPS SCRIPT
// Paste this entire file into Code.gs in the Apps Script editor.
// ============================================================================


// ============================================================================
// FILE: Config.gs
// ============================================================================

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
 *  2  Starting FBA available units
 *  3  Starting FBA in processing/receiving units
 *  4  FBA check-in and processing delay (days)
 *  5  Starting FBM on-hand units
 *  6  Daily sales velocity (units/day)
 *  7  Velocity override: start date
 *  8  Velocity override: end date
 *  9  Velocity override: units/day
 * 10  Conversion rate (%)
 * 11  Conversion rate override: start date
 * 12  Conversion rate override: end date
 * 13  Conversion rate override: value (%)
 * 14  SPD shipment: units
 * 15  SPD shipment: send date
 * 16  SPD shipment: transit time (days)
 * 17  LTL shipment: units
 * 18  LTL shipment: send date
 * 19  LTL shipment: transit time (days)
 * 20  DTC bridge: units
 * 21  DTC bridge: start date
 * 22  DTC bridge: end date
 * 23  Ad-hoc shipment: units
 * 24  Ad-hoc shipment: send date
 * 25  Ad-hoc shipment: transit time (days)
 * 26  (blank spacer row)
 */
var ROWS_PER_SKU_BLOCK = 26;

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


// ============================================================================
// FILE: SettingsTab.gs
// ============================================================================

/**
 * SettingsTab.gs
 * ---------------------------------------------------------------------------
 * Creates and formats the Settings tab with per-SKU input sections.
 * Each SKU gets its own clearly labeled block of input rows.
 * All values are stored in column B; labels are in column A.
 * Named ranges are created for every input so other modules can read them
 * by name (e.g., "SB_HW_100W_FBA__FBA_AVAILABLE").
 * ---------------------------------------------------------------------------
 */

/**
 * Master list defining every input row within a SKU block.
 * "key" is appended to the SKU-based named range prefix.
 * "label" is the human-readable text in column A.
 * "default" is the initial value written to column B.
 * "format" is optional: "date" applies date formatting, "percent" applies %.
 */
var SKU_INPUT_ROWS = [
  { key: 'FBA_AVAILABLE',         label: 'Starting FBA available units',               defaultVal: 0,    format: 'number'  },
  { key: 'FBA_PROCESSING',        label: 'Starting FBA in processing/receiving units',  defaultVal: 0,    format: 'number'  },
  { key: 'FBA_CHECKIN_DELAY',     label: 'FBA check-in and processing delay (days)',    defaultVal: 7,    format: 'number'  },
  { key: 'FBM_ONHAND',            label: 'Starting FBM on-hand units',                  defaultVal: 0,    format: 'number'  },
  { key: 'DAILY_VELOCITY',        label: 'Daily sales velocity (units/day)',             defaultVal: 0,    format: 'number'  },
  { key: 'VEL_OVERRIDE_START',    label: 'Velocity override: start date',               defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_END',      label: 'Velocity override: end date',                 defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_VALUE',    label: 'Velocity override: units/day',                defaultVal: '',   format: 'number'  },
  { key: 'CONVERSION_RATE',       label: 'Conversion rate (%)',                          defaultVal: 100,  format: 'percent' },
  { key: 'CR_OVERRIDE_START',     label: 'Conversion rate override: start date',         defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_END',       label: 'Conversion rate override: end date',           defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_VALUE',     label: 'Conversion rate override: value (%)',          defaultVal: '',   format: 'percent' },
  { key: 'SPD_UNITS',             label: 'SPD shipment: units',                          defaultVal: 0,    format: 'number'  },
  { key: 'SPD_SEND_DATE',         label: 'SPD shipment: send date',                      defaultVal: '',   format: 'date'    },
  { key: 'SPD_TRANSIT_DAYS',      label: 'SPD shipment: transit time (days)',             defaultVal: 5,    format: 'number'  },
  { key: 'LTL_UNITS',             label: 'LTL shipment: units',                          defaultVal: 0,    format: 'number'  },
  { key: 'LTL_SEND_DATE',         label: 'LTL shipment: send date',                      defaultVal: new Date(2026, 2, 13), format: 'date' },
  { key: 'LTL_TRANSIT_DAYS',      label: 'LTL shipment: transit time (days)',             defaultVal: 14,   format: 'number'  },
  { key: 'DTC_UNITS',             label: 'DTC bridge: units',                             defaultVal: 0,    format: 'number'  },
  { key: 'DTC_START_DATE',        label: 'DTC bridge: start date',                        defaultVal: '',   format: 'date'    },
  { key: 'DTC_END_DATE',          label: 'DTC bridge: end date',                           defaultVal: '',   format: 'date'    },
  { key: 'ADHOC_UNITS',           label: 'Ad-hoc shipment: units',                        defaultVal: 0,    format: 'number'  },
  { key: 'ADHOC_SEND_DATE',       label: 'Ad-hoc shipment: send date',                    defaultVal: '',   format: 'date'    },
  { key: 'ADHOC_TRANSIT_DAYS',    label: 'Ad-hoc shipment: transit time (days)',           defaultVal: 5,    format: 'number'  }
];

/**
 * Converts a SKU id into a valid named-range prefix by replacing hyphens with
 * underscores (named ranges cannot contain hyphens).
 * @param {string} skuId  e.g. "SB-HW-100W-FBA"
 * @return {string}       e.g. "SB_HW_100W_FBA"
 */
function namedRangePrefix(skuId) {
  return skuId.replace(/-/g, '_');
}

/**
 * Builds (or rebuilds) the Settings tab from scratch.
 * Clears existing content, writes all SKU blocks, applies formatting,
 * and creates named ranges for every input cell.
 */
function buildSettingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create the Settings sheet
  var sheet = ss.getSheetByName(SETTINGS_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_TAB_NAME, 0);
  }
  sheet.clear();
  sheet.clearFormats();

  // Remove existing named ranges that belong to our SKU inputs
  var existingRanges = ss.getNamedRanges();
  for (var r = 0; r < existingRanges.length; r++) {
    var rName = existingRanges[r].getName();
    for (var s = 0; s < SKUS.length; s++) {
      if (rName.indexOf(namedRangePrefix(SKUS[s].id)) === 0) {
        existingRanges[r].remove();
        break;
      }
    }
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 350); // Labels
  sheet.setColumnWidth(2, 200); // Values

  // ── Title row ──
  var row = 1;
  sheet.getRange(row, 1).setValue('Inventory Forecasting Calendar — Settings')
       .setFontSize(14).setFontWeight('bold');
  sheet.getRange(row, 1, 1, 2).setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_FG);
  row += 1;
  sheet.getRange(row, 1).setValue('All inputs below are per SKU. Change any value to recalculate.')
       .setFontSize(10).setFontStyle('italic').setFontColor('#555555');
  row += 2; // blank spacer

  // ── Build each SKU block ──
  for (var i = 0; i < SKUS.length; i++) {
    var sku = SKUS[i];
    var prefix = namedRangePrefix(sku.id);

    // SKU header row
    sheet.getRange(row, 1, 1, 2).merge()
         .setValue(sku.id + '  —  ' + sku.name)
         .setBackground(COLORS.SECTION_BG)
         .setFontWeight('bold')
         .setFontSize(11)
         .setBorder(true, true, true, true, false, false);
    row += 1;

    // Input rows
    for (var j = 0; j < SKU_INPUT_ROWS.length; j++) {
      var input = SKU_INPUT_ROWS[j];

      // Label in column A
      sheet.getRange(row, SETTINGS_LABEL_COL)
           .setValue(input.label)
           .setFontSize(10);

      // Default value in column B
      var valueCell = sheet.getRange(row, SETTINGS_VALUE_COL);
      if (input.defaultVal !== '' && input.defaultVal !== null) {
        valueCell.setValue(input.defaultVal);
      }

      // Apply formatting
      if (input.format === 'date') {
        valueCell.setNumberFormat('m/d/yyyy');
      } else if (input.format === 'percent') {
        valueCell.setNumberFormat('0.0"%"');
      } else {
        valueCell.setNumberFormat('#,##0');
      }

      // Borders
      sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);

      // Create named range: PREFIX__KEY → cell B<row>
      var rangeName = prefix + '__' + input.key;
      ss.setNamedRange(rangeName, valueCell);

      row += 1;
    }

    // Bottom border for block
    sheet.getRange(row - 1, 1, 1, 2).setBorder(null, null, true, null, false, false);

    // Spacer row between SKU blocks
    row += 2;
  }

  // Freeze the title rows and protect structure
  sheet.setFrozenRows(2);

  // Add data validation hint: conversion rate 0-100
  // (Applied via named ranges later if needed)

  SpreadsheetApp.flush();
}


// ============================================================================
// FILE: SettingsReader.gs
// ============================================================================

/**
 * SettingsReader.gs
 * ---------------------------------------------------------------------------
 * Reads all per-SKU inputs from the Settings tab named ranges and returns
 * a structured object for consumption by the waterfall engine.
 * ---------------------------------------------------------------------------
 */

/**
 * Reads a single named range value from the spreadsheet.
 * Returns null if the range is empty or doesn't exist.
 * @param {Spreadsheet} ss
 * @param {string} rangeName
 * @return {*}
 */
function readNamedRange(ss, rangeName) {
  var namedRange = ss.getRangeByName(rangeName);
  if (!namedRange) return null;
  var val = namedRange.getValue();
  if (val === '' || val === undefined) return null;
  return val;
}

/**
 * Reads a named range value as a number, defaulting to 0 if empty.
 * @param {Spreadsheet} ss
 * @param {string} rangeName
 * @return {number}
 */
function readNum(ss, rangeName) {
  var val = readNamedRange(ss, rangeName);
  if (val === null || val === '') return 0;
  return Number(val) || 0;
}

/**
 * Reads a named range value as a Date, returning null if empty.
 * @param {Spreadsheet} ss
 * @param {string} rangeName
 * @return {Date|null}
 */
function readDate(ss, rangeName) {
  var val = readNamedRange(ss, rangeName);
  if (val === null) return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Reads all Settings inputs for a single SKU and returns a structured object.
 *
 * @param {string} skuId  e.g. "SB-HW-100W-FBA"
 * @return {Object} settings object with all fields
 */
function readSkuSettings(skuId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p  = namedRangePrefix(skuId); // e.g. "SB_HW_100W_FBA"

  return {
    skuId:             skuId,

    // ── Core inventory buckets ──
    fbaAvailable:      readNum(ss,  p + '__FBA_AVAILABLE'),
    fbaProcessing:     readNum(ss,  p + '__FBA_PROCESSING'),
    fbaCheckinDelay:   readNum(ss,  p + '__FBA_CHECKIN_DELAY'),
    fbmOnHand:         readNum(ss,  p + '__FBM_ONHAND'),

    // ── Sales velocity ──
    dailyVelocity:     readNum(ss,  p + '__DAILY_VELOCITY'),
    velOverrideStart:  readDate(ss, p + '__VEL_OVERRIDE_START'),
    velOverrideEnd:    readDate(ss, p + '__VEL_OVERRIDE_END'),
    velOverrideValue:  readNamedRange(ss, p + '__VEL_OVERRIDE_VALUE'), // null if not set, number if set (including 0)

    // ── Conversion rate ──
    conversionRate:    readNum(ss,  p + '__CONVERSION_RATE'),
    crOverrideStart:   readDate(ss, p + '__CR_OVERRIDE_START'),
    crOverrideEnd:     readDate(ss, p + '__CR_OVERRIDE_END'),
    crOverrideValue:   readNamedRange(ss, p + '__CR_OVERRIDE_VALUE'), // null if not set, number if set (including 0)

    // ── SPD shipment ──
    spdUnits:          readNum(ss,  p + '__SPD_UNITS'),
    spdSendDate:       readDate(ss, p + '__SPD_SEND_DATE'),
    spdTransitDays:    readNum(ss,  p + '__SPD_TRANSIT_DAYS'),

    // ── LTL shipment ──
    ltlUnits:          readNum(ss,  p + '__LTL_UNITS'),
    ltlSendDate:       readDate(ss, p + '__LTL_SEND_DATE'),
    ltlTransitDays:    readNum(ss,  p + '__LTL_TRANSIT_DAYS'),

    // ── DTC bridge ──
    dtcUnits:          readNum(ss,  p + '__DTC_UNITS'),
    dtcStartDate:      readDate(ss, p + '__DTC_START_DATE'),
    dtcEndDate:        readDate(ss, p + '__DTC_END_DATE'),

    // ── Ad-hoc shipment ──
    adhocUnits:        readNum(ss,  p + '__ADHOC_UNITS'),
    adhocSendDate:     readDate(ss, p + '__ADHOC_SEND_DATE'),
    adhocTransitDays:  readNum(ss,  p + '__ADHOC_TRANSIT_DAYS')
  };
}

/**
 * Reads settings for all SKUs and returns an array of settings objects.
 * @return {Object[]}
 */
function readAllSkuSettings() {
  var results = [];
  for (var i = 0; i < SKUS.length; i++) {
    results.push(readSkuSettings(SKUS[i].id));
  }
  return results;
}


// ============================================================================
// FILE: WaterfallEngine.gs
// ============================================================================

/**
 * WaterfallEngine.gs
 * ---------------------------------------------------------------------------
 * The heart of the forecasting system. For a given SKU's settings, runs
 * a day-by-day sequential simulation from START_DATE to END_DATE.
 *
 * Each day evaluates inventory buckets in priority order:
 *   1. FBA available → sell from here first
 *   2. FBM on-hand  → activates when FBA available hits zero
 *   3. DTC bridge   → manual override within user-specified date range
 *   4. TRUE OOS     → nothing left
 *
 * Inbound shipments (SPD, LTL, ad-hoc) arrive into FBA processing on their
 * calculated arrival date. FBA processing inventory becomes FBA available
 * after the SKU's configurable check-in delay.
 *
 * Returns a structured array of daily snapshots used by both the detail tab
 * renderer and the summary tab builder.
 * ---------------------------------------------------------------------------
 */

/**
 * @typedef {Object} DaySnapshot
 * @property {Date}     date              - Calendar date
 * @property {number}   fbaAvailableStart - FBA available at start of day
 * @property {number}   fbaProcessingStart- FBA processing at start of day
 * @property {number}   fbmOnHandStart    - FBM on-hand at start of day
 * @property {number}   dtcStart          - DTC remaining at start of day
 * @property {number}   unitsSold         - Units sold this day
 * @property {string}   soldFrom          - Channel that fulfilled: FBA|FBM|DTC|OOS
 * @property {number}   conversionRate    - Conversion rate applied this day (%)
 * @property {number}   fbaAvailableEnd   - FBA available at end of day
 * @property {number}   fbaProcessingEnd  - FBA processing at end of day
 * @property {number}   fbmOnHandEnd      - FBM on-hand at end of day
 * @property {number}   dtcEnd            - DTC remaining at end of day
 * @property {string}   channel           - Display channel label
 * @property {string[]} events            - Event log entries for this day
 */

/**
 * Runs the waterfall simulation for one SKU.
 *
 * @param {Object} cfg  Settings object from readSkuSettings()
 * @return {DaySnapshot[]}  Array of daily snapshots, one per forecast day
 */
function runWaterfall(cfg) {
  var dates    = forecastDates();
  var numDays  = dates.length;
  var results  = [];

  // ── Initialize running inventory balances ──
  var fbaAvail     = cfg.fbaAvailable;
  var fbaProc      = cfg.fbaProcessing;
  var fbmOnHand    = cfg.fbmOnHand;
  var dtcRemaining = cfg.dtcUnits;

  // ── Pre-compute shipment arrival dates ──
  var spdArrival   = computeArrivalDate(cfg.spdSendDate,   cfg.spdTransitDays);
  var ltlArrival   = computeArrivalDate(cfg.ltlSendDate,   cfg.ltlTransitDays);
  var adhocArrival = computeArrivalDate(cfg.adhocSendDate,  cfg.adhocTransitDays);

  // ── Track FBA processing batches (each has an "available on" date) ──
  // A batch: { units: N, availableOn: Date }
  var procBatches = [];

  // If there is starting FBA processing inventory, it becomes available
  // after the check-in delay from the start date
  if (fbaProc > 0) {
    var initAvailDate = addDays(START_DATE, cfg.fbaCheckinDelay);
    procBatches.push({ units: fbaProc, availableOn: initAvailDate });
  }

  // ── Day-by-day simulation ──
  for (var d = 0; d < numDays; d++) {
    var today  = dates[d];
    var events = [];

    // Snapshot starting balances (before any events today)
    var dayStartFbaAvail = fbaAvail;
    var dayStartFbaProc  = fbaProc;
    var dayStartFbm      = fbmOnHand;
    var dayStartDtc      = dtcRemaining;

    // ── Step 1: Check if any shipments ARRIVE today ──
    // SPD arrival
    if (spdArrival && sameDay(today, spdArrival) && cfg.spdUnits > 0) {
      fbaProc += cfg.spdUnits;
      var spdAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.spdUnits, availableOn: spdAvailOn });
      events.push('SPD shipment arrived (' + cfg.spdUnits + ' units) — enters FBA processing');
    }

    // LTL arrival
    if (ltlArrival && sameDay(today, ltlArrival) && cfg.ltlUnits > 0) {
      fbaProc += cfg.ltlUnits;
      var ltlAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.ltlUnits, availableOn: ltlAvailOn });
      events.push('LTL shipment arrived (' + cfg.ltlUnits + ' units) — enters FBA processing');
    }

    // Ad-hoc arrival
    if (adhocArrival && sameDay(today, adhocArrival) && cfg.adhocUnits > 0) {
      fbaProc += cfg.adhocUnits;
      var adhocAvailOn = addDays(today, cfg.fbaCheckinDelay);
      procBatches.push({ units: cfg.adhocUnits, availableOn: adhocAvailOn });
      events.push('Ad-hoc shipment arrived (' + cfg.adhocUnits + ' units) — enters FBA processing');
    }

    // ── Step 2: Check if any processing batches clear check-in today ──
    for (var b = procBatches.length - 1; b >= 0; b--) {
      if (sameDay(today, procBatches[b].availableOn) || today > procBatches[b].availableOn) {
        // Only process batches that become available exactly today
        // (batches from before START_DATE that should already be available are handled too)
        if (sameDay(today, procBatches[b].availableOn)) {
          var batchUnits = procBatches[b].units;
          fbaAvail += batchUnits;
          fbaProc  -= batchUnits;
          if (fbaProc < 0) fbaProc = 0; // safety clamp
          events.push(batchUnits + ' units cleared FBA processing — now FBA available');
          procBatches.splice(b, 1);
        } else if (today > procBatches[b].availableOn && d === 0) {
          // Edge case: batch was supposed to clear before forecast started
          var batchUnits2 = procBatches[b].units;
          fbaAvail += batchUnits2;
          fbaProc  -= batchUnits2;
          if (fbaProc < 0) fbaProc = 0;
          events.push(batchUnits2 + ' units cleared FBA processing (pre-forecast)');
          procBatches.splice(b, 1);
        }
      }
    }

    // ── Step 3: Determine today's velocity and conversion rate ──
    var velocity = cfg.dailyVelocity;
    if (cfg.velOverrideStart && cfg.velOverrideEnd && cfg.velOverrideValue !== null) {
      if (dateInRange(today, cfg.velOverrideStart, cfg.velOverrideEnd)) {
        velocity = Number(cfg.velOverrideValue) || 0;
      }
    }

    var convRate = cfg.conversionRate;
    if (cfg.crOverrideStart && cfg.crOverrideEnd && cfg.crOverrideValue !== null) {
      if (dateInRange(today, cfg.crOverrideStart, cfg.crOverrideEnd)) {
        convRate = Number(cfg.crOverrideValue) || 0;
      }
    }

    // Apply conversion rate to velocity
    var effectiveVelocity = velocity * (convRate / 100);
    // Round to avoid floating point drift — use banker's rounding to nearest 0.01
    effectiveVelocity = Math.round(effectiveVelocity * 100) / 100;

    // ── Step 4: Fulfillment waterfall — sequential if/else ──
    var unitsSold = 0;
    var channel   = 'OOS';

    // Check if DTC bridge is active as a manual override for this date
    var dtcActive = cfg.dtcStartDate && cfg.dtcEndDate &&
                    dateInRange(today, cfg.dtcStartDate, cfg.dtcEndDate) &&
                    dtcRemaining > 0;

    if (fbaAvail > 0) {
      // Priority 1: FBA available
      channel = 'FBA';
      unitsSold = Math.min(effectiveVelocity, fbaAvail);
      fbaAvail = roundInv(fbaAvail - unitsSold);
    } else if (fbmOnHand > 0) {
      // Priority 2: FBM on-hand (activates the moment FBA hits zero)
      channel = 'FBM';
      unitsSold = Math.min(effectiveVelocity, fbmOnHand);
      fbmOnHand = roundInv(fbmOnHand - unitsSold);
      if (dayStartFbaAvail > 0) {
        events.push('FBA stock depleted — switched to FBM');
      }
    } else if (dtcActive) {
      // Priority 3: DTC bridge (manual override within date range)
      channel = 'DTC';
      unitsSold = Math.min(effectiveVelocity, dtcRemaining);
      dtcRemaining = roundInv(dtcRemaining - unitsSold);
      events.push('Fulfilling from DTC bridge');
    } else {
      // Priority 4: TRUE OOS
      channel = 'OOS';
      unitsSold = 0;
      events.push('TRUE OOS — no inventory available on any channel');
    }

    // Check for channel switch back to FBA (if we were on FBM/DTC and FBA became available)
    if ((dayStartFbaAvail === 0) && (channel === 'FBA')) {
      events.push('FBA stock replenished — switched back to FBA');
    }

    // Log shipment send events (informational)
    if (cfg.spdSendDate && sameDay(today, cfg.spdSendDate) && cfg.spdUnits > 0) {
      events.push('SPD shipment sent (' + cfg.spdUnits + ' units)');
    }
    if (cfg.ltlSendDate && sameDay(today, cfg.ltlSendDate) && cfg.ltlUnits > 0) {
      events.push('LTL shipment sent (' + cfg.ltlUnits + ' units)');
    }
    if (cfg.adhocSendDate && sameDay(today, cfg.adhocSendDate) && cfg.adhocUnits > 0) {
      events.push('Ad-hoc shipment sent (' + cfg.adhocUnits + ' units)');
    }

    // ── Step 5: Record snapshot ──
    results.push({
      date:              today,
      fbaAvailableStart: dayStartFbaAvail,
      fbaProcessingStart:dayStartFbaProc,
      fbmOnHandStart:    dayStartFbm,
      dtcStart:          dayStartDtc,
      unitsSold:         unitsSold,
      soldFrom:          channel,
      conversionRate:    convRate,
      effectiveVelocity: effectiveVelocity,
      fbaAvailableEnd:   fbaAvail,
      fbaProcessingEnd:  fbaProc,
      fbmOnHandEnd:      fbmOnHand,
      dtcEnd:            dtcRemaining,
      channel:           channel,
      events:            events
    });
  }

  return results;
}

/**
 * Computes the arrival date given a send date and transit time in days.
 * Returns null if send date is not set.
 * @param {Date|null} sendDate
 * @param {number}    transitDays
 * @return {Date|null}
 */
function computeArrivalDate(sendDate, transitDays) {
  if (!sendDate) return null;
  return addDays(sendDate, transitDays);
}

/**
 * Rounds an inventory balance to 2 decimal places and clamps to zero.
 * Prevents floating-point drift from accumulating over 71 days of
 * fractional subtraction (e.g., velocity 7 * conversion 33% = 2.31/day).
 * @param {number} val
 * @return {number}
 */
function roundInv(val) {
  var rounded = Math.round(val * 100) / 100;
  return rounded < 0 ? 0 : rounded;
}

/**
 * Adds a number of days to a date and returns a new Date.
 * @param {Date}   d
 * @param {number} days
 * @return {Date}
 */
function addDays(d, days) {
  var result = new Date(d);
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * Extracts milestone events from a waterfall result set.
 * Used by the Summary tab to show the milestone table.
 *
 * @param {DaySnapshot[]} snapshots
 * @param {Object}        cfg  - SKU settings for shipment dates
 * @return {Object} milestones
 */
function extractMilestones(snapshots, cfg) {
  var milestones = {
    lastFbaDay:     null,
    firstFbmDay:    null,
    spdArrival:     computeArrivalDate(cfg.spdSendDate, cfg.spdTransitDays),
    ltlArrival:     computeArrivalDate(cfg.ltlSendDate, cfg.ltlTransitDays),
    firstBackOnFba: null,
    oosGaps:        []     // array of {start: Date, end: Date}
  };

  var initialFbaEnded = false;  // true once the initial FBA run ends
  var wasOos          = false;
  var oosStart        = null;
  var foundFirstFbm   = false;
  var foundBackOnFba  = false;

  for (var i = 0; i < snapshots.length; i++) {
    var snap = snapshots[i];

    // Track the last day of the INITIAL FBA stock run (before first switch away)
    if (!initialFbaEnded && snap.channel === 'FBA') {
      milestones.lastFbaDay = snap.date;
    } else if (!initialFbaEnded && snap.channel !== 'FBA' && milestones.lastFbaDay !== null) {
      initialFbaEnded = true; // initial FBA run has ended
    }

    // First day on FBM
    if (!foundFirstFbm && snap.channel === 'FBM') {
      milestones.firstFbmDay = snap.date;
      foundFirstFbm = true;
    }

    // First day back on FBA (after being on FBM/DTC/OOS)
    if (!foundBackOnFba && initialFbaEnded && snap.channel === 'FBA') {
      milestones.firstBackOnFba = snap.date;
      foundBackOnFba = true;
    }

    // OOS gaps
    if (snap.channel === 'OOS') {
      if (!wasOos) {
        oosStart = snap.date;
        wasOos = true;
      }
    } else {
      if (wasOos && oosStart) {
        milestones.oosGaps.push({ start: oosStart, end: snapshots[i - 1].date });
        wasOos = false;
        oosStart = null;
      }
    }
  }

  // Close any open OOS gap
  if (wasOos && oosStart) {
    milestones.oosGaps.push({ start: oosStart, end: snapshots[snapshots.length - 1].date });
  }

  return milestones;
}


// ============================================================================
// FILE: DetailTab.gs
// ============================================================================

/**
 * DetailTab.gs
 * ---------------------------------------------------------------------------
 * Renders one SKU detail tab as a monthly calendar grid view — similar to
 * a phone calendar app. Each month (Feb, Mar, Apr 2026) is displayed as
 * a 7-column grid (Sun–Sat). Each day cell is color-coded by fulfillment
 * channel and shows key inventory data at a glance.
 *
 * Day cell contents:
 *   Line 1: Day number
 *   Line 2: Channel (FBA / FBM / DTC / OOS)
 *   Line 3: Sold X units
 *   Line 4: End FBA: X | FBM: X
 *   Line 5: Event (if any)
 * ---------------------------------------------------------------------------
 */

var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

/**
 * Returns the background color hex for a given fulfillment channel.
 * @param {string} channel
 * @return {string} hex color
 */
function getChannelColor(channel) {
  switch (channel) {
    case 'FBA':  return COLORS.FBA;
    case 'FBM':  return COLORS.FBM;
    case 'DTC':  return COLORS.DTC;
    case 'OOS':  return COLORS.OOS;
    default:     return COLORS.WHITE;
  }
}

/**
 * Builds a lookup map from date string to snapshot for fast access.
 * @param {DaySnapshot[]} data
 * @return {Object} map of "YYYY-M-D" -> DaySnapshot
 */
function buildSnapshotMap(data) {
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var d = data[i].date;
    var key = d.getFullYear() + '-' + d.getMonth() + '-' + d.getDate();
    map[key] = data[i];
  }
  return map;
}

/**
 * Formats a number compactly: shows integer if whole, one decimal otherwise.
 * @param {number} n
 * @return {string}
 */
function fmtNum(n) {
  if (n === Math.floor(n)) return String(Math.round(n));
  return (Math.round(n * 10) / 10).toString();
}

/**
 * Builds (or rebuilds) a single SKU detail tab as a calendar grid.
 *
 * @param {Object} skuDef      SKU definition from SKUS array
 * @param {DaySnapshot[]} data Waterfall results from runWaterfall()
 */
function buildDetailTab(skuDef, data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(skuDef.tab);

  if (!sheet) {
    sheet = ss.insertSheet(skuDef.tab);
  }
  sheet.clear();
  sheet.clearFormats();

  // Build snapshot lookup
  var snapMap = buildSnapshotMap(data);

  // Determine which months to render
  // Forecast: Feb 19 – Apr 30, 2026
  var months = [
    { year: 2026, month: 1, name: 'February 2026' },  // month is 0-indexed
    { year: 2026, month: 2, name: 'March 2026' },
    { year: 2026, month: 3, name: 'April 2026' }
  ];

  var row = 1;
  var numCols = 7; // Sun–Sat

  // ── Title ──
  sheet.getRange(row, 1, 1, numCols).merge()
       .setValue(skuDef.id + '  —  ' + skuDef.name)
       .setFontSize(13)
       .setFontWeight('bold')
       .setBackground(COLORS.HEADER)
       .setFontColor(COLORS.HEADER_FG)
       .setHorizontalAlignment('left');
  row += 1;

  // ── Legend row ──
  var legendItems = [
    { label: 'FBA', color: COLORS.FBA },
    { label: 'FBM', color: COLORS.FBM },
    { label: 'DTC', color: COLORS.DTC },
    { label: 'OOS', color: COLORS.OOS },
    { label: 'ARRIVING', color: COLORS.PROCESSING },
    { label: 'PROCESSING', color: COLORS.GRAY }
  ];
  for (var li = 0; li < legendItems.length; li++) {
    sheet.getRange(row, li + 1)
         .setValue(legendItems[li].label)
         .setBackground(legendItems[li].color)
         .setHorizontalAlignment('center')
         .setFontSize(8)
         .setFontWeight('bold')
         .setBorder(true, true, true, true, false, false);
  }
  // Fill remaining cell in 7-col row
  sheet.getRange(row, 7)
       .setBackground(COLORS.WHITE)
       .setBorder(true, true, true, true, false, false);
  row += 2;

  // ── Column widths ──
  for (var cw = 1; cw <= numCols; cw++) {
    sheet.setColumnWidth(cw, 145);
  }

  // ── Render each month ──
  for (var mi = 0; mi < months.length; mi++) {
    var mo = months[mi];

    // Month header
    sheet.getRange(row, 1, 1, numCols).merge()
         .setValue(mo.name)
         .setFontSize(12)
         .setFontWeight('bold')
         .setBackground('#D6E4F0')
         .setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false);
    row += 1;

    // Day-of-week header
    for (var dh = 0; dh < numCols; dh++) {
      sheet.getRange(row, dh + 1)
           .setValue(DAY_NAMES[dh])
           .setFontWeight('bold')
           .setFontSize(9)
           .setBackground('#E2EFDA')
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
    }
    row += 1;

    // Determine calendar grid for this month
    var firstOfMonth = new Date(mo.year, mo.month, 1);
    var daysInMonth  = new Date(mo.year, mo.month + 1, 0).getDate();
    var startDow     = firstOfMonth.getDay(); // 0=Sun

    var day = 1;
    var weekRow = row;

    // Fill weeks
    while (day <= daysInMonth) {
      var cellValues = [];
      var cellBgs    = [];

      for (var col = 0; col < numCols; col++) {
        if ((weekRow === row && col < startDow) || day > daysInMonth) {
          // Empty cell (before first day or after last day of month)
          cellValues.push('');
          cellBgs.push('#F5F5F5');
        } else {
          var cellDate = new Date(mo.year, mo.month, day);
          var key = mo.year + '-' + mo.month + '-' + day;
          var snap = snapMap[key];

          if (snap) {
            // Build cell content
            var lines = [];
            lines.push(String(day));
            lines.push(snap.channel);
            lines.push('Sold: ' + fmtNum(snap.unitsSold));

            // Show ending inventory for active buckets
            var invParts = [];
            if (snap.fbaAvailableEnd > 0 || snap.channel === 'FBA') {
              invParts.push('FBA: ' + fmtNum(snap.fbaAvailableEnd));
            }
            if (snap.fbmOnHandEnd > 0 || snap.channel === 'FBM') {
              invParts.push('FBM: ' + fmtNum(snap.fbmOnHandEnd));
            }
            if (snap.dtcEnd > 0 || snap.channel === 'DTC') {
              invParts.push('DTC: ' + fmtNum(snap.dtcEnd));
            }
            if (snap.fbaProcessingEnd > 0) {
              invParts.push('Proc: ' + fmtNum(snap.fbaProcessingEnd));
            }
            if (invParts.length > 0) {
              lines.push(invParts.join(' | '));
            }

            // Show first event if any (truncated)
            if (snap.events.length > 0) {
              var evt = snap.events[0];
              if (evt.length > 30) evt = evt.substring(0, 28) + '..';
              lines.push(evt);
            }

            cellValues.push(lines.join('\n'));

            // Determine background color
            var bg = getChannelColor(snap.channel);
            // Check for arrival event
            var hasArrival = false;
            for (var ei = 0; ei < snap.events.length; ei++) {
              if (snap.events[ei].indexOf('arrived') > -1) {
                hasArrival = true;
                break;
              }
            }
            if (hasArrival) bg = COLORS.PROCESSING;

            cellBgs.push(bg);
          } else {
            // Date is outside the forecast range
            cellValues.push(String(day) + '\n—');
            cellBgs.push('#F5F5F5');
          }
          day += 1;
        }
      }

      // Write this week's row
      for (var wc = 0; wc < numCols; wc++) {
        var cell = sheet.getRange(weekRow, wc + 1);
        cell.setValue(cellValues[wc])
            .setBackground(cellBgs[wc])
            .setFontSize(8)
            .setVerticalAlignment('top')
            .setWrap(true)
            .setBorder(true, true, true, true, false, false);
      }

      // Set row height to fit calendar cells
      sheet.setRowHeight(weekRow, 85);
      weekRow += 1;
    }

    row = weekRow + 1; // spacer between months
  }

  // ── Bold the day numbers in each cell ──
  // (Can't selectively bold within a cell via script without RichTextValue;
  //  we handle this by making the first line stand out via the cell layout)

  // ── Freeze title ──
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}


// ============================================================================
// FILE: SummaryTab.gs
// ============================================================================

/**
 * SummaryTab.gs
 * ---------------------------------------------------------------------------
 * Builds the Summary tab with two sections:
 *
 * 1. MILESTONE TABLE — Per-SKU row showing key dates:
 *    Last FBA Day, First FBM Day, SPD Arrival, LTL Arrival,
 *    First Day Back on FBA, OOS Gaps
 *    Uses merged cells so each milestone field spans multiple calendar-width
 *    columns, avoiding the column-width conflict between the milestone table
 *    and the narrow calendar grid below.
 *
 * 2. CALENDAR GRID — Rows = SKUs, Columns = each calendar day.
 *    Each cell shows a short label (FBA, FBM, DTC, OOS) and is
 *    color-coded per the palette in Config.gs.
 *
 * Designed for clean, professional sharing with supply chain partners.
 * ---------------------------------------------------------------------------
 */

/**
 * Milestone field definitions.
 * Each field spans a number of 48px calendar columns so the merged cell is
 * wide enough to display the header and value text.
 *   colSpan: how many 48px columns to merge for this field.
 */
var MILESTONE_FIELDS = [
  { header: 'SKU',                     colSpan: 5 },  // 5 * 48 = 240px
  { header: 'Last FBA Day',            colSpan: 3 },  // 3 * 48 = 144px
  { header: 'First FBM Day',           colSpan: 3 },  // 144px
  { header: 'SPD Arrival',             colSpan: 3 },  // 144px
  { header: 'LTL Arrival',             colSpan: 3 },  // 144px
  { header: 'First Day Back on FBA',   colSpan: 3 },  // 144px
  { header: 'OOS Gaps',                colSpan: 6 },  // 6 * 48 = 288px
  { header: 'Total OOS Days',          colSpan: 2 }   // 2 * 48 = 96px
];

/**
 * Builds the Summary tab from waterfall results for all SKUs.
 *
 * @param {Object[]} allResults  Array of { skuDef, cfg, snapshots, milestones }
 */
function buildSummaryTab(allResults) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SUMMARY_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_TAB_NAME, 1); // position after Settings
  }
  sheet.clear();
  sheet.clearFormats();

  var dates   = forecastDates();
  var numDays = dates.length;
  var numSkus = allResults.length;

  // Compute the total column span of the milestone table
  var msTotalCols = 0;
  for (var mf = 0; mf < MILESTONE_FIELDS.length; mf++) {
    msTotalCols += MILESTONE_FIELDS[mf].colSpan;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 1: MILESTONE TABLE (uses merged cells across calendar columns)
  // ══════════════════════════════════════════════════════════════════════════

  var msRow = 1;

  // Title — span across the milestone field columns
  sheet.getRange(msRow, 1, 1, msTotalCols).merge()
       .setValue('Inventory Forecasting Calendar — Milestone Summary')
       .setFontSize(13)
       .setFontWeight('bold')
       .setBackground(COLORS.HEADER)
       .setFontColor(COLORS.HEADER_FG);
  msRow += 1;

  // ── Milestone headers (merged cells) ──
  var col = 1;
  for (var mh = 0; mh < MILESTONE_FIELDS.length; mh++) {
    var field = MILESTONE_FIELDS[mh];
    var headerRange = sheet.getRange(msRow, col, 1, field.colSpan);
    headerRange.merge()
               .setValue(field.header)
               .setFontWeight('bold')
               .setFontSize(9)
               .setBackground('#E2EFDA')
               .setBorder(true, true, true, true, false, false)
               .setHorizontalAlignment('center')
               .setWrap(true);
    col += field.colSpan;
  }
  msRow += 1;

  // ── Milestone data rows (merged cells per field) ──
  for (var s = 0; s < numSkus; s++) {
    var r   = allResults[s];
    var ms  = r.milestones;
    var oosGapStr = '';
    var totalOosDays = 0;

    if (ms.oosGaps.length > 0) {
      var parts = [];
      for (var g = 0; g < ms.oosGaps.length; g++) {
        var gap = ms.oosGaps[g];
        var gapDays = Math.round((gap.end - gap.start) / 86400000) + 1;
        totalOosDays += gapDays;
        parts.push(fmtDate(gap.start) + ' to ' + fmtDate(gap.end) + ' (' + gapDays + 'd)');
      }
      oosGapStr = parts.join(', ');
    } else {
      oosGapStr = 'None';
    }

    var msValues = [
      r.skuDef.id,
      ms.lastFbaDay     ? fmtDate(ms.lastFbaDay)     : 'N/A',
      ms.firstFbmDay    ? fmtDate(ms.firstFbmDay)    : 'N/A',
      ms.spdArrival     ? fmtDate(ms.spdArrival)     : 'N/A',
      ms.ltlArrival     ? fmtDate(ms.ltlArrival)     : 'N/A',
      ms.firstBackOnFba ? fmtDate(ms.firstBackOnFba) : 'N/A',
      oosGapStr,
      totalOosDays
    ];

    col = 1;
    for (var mv = 0; mv < MILESTONE_FIELDS.length; mv++) {
      var fld = MILESTONE_FIELDS[mv];
      var cellRange = sheet.getRange(msRow, col, 1, fld.colSpan);
      cellRange.merge()
               .setValue(msValues[mv])
               .setFontSize(9)
               .setBorder(true, true, true, true, false, false)
               .setHorizontalAlignment(mv === 0 ? 'left' : 'center')
               .setWrap(true);

      // Highlight OOS gaps in red
      if (mv === 6 && totalOosDays > 0) {
        cellRange.setBackground(COLORS.OOS);
      }
      if (mv === 7 && totalOosDays > 0) {
        cellRange.setBackground(COLORS.OOS);
      }

      col += fld.colSpan;
    }

    msRow += 1;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 2: COLOR LEGEND
  // ══════════════════════════════════════════════════════════════════════════

  var legendRow = msRow + 1;
  sheet.getRange(legendRow, 1, 1, 12).merge()
       .setValue('Color Legend')
       .setFontWeight('bold')
       .setFontSize(10)
       .setBackground('#F2F2F2');
  legendRow += 1;

  var legendItems = [
    { label: 'FBA',            color: COLORS.FBA },
    { label: 'FBM',            color: COLORS.FBM },
    { label: 'DTC',            color: COLORS.DTC },
    { label: 'OOS',            color: COLORS.OOS },
    { label: 'ARRIVING',       color: COLORS.PROCESSING },
    { label: 'PROCESSING',     color: COLORS.GRAY }
  ];

  // Each legend item spans 2 columns for readability
  for (var li = 0; li < legendItems.length; li++) {
    var legendCol = 1 + (li * 2);
    sheet.getRange(legendRow, legendCol, 1, 2).merge()
         .setValue(legendItems[li].label)
         .setBackground(legendItems[li].color)
         .setHorizontalAlignment('center')
         .setFontSize(9)
         .setBorder(true, true, true, true, false, false);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 3: CALENDAR GRID
  // ══════════════════════════════════════════════════════════════════════════

  var calStartRow = legendRow + 2;

  // Calendar title
  sheet.getRange(calStartRow, 1, 1, numDays + 1).merge()
       .setValue('Daily Fulfillment Channel Calendar — ' + fmtDate(START_DATE) + ' through ' + fmtDate(END_DATE))
       .setFontSize(12)
       .setFontWeight('bold')
       .setBackground(COLORS.HEADER)
       .setFontColor(COLORS.HEADER_FG);
  calStartRow += 1;

  // ── Column headers: "SKU" + each date ──
  var calHeaders = ['SKU'];
  for (var d = 0; d < numDays; d++) {
    calHeaders.push(fmtDateShort(dates[d]));
  }
  var calHeaderRange = sheet.getRange(calStartRow, 1, 1, numDays + 1);
  calHeaderRange.setValues([calHeaders])
                .setFontWeight('bold')
                .setFontSize(8)
                .setBackground('#E2EFDA')
                .setBorder(true, true, true, true, true, true)
                .setHorizontalAlignment('center')
                .setWrap(true);

  // ── Column widths ──
  // All columns use the same 48px width for the calendar grid.
  // Column 1 (SKU) is wider. The milestone table above uses merged cells
  // so it is not affected by these column widths.
  sheet.setColumnWidth(1, 220);
  for (var dw = 2; dw <= numDays + 1; dw++) {
    sheet.setColumnWidth(dw, 48);
  }
  calStartRow += 1;

  // ── Data rows: one per SKU ──
  var calData    = [];
  var calBgAll   = [];
  var calFgAll   = [];

  for (var si = 0; si < numSkus; si++) {
    var result    = allResults[si];
    var snapshots = result.snapshots;
    var rowData   = [result.skuDef.id];
    var rowBg     = [COLORS.WHITE];
    var rowFg     = ['#000000'];

    for (var di = 0; di < numDays; di++) {
      var snap = snapshots[di];
      var label = snap.channel;

      // Check if a shipment arrived today — show PROCESSING label
      var isArrival = false;
      for (var ei = 0; ei < snap.events.length; ei++) {
        if (snap.events[ei].indexOf('arrived') > -1) {
          isArrival = true;
          break;
        }
      }

      if (isArrival && snap.channel !== 'FBA') {
        // Shipment arrived but not yet FBA available — mark with asterisk
        label = label + '*';
      }

      rowData.push(label);
      rowBg.push(getChannelColor(snap.channel));
      rowFg.push('#000000');

      // Override color for arrival days
      if (isArrival) {
        rowBg[rowBg.length - 1] = COLORS.PROCESSING;
      }
    }

    calData.push(rowData);
    calBgAll.push(rowBg);
    calFgAll.push(rowFg);
  }

  // Write calendar grid in one batch
  var calDataRange = sheet.getRange(calStartRow, 1, numSkus, numDays + 1);
  calDataRange.setValues(calData);
  calDataRange.setBackgrounds(calBgAll);
  calDataRange.setFontColors(calFgAll);
  calDataRange.setFontSize(8);
  calDataRange.setHorizontalAlignment('center');
  calDataRange.setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

  // SKU column in calendar should be left-aligned and wider
  sheet.getRange(calStartRow, 1, numSkus, 1)
       .setHorizontalAlignment('left')
       .setFontWeight('bold')
       .setFontSize(8);

  // ── Freeze ──
  sheet.setFrozenRows(calStartRow - 1); // Freeze everything above the data
  sheet.setFrozenColumns(1);

  SpreadsheetApp.flush();
}


// ============================================================================
// FILE: Main.gs
// ============================================================================

/**
 * Main.gs
 * ---------------------------------------------------------------------------
 * Orchestration layer — ties together all modules.
 *
 * Entry points:
 *   onOpen()          — Adds a custom menu to the spreadsheet
 *   onEditTrigger(e)  — Installable onEdit trigger; recalculates when
 *                        Settings tab values change
 *   initialSetup()    — First-time setup: builds Settings tab, runs forecast
 *   recalculateAll()  — Full recalculation of all SKU detail tabs + Summary
 *   installTrigger()  — Creates the installable onEdit trigger
 * ---------------------------------------------------------------------------
 */

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Forecast')
    .addItem('Initial Setup (first time)', 'initialSetup')
    .addItem('Recalculate All', 'recalculateAll')
    .addItem('Install Auto-Refresh Trigger', 'installTrigger')
    .addToUi();
}

/**
 * First-time setup: builds the Settings tab with defaults, then runs
 * a full forecast calculation.
 */
function initialSetup() {
  var ui = SpreadsheetApp.getUi();

  // Build the Settings tab with all per-SKU input sections
  buildSettingsTab();

  // Run the full forecast
  recalculateAll();

  // Install the onEdit trigger
  installTrigger();

  ui.alert(
    'Setup Complete',
    'The Settings tab has been created with default values for all 6 SKUs.\n\n' +
    'Detail tabs and the Summary tab have been generated.\n\n' +
    'An onEdit trigger has been installed — any change to the Settings tab ' +
    'will automatically recalculate all forecasts.\n\n' +
    'Enter your actual inventory numbers in the Settings tab to begin forecasting.',
    ui.ButtonSet.OK
  );
}

/**
 * Full recalculation: reads all SKU settings, runs the waterfall engine
 * for each SKU, renders all detail tabs, and rebuilds the Summary tab.
 */
function recalculateAll() {
  var allSettings = readAllSkuSettings();
  var allResults  = [];

  for (var i = 0; i < SKUS.length; i++) {
    var skuDef    = SKUS[i];
    var cfg       = allSettings[i];
    var snapshots = runWaterfall(cfg);
    var milestones = extractMilestones(snapshots, cfg);

    // Build the detail tab for this SKU
    buildDetailTab(skuDef, snapshots);

    allResults.push({
      skuDef:     skuDef,
      cfg:        cfg,
      snapshots:  snapshots,
      milestones: milestones
    });
  }

  // Build the Summary tab with all results
  buildSummaryTab(allResults);

  // Move Summary tab to position 2 (after Settings)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) {
    ss.setActiveSheet(summarySheet);
    ss.moveActiveSheet(2);
  }

  // Ensure Settings is first
  var settingsSheet = ss.getSheetByName(SETTINGS_TAB_NAME);
  if (settingsSheet) {
    ss.setActiveSheet(settingsSheet);
    ss.moveActiveSheet(1);
  }

  SpreadsheetApp.flush();
}

/**
 * Installable onEdit trigger handler.
 * Fires on every edit; only recalculates if the edit was on the Settings tab.
 *
 * @param {Object} e  Event object from the trigger
 */
function onEditTrigger(e) {
  // Guard: only recalculate if the edit happened on the Settings tab
  if (!e || !e.range) return;

  // Use e.range.getSheet() which is authoritative for the edited cell,
  // unlike e.source.getActiveSheet() which can be unreliable if the user
  // switches tabs quickly after editing.
  var editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== SETTINGS_TAB_NAME) return;

  // Only recalculate if the edit was in the values column (column B).
  // Edits to column A (labels) should not trigger a recalculation.
  if (e.range.getColumn() !== SETTINGS_VALUE_COL) return;

  // Debounce: use a lock to prevent overlapping recalculations
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(2000); // wait up to 2 seconds
  if (!acquired) return; // another recalc is already running

  try {
    recalculateAll();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Creates an installable onEdit trigger for the active spreadsheet.
 * Removes any existing triggers with the same handler name first to
 * avoid duplicates.
 */
function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new installable onEdit trigger
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Utility: Deletes all SKU detail tabs and the Summary tab.
 * Useful for a clean rebuild during development.
 */
function deleteGeneratedTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete SKU detail tabs
  for (var i = 0; i < SKUS.length; i++) {
    var sheet = ss.getSheetByName(SKUS[i].tab);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  // Delete Summary tab
  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) {
    ss.deleteSheet(summarySheet);
  }
}

