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

  // ── Read velocity overrides (up to 3) ──
  var velOverrides = [];
  for (var v = 1; v <= 3; v++) {
    var vStart = readDate(ss, p + '__VEL_OVERRIDE_' + v + '_START');
    var vEnd   = readDate(ss, p + '__VEL_OVERRIDE_' + v + '_END');
    var vVal   = readNamedRange(ss, p + '__VEL_OVERRIDE_' + v + '_VALUE');
    if (vStart && vEnd && vVal !== null) {
      velOverrides.push({ start: vStart, end: vEnd, value: Number(vVal) || 0 });
    }
  }

  // ── Read conversion rate overrides (up to 3) ──
  var crOverrides = [];
  for (var c = 1; c <= 3; c++) {
    var cStart = readDate(ss, p + '__CR_OVERRIDE_' + c + '_START');
    var cEnd   = readDate(ss, p + '__CR_OVERRIDE_' + c + '_END');
    var cVal   = readNamedRange(ss, p + '__CR_OVERRIDE_' + c + '_VALUE');
    if (cStart && cEnd && cVal !== null) {
      crOverrides.push({ start: cStart, end: cEnd, value: Number(cVal) || 0 });
    }
  }

  return {
    skuId:             skuId,

    // ── Core inventory buckets ──
    fbaAvailable:      readNum(ss,  p + '__FBA_AVAILABLE'),
    fbaProcessing:     readNum(ss,  p + '__FBA_PROCESSING'),
    fbaCheckinDelay:   readNum(ss,  p + '__FBA_CHECKIN_DELAY'),
    fbmOnHand:         readNum(ss,  p + '__FBM_ONHAND'),
    reserved:          readNum(ss,  p + '__RESERVED'),
    researching:       readNum(ss,  p + '__RESEARCHING'),
    unfulfillable:     readNum(ss,  p + '__UNFULFILLABLE'),

    // ── Sales velocity ──
    dailyVelocity:     readNum(ss,  p + '__DAILY_VELOCITY'),
    velOverrides:      velOverrides,

    // ── Conversion rate ──
    conversionRate:    readNum(ss,  p + '__CONVERSION_RATE'),
    crOverrides:       crOverrides,

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
