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

  // ── Read FBA override (manual override for total FBA available) ──
  var fbaOverride = readNamedRange(ss, p + '__FBA_OVERRIDE');

  // ── Read granular inventory fields (mirrors Seller Central) ──
  var inboundWorking    = readNum(ss, p + '__INBOUND_WORKING');
  var inboundShipped    = readNum(ss, p + '__INBOUND_SHIPPED');
  var inboundReceiving  = readNum(ss, p + '__INBOUND_RECEIVING');
  var onhandAvailable   = readNum(ss, p + '__ONHAND_AVAILABLE');
  var onhandFcTransfer  = readNum(ss, p + '__ONHAND_FC_TRANSFER');
  var reservedCustOrder = readNum(ss, p + '__RESERVED_CUSTOMER_ORDER');
  var reservedFcProc    = readNum(ss, p + '__RESERVED_FC_PROCESSING');
  var researching       = readNum(ss, p + '__RESEARCHING');
  var unfulfillWhDmg    = readNum(ss, p + '__UNFULFILLABLE_WAREHOUSE_DAMAGED');
  var unfulfillDefect   = readNum(ss, p + '__UNFULFILLABLE_DEFECTIVE');
  var unfulfillExpired  = readNum(ss, p + '__UNFULFILLABLE_EXPIRED');
  var unfulfillCustDmg  = readNum(ss, p + '__UNFULFILLABLE_CUSTOMER_DAMAGED');
  var unfulfillCarrDmg  = readNum(ss, p + '__UNFULFILLABLE_CARRIER_DAMAGED');
  var unfulfillDistDmg  = readNum(ss, p + '__UNFULFILLABLE_DISTRIBUTOR_DAMAGED');

  // ── Read financial fields ──
  var sellingPriceRaw = readNamedRange(ss, p + '__SELLING_PRICE');
  var sellingPrice    = (sellingPriceRaw !== null) ? Number(sellingPriceRaw) || 0 : 0;
  var dppMarginRaw    = readNamedRange(ss, p + '__DPP_MARGIN');
  var dppMargin       = (dppMarginRaw !== null) ? Number(dppMarginRaw) || 0 : 0;

  // FBA available: use manual override if set, otherwise use On-hand Available
  var effectiveFbaAvailable = (fbaOverride !== null) ? Number(fbaOverride) || 0 : onhandAvailable;

  return {
    skuId:             skuId,

    // ── Engine inventory buckets (derived from granular fields) ──
    // FBA override takes precedence; otherwise On-hand Available is used
    fbaAvailable:      effectiveFbaAvailable,
    // Inbound Receiving → units at Amazon being checked in (becomes available after delay)
    fbaProcessing:     inboundReceiving,
    fbaCheckinDelay:   readNum(ss,  p + '__FBA_CHECKIN_DELAY'),
    fbmOnHand:         readNum(ss,  p + '__FBM_ONHAND'),
    // Customer orders → subtracted from FBA available on day 1 (already spoken for)
    customerOrders:    reservedCustOrder,

    // ── Granular inventory (informational — stored for display) ──
    inboundWorking:    inboundWorking,
    inboundShipped:    inboundShipped,
    inboundReceiving:  inboundReceiving,
    onhandAvailable:   onhandAvailable,
    onhandFcTransfer:  onhandFcTransfer,
    reservedCustOrder: reservedCustOrder,
    reservedFcProc:    reservedFcProc,
    researching:       researching,
    unfulfillWhDmg:    unfulfillWhDmg,
    unfulfillDefect:   unfulfillDefect,
    unfulfillExpired:  unfulfillExpired,
    unfulfillCustDmg:  unfulfillCustDmg,
    unfulfillCarrDmg:  unfulfillCarrDmg,
    unfulfillDistDmg:  unfulfillDistDmg,

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
    adhocTransitDays:  readNum(ss,  p + '__ADHOC_TRANSIT_DAYS'),

    // ── Financials ──
    sellingPrice:      sellingPrice,
    dppMargin:         dppMargin,
    pastOosDays:       readNum(ss,  p + '__PAST_OOS_DAYS')
  };
}

/**
 * Reads settings for all SKUs and returns an array of settings objects.
 * @return {Object[]}
 */
function readAllSkuSettings() {
  var skus = getSkus();
  var results = [];
  for (var i = 0; i < skus.length; i++) {
    results.push(readSkuSettings(skus[i].id));
  }
  return results;
}
