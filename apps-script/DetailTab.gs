/**
 * DetailTab.gs
 * ---------------------------------------------------------------------------
 * Renders one SKU detail tab showing the full day-by-day waterfall breakdown.
 * Columns:
 *   A  Date
 *   B  FBA Available (Start)
 *   C  FBA Processing (Start)
 *   D  FBM On-Hand (Start)
 *   E  DTC Remaining (Start)
 *   F  Units Sold
 *   G  Sold From (channel)
 *   H  Conversion Rate
 *   I  Effective Velocity
 *   J  FBA Available (End)
 *   K  FBA Processing (End)
 *   L  FBM On-Hand (End)
 *   M  DTC Remaining (End)
 *   N  Channel
 *   O  Events
 * ---------------------------------------------------------------------------
 */

/** Column headers for SKU detail tabs */
var DETAIL_HEADERS = [
  'Date',
  'FBA Avail (Start)',
  'FBA Processing (Start)',
  'FBM On-Hand (Start)',
  'DTC Remaining (Start)',
  'Units Sold',
  'Sold From',
  'Conv. Rate (%)',
  'Eff. Velocity',
  'FBA Avail (End)',
  'FBA Processing (End)',
  'FBM On-Hand (End)',
  'DTC Remaining (End)',
  'Channel',
  'Events'
];

/**
 * Builds (or rebuilds) a single SKU detail tab.
 *
 * @param {Object} skuDef      SKU definition from SKUS array (has .id, .name, .tab)
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

  var numCols = DETAIL_HEADERS.length;
  var numRows = data.length;

  // ── Title row ──
  sheet.getRange(1, 1, 1, numCols).merge()
       .setValue(skuDef.id + '  —  ' + skuDef.name + '  —  Day-by-Day Waterfall')
       .setFontSize(12)
       .setFontWeight('bold')
       .setBackground(COLORS.HEADER)
       .setFontColor(COLORS.HEADER_FG)
       .setHorizontalAlignment('left');

  // ── Header row ──
  var headerRange = sheet.getRange(2, 1, 1, numCols);
  headerRange.setValues([DETAIL_HEADERS])
             .setFontWeight('bold')
             .setBackground('#E2EFDA')
             .setBorder(true, true, true, true, true, true)
             .setHorizontalAlignment('center')
             .setWrap(true);

  // ── Data rows ──
  var output     = [];
  var bgColors   = [];
  var fgColors   = [];

  for (var i = 0; i < numRows; i++) {
    var snap = data[i];
    var row  = [
      snap.date,
      snap.fbaAvailableStart,
      snap.fbaProcessingStart,
      snap.fbmOnHandStart,
      snap.dtcStart,
      Math.round(snap.unitsSold * 100) / 100,
      snap.soldFrom,
      snap.conversionRate,
      Math.round(snap.effectiveVelocity * 100) / 100,
      snap.fbaAvailableEnd,
      snap.fbaProcessingEnd,
      snap.fbmOnHandEnd,
      snap.dtcEnd,
      snap.channel,
      snap.events.join('; ')
    ];
    output.push(row);

    // ── Row background color based on channel ──
    var bg = getChannelColor(snap.channel);
    // Check if a shipment arrived today (blue highlight for the channel column)
    var hasArrival = false;
    for (var e = 0; e < snap.events.length; e++) {
      if (snap.events[e].indexOf('arrived') > -1) {
        hasArrival = true;
        break;
      }
    }

    var rowBg = [];
    var rowFg = [];
    for (var c = 0; c < numCols; c++) {
      rowBg.push(bg);
      rowFg.push('#000000');
    }

    // Override processing columns with gray if FBA processing > 0
    if (snap.fbaProcessingStart > 0 || snap.fbaProcessingEnd > 0) {
      rowBg[2]  = COLORS.GRAY;  // FBA Processing (Start)
      rowBg[10] = COLORS.GRAY;  // FBA Processing (End)
    }

    // Override events column with blue if shipment arrived
    if (hasArrival) {
      rowBg[14] = COLORS.PROCESSING; // Events column
    }

    bgColors.push(rowBg);
    fgColors.push(rowFg);
  }

  // Write data in one batch
  var dataRange = sheet.getRange(3, 1, numRows, numCols);
  dataRange.setValues(output);
  dataRange.setBackgrounds(bgColors);
  dataRange.setFontColors(fgColors);
  dataRange.setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

  // ── Formatting ──
  // Date column
  sheet.getRange(3, 1, numRows, 1).setNumberFormat('m/d/yyyy');
  // Number columns
  for (var nc = 2; nc <= 6; nc++) {
    sheet.getRange(3, nc, numRows, 1).setNumberFormat('#,##0.00');
  }
  sheet.getRange(3, 8, numRows, 1).setNumberFormat('#,##0.0"%"');
  sheet.getRange(3, 9, numRows, 1).setNumberFormat('#,##0.00');
  for (var nc2 = 10; nc2 <= 13; nc2++) {
    sheet.getRange(3, nc2, numRows, 1).setNumberFormat('#,##0.00');
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 100);  // Date
  for (var w = 2; w <= 13; w++) {
    sheet.setColumnWidth(w, 110);
  }
  sheet.setColumnWidth(14, 90);  // Channel
  sheet.setColumnWidth(15, 400); // Events

  // ── Freeze headers ──
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(1);

  // ── Alternating row borders for readability ──
  sheet.getRange(2, 1, numRows + 1, numCols).setVerticalAlignment('middle');
}

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
