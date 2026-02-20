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
