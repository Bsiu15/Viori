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
 *   Line 6: Lost: $X rev (OOS days only, when financials are set)
 *
 * Below the calendar grid, a unified Financial Impact table shows:
 *   Per-month: OOS Days, Lost Rev, FBM Days, FBM Saved, FBM Cost, Net Impact
 * ---------------------------------------------------------------------------
 */

var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

/**
 * Returns the background color hex for a given fulfillment channel.
 * For split channels like 'FBA→FBM', uses the last channel's color
 * (the most notable/terminal state).
 * @param {string} channel
 * @return {string} hex color
 */
function getChannelColor(channel) {
  // For split channels, use the last segment's color
  var key = channel;
  if (channel.indexOf('→') > -1) {
    var parts = channel.split('→');
    key = parts[parts.length - 1];
  }
  switch (key) {
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
 * Formats a dollar amount compactly.
 * @param {number} n
 * @return {string}
 */
function fmtDollar(n) {
  if (n >= 1000) return '$' + (Math.round(n / 100) / 10).toFixed(1) + 'k';
  return '$' + (Math.round(n * 100) / 100).toFixed(2);
}

/**
 * Builds (or rebuilds) a single SKU detail tab as a calendar grid.
 *
 * @param {Object} skuDef      SKU definition from SKUS array
 * @param {DaySnapshot[]} data Waterfall results from runWaterfall()
 * @param {Object} [cfg]       SKU settings (for sellingPrice and dppMargin)
 */
function buildDetailTab(skuDef, data, cfg) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(skuDef.tab);

  if (!sheet) {
    sheet = ss.insertSheet(skuDef.tab);
  }
  sheet.clear();
  sheet.clearFormats();

  // Financial fields
  var sellingPrice = (cfg && cfg.sellingPrice) ? cfg.sellingPrice : 0;
  var hasFinancials = sellingPrice > 0;
  var fbmPrice = (cfg && cfg.fbmSellingPrice > 0) ? cfg.fbmSellingPrice : sellingPrice;

  // Build snapshot lookup
  var snapMap = buildSnapshotMap(data);

  // Determine which months to render (dynamic from START_DATE → END_DATE)
  var months = forecastMonths();

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

  // Track OOS and FBM financial data per month for the summary table
  // monthKey -> { oosDays, lostRevenue, fbmDays, fbmRevenue, fbmCost }
  var monthFinancials = {};
  for (var mi2 = 0; mi2 < months.length; mi2++) {
    monthFinancials[months[mi2].name] = {
      oosDays: 0, lostRevenue: 0,
      fbmDays: 0, fbmRevenue: 0, fbmCost: 0
    };
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

    // Day-of-week header (batch)
    var dowRange = sheet.getRange(row, 1, 1, numCols);
    dowRange.setValues([DAY_NAMES])
            .setFontWeight('bold')
            .setFontSize(9)
            .setBackground('#E2EFDA')
            .setHorizontalAlignment('center')
            .setBorder(true, true, true, true, false, false);
    row += 1;

    // Determine calendar grid for this month
    var firstOfMonth = new Date(mo.year, mo.month, 1);
    var daysInMonth  = new Date(mo.year, mo.month + 1, 0).getDate();
    var startDow     = firstOfMonth.getDay(); // 0=Sun

    var day = 1;
    var weekStartRow = row;
    var isFirstWeek = true;

    // Collect all weeks for this month, then write in batch
    var allWeekValues = [];
    var allWeekBgs    = [];

    while (day <= daysInMonth) {
      var cellValues = [];
      var cellBgs    = [];

      for (var col = 0; col < numCols; col++) {
        if ((isFirstWeek && col < startDow) || day > daysInMonth) {
          // Empty cell (before first day or after last day of month)
          cellValues.push('');
          cellBgs.push('#F5F5F5');
        } else {
          var key = mo.year + '-' + mo.month + '-' + day;
          var snap = snapMap[key];

          if (snap) {
            // Build cell content
            var lines = [];
            lines.push(String(day));
            lines.push(snap.channel);

            // Show sold breakdown for split days, simple total for single-channel
            if (snap.channel.indexOf('→') > -1 && snap.unitsSold > 0) {
              var soldParts = [];
              if (snap.soldFromFba > 0) soldParts.push(fmtNum(snap.soldFromFba) + ' FBA');
              if (snap.soldFromFbm > 0) soldParts.push(fmtNum(snap.soldFromFbm) + ' FBM');
              if (snap.soldFromDtc > 0) soldParts.push(fmtNum(snap.soldFromDtc) + ' DTC');
              lines.push('Sold: ' + fmtNum(snap.unitsSold) + ' (' + soldParts.join('+') + ')');
            } else {
              lines.push('Sold: ' + fmtNum(snap.unitsSold));
            }

            // Show ending inventory for active buckets
            var invParts = [];
            if (snap.fbaAvailableEnd > 0 || snap.soldFromFba > 0) {
              invParts.push('FBA: ' + fmtNum(snap.fbaAvailableEnd));
            }
            if (snap.fbmOnHandEnd > 0 || snap.soldFromFbm > 0) {
              invParts.push('FBM: ' + fmtNum(snap.fbmOnHandEnd));
            }
            if (snap.dtcEnd > 0 || snap.soldFromDtc > 0) {
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

            // OOS financial impact line — includes partial OOS (split days with unfulfilled demand)
            if (snap.unfulfilledUnits > 0 && hasFinancials) {
              var dayLostRev = snap.unfulfilledUnits * (snap.effectivePrice || sellingPrice);
              lines.push('Lost: ' + fmtDollar(dayLostRev) + ' rev');

              // Accumulate for monthly summary
              monthFinancials[mo.name].oosDays += 1; // any unfulfilled demand = 1 whole OOS day
              monthFinancials[mo.name].lostRevenue += dayLostRev;
            }

            // FBM cost impact line — uses per-channel sold amounts for accuracy
            if (snap.soldFromFbm > 0) {
              monthFinancials[mo.name].fbmDays += 1;
              monthFinancials[mo.name].fbmRevenue += snap.soldFromFbm * fbmPrice;
              if (snap.fbmCostPerUnit > 0) {
                var dayFbmCost = snap.soldFromFbm * snap.fbmCostPerUnit;
                lines.push('FBM cost: ' + fmtDollar(dayFbmCost));
                monthFinancials[mo.name].fbmCost += dayFbmCost;
              }
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

      allWeekValues.push(cellValues);
      allWeekBgs.push(cellBgs);
      isFirstWeek = false;
    }

    // Write entire month's calendar grid in one batch
    var numWeeks = allWeekValues.length;
    var monthRange = sheet.getRange(weekStartRow, 1, numWeeks, numCols);
    monthRange.setValues(allWeekValues);
    monthRange.setBackgrounds(allWeekBgs);
    monthRange.setFontSize(8);
    monthRange.setVerticalAlignment('top');
    monthRange.setWrap(true);
    monthRange.setBorder(true, true, true, true, false, false);

    // Set row heights for calendar cells
    for (var rh = 0; rh < numWeeks; rh++) {
      sheet.setRowHeight(weekStartRow + rh, 95);
    }

    row = weekStartRow + numWeeks + 1; // spacer between months
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FINANCIAL IMPACT TABLE (single unified table below calendar)
  // ══════════════════════════════════════════════════════════════════════════

  if (hasFinancials) {
    row += 1;

    // Title
    sheet.getRange(row, 1, 1, numCols).merge()
         .setValue('Inventory Gap Financial Impact — ' + skuDef.id)
         .setFontSize(12)
         .setFontWeight('bold')
         .setBackground(COLORS.HEADER)
         .setFontColor(COLORS.HEADER_FG)
         .setHorizontalAlignment('left');
    row += 1;

    // Headers: 7 columns matching the calendar grid width
    var finHeaders = ['Month', 'OOS Days', 'Lost Rev', 'FBM Days', 'FBM Saved', 'FBM Cost', 'Net Impact'];
    var finHdrBgs  = ['#E2EFDA', '#E2EFDA', COLORS.OOS, '#E2EFDA', '#C6EFCE', '#FFF2CC', '#D6E4F0'];
    for (var fh = 0; fh < finHeaders.length; fh++) {
      sheet.getRange(row, fh + 1)
           .setValue(finHeaders[fh])
           .setFontWeight('bold')
           .setFontSize(10)
           .setBackground(finHdrBgs[fh])
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
    }
    row += 1;

    // Monthly rows
    var totalOosDays = 0, totalLostRev = 0;
    var totalFbmDays = 0, totalFbmSaved = 0, totalFbmCost = 0;

    for (var fm = 0; fm < months.length; fm++) {
      var mName = months[fm].name;
      var mf = monthFinancials[mName];
      totalOosDays += mf.oosDays;
      totalLostRev += mf.lostRevenue;
      totalFbmDays += mf.fbmDays;
      totalFbmSaved += mf.fbmRevenue;
      totalFbmCost += mf.fbmCost;

      var mNetImpact = -(mf.lostRevenue + mf.fbmCost);

      sheet.getRange(row, 1).setValue(mName)
           .setFontSize(10).setHorizontalAlignment('left')
           .setBorder(true, true, true, true, false, false);

      var oosDaysCell = sheet.getRange(row, 2).setValue(mf.oosDays)
           .setFontSize(10).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
      if (mf.oosDays > 0) oosDaysCell.setBackground(COLORS.OOS);

      var revCell = sheet.getRange(row, 3).setValue(mf.lostRevenue)
           .setNumberFormat('$#,##0')
           .setFontSize(10).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (mf.lostRevenue > 0) revCell.setBackground(COLORS.OOS);

      sheet.getRange(row, 4).setValue(mf.fbmDays)
           .setFontSize(10).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);

      var fbmSavedCell = sheet.getRange(row, 5).setValue(mf.fbmRevenue)
           .setNumberFormat('$#,##0')
           .setFontSize(10).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (mf.fbmRevenue > 0) fbmSavedCell.setBackground('#C6EFCE');

      var fbmCostCell = sheet.getRange(row, 6).setValue(mf.fbmCost)
           .setNumberFormat('$#,##0')
           .setFontSize(10).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (mf.fbmCost > 0) fbmCostCell.setBackground('#FFF2CC');

      sheet.getRange(row, 7).setValue(mNetImpact)
           .setNumberFormat('$#,##0;-$#,##0;$0')
           .setFontSize(10).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false)
           .setBackground(mNetImpact < 0 ? '#D6E4F0' : null);

      row += 1;
    }

    // ── Already Lost row (from user-entered past OOS days) ──
    var pastOosDays = (cfg && cfg.pastOosDays) ? cfg.pastOosDays : 0;
    var pastEffVel  = (cfg && cfg.dailyVelocity) ? cfg.dailyVelocity : 0;
    var pastLostRev = pastOosDays * pastEffVel * sellingPrice;
    var pastNet = -pastLostRev;

    sheet.getRange(row, 1).setValue('ALREADY LOST')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#FFF2CC');
    sheet.getRange(row, 2).setValue(pastOosDays)
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#FFF2CC');
    sheet.getRange(row, 3).setValue(pastLostRev)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#FFF2CC');
    // FBM Days, Saved, Cost = N/A for past
    for (var pc = 4; pc <= 6; pc++) {
      sheet.getRange(row, pc).setValue('—')
           .setFontSize(10).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false)
           .setBackground('#FFF2CC');
    }
    sheet.getRange(row, 7).setValue(pastNet)
         .setNumberFormat('$#,##0;-$#,##0;$0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#FFF2CC');
    row += 1;

    // ── Projected row (simulation totals) ──
    var projNet = -(totalLostRev + totalFbmCost);

    sheet.getRange(row, 1).setValue('PROJECTED')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#F2F2F2');

    var projOosDaysC = sheet.getRange(row, 2).setValue(totalOosDays)
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false);
    if (totalOosDays > 0) projOosDaysC.setBackground(COLORS.OOS);
    else projOosDaysC.setBackground('#F2F2F2');

    var projRevC = sheet.getRange(row, 3).setValue(totalLostRev)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (totalLostRev > 0) projRevC.setBackground(COLORS.OOS);
    else projRevC.setBackground('#F2F2F2');

    sheet.getRange(row, 4).setValue(totalFbmDays)
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#F2F2F2');

    var projFbmSavedC = sheet.getRange(row, 5).setValue(totalFbmSaved)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (totalFbmSaved > 0) projFbmSavedC.setBackground('#C6EFCE');
    else projFbmSavedC.setBackground('#F2F2F2');

    var projFbmCostC = sheet.getRange(row, 6).setValue(totalFbmCost)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (totalFbmCost > 0) projFbmCostC.setBackground('#FFF2CC');
    else projFbmCostC.setBackground('#F2F2F2');

    sheet.getRange(row, 7).setValue(projNet)
         .setNumberFormat('$#,##0;-$#,##0;$0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false)
         .setBackground(projNet < 0 ? '#D6E4F0' : '#F2F2F2');
    row += 1;

    // ── Total Impact row (past + projected combined) ──
    var grandDays = pastOosDays + totalOosDays;
    var grandRev  = pastLostRev + totalLostRev;
    var grandNet  = -(grandRev + totalFbmCost);

    sheet.getRange(row, 1).setValue('TOTAL IMPACT')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#D6E4F0');
    var totalDaysC = sheet.getRange(row, 2).setValue(grandDays)
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false);
    if (grandDays > 0) totalDaysC.setBackground(COLORS.OOS);
    else totalDaysC.setBackground('#D6E4F0');

    var totalRevC = sheet.getRange(row, 3).setValue(grandRev)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (grandRev > 0) totalRevC.setBackground(COLORS.OOS);
    else totalRevC.setBackground('#D6E4F0');

    sheet.getRange(row, 4).setValue(totalFbmDays)
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#D6E4F0');

    var totalFbmSavedC = sheet.getRange(row, 5).setValue(totalFbmSaved)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (totalFbmSaved > 0) totalFbmSavedC.setBackground('#C6EFCE');
    else totalFbmSavedC.setBackground('#D6E4F0');

    var totalFbmCostC = sheet.getRange(row, 6).setValue(totalFbmCost)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (totalFbmCost > 0) totalFbmCostC.setBackground('#FFF2CC');
    else totalFbmCostC.setBackground('#D6E4F0');

    sheet.getRange(row, 7).setValue(grandNet)
         .setNumberFormat('$#,##0;-$#,##0;$0')
         .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false)
         .setBackground(grandNet < 0 ? COLORS.OOS : '#D6E4F0');
  }

  // ── Bold the day numbers in each cell ──
  // (Can't selectively bold within a cell via script without RichTextValue;
  //  we handle this by making the first line stand out via the cell layout)

  // ── Freeze title ──
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}
