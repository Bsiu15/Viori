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
  // FINANCIAL IMPACT — split into Stock-Out / FBM / Net sections
  // ══════════════════════════════════════════════════════════════════════════

  if (hasFinancials) {

    // ── Aggregate totals across months ──
    var totalOosDays = 0, totalLostRev = 0;
    var totalFbmDays = 0, totalFbmSaved = 0, totalFbmCost = 0;
    var oosMonths = [];   // months with OOS activity
    var fbmMonths = [];   // months with FBM activity

    for (var fm = 0; fm < months.length; fm++) {
      var mName = months[fm].name;
      var mf = monthFinancials[mName];
      totalOosDays += mf.oosDays;
      totalLostRev += mf.lostRevenue;
      totalFbmDays += mf.fbmDays;
      totalFbmSaved += mf.fbmRevenue;
      totalFbmCost += mf.fbmCost;
      if (mf.oosDays > 0 || mf.lostRevenue > 0) { mf.name = mName; oosMonths.push(mf); }
      if (mf.fbmDays > 0 || mf.fbmRevenue > 0 || mf.fbmCost > 0) { mf.name = mName; fbmMonths.push(mf); }
    }

    // Already-lost (user-entered past OOS)
    var pastOosDays = (cfg && cfg.pastOosDays) ? cfg.pastOosDays : 0;
    var pastEffVel  = (cfg && cfg.dailyVelocity) ? cfg.dailyVelocity : 0;
    var pastLostRev = pastOosDays * pastEffVel * sellingPrice;

    // Grand totals
    var grandDays = pastOosDays + totalOosDays;
    var grandRev  = pastLostRev + totalLostRev;
    var grandNet  = -(grandRev + totalFbmCost);

    // ── Helper: write a section title row (full-width merge) ──
    function writeSectionTitle(title) {
      row += 1;
      sheet.getRange(row, 1, 1, numCols).merge()
           .setValue(title)
           .setFontSize(11)
           .setFontWeight('bold')
           .setBackground(COLORS.HEADER)
           .setFontColor(COLORS.HEADER_FG)
           .setHorizontalAlignment('left');
      row += 1;
    }

    // ── Helper: set cell with common formatting ──
    function fmtCell(r, c, val, opts) {
      var cell = sheet.getRange(r, c).setValue(val)
           .setFontSize(10)
           .setHorizontalAlignment(opts.align || 'center')
           .setBorder(true, true, true, true, false, false);
      if (opts.bold) cell.setFontWeight('bold');
      if (opts.bg) cell.setBackground(opts.bg);
      if (opts.fmt) cell.setNumberFormat(opts.fmt);
      return cell;
    }

    // ────────────────────────────────────────────────────────────
    // SECTION 1 — Stock-Out Impact
    // ────────────────────────────────────────────────────────────
    writeSectionTitle('Stock-Out Impact — ' + skuDef.id);

    // Headers (3 columns)
    var oosHeaders = ['Month', 'OOS Days', 'Lost Revenue'];
    var oosHdrBgs  = ['#E2EFDA', '#E2EFDA', COLORS.OOS];
    for (var oh = 0; oh < oosHeaders.length; oh++) {
      fmtCell(row, oh + 1, oosHeaders[oh], { bold: true, bg: oosHdrBgs[oh] });
    }
    // Blank out remaining columns on header row
    if (numCols > oosHeaders.length) {
      sheet.getRange(row, oosHeaders.length + 1, 1, numCols - oosHeaders.length)
           .setBackground(null).setBorder(false, false, false, false, false, false);
    }
    row += 1;

    // Monthly rows (only months with OOS)
    for (var oi = 0; oi < oosMonths.length; oi++) {
      var om = oosMonths[oi];
      fmtCell(row, 1, om.name, { align: 'left' });
      fmtCell(row, 2, om.oosDays, { bg: om.oosDays > 0 ? COLORS.OOS : null });
      fmtCell(row, 3, om.lostRevenue, { align: 'right', fmt: '$#,##0', bg: om.lostRevenue > 0 ? COLORS.OOS : null });
      row += 1;
    }
    if (oosMonths.length === 0 && pastOosDays === 0) {
      // No OOS at all — show a clean "None" row
      sheet.getRange(row, 1, 1, 3).merge()
           .setValue('No stock-outs detected')
           .setFontSize(10).setFontColor('#5f6368')
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
      row += 1;
    }

    // Already Lost
    fmtCell(row, 1, 'ALREADY LOST', { bold: true, align: 'left', bg: '#FFF2CC' });
    fmtCell(row, 2, pastOosDays, { bold: true, bg: '#FFF2CC' });
    fmtCell(row, 3, pastLostRev, { bold: true, align: 'right', fmt: '$#,##0', bg: '#FFF2CC' });
    row += 1;

    // Projected
    fmtCell(row, 1, 'PROJECTED', { bold: true, align: 'left', bg: '#F2F2F2' });
    fmtCell(row, 2, totalOosDays, { bold: true, bg: totalOosDays > 0 ? COLORS.OOS : '#F2F2F2' });
    fmtCell(row, 3, totalLostRev, { bold: true, align: 'right', fmt: '$#,##0', bg: totalLostRev > 0 ? COLORS.OOS : '#F2F2F2' });
    row += 1;

    // Total
    fmtCell(row, 1, 'TOTAL', { bold: true, align: 'left', bg: '#D6E4F0' });
    fmtCell(row, 2, grandDays, { bold: true, bg: grandDays > 0 ? COLORS.OOS : '#D6E4F0' });
    fmtCell(row, 3, grandRev, { bold: true, align: 'right', fmt: '$#,##0', bg: grandRev > 0 ? COLORS.OOS : '#D6E4F0' });
    row += 1;

    // ────────────────────────────────────────────────────────────
    // SECTION 2 — FBM Backup Performance
    // ────────────────────────────────────────────────────────────
    writeSectionTitle('FBM Backup Performance — ' + skuDef.id);

    // Headers (4 columns)
    var fbmHeaders = ['Month', 'FBM Days', 'Revenue Saved', 'Fulfillment Cost'];
    var fbmHdrBgs  = ['#E2EFDA', '#E2EFDA', '#C6EFCE', '#FFF2CC'];
    for (var fhi = 0; fhi < fbmHeaders.length; fhi++) {
      fmtCell(row, fhi + 1, fbmHeaders[fhi], { bold: true, bg: fbmHdrBgs[fhi] });
    }
    if (numCols > fbmHeaders.length) {
      sheet.getRange(row, fbmHeaders.length + 1, 1, numCols - fbmHeaders.length)
           .setBackground(null).setBorder(false, false, false, false, false, false);
    }
    row += 1;

    // Monthly rows (only months with FBM activity)
    for (var fi = 0; fi < fbmMonths.length; fi++) {
      var fb = fbmMonths[fi];
      fmtCell(row, 1, fb.name, { align: 'left' });
      fmtCell(row, 2, fb.fbmDays, {});
      fmtCell(row, 3, fb.fbmRevenue, { align: 'right', fmt: '$#,##0', bg: fb.fbmRevenue > 0 ? '#C6EFCE' : null });
      fmtCell(row, 4, fb.fbmCost, { align: 'right', fmt: '$#,##0', bg: fb.fbmCost > 0 ? '#FFF2CC' : null });
      row += 1;
    }
    if (fbmMonths.length === 0) {
      sheet.getRange(row, 1, 1, 4).merge()
           .setValue('No FBM backup activity')
           .setFontSize(10).setFontColor('#5f6368')
           .setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
      row += 1;
    }

    // Projected (no "Already Lost" for FBM — it's forward-looking only)
    fmtCell(row, 1, 'PROJECTED', { bold: true, align: 'left', bg: '#F2F2F2' });
    fmtCell(row, 2, totalFbmDays, { bold: true, bg: '#F2F2F2' });
    fmtCell(row, 3, totalFbmSaved, { bold: true, align: 'right', fmt: '$#,##0', bg: totalFbmSaved > 0 ? '#C6EFCE' : '#F2F2F2' });
    fmtCell(row, 4, totalFbmCost, { bold: true, align: 'right', fmt: '$#,##0', bg: totalFbmCost > 0 ? '#FFF2CC' : '#F2F2F2' });
    row += 1;

    // ────────────────────────────────────────────────────────────
    // SECTION 3 — Net Financial Impact (vertical summary)
    // ────────────────────────────────────────────────────────────
    row += 1;
    sheet.getRange(row, 1, 1, numCols).merge()
         .setValue('Net Financial Impact — ' + skuDef.id)
         .setFontSize(11)
         .setFontWeight('bold')
         .setBackground(COLORS.HEADER)
         .setFontColor(COLORS.HEADER_FG)
         .setHorizontalAlignment('left');
    row += 1;

    var netRows = [
      { label: 'Total Lost Revenue',     value: grandRev,      bg: grandRev > 0 ? COLORS.OOS : '#F2F2F2',  fmt: '$#,##0' },
      { label: 'FBM Revenue Saved',      value: totalFbmSaved, bg: totalFbmSaved > 0 ? '#C6EFCE' : '#F2F2F2', fmt: '$#,##0' },
      { label: 'FBM Fulfillment Cost',   value: totalFbmCost,  bg: totalFbmCost > 0 ? '#FFF2CC' : '#F2F2F2',  fmt: '$#,##0' },
      { label: 'NET IMPACT',             value: grandNet,      bg: '#D6E4F0', fmt: '$#,##0;-$#,##0;$0' }
    ];

    for (var ni = 0; ni < netRows.length; ni++) {
      var nr = netRows[ni];
      var isLast = (ni === netRows.length - 1);
      fmtCell(row, 1, nr.label, { bold: true, align: 'left', bg: nr.bg });
      fmtCell(row, 2, nr.value, { bold: isLast, align: 'right', fmt: nr.fmt, bg: nr.bg });
      // Extend the background across remaining columns for a clean look
      if (numCols > 2) {
        sheet.getRange(row, 3, 1, numCols - 2).merge()
             .setBackground(nr.bg)
             .setBorder(true, true, true, true, false, false);
      }
      row += 1;
    }
  }

  // ── Bold the day numbers in each cell ──
  // (Can't selectively bold within a cell via script without RichTextValue;
  //  we handle this by making the first line stand out via the cell layout)

  // ── Freeze title ──
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}
