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

  // ── Legend row (batch) ──
  var legendLabels = ['FBA', 'FBM', 'DTC', 'OOS', 'ARRIVING', 'PROCESSING', ''];
  var legendBgs    = [COLORS.FBA, COLORS.FBM, COLORS.DTC, COLORS.OOS, COLORS.PROCESSING, COLORS.GRAY, COLORS.WHITE];
  var legendRange  = sheet.getRange(row, 1, 1, numCols);
  legendRange.setValues([legendLabels]);
  legendRange.setBackgrounds([legendBgs]);
  legendRange.setHorizontalAlignment('center');
  legendRange.setFontSize(8);
  legendRange.setFontWeight('bold');
  legendRange.setBorder(true, true, true, true, true, true);
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
      fbmDays: 0, fbmRevenue: 0, fbmCost: 0, fbmFbaEquivRevenue: 0
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

              // FBA-equivalent revenue: what FBA would have earned on this FBM day
              if (snap.soldFromFba === 0) {
                // FBM was primary channel — full demand was at FBM rate, not FBA rate
                monthFinancials[mo.name].fbmFbaEquivRevenue += snap.fbaEffVelocity * sellingPrice;
              } else {
                // Split day — FBA-equiv = what FBA would have earned in the remaining fraction
                var fbaRemainder = snap.fbaEffVelocity - snap.soldFromFba;
                monthFinancials[mo.name].fbmFbaEquivRevenue += fbaRemainder * sellingPrice;
              }

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

    // ── Highlight today's cell with a thick border ──
    var todayDate = new Date();
    todayDate.setHours(0, 0, 0, 0);
    if (todayDate.getFullYear() === mo.year && todayDate.getMonth() === mo.month) {
      var todayDay = todayDate.getDate();
      var todayOffset = todayDay - 1 + startDow; // 0-indexed position in the grid
      var todayWeekRow = Math.floor(todayOffset / 7);
      var todayColIdx  = (todayOffset % 7) + 1;  // 1-indexed column
      var todayCell = sheet.getRange(weekStartRow + todayWeekRow, todayColIdx);
      todayCell.setBorder(true, true, true, true, false, false,
        '#1a1a1a', SpreadsheetApp.BorderStyle.SOLID_THICK);
    }

    row = weekStartRow + numWeeks + 1; // spacer between months
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FINANCIAL IMPACT — split into Stock-Out / FBM / Net sections
  // ══════════════════════════════════════════════════════════════════════════

  if (hasFinancials) {

    // ── Aggregate totals across months ──
    var totalOosDays = 0, totalLostRev = 0;
    var totalFbmDays = 0, totalFbmSaved = 0, totalFbmCost = 0, totalFbmFbaEquiv = 0;
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
      totalFbmFbaEquiv += mf.fbmFbaEquivRevenue;
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
    // FBM demand penalty = what FBA would have earned − what FBM actually earned
    var fbmDemandPenalty = Math.max(0, totalFbmFbaEquiv - totalFbmSaved);
    // NET = all losses: OOS lost rev + FBM demand gap + FBM fulfillment cost
    var grandNet  = -(grandRev + fbmDemandPenalty + totalFbmCost);

    // ── FBA In-Stock Rate ──
    var totalForecastDays = data.length;
    var fbaDays = 0;
    for (var isd = 0; isd < data.length; isd++) {
      if (data[isd].soldFromFba > 0) fbaDays++;
    }
    var fbaInStockRate = totalForecastDays > 0 ? Math.round((fbaDays / totalForecastDays) * 100) : 0;

    // ── Helper: fill a 7-wide row array (pad remaining cols with fillVal) ──
    function padRow(arr, fillVal) {
      var r = arr.slice();
      while (r.length < numCols) r.push(fillVal !== undefined ? fillVal : '');
      return r;
    }

    // ── Build all financial rows in memory, then write in batch ──
    // Each row tracks: values, backgrounds, font weights, alignments, number formats
    var finValues = [], finBgs = [], finWeights = [], finAligns = [], finFormats = [];
    var fillNull = null; // default bg

    // Helper to push a row into all arrays at once
    function pushRow(vals, bgs, weights, aligns, formats) {
      finValues.push(padRow(vals, ''));
      finBgs.push(padRow(bgs, fillNull));
      finWeights.push(padRow(weights, 'normal'));
      finAligns.push(padRow(aligns, 'center'));
      finFormats.push(padRow(formats, ''));
    }

    // ────────────────────────────────────────────────────────────
    // SECTION 1 — Stock-Out Impact
    // ────────────────────────────────────────────────────────────

    // Spacer row
    pushRow(['','','','','','',''], [null,null,null,null,null,null,null], ['normal','normal','normal','normal','normal','normal','normal'], ['center','center','center','center','center','center','center'], ['','','','','','','']);

    // Title row (will be merged after write)
    pushRow(
      ['Stock-Out Impact — ' + skuDef.id, '', '', '', '', '', ''],
      [COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER],
      ['bold','bold','bold','bold','bold','bold','bold'],
      ['left','left','left','left','left','left','left'],
      ['','','','','','','']
    );
    var oosTitleIdx = finValues.length - 1;

    // Headers
    pushRow(
      ['Month', 'OOS Days', 'Lost Revenue', '', '', '', ''],
      ['#E2EFDA', '#E2EFDA', COLORS.OOS, null, null, null, null],
      ['bold', 'bold', 'bold', 'normal', 'normal', 'normal', 'normal'],
      ['center', 'center', 'center', 'center', 'center', 'center', 'center'],
      ['', '', '', '', '', '', '']
    );

    // Monthly rows (only months with OOS)
    if (oosMonths.length === 0 && pastOosDays === 0) {
      pushRow(
        ['No stock-outs detected', '', '', '', '', '', ''],
        [null, null, null, null, null, null, null],
        ['normal','normal','normal','normal','normal','normal','normal'],
        ['center','center','center','center','center','center','center'],
        ['','','','','','','']
      );
      var noOosMergeIdx = finValues.length - 1;
    }
    for (var oi = 0; oi < oosMonths.length; oi++) {
      var om = oosMonths[oi];
      pushRow(
        [om.name, om.oosDays, om.lostRevenue, '', '', '', ''],
        [null, om.oosDays > 0 ? COLORS.OOS : null, om.lostRevenue > 0 ? COLORS.OOS : null, null, null, null, null],
        ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal'],
        ['left', 'center', 'right', 'center', 'center', 'center', 'center'],
        ['', '', '$#,##0', '', '', '', '']
      );
    }

    // Already Lost
    pushRow(
      ['ALREADY LOST', pastOosDays, pastLostRev, '', '', '', ''],
      ['#FFF2CC', '#FFF2CC', '#FFF2CC', null, null, null, null],
      ['bold', 'bold', 'bold', 'normal', 'normal', 'normal', 'normal'],
      ['left', 'center', 'right', 'center', 'center', 'center', 'center'],
      ['', '', '$#,##0', '', '', '', '']
    );

    // Projected
    pushRow(
      ['PROJECTED', totalOosDays, totalLostRev, '', '', '', ''],
      ['#F2F2F2', totalOosDays > 0 ? COLORS.OOS : '#F2F2F2', totalLostRev > 0 ? COLORS.OOS : '#F2F2F2', null, null, null, null],
      ['bold', 'bold', 'bold', 'normal', 'normal', 'normal', 'normal'],
      ['left', 'center', 'right', 'center', 'center', 'center', 'center'],
      ['', '', '$#,##0', '', '', '', '']
    );

    // Total
    pushRow(
      ['TOTAL', grandDays, grandRev, '', '', '', ''],
      ['#D6E4F0', grandDays > 0 ? COLORS.OOS : '#D6E4F0', grandRev > 0 ? COLORS.OOS : '#D6E4F0', null, null, null, null],
      ['bold', 'bold', 'bold', 'normal', 'normal', 'normal', 'normal'],
      ['left', 'center', 'right', 'center', 'center', 'center', 'center'],
      ['', '', '$#,##0', '', '', '', '']
    );

    // ────────────────────────────────────────────────────────────
    // SECTION 2 — FBM Fallback Cost (NOT on FBA = bad)
    // ────────────────────────────────────────────────────────────

    // Spacer
    pushRow(['','','','','','',''], [null,null,null,null,null,null,null], ['normal','normal','normal','normal','normal','normal','normal'], ['center','center','center','center','center','center','center'], ['','','','','','','']);

    // Title
    pushRow(
      ['FBM Fallback Cost — ' + skuDef.id, '', '', '', '', '', ''],
      [COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER],
      ['bold','bold','bold','bold','bold','bold','bold'],
      ['left','left','left','left','left','left','left'],
      ['','','','','','','']
    );
    var fbmTitleIdx = finValues.length - 1;

    // Headers — everything framed as loss/cost
    pushRow(
      ['Month', 'Days Off FBA', 'FBA Would Earn', 'FBM Earned', 'Revenue Gap', 'Fulfill. Cost', ''],
      ['#E2EFDA', COLORS.OOS, '#BDD7EE', '#FFF2CC', COLORS.OOS, '#FFF2CC', null],
      ['bold', 'bold', 'bold', 'bold', 'bold', 'bold', 'normal'],
      ['center', 'center', 'center', 'center', 'center', 'center', 'center'],
      ['', '', '', '', '', '', '']
    );

    // Monthly rows
    if (fbmMonths.length === 0) {
      pushRow(
        ['All days on FBA — no FBM fallback needed', '', '', '', '', '', ''],
        [null, null, null, null, null, null, null],
        ['normal','normal','normal','normal','normal','normal','normal'],
        ['center','center','center','center','center','center','center'],
        ['','','','','','','']
      );
      var noFbmMergeIdx = finValues.length - 1;
    }
    for (var fi = 0; fi < fbmMonths.length; fi++) {
      var fb = fbmMonths[fi];
      var fbPenalty = Math.max(0, fb.fbmFbaEquivRevenue - fb.fbmRevenue);
      pushRow(
        [fb.name, fb.fbmDays, fb.fbmFbaEquivRevenue, fb.fbmRevenue, fbPenalty, fb.fbmCost, ''],
        [null, fb.fbmDays > 0 ? COLORS.OOS : null, null, fb.fbmRevenue > 0 ? '#FFF2CC' : null, fbPenalty > 0 ? COLORS.OOS : null, fb.fbmCost > 0 ? '#FFF2CC' : null, null],
        ['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal'],
        ['left', 'center', 'right', 'right', 'right', 'right', 'center'],
        ['', '', '$#,##0', '$#,##0', '$#,##0', '$#,##0', '']
      );
    }

    // Projected
    pushRow(
      ['PROJECTED', totalFbmDays, totalFbmFbaEquiv, totalFbmSaved, fbmDemandPenalty, totalFbmCost, ''],
      ['#F2F2F2', totalFbmDays > 0 ? COLORS.OOS : '#F2F2F2', '#F2F2F2', totalFbmSaved > 0 ? '#FFF2CC' : '#F2F2F2', fbmDemandPenalty > 0 ? COLORS.OOS : '#F2F2F2', totalFbmCost > 0 ? '#FFF2CC' : '#F2F2F2', null],
      ['bold', 'bold', 'bold', 'bold', 'bold', 'bold', 'normal'],
      ['left', 'center', 'right', 'right', 'right', 'right', 'center'],
      ['', '', '$#,##0', '$#,##0', '$#,##0', '$#,##0', '']
    );

    // ────────────────────────────────────────────────────────────
    // SECTION 3 — Net Financial Impact
    // ────────────────────────────────────────────────────────────

    // Spacer
    pushRow(['','','','','','',''], [null,null,null,null,null,null,null], ['normal','normal','normal','normal','normal','normal','normal'], ['center','center','center','center','center','center','center'], ['','','','','','','']);

    // Title
    pushRow(
      ['Net Financial Impact — ' + skuDef.id, '', '', '', '', '', ''],
      [COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER, COLORS.HEADER],
      ['bold','bold','bold','bold','bold','bold','bold'],
      ['left','left','left','left','left','left','left'],
      ['','','','','','','']
    );
    var netTitleIdx = finValues.length - 1;

    // Net summary rows — everything framed as cost / loss (no green FBM line)
    var netDefs = [
      { label: 'Lost Revenue (Stock-Outs)',  value: -grandRev,           bg: grandRev > 0 ? COLORS.OOS : '#F2F2F2',              fmt: '$#,##0', bold: true },
      { label: 'FBM Revenue Gap (vs FBA)',   value: -fbmDemandPenalty,   bg: fbmDemandPenalty > 0 ? COLORS.OOS : '#F2F2F2',      fmt: '$#,##0', bold: true },
      { label: 'FBM Fulfillment Cost',       value: -totalFbmCost,       bg: totalFbmCost > 0 ? '#FFF2CC' : '#F2F2F2',           fmt: '$#,##0', bold: true },
      { label: 'TOTAL NOT-ON-FBA COST',      value: grandNet,            bg: '#D6E4F0',                                           fmt: '$#,##0', bold: true },
      { label: 'FBA In-Stock Rate',          value: fbaInStockRate + '%', bg: fbaInStockRate >= 95 ? '#C6EFCE' : (fbaInStockRate >= 80 ? '#FFF2CC' : COLORS.OOS), fmt: '', bold: true }
    ];
    var netMergeStartIdx = finValues.length;
    for (var ni = 0; ni < netDefs.length; ni++) {
      var nr = netDefs[ni];
      pushRow(
        [nr.label, nr.value, '', '', '', '', ''],
        [nr.bg, nr.bg, nr.bg, nr.bg, nr.bg, nr.bg, nr.bg],
        [nr.bold ? 'bold' : 'normal', ni === netDefs.length - 1 ? 'bold' : 'normal', 'normal', 'normal', 'normal', 'normal', 'normal'],
        ['left', 'right', 'center', 'center', 'center', 'center', 'center'],
        ['', nr.fmt, '', '', '', '', '']
      );
    }

    // ── Write entire financial block in one batch ──
    var finStartRow = row;
    var finNumRows  = finValues.length;
    var finRange    = sheet.getRange(finStartRow, 1, finNumRows, numCols);
    finRange.setValues(finValues);
    finRange.setFontSize(10);
    finRange.setBorder(true, true, true, true, false, false);
    finRange.setFontWeights(finWeights);
    finRange.setHorizontalAlignments(finAligns);
    finRange.setNumberFormats(finFormats);

    // Apply backgrounds (only non-null cells — setBackgrounds accepts null to skip)
    finRange.setBackgrounds(finBgs);

    // Set white font color on title rows
    var titleIdxes = [oosTitleIdx, fbmTitleIdx, netTitleIdx];
    for (var ti = 0; ti < titleIdxes.length; ti++) {
      var tRow = finStartRow + titleIdxes[ti];
      var titleRange = sheet.getRange(tRow, 1, 1, numCols);
      titleRange.merge().setFontSize(11).setFontColor(COLORS.HEADER_FG);
    }

    // Merge "no activity" rows if applicable
    if (typeof noOosMergeIdx !== 'undefined') {
      sheet.getRange(finStartRow + noOosMergeIdx, 1, 1, 3).merge()
           .setFontColor('#5f6368');
    }
    if (typeof noFbmMergeIdx !== 'undefined') {
      sheet.getRange(finStartRow + noFbmMergeIdx, 1, 1, 6).merge()
           .setFontColor('#5f6368');
    }

    // Merge net summary rows (cols 3-7)
    for (var nm = 0; nm < netDefs.length; nm++) {
      var nmRow = finStartRow + netMergeStartIdx + nm;
      if (numCols > 2) {
        sheet.getRange(nmRow, 3, 1, numCols - 2).merge();
      }
    }

    row = finStartRow + finNumRows;
  }

  // ── Bold the day numbers in each cell ──
  // (Can't selectively bold within a cell via script without RichTextValue;
  //  we handle this by making the first line stand out via the cell layout)

  // ── Freeze title ──
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}
