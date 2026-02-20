/**
 * SummaryTab.gs
 * ---------------------------------------------------------------------------
 * Builds the Summary tab with four sections:
 *
 * 1. MILESTONE TABLE — Per-SKU row showing key dates:
 *    Last FBA Day, First FBM Day, SPD Arrival, LTL Arrival,
 *    First Day Back on FBA, OOS Gaps
 *    Uses merged cells so each milestone field spans multiple calendar-width
 *    columns, avoiding the column-width conflict between the milestone table
 *    and the narrow calendar grid below.
 *
 * 2. COLOR LEGEND — Channel color references.
 *
 * 3. OOS FINANCIAL IMPACT TABLE — Monthly breakdown of lost Revenue and DPP
 *    per SKU, with a totals row. Only shown if any SKU has financial data.
 *
 * 4. CALENDAR GRID — Rows = SKUs, Columns = each calendar day.
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

  // ── Set column widths upfront (calendar grid needs 48px day columns) ──
  // Doing this early avoids 71 individual API calls later in the function
  // which can contribute to script timeout on the calendar grid section.
  sheet.setColumnWidth(1, 220);
  for (var cw = 2; cw <= numDays + 1; cw++) {
    sheet.setColumnWidth(cw, 48);
  }

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
  // SECTION 3: OOS FINANCIAL IMPACT TABLE
  // ══════════════════════════════════════════════════════════════════════════

  // Check if any SKU has financial data
  var anyFinancials = false;
  for (var fc = 0; fc < numSkus; fc++) {
    if (allResults[fc].cfg && allResults[fc].cfg.sellingPrice > 0) {
      anyFinancials = true;
      break;
    }
  }

  var finEndRow = legendRow; // track where financials end for calendar placement

  if (anyFinancials) {
    var finRow = legendRow + 2;

    // Define months for financial breakdown
    var finMonths = [
      { year: 2026, month: 1, name: 'Feb' },
      { year: 2026, month: 2, name: 'Mar' },
      { year: 2026, month: 3, name: 'Apr' }
    ];

    // Title
    sheet.getRange(finRow, 1, 1, msTotalCols).merge()
         .setValue('OOS Financial Impact')
         .setFontSize(12)
         .setFontWeight('bold')
         .setBackground(COLORS.HEADER)
         .setFontColor(COLORS.HEADER_FG);
    finRow += 1;

    // Headers: SKU | Feb OOS Days | Feb Lost Rev | Feb Lost DPP | Mar... | Apr... | TOTAL ...
    var finHeaders = ['SKU'];
    for (var fmi = 0; fmi < finMonths.length; fmi++) {
      finHeaders.push(finMonths[fmi].name + ' OOS Days');
      finHeaders.push(finMonths[fmi].name + ' Lost Rev');
      finHeaders.push(finMonths[fmi].name + ' Lost DPP');
    }
    finHeaders.push('Total OOS Days');
    finHeaders.push('Total Lost Rev');
    finHeaders.push('Total Lost DPP');

    // Use merged cells: SKU gets 5 cols, each data field gets 2 cols
    var finColSpans = [5]; // SKU
    for (var fhi = 1; fhi < finHeaders.length; fhi++) {
      finColSpans.push(2);
    }

    // Write headers
    var hCol = 1;
    for (var fh = 0; fh < finHeaders.length; fh++) {
      var hSpan = finColSpans[fh];
      sheet.getRange(finRow, hCol, 1, hSpan).merge()
           .setValue(finHeaders[fh])
           .setFontWeight('bold')
           .setFontSize(8)
           .setBackground('#E2EFDA')
           .setBorder(true, true, true, true, false, false)
           .setHorizontalAlignment('center')
           .setWrap(true);
      hCol += hSpan;
    }
    finRow += 1;

    // Grand totals for "Total" row
    var grandOosDays = 0, grandLostRev = 0, grandLostDpp = 0;
    var grandMonthly = {};
    for (var gm = 0; gm < finMonths.length; gm++) {
      grandMonthly[finMonths[gm].name] = { oosDays: 0, lostRevenue: 0, lostDpp: 0 };
    }

    // Data rows — one per SKU
    for (var si = 0; si < numSkus; si++) {
      var skuResult = allResults[si];
      var skuCfg = skuResult.cfg;
      var price = (skuCfg && skuCfg.sellingPrice) ? skuCfg.sellingPrice : 0;
      var dpp   = (skuCfg && skuCfg.dppMargin) ? skuCfg.dppMargin : 0;
      var snaps = skuResult.snapshots;

      // Calculate per-month OOS impact
      var skuMonthly = {};
      for (var sm = 0; sm < finMonths.length; sm++) {
        skuMonthly[finMonths[sm].name] = { oosDays: 0, lostRevenue: 0, lostDpp: 0 };
      }

      var skuTotalOos = 0, skuTotalRev = 0, skuTotalDpp = 0;

      for (var di = 0; di < snaps.length; di++) {
        var snap = snaps[di];
        if (snap.channel === 'OOS' && price > 0) {
          var snapMonth = snap.date.getMonth(); // 0-indexed
          var mKey = null;
          for (var mk = 0; mk < finMonths.length; mk++) {
            if (finMonths[mk].month === snapMonth) {
              mKey = finMonths[mk].name;
              break;
            }
          }
          if (mKey) {
            var dayRev = snap.effectiveVelocity * price;
            var dayDpp = dayRev * (dpp / 100);
            skuMonthly[mKey].oosDays += 1;
            skuMonthly[mKey].lostRevenue += dayRev;
            skuMonthly[mKey].lostDpp += dayDpp;
            skuTotalOos += 1;
            skuTotalRev += dayRev;
            skuTotalDpp += dayDpp;
          }
        }
      }

      // Write row
      var dCol = 1;
      // SKU name
      sheet.getRange(finRow, dCol, 1, finColSpans[0]).merge()
           .setValue(skuResult.skuDef.id)
           .setFontSize(8).setHorizontalAlignment('left')
           .setBorder(true, true, true, true, false, false);
      dCol += finColSpans[0];

      // Monthly values
      var colIdx = 1;
      for (var wm = 0; wm < finMonths.length; wm++) {
        var mData = skuMonthly[finMonths[wm].name];

        // OOS Days
        sheet.getRange(finRow, dCol, 1, 2).merge()
             .setValue(mData.oosDays)
             .setFontSize(8).setHorizontalAlignment('center')
             .setBorder(true, true, true, true, false, false);
        dCol += 2;

        // Lost Revenue
        var revRange = sheet.getRange(finRow, dCol, 1, 2).merge()
             .setValue(mData.lostRevenue)
             .setNumberFormat('$#,##0')
             .setFontSize(8).setHorizontalAlignment('right')
             .setBorder(true, true, true, true, false, false);
        if (mData.lostRevenue > 0) revRange.setBackground(COLORS.OOS);
        dCol += 2;

        // Lost DPP
        var dppRange = sheet.getRange(finRow, dCol, 1, 2).merge()
             .setValue(mData.lostDpp)
             .setNumberFormat('$#,##0')
             .setFontSize(8).setHorizontalAlignment('right')
             .setBorder(true, true, true, true, false, false);
        if (mData.lostDpp > 0) dppRange.setBackground(COLORS.OOS);
        dCol += 2;

        // Accumulate grand totals
        grandMonthly[finMonths[wm].name].oosDays += mData.oosDays;
        grandMonthly[finMonths[wm].name].lostRevenue += mData.lostRevenue;
        grandMonthly[finMonths[wm].name].lostDpp += mData.lostDpp;
      }

      // Total columns
      sheet.getRange(finRow, dCol, 1, 2).merge()
           .setValue(skuTotalOos)
           .setFontSize(8).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
      dCol += 2;

      var totalRevR = sheet.getRange(finRow, dCol, 1, 2).merge()
           .setValue(skuTotalRev)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (skuTotalRev > 0) totalRevR.setBackground(COLORS.OOS);
      dCol += 2;

      var totalDppR = sheet.getRange(finRow, dCol, 1, 2).merge()
           .setValue(skuTotalDpp)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (skuTotalDpp > 0) totalDppR.setBackground(COLORS.OOS);

      grandOosDays += skuTotalOos;
      grandLostRev += skuTotalRev;
      grandLostDpp += skuTotalDpp;

      finRow += 1;
    }

    // ── Totals row ──
    var tCol = 1;
    sheet.getRange(finRow, tCol, 1, finColSpans[0]).merge()
         .setValue('ALL SKUs TOTAL')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('left')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#F2F2F2');
    tCol += finColSpans[0];

    for (var tm = 0; tm < finMonths.length; tm++) {
      var gm2 = grandMonthly[finMonths[tm].name];

      sheet.getRange(finRow, tCol, 1, 2).merge()
           .setValue(gm2.oosDays)
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false)
           .setBackground('#F2F2F2');
      tCol += 2;

      var gmRevR = sheet.getRange(finRow, tCol, 1, 2).merge()
           .setValue(gm2.lostRevenue)
           .setNumberFormat('$#,##0')
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (gm2.lostRevenue > 0) gmRevR.setBackground(COLORS.OOS);
      else gmRevR.setBackground('#F2F2F2');
      tCol += 2;

      var gmDppR = sheet.getRange(finRow, tCol, 1, 2).merge()
           .setValue(gm2.lostDpp)
           .setNumberFormat('$#,##0')
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (gm2.lostDpp > 0) gmDppR.setBackground(COLORS.OOS);
      else gmDppR.setBackground('#F2F2F2');
      tCol += 2;
    }

    // Grand totals
    sheet.getRange(finRow, tCol, 1, 2).merge()
         .setValue(grandOosDays)
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#F2F2F2');
    tCol += 2;

    var grRevR = sheet.getRange(finRow, tCol, 1, 2).merge()
         .setValue(grandLostRev)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (grandLostRev > 0) grRevR.setBackground(COLORS.OOS);
    else grRevR.setBackground('#F2F2F2');
    tCol += 2;

    var grDppR = sheet.getRange(finRow, tCol, 1, 2).merge()
         .setValue(grandLostDpp)
         .setNumberFormat('$#,##0')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (grandLostDpp > 0) grDppR.setBackground(COLORS.OOS);
    else grDppR.setBackground('#F2F2F2');

    finEndRow = finRow;
  }

  // Flush to ensure milestone + financial sections are committed
  // before the calendar grid (protects against script timeout)
  SpreadsheetApp.flush();

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 4: CALENDAR GRID
  // ══════════════════════════════════════════════════════════════════════════

  var calStartRow = finEndRow + 2;

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

  // Column widths already set at the top of the function
  calStartRow += 1;

  // ── Data rows: one per SKU ──
  var calData    = [];
  var calBgAll   = [];
  var calFgAll   = [];

  for (var si2 = 0; si2 < numSkus; si2++) {
    var result    = allResults[si2];
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
  // Freeze just the title row so the rest of the sheet scrolls freely
  sheet.setFrozenRows(1);

  SpreadsheetApp.flush();
}
