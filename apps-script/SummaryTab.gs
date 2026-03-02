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

  // Ensure the sheet has enough columns for the calendar grid (1 SKU label + numDays)
  var requiredCols = numDays + 1;
  var currentCols = sheet.getMaxColumns();
  if (requiredCols > currentCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  } else if (requiredCols < currentCols) {
    sheet.deleteColumns(requiredCols + 1, currentCols - requiredCols);
  }

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
  // SECTION 2B: SHIPMENT DATA WARNINGS
  // ══════════════════════════════════════════════════════════════════════════

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var warnings = []; // array of { skuId, type, message }

  for (var ws = 0; ws < numSkus; ws++) {
    var wCfg = allResults[ws].cfg;
    var wSkuId = allResults[ws].skuDef.id;

    // Check each shipment type: SPD, LTL, Ad-hoc
    var shipments = [
      { name: 'SPD',    units: wCfg.spdUnits,   sendDate: wCfg.spdSendDate,   transitDays: wCfg.spdTransitDays },
      { name: 'LTL',    units: wCfg.ltlUnits,   sendDate: wCfg.ltlSendDate,   transitDays: wCfg.ltlTransitDays },
      { name: 'Ad-hoc', units: wCfg.adhocUnits,  sendDate: wCfg.adhocSendDate,  transitDays: wCfg.adhocTransitDays }
    ];

    for (var wsi = 0; wsi < shipments.length; wsi++) {
      var ship = shipments[wsi];
      if (!ship.sendDate || !ship.units || ship.units <= 0) continue;

      var estArrival = addDays(ship.sendDate, ship.transitDays || 0);

      if (estArrival < today) {
        // Arrival date has passed — likely already received
        warnings.push({
          skuId: wSkuId,
          type: 'STALE',
          message: ship.name + ' (' + ship.units + ' units) — expected arrival ' +
            fmtDate(estArrival) + ' has passed. If received, these units now appear ' +
            'in Inbound Receiving or Available. Remove this entry to avoid double-counting.'
        });
      } else if (ship.sendDate < today) {
        // Send date passed but arrival still in future — just a heads-up
        warnings.push({
          skuId: wSkuId,
          type: 'CHECK',
          message: ship.name + ' (' + ship.units + ' units) — send date ' +
            fmtDate(ship.sendDate) + ' has passed. If shipped, these units may now ' +
            'show as Inbound Shipped in Gorilla. Verify transit time is still accurate.'
        });
      }
    }
  }

  var warnEndRow = legendRow;
  if (warnings.length > 0) {
    var warnRow = legendRow + 2;

    // Warning title — span across enough cols to be readable
    var warnCols = 26; // SKU(5) + message(21)
    sheet.getRange(warnRow, 1, 1, warnCols).merge()
         .setValue('Shipment Data Warnings')
         .setFontSize(11)
         .setFontWeight('bold')
         .setBackground('#FFF3CD')
         .setFontColor('#856404');
    warnRow += 1;

    // Headers
    var whCol = 1;
    sheet.getRange(warnRow, whCol, 1, 5).merge()
         .setValue('SKU')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBackground('#FFF3CD').setBorder(true, true, true, true, false, false);
    whCol += 5;
    sheet.getRange(warnRow, whCol, 1, 3).merge()
         .setValue('Status')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBackground('#FFF3CD').setBorder(true, true, true, true, false, false);
    whCol += 3;
    sheet.getRange(warnRow, whCol, 1, 18).merge()
         .setValue('Action Needed')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBackground('#FFF3CD').setBorder(true, true, true, true, false, false);
    warnRow += 1;

    for (var wi = 0; wi < warnings.length; wi++) {
      var warn = warnings[wi];
      var wdCol = 1;

      // SKU
      sheet.getRange(warnRow, wdCol, 1, 5).merge()
           .setValue(warn.skuId)
           .setFontSize(8).setHorizontalAlignment('left')
           .setBorder(true, true, true, true, false, false);
      wdCol += 5;

      // Status badge
      var statusLabel = warn.type === 'STALE' ? 'STALE' : 'VERIFY';
      var statusBg    = warn.type === 'STALE' ? COLORS.OOS : '#FFF3CD';
      sheet.getRange(warnRow, wdCol, 1, 3).merge()
           .setValue(statusLabel)
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
           .setBackground(statusBg)
           .setBorder(true, true, true, true, false, false);
      wdCol += 3;

      // Message
      sheet.getRange(warnRow, wdCol, 1, 18).merge()
           .setValue(warn.message)
           .setFontSize(8).setHorizontalAlignment('left')
           .setWrap(true)
           .setBorder(true, true, true, true, false, false);

      warnRow += 1;
    }

    warnEndRow = warnRow;
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

  var finEndRow = warnEndRow; // track where financials end for calendar placement

  if (anyFinancials) {
    var finRow = warnEndRow + 2;

    // Define months for financial breakdown (dynamic from START_DATE → END_DATE)
    var finMonths = forecastMonthsShort();
    // Total columns: SKU(5) + per-month(6 each) + totals(6)
    var finTotalCols = 5 + finMonths.length * 6 + 6;

    // Title
    sheet.getRange(finRow, 1, 1, finTotalCols).merge()
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
          var snapYear  = snap.date.getFullYear();
          var mKey = null;
          for (var mk = 0; mk < finMonths.length; mk++) {
            if (finMonths[mk].month === snapMonth && finMonths[mk].year === snapYear) {
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

    // ════════════════════════════════════════════════════════════════════════
    // SECTION 3B: TOTAL IMPACT SUMMARY (Already Lost + Projected + Total)
    // ════════════════════════════════════════════════════════════════════════

    var impRow = finRow + 2;

    // Title
    // Total Impact cols: SKU(5) + 3 groups × 3 fields × 2 cols = 5+18 = 23
    var impTotalCols = 5 + 9 * 2;
    sheet.getRange(impRow, 1, 1, impTotalCols).merge()
         .setValue('Total Impact Summary')
         .setFontSize(12)
         .setFontWeight('bold')
         .setBackground(COLORS.HEADER)
         .setFontColor(COLORS.HEADER_FG);
    impRow += 1;

    // Headers: SKU | Already Lost (Days, Rev, DPP) | Projected (Days, Rev, DPP) | Total (Days, Rev, DPP)
    var impHeaders = [
      'SKU',
      'Past Days', 'Past Rev', 'Past DPP',
      'Projected Days', 'Projected Rev', 'Projected DPP',
      'Total Days', 'Total Rev', 'Total DPP'
    ];
    var impColSpans = [5, 2, 2, 2, 2, 2, 2, 2, 2, 2]; // SKU=5, rest=2 each

    var ihCol = 1;
    for (var ih = 0; ih < impHeaders.length; ih++) {
      var ihSpan = impColSpans[ih];
      var ihRange = sheet.getRange(impRow, ihCol, 1, ihSpan).merge()
           .setValue(impHeaders[ih])
           .setFontWeight('bold')
           .setFontSize(8)
           .setBorder(true, true, true, true, false, false)
           .setHorizontalAlignment('center')
           .setWrap(true);
      // Color-code the header groups
      if (ih >= 1 && ih <= 3) ihRange.setBackground('#FFF2CC'); // Past = warm yellow
      else if (ih >= 4 && ih <= 6) ihRange.setBackground('#E2EFDA'); // Projected = green
      else if (ih >= 7) ihRange.setBackground('#D6E4F0'); // Total = blue
      else ihRange.setBackground('#E2EFDA');
      ihCol += ihSpan;
    }
    impRow += 1;

    // Grand totals for impact summary
    var gPastDays = 0, gPastRev = 0, gPastDpp = 0;
    var gProjDays = 0, gProjRev = 0, gProjDpp = 0;

    // Data rows — one per SKU
    for (var ip = 0; ip < numSkus; ip++) {
      var ipResult = allResults[ip];
      var ipCfg = ipResult.cfg;
      var ipPrice = (ipCfg && ipCfg.sellingPrice) ? ipCfg.sellingPrice : 0;
      var ipDpp   = (ipCfg && ipCfg.dppMargin) ? ipCfg.dppMargin : 0;

      // Past losses: from user-entered pastOosDays
      var pastDays = (ipCfg && ipCfg.pastOosDays) ? ipCfg.pastOosDays : 0;
      var pastEffVel = (ipCfg && ipCfg.dailyVelocity) ? ipCfg.dailyVelocity : 0;
      var pastRev = pastDays * pastEffVel * ipPrice;
      var pastDppVal = pastRev * (ipDpp / 100);

      // Projected losses: sum from monthly financial data (already computed above)
      var projDays = 0, projRev = 0, projDppVal2 = 0;
      var ipSnaps = ipResult.snapshots;
      for (var ipd = 0; ipd < ipSnaps.length; ipd++) {
        if (ipSnaps[ipd].channel === 'OOS' && ipPrice > 0) {
          projDays += 1;
          var ipDayRev = ipSnaps[ipd].effectiveVelocity * ipPrice;
          projRev += ipDayRev;
          projDppVal2 += ipDayRev * (ipDpp / 100);
        }
      }

      var totalDays = pastDays + projDays;
      var totalRev  = pastRev + projRev;
      var totalDpp2 = pastDppVal + projDppVal2;

      // Write row
      var ipCol = 1;

      // SKU
      sheet.getRange(impRow, ipCol, 1, impColSpans[0]).merge()
           .setValue(ipResult.skuDef.id)
           .setFontSize(8).setHorizontalAlignment('left')
           .setBorder(true, true, true, true, false, false);
      ipCol += impColSpans[0];

      // Past Days
      sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(pastDays)
           .setFontSize(8).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false)
           .setBackground(pastDays > 0 ? '#FFF2CC' : null);
      ipCol += 2;

      // Past Rev
      var pRevR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(pastRev)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (pastRev > 0) pRevR.setBackground('#FFF2CC');
      ipCol += 2;

      // Past DPP
      var pDppR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(pastDppVal)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (pastDppVal > 0) pDppR.setBackground('#FFF2CC');
      ipCol += 2;

      // Projected Days
      sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(projDays)
           .setFontSize(8).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false);
      ipCol += 2;

      // Projected Rev
      var prRevR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(projRev)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (projRev > 0) prRevR.setBackground(COLORS.OOS);
      ipCol += 2;

      // Projected DPP
      var prDppR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(projDppVal2)
           .setNumberFormat('$#,##0')
           .setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (projDppVal2 > 0) prDppR.setBackground(COLORS.OOS);
      ipCol += 2;

      // Total Days
      sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(totalDays)
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
           .setBorder(true, true, true, true, false, false)
           .setBackground('#D6E4F0');
      ipCol += 2;

      // Total Rev
      var tRevR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(totalRev)
           .setNumberFormat('$#,##0')
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (totalRev > 0) tRevR.setBackground(COLORS.OOS);
      else tRevR.setBackground('#D6E4F0');
      ipCol += 2;

      // Total DPP
      var tDppR = sheet.getRange(impRow, ipCol, 1, 2).merge()
           .setValue(totalDpp2)
           .setNumberFormat('$#,##0')
           .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
           .setBorder(true, true, true, true, false, false);
      if (totalDpp2 > 0) tDppR.setBackground(COLORS.OOS);
      else tDppR.setBackground('#D6E4F0');

      // Accumulate grand totals
      gPastDays += pastDays;  gPastRev += pastRev;  gPastDpp += pastDppVal;
      gProjDays += projDays;  gProjRev += projRev;  gProjDpp += projDppVal2;

      impRow += 1;
    }

    // ── Impact totals row ──
    var itCol = 1;
    sheet.getRange(impRow, itCol, 1, impColSpans[0]).merge()
         .setValue('ALL SKUs TOTAL')
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('left')
         .setBorder(true, true, true, true, false, false)
         .setBackground('#F2F2F2');
    itCol += impColSpans[0];

    // Past totals
    sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gPastDays)
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false).setBackground('#F2F2F2');
    itCol += 2;
    var gPastRevR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gPastRev)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gPastRev > 0) gPastRevR.setBackground('#FFF2CC'); else gPastRevR.setBackground('#F2F2F2');
    itCol += 2;
    var gPastDppR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gPastDpp)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gPastDpp > 0) gPastDppR.setBackground('#FFF2CC'); else gPastDppR.setBackground('#F2F2F2');
    itCol += 2;

    // Projected totals
    sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gProjDays)
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false).setBackground('#F2F2F2');
    itCol += 2;
    var gProjRevR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gProjRev)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gProjRev > 0) gProjRevR.setBackground(COLORS.OOS); else gProjRevR.setBackground('#F2F2F2');
    itCol += 2;
    var gProjDppR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gProjDpp)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gProjDpp > 0) gProjDppR.setBackground(COLORS.OOS); else gProjDppR.setBackground('#F2F2F2');
    itCol += 2;

    // Grand totals
    var gTotalDays = gPastDays + gProjDays;
    var gTotalRev  = gPastRev + gProjRev;
    var gTotalDpp  = gPastDpp + gProjDpp;

    sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gTotalDays)
         .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
         .setBorder(true, true, true, true, false, false).setBackground('#D6E4F0');
    itCol += 2;
    var gTotRevR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gTotalRev)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gTotalRev > 0) gTotRevR.setBackground(COLORS.OOS); else gTotRevR.setBackground('#D6E4F0');
    itCol += 2;
    var gTotDppR = sheet.getRange(impRow, itCol, 1, 2).merge().setValue(gTotalDpp)
         .setNumberFormat('$#,##0').setFontWeight('bold').setFontSize(8).setHorizontalAlignment('right')
         .setBorder(true, true, true, true, false, false);
    if (gTotalDpp > 0) gTotDppR.setBackground(COLORS.OOS); else gTotDppR.setBackground('#D6E4F0');

    finEndRow = impRow;
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
