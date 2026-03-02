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
 * 3. INVENTORY GAP FINANCIAL IMPACT — Single compact table showing per-SKU
 *    totals: OOS Days, Lost Revenue, FBM Saved, FBM Cost, Net Impact.
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
  // Split-day note (after the 6 legend swatches)
  sheet.getRange(legendRow, 13, 1, 6).merge()
       .setValue('F→M = split day (FBA ran out, FBM picked up overflow)')
       .setFontSize(8)
       .setFontColor('#666666')
       .setHorizontalAlignment('left');

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
  // SECTION 3: INVENTORY GAP FINANCIAL IMPACT (single compact table)
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

    // Column layout: SKU(5) + OOS Days(2) + Lost Rev(3) + FBM Saved(3) + FBM Cost(3) + Net Impact(3) = 19
    var finTotalCols = 19;

    // Title
    sheet.getRange(finRow, 1, 1, finTotalCols).merge()
         .setValue('Inventory Gap Financial Impact')
         .setFontSize(12)
         .setFontWeight('bold')
         .setBackground(COLORS.HEADER)
         .setFontColor(COLORS.HEADER_FG);
    finRow += 1;

    // Headers
    var finHeaders = ['SKU', 'OOS Days', 'Lost Revenue', 'FBM Saved', 'FBM Cost', 'Net Impact'];
    var finColSpans = [5, 2, 3, 3, 3, 3];
    var finHdrBgs   = ['#E2EFDA', '#E2EFDA', COLORS.OOS, '#C6EFCE', '#FFF2CC', '#D6E4F0'];

    var hCol = 1;
    for (var fh = 0; fh < finHeaders.length; fh++) {
      var hSpan = finColSpans[fh];
      sheet.getRange(finRow, hCol, 1, hSpan).merge()
           .setValue(finHeaders[fh])
           .setFontWeight('bold')
           .setFontSize(9)
           .setBackground(finHdrBgs[fh])
           .setBorder(true, true, true, true, false, false)
           .setHorizontalAlignment('center')
           .setWrap(true);
      hCol += hSpan;
    }
    finRow += 1;

    // Grand totals accumulators
    var grandOosDays = 0, grandLostRev = 0, grandFbmSaved = 0, grandFbmCost = 0;

    // ── Compute all SKU financial data in memory first ──
    var finRowData = []; // array of { values: [...], bgs: [...], bold: bool }
    for (var si = 0; si < numSkus; si++) {
      var skuResult = allResults[si];
      var skuCfg = skuResult.cfg;
      var price = (skuCfg && skuCfg.sellingPrice) ? skuCfg.sellingPrice : 0;
      var snaps = skuResult.snapshots;

      var pastDays = (skuCfg && skuCfg.pastOosDays) ? skuCfg.pastOosDays : 0;
      var pastEffVel = (skuCfg && skuCfg.dailyVelocity) ? skuCfg.dailyVelocity : 0;
      var pastRev = pastDays * pastEffVel * price;

      var projOosDays = 0, projLostRev = 0, projFbmSaved = 0, projFbmCost = 0;
      var fbmPrice = (skuCfg && skuCfg.fbmSellingPrice > 0) ? skuCfg.fbmSellingPrice : price;

      for (var di = 0; di < snaps.length; di++) {
        var snap = snaps[di];
        if (snap.unfulfilledUnits > 0 && price > 0) {
          projOosDays += 1;
          var dayPrice = (snap.effectivePrice > 0) ? snap.effectivePrice : price;
          projLostRev += snap.unfulfilledUnits * dayPrice;
        }
        if (snap.soldFromFbm > 0) {
          projFbmSaved += snap.soldFromFbm * fbmPrice;
          projFbmCost += snap.soldFromFbm * (snap.fbmCostPerUnit || 0);
        }
      }

      var skuOosDays = pastDays + projOosDays;
      var skuLostRev = pastRev + projLostRev;
      var skuNetImpact = -(skuLostRev + projFbmCost);

      grandOosDays += skuOosDays;
      grandLostRev += skuLostRev;
      grandFbmSaved += projFbmSaved;
      grandFbmCost += projFbmCost;

      finRowData.push({
        values: [skuResult.skuDef.id, skuOosDays, skuLostRev, projFbmSaved, projFbmCost, skuNetImpact],
        bgs: [
          null,
          skuOosDays > 0 ? COLORS.OOS : null,
          skuLostRev > 0 ? COLORS.OOS : null,
          projFbmSaved > 0 ? '#C6EFCE' : null,
          projFbmCost > 0 ? '#FFF2CC' : null,
          skuNetImpact < 0 ? '#D6E4F0' : '#C6EFCE'
        ],
        bold: false
      });
    }

    // Totals row
    var grandNet = -(grandLostRev + grandFbmCost);
    finRowData.push({
      values: ['ALL SKUs TOTAL', grandOosDays, grandLostRev, grandFbmSaved, grandFbmCost, grandNet],
      bgs: [
        '#F2F2F2',
        grandOosDays > 0 ? COLORS.OOS : '#F2F2F2',
        grandLostRev > 0 ? COLORS.OOS : '#F2F2F2',
        grandFbmSaved > 0 ? '#C6EFCE' : '#F2F2F2',
        grandFbmCost > 0 ? '#FFF2CC' : '#F2F2F2',
        grandNet < 0 ? '#D6E4F0' : '#C6EFCE'
      ],
      bold: true
    });

    // ── Write data + totals: merge fields then batch format ──
    var dataStartRow = finRow;
    for (var rd = 0; rd < finRowData.length; rd++) {
      var rowInfo = finRowData[rd];
      var dCol = 1;
      for (var fd = 0; fd < finHeaders.length; fd++) {
        var span = finColSpans[fd];
        var mergedRange = sheet.getRange(finRow, dCol, 1, span).merge();
        mergedRange.setValue(rowInfo.values[fd]);
        if (rowInfo.bgs[fd]) mergedRange.setBackground(rowInfo.bgs[fd]);
        dCol += span;
      }
      finRow += 1;
    }

    // Batch formatting on entire data block (all SKU rows + totals)
    var dataRows = finRowData.length;
    var dataRange = sheet.getRange(dataStartRow, 1, dataRows, finTotalCols);
    dataRange.setFontSize(9);
    dataRange.setBorder(true, true, true, true, false, false);

    // Build alignment + format + weight arrays per-cell
    var allAligns  = [];
    var allFormats = [];
    var allWeights = [];
    var fieldAligns  = ['left', 'center', 'right', 'right', 'right', 'right'];
    var fieldFormats = ['', '', '$#,##0', '$#,##0', '$#,##0', '$#,##0;-$#,##0;$0'];

    for (var ar = 0; ar < dataRows; ar++) {
      var rowAligns  = [];
      var rowFormats = [];
      var rowWeights = [];
      var isBold = finRowData[ar].bold;
      for (var af = 0; af < finHeaders.length; af++) {
        for (var ac = 0; ac < finColSpans[af]; ac++) {
          rowAligns.push(fieldAligns[af]);
          rowFormats.push(fieldFormats[af]);
          rowWeights.push(isBold || af === finHeaders.length - 1 ? 'bold' : 'normal');
        }
      }
      allAligns.push(rowAligns);
      allFormats.push(rowFormats);
      allWeights.push(rowWeights);
    }
    dataRange.setHorizontalAlignments(allAligns);
    dataRange.setNumberFormats(allFormats);
    dataRange.setFontWeights(allWeights);

    finEndRow = finRow - 1;
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

      // Abbreviate split labels for 48px calendar cells (FBA→FBM → F→M)
      if (label.indexOf('→') > -1) {
        var ABBREV = { 'FBA': 'F', 'FBM': 'M', 'DTC': 'D', 'OOS': 'O' };
        var parts = label.split('→');
        var compact = [];
        for (var ci = 0; ci < parts.length; ci++) {
          compact.push(ABBREV[parts[ci]] || parts[ci]);
        }
        label = compact.join('→');
      }

      // Check if a shipment arrived today — show PROCESSING label
      var isArrival = false;
      for (var ei = 0; ei < snap.events.length; ei++) {
        if (snap.events[ei].indexOf('arrived') > -1) {
          isArrival = true;
          break;
        }
      }

      if (isArrival && snap.channel.indexOf('FBA') !== 0) {
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
