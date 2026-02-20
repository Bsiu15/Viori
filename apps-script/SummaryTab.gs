/**
 * SummaryTab.gs
 * ---------------------------------------------------------------------------
 * Builds the Summary tab with two sections:
 *
 * 1. MILESTONE TABLE — Per-SKU row showing key dates:
 *    Last FBA Day, First FBM Day, SPD Arrival, LTL Arrival,
 *    First Day Back on FBA, OOS Gaps
 *
 * 2. CALENDAR GRID — Rows = SKUs, Columns = each calendar day.
 *    Each cell shows a short label (FBA, FBM, DTC, OOS) and is
 *    color-coded per the palette in Config.gs.
 *
 * Designed for clean, professional sharing with supply chain partners.
 * ---------------------------------------------------------------------------
 */

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

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 1: MILESTONE TABLE
  // ══════════════════════════════════════════════════════════════════════════

  var msRow = 1;

  // Title
  sheet.getRange(msRow, 1, 1, 8).merge()
       .setValue('Inventory Forecasting Calendar — Milestone Summary')
       .setFontSize(13)
       .setFontWeight('bold')
       .setBackground(COLORS.HEADER)
       .setFontColor(COLORS.HEADER_FG);
  msRow += 1;

  // Milestone headers
  var msHeaders = [
    'SKU',
    'Last FBA Day',
    'First FBM Day',
    'SPD Arrival',
    'LTL Arrival',
    'First Day Back on FBA',
    'OOS Gaps',
    'Total OOS Days'
  ];
  var msHeaderRange = sheet.getRange(msRow, 1, 1, msHeaders.length);
  msHeaderRange.setValues([msHeaders])
               .setFontWeight('bold')
               .setBackground('#E2EFDA')
               .setBorder(true, true, true, true, true, true)
               .setHorizontalAlignment('center')
               .setWrap(true);
  msRow += 1;

  // Milestone data rows
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

    var msRowData = [
      r.skuDef.id,
      ms.lastFbaDay   ? fmtDate(ms.lastFbaDay)     : 'N/A',
      ms.firstFbmDay  ? fmtDate(ms.firstFbmDay)    : 'N/A',
      ms.spdArrival   ? fmtDate(ms.spdArrival)     : 'N/A',
      ms.ltlArrival   ? fmtDate(ms.ltlArrival)     : 'N/A',
      ms.firstBackOnFba ? fmtDate(ms.firstBackOnFba) : 'N/A',
      oosGapStr,
      totalOosDays
    ];

    var dataRowRange = sheet.getRange(msRow, 1, 1, msHeaders.length);
    dataRowRange.setValues([msRowData])
                .setBorder(true, true, true, true, true, true, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);

    // Highlight OOS rows
    if (totalOosDays > 0) {
      sheet.getRange(msRow, 7, 1, 2).setBackground(COLORS.OOS);
    }

    msRow += 1;
  }

  // Column widths for milestone section
  sheet.setColumnWidth(1, 220); // SKU
  for (var mw = 2; mw <= 6; mw++) {
    sheet.setColumnWidth(mw, 140);
  }
  sheet.setColumnWidth(7, 280); // OOS Gaps
  sheet.setColumnWidth(8, 110); // Total OOS Days

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 2: COLOR LEGEND
  // ══════════════════════════════════════════════════════════════════════════

  var legendRow = msRow + 1;
  sheet.getRange(legendRow, 1, 1, 8).merge()
       .setValue('Color Legend')
       .setFontWeight('bold')
       .setFontSize(10)
       .setBackground('#F2F2F2');
  legendRow += 1;

  var legendItems = [
    { label: 'FBA',        color: COLORS.FBA },
    { label: 'FBM',        color: COLORS.FBM },
    { label: 'DTC',        color: COLORS.DTC },
    { label: 'OOS',        color: COLORS.OOS },
    { label: 'PROCESSING', color: COLORS.PROCESSING },
    { label: 'IN PROCESSING (gray)', color: COLORS.GRAY }
  ];

  for (var li = 0; li < legendItems.length; li++) {
    sheet.getRange(legendRow, 1 + li)
         .setValue(legendItems[li].label)
         .setBackground(legendItems[li].color)
         .setHorizontalAlignment('center')
         .setFontSize(9)
         .setBorder(true, true, true, true, false, false);
  }

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 3: CALENDAR GRID
  // ══════════════════════════════════════════════════════════════════════════

  var calStartRow = legendRow + 2;

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

  // Set date column widths (narrow)
  for (var dw = 2; dw <= numDays + 1; dw++) {
    sheet.setColumnWidth(dw, 48);
  }
  calStartRow += 1;

  // ── Data rows: one per SKU ──
  var calData    = [];
  var calBgAll   = [];
  var calFgAll   = [];

  for (var si = 0; si < numSkus; si++) {
    var result    = allResults[si];
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

      // Check if FBA processing has inventory (gray indicator)
      var hasProcessing = snap.fbaProcessingEnd > 0;

      if (isArrival && snap.channel !== 'FBA') {
        // Shipment arrived but not yet FBA available — show as arrival event
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
  sheet.setFrozenRows(calStartRow - 1); // Freeze everything above the data
  sheet.setFrozenColumns(1);

  SpreadsheetApp.flush();
}
