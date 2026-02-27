/**
 * GorillaDataTab.gs
 * ---------------------------------------------------------------------------
 * Builds a "Gorilla Data" tab that uses Gorilla ROI custom functions to pull
 * live inventory, sales, and pricing data from Amazon Seller Central.
 *
 * This tab acts as a data feed layer. The Settings tab references these cells
 * so that inventory levels, velocity, and pricing auto-populate from Amazon.
 * Users can override any Settings value by typing over the formula.
 *
 * Requires the Gorilla ROI Google Sheets add-on to be installed and connected
 * to a Seller Central account. Without it, formulas will show #NAME? errors
 * which are handled gracefully (treated as 0).
 * ---------------------------------------------------------------------------
 */

/**
 * Maps Settings tab field keys to Gorilla Data tab column letters.
 * Used by SettingsTab.gs to create formula references.
 */
var GORILLA_LINK_MAP = {
  'ONHAND_AVAILABLE':        'B',
  'INBOUND_WORKING':         'C',
  'INBOUND_SHIPPED':         'D',
  'INBOUND_RECEIVING':       'E',
  'RESERVED_CUSTOMER_ORDER': 'F',
  'DAILY_VELOCITY':          'H',
  'SELLING_PRICE':           'I'
};

/**
 * Builds (or rebuilds) the Gorilla Data tab with GORILLA_* formulas.
 * Reads config from the GLOBAL__GORILLA_* named ranges on the Settings tab.
 * Each SKU gets one row with formulas for inventory, sales, and pricing.
 *
 * Does nothing if the Gorilla Seller ID is not configured.
 */
function buildGorillaDataTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read Gorilla config from named ranges
  var sellerIdRange = ss.getRangeByName('GLOBAL__GORILLA_SELLER_ID');
  if (!sellerIdRange) return;
  var sellerId = sellerIdRange.getValue();
  if (!sellerId || sellerId === '') return;

  var skus  = getSkus();
  var sheet = ss.getSheetByName(GORILLA_DATA_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(GORILLA_DATA_TAB_NAME);
  }
  sheet.clear();
  sheet.clearFormats();

  // Ensure enough columns and rows
  var requiredCols = 9;
  var requiredRows = skus.length + 1;
  var currentCols  = sheet.getMaxColumns();
  var currentRows  = sheet.getMaxRows();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }

  // ── Headers ──
  var headers = [
    'SKU',
    'Available',
    'Inbound Working',
    'Inbound Shipped',
    'Inbound Receiving',
    'Reserved',
    'Sales (lookback)',
    'Daily Velocity',
    'Selling Price'
  ];
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setFontWeight('bold')
       .setFontSize(9)
       .setBackground('#E2EFDA')
       .setBorder(true, true, true, true, false, false)
       .setHorizontalAlignment('center');

  // Named range references used in formulas
  var sellerRef = 'GLOBAL__GORILLA_SELLER_ID';
  var mktRef    = 'GLOBAL__GORILLA_MARKETPLACE';
  var lookRef   = 'GLOBAL__GORILLA_LOOKBACK_DAYS';

  // ── Data rows — one per SKU ──
  for (var i = 0; i < skus.length; i++) {
    var row   = i + 2;
    var skuId = skus[i].id;

    // Col A: SKU ID
    sheet.getRange(row, 1).setValue(skuId);

    // Col B: Available (fulfillable)
    sheet.getRange(row, 2).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "fulfillable"), 0)'
    );

    // Col C: Inbound Working
    sheet.getRange(row, 3).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "inbound_working"), 0)'
    );

    // Col D: Inbound Shipped
    sheet.getRange(row, 4).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "inbound_shipped"), 0)'
    );

    // Col E: Inbound Receiving
    sheet.getRange(row, 5).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "inbound_receiving"), 0)'
    );

    // Col F: Reserved
    sheet.getRange(row, 6).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "reserved"), 0)'
    );

    // Col G: Sales Count (last N days)
    sheet.getRange(row, 7).setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', A' + row +
      ', "Shipped", "NO", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );

    // Col H: Daily Velocity (derived from sales / lookback days)
    sheet.getRange(row, 8).setFormula(
      '=IFERROR(G' + row + '/' + lookRef + ', 0)'
    );

    // Col I: Selling Price
    sheet.getRange(row, 9).setFormula(
      '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', A' + row + ', ' + mktRef + '), 0)'
    );
  }

  // ── Formatting ──
  if (skus.length > 0) {
    sheet.getRange(2, 1, skus.length, 1).setFontWeight('bold');
    sheet.getRange(2, 2, skus.length, 6).setNumberFormat('#,##0');
    sheet.getRange(2, 8, skus.length, 1).setNumberFormat('#,##0.0');
    sheet.getRange(2, 9, skus.length, 1).setNumberFormat('$#,##0.00');
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 9; c++) {
    sheet.setColumnWidth(c, 130);
  }

  // ── Info note ──
  sheet.getRange(1, 1).setNote(
    'This tab auto-populates from Gorilla ROI.\n' +
    'Requires the Gorilla ROI add-on connected to Seller Central.\n' +
    'Do not edit these cells — edit values on the Settings tab instead.'
  );

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}
