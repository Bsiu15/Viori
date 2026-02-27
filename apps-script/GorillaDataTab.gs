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
  // NOTE: Reserved (col F) is intentionally NOT linked to Settings.
  // Gorilla's "reserved" is the TOTAL of all reserved types (customer orders +
  // FC processing + FC transfer). Amazon's "fulfillable" already excludes reserved
  // units, so linking this would double-subtract from the forecast. The column
  // stays on the Gorilla Data tab as a reference number only.
  'ONHAND_FC_TRANSFER':      'G',
  'DAILY_VELOCITY':          'J',
  'SELLING_PRICE':           'K'
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
  var requiredCols = 11;
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
    'FC Transfer',
    'Unsellable',
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

    // Col G: FC Transfer
    sheet.getRange(row, 7).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "transfer"), 0)'
    );

    // Col H: Unsellable (total unfulfillable)
    sheet.getRange(row, 8).setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', A' + row + ', ' + mktRef + ', "unsellable"), 0)'
    );

    // Col I: Sales Count (last N days)
    sheet.getRange(row, 9).setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', A' + row +
      ', "Shipped", "NO", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );

    // Col J: Daily Velocity (derived from sales / lookback days)
    sheet.getRange(row, 10).setFormula(
      '=IFERROR(I' + row + '/' + lookRef + ', 0)'
    );

    // Col K: Selling Price
    sheet.getRange(row, 11).setFormula(
      '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', A' + row + ', ' + mktRef + '), 0)'
    );
  }

  // ── Formatting ──
  if (skus.length > 0) {
    sheet.getRange(2, 1, skus.length, 1).setFontWeight('bold');
    // Cols B-I: inventory counts + sales count (integers)
    sheet.getRange(2, 2, skus.length, 8).setNumberFormat('#,##0');
    // Col J: Daily Velocity (one decimal)
    sheet.getRange(2, 10, skus.length, 1).setNumberFormat('#,##0.0');
    // Col K: Selling Price (currency)
    sheet.getRange(2, 11, skus.length, 1).setNumberFormat('$#,##0.00');
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 11; c++) {
    sheet.setColumnWidth(c, 130);
  }

  // ── Header notes (ELI5 + Settings mapping) ──
  sheet.getRange(1, 1).setNote(
    'SKU = Your product ID in Amazon Seller Central.\n' +
    'This tab auto-populates from Gorilla ROI.\n' +
    'Requires the Gorilla ROI add-on connected to Seller Central.\n' +
    'Do not edit these cells — edit values on the Settings tab instead.'
  );
  sheet.getRange(1, 2).setNote(
    'Available = Units that customers can buy right now.\n' +
    'This is what Amazon considers "fulfillable" — sitting on the shelf, ready to ship.\n\n' +
    'Settings field: "On-hand: Available"'
  );
  sheet.getRange(1, 3).setNote(
    'Inbound Working = Units in a shipment you created but haven\'t sent yet.\n' +
    'You made the shipping plan in Seller Central, but the boxes haven\'t left your door.\n\n' +
    'Settings field: "Inbound: Working"'
  );
  sheet.getRange(1, 4).setNote(
    'Inbound Shipped = Units on the truck heading to Amazon\'s warehouse.\n' +
    'You shipped them out, but Amazon hasn\'t checked them in yet.\n\n' +
    'Settings field: "Inbound: Shipped"'
  );
  sheet.getRange(1, 5).setNote(
    'Inbound Receiving = Units that arrived at Amazon and are being checked in.\n' +
    'Amazon has the boxes but is still scanning and shelving them.\n\n' +
    'Settings field: "Inbound: Receiving"'
  );
  sheet.getRange(1, 6).setNote(
    'Reserved = Units set aside and temporarily unavailable for new orders.\n' +
    'Includes units in customer orders being packed, being moved between\n' +
    'warehouses, or being processed. They\'re yours, just busy.\n\n' +
    'REFERENCE ONLY — not linked to Settings.\n' +
    'Why: Gorilla gives the TOTAL reserved (customer orders + FC processing +\n' +
    'FC transfer combined). The "Available" column already excludes these units,\n' +
    'so linking this would double-subtract from your forecast.\n\n' +
    'Leave "Reserved: Customer order" at 0 in Settings when using Gorilla.\n' +
    'FC Transfer is handled separately (col G) and flows into the forecast\n' +
    'as units that become available after the check-in delay.'
  );
  sheet.getRange(1, 7).setNote(
    'FC Transfer = Units being moved from one Amazon warehouse to another.\n' +
    'Amazon shuffles your inventory around to be closer to buyers.\n' +
    'These units can\'t be sold while in transit between fulfillment centers.\n\n' +
    'Settings field: "On-hand: FC transfer"'
  );
  sheet.getRange(1, 8).setNote(
    'Unsellable = Units Amazon won\'t sell because they\'re damaged, defective, or expired.\n' +
    'Could be warehouse damage, customer returns in bad shape, carrier damage, etc.\n' +
    'You\'ll want to create a removal order or investigate.\n\n' +
    'Settings fields: This is the TOTAL of all 6 unfulfillable categories\n' +
    '(warehouse damaged, defective, expired, customer damaged, carrier damaged,\n' +
    'distributor damaged). Gorilla only gives the combined number — enter the\n' +
    'breakdown on the Settings tab manually if needed.'
  );
  sheet.getRange(1, 9).setNote(
    'Sales (lookback) = Total units sold in the last N days.\n' +
    'N is your "Lookback Days" setting (e.g. 30 days).\n' +
    'This is the raw number used to calculate your daily velocity.\n\n' +
    'Not directly linked to Settings — used to derive Daily Velocity.'
  );
  sheet.getRange(1, 10).setNote(
    'Daily Velocity = How many units you sell per day on average.\n' +
    'Calculated as: Sales (lookback) / Lookback Days.\n' +
    'This is the #1 driver of your reorder forecast — it tells the calendar\n' +
    'when you\'ll run out of stock.\n\n' +
    'Settings field: "Daily Velocity"'
  );
  sheet.getRange(1, 11).setNote(
    'Selling Price = Your current listing price on Amazon.\n' +
    'What the customer sees. Used in the forecast to estimate revenue impact\n' +
    'and prioritize which SKUs to reorder first.\n\n' +
    'Settings field: "Selling Price"'
  );

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}

/**
 * Lightweight linking: sets IFERROR formulas on existing Settings named ranges
 * to point at the Gorilla Data tab. Only touches the fields in GORILLA_LINK_MAP
 * (7 fields × N SKUs), so it runs in seconds instead of minutes.
 *
 * Skips any cell that already has a user-typed value (no formula) to preserve
 * manual overrides.
 */
function linkSettingsToGorilla() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var skus = getSkus();

  // Build set of fields that risk double-counting — maps key to warning text
  var warningKeys = {};
  for (var w = 0; w < SKU_INPUT_ROWS.length; w++) {
    if (SKU_INPUT_ROWS[w].gorillaWarning) {
      warningKeys[SKU_INPUT_ROWS[w].key] = SKU_INPUT_ROWS[w].gorillaWarning;
    }
  }

  for (var i = 0; i < skus.length; i++) {
    var prefix    = namedRangePrefix(skus[i].id);
    var gorillaRow = i + 2; // SKU index 0 → Gorilla Data row 2

    // Link Gorilla fields
    for (var key in GORILLA_LINK_MAP) {
      if (!GORILLA_LINK_MAP.hasOwnProperty(key)) continue;

      var rangeName = prefix + '__' + key;
      var cell = ss.getRangeByName(rangeName);
      if (!cell) continue;

      // If the cell already has a Gorilla formula, skip (already linked)
      var existingFormula = cell.getFormula();
      if (existingFormula && existingFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1) continue;

      // If the cell has a non-zero user-entered value (no formula), skip to preserve override
      if (!existingFormula) {
        var val = cell.getValue();
        if (val !== '' && val !== 0 && val !== null && val !== undefined) continue;
      }

      var gorillaCol = GORILLA_LINK_MAP[key];
      cell.setFormula("=IFERROR('" + GORILLA_DATA_TAB_NAME + "'!" + gorillaCol + gorillaRow + ", 0)");
      cell.setBackground('#E8F0FE');
      cell.setNote('Auto-populated from Gorilla ROI. Type a number to override.');
    }

    // Link velocity override value formulas (auto-calc from last year's sales)
    for (var ovr = 1; ovr <= 3; ovr++) {
      var ovrKey  = 'VEL_OVERRIDE_' + ovr + '_VALUE';
      var ovrCell = ss.getRangeByName(prefix + '__' + ovrKey);
      if (!ovrCell) continue;

      // Skip if already has a Gorilla-linked formula
      var ovrFormula = ovrCell.getFormula();
      if (ovrFormula && (ovrFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
                         ovrFormula.indexOf('GORILLA_') > -1)) continue;

      // Skip if the user has a non-empty manual override
      if (!ovrFormula) {
        var ovrVal = ovrCell.getValue();
        if (ovrVal !== '' && ovrVal !== 0 && ovrVal !== null && ovrVal !== undefined) continue;
      }

      ovrCell.setFormula(buildVelAutoCalcFormula(prefix, skus[i].id, ovr));
      ovrCell.setBackground('#E8F0FE');
      ovrCell.setNote(
        'Auto-calculated from Gorilla ROI historical data.\n' +
        'Pulls last year\'s shipped sales for the same date range\n' +
        'and divides by the number of days to get daily velocity.\n\n' +
        'Just enter the start and end dates — this value fills in automatically.\n' +
        'Type a number to manually override.'
      );
    }

    // Default double-counting fields to 0 if empty, but respect manual overrides.
    for (var wKey in warningKeys) {
      if (!warningKeys.hasOwnProperty(wKey)) continue;
      var wRange = ss.getRangeByName(prefix + '__' + wKey);
      if (!wRange) continue;
      var wVal = wRange.getValue();
      if (wVal === '' || wVal === null || wVal === undefined) {
        wRange.setValue(0);
      }
      wRange.setBackground('#FFF3CD');
      wRange.setNote(warningKeys[wKey]);
    }
  }

  SpreadsheetApp.flush();
}
