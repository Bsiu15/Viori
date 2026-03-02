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
 * Maps Settings tab field keys to Gorilla Data tab column letters (FBA fields).
 * Used by SettingsTab.gs to create formula references.
 */
var GORILLA_LINK_MAP = {
  'ONHAND_AVAILABLE':        'B',
  'INBOUND_WORKING':         'C',
  'INBOUND_SHIPPED':         'D',
  'INBOUND_RECEIVING':       'E',
  'RESERVED_TOTAL':          'F',   // Display-only — Available already excludes reserved
  'ONHAND_FC_TRANSFER':      'G',
  'UNSELLABLE_TOTAL':        'H',   // Display-only — total unfulfillable units
  'DAILY_VELOCITY':          'J',
  'SELLING_PRICE':           'K'
};

/**
 * Maps Settings tab FBM field keys to Gorilla Data tab column letters.
 * FBM data lives in columns M-P, pulled using the FBM SKU ID.
 */
var GORILLA_FBM_LINK_MAP = {
  'FBM_ONHAND':              'N',   // FBM available/on-hand
  'FBM_DAILY_VELOCITY':      'P',   // FBM daily velocity (derived)
  'FBM_SELLING_PRICE':       'Q'    // FBM selling price
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

  // Ensure enough columns and rows (11 FBA cols + 1 spacer + 6 FBM cols = 18)
  var requiredCols = 18;
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

  // ── Data rows ──
  // Per Gorilla ROI docs, use SKU RANGES (A2:A<n>) instead of individual
  // cell refs so each column makes ONE bulk API call. This drops the total
  // from 9 × N_SKUs calls down to just 9 calls and avoids the Google
  // quota error ("unusually high number of requests").

  // Col A: SKU IDs (plain values)
  var skuValues = [];
  for (var i = 0; i < skus.length; i++) {
    skuValues.push([skus[i].id]);
  }
  sheet.getRange(2, 1, skus.length, 1).setValues(skuValues);

  var lastRow  = skus.length + 1;
  var skuRange = 'A2:A' + lastRow;

  // Col B: Available (fulfillable)
  sheet.getRange(2, 2).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "fulfillable"), 0)'
  );

  // Col C: Inbound Working
  sheet.getRange(2, 3).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "inbound_working"), 0)'
  );

  // Col D: Inbound Shipped
  sheet.getRange(2, 4).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "inbound_shipped"), 0)'
  );

  // Col E: Inbound Receiving
  sheet.getRange(2, 5).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "inbound_receiving"), 0)'
  );

  // Col F: Reserved
  sheet.getRange(2, 6).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "reserved"), 0)'
  );

  // Col G: FC Transfer
  sheet.getRange(2, 7).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "transfer"), 0)'
  );

  // Col H: Unsellable (total unfulfillable)
  sheet.getRange(2, 8).setFormula(
    '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuRange + ', ' + mktRef + ', "unsellable"), 0)'
  );

  // Col I: Sales Count (last N days)
  // MCF param: "Exclude" omits Multi-Channel Fulfillment orders (FBA only)
  // NOTE: GORILLA_SALESCOUNT does not spill results when given a SKU range
  // (unlike GORILLA_INVENTORY), so we write one formula per row.
  for (var s = 0; s < skus.length; s++) {
    var skuCell = 'A' + (s + 2);
    sheet.getRange(s + 2, 9).setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + skuCell +
      ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );
  }

  // Col J: Daily Velocity (derived: sales / lookback days)
  sheet.getRange(2, 10).setFormula(
    '=ARRAYFORMULA(IFERROR(I2:I' + lastRow + '/' + lookRef + ', 0))'
  );

  // Col K: Selling Price
  sheet.getRange(2, 11).setFormula(
    '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + skuRange + ', ' + mktRef + '), 0)'
  );

  // ── Formatting (FBA columns) ──
  if (skus.length > 0) {
    sheet.getRange(2, 1, skus.length, 1).setFontWeight('bold');
    // Cols B-I: inventory counts + sales count (integers)
    sheet.getRange(2, 2, skus.length, 8).setNumberFormat('#,##0');
    // Col J: Daily Velocity (one decimal)
    sheet.getRange(2, 10, skus.length, 1).setNumberFormat('#,##0.0');
    // Col K: Selling Price (currency)
    sheet.getRange(2, 11, skus.length, 1).setNumberFormat('$#,##0.00');
  }

  // ══════════════════════════════════════════════════════════════════════════
  // FBM SECTION (columns M-Q): Pulled using each SKU's optional FBM SKU ID
  // ══════════════════════════════════════════════════════════════════════════

  // Col L: Spacer
  sheet.getRange(1, 12).setValue('').setBackground('#F5F5F5');

  // FBM Headers (cols M-Q)
  var fbmHeaders = ['FBM SKU', 'FBM Available', 'FBM Sales (lookback)', 'FBM Velocity', 'FBM Price'];
  sheet.getRange(1, 13, 1, fbmHeaders.length)
       .setValues([fbmHeaders])
       .setFontWeight('bold')
       .setFontSize(9)
       .setBackground('#FFF2CC')
       .setBorder(true, true, true, true, false, false)
       .setHorizontalAlignment('center');

  // Col M: FBM SKU IDs (formula referencing Settings named range)
  for (var fi = 0; fi < skus.length; fi++) {
    var fbmPrefix = namedRangePrefix(skus[fi].id);
    var fbmSkuRef = fbmPrefix + '__FBM_SKU_ID';
    sheet.getRange(fi + 2, 13).setFormula(
      '=IFERROR(' + fbmSkuRef + ', "")'
    );
  }

  // Cols N-Q: Individual formulas per row (only compute when FBM SKU is set)
  for (var fbi = 0; fbi < skus.length; fbi++) {
    var fbmSkuCell = 'M' + (fbi + 2);

    // Col N: FBM Available (fulfillable)
    sheet.getRange(fbi + 2, 14).setFormula(
      '=IF(' + fbmSkuCell + '<>"", IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ', "fulfillable"), 0), "")'
    );

    // Col O: FBM Sales Count (lookback)
    sheet.getRange(fbi + 2, 15).setFormula(
      '=IF(' + fbmSkuCell + '<>"", IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + fbmSkuCell +
      ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0), "")'
    );

    // Col P: FBM Daily Velocity (derived: sales / lookback days)
    sheet.getRange(fbi + 2, 16).setFormula(
      '=IF(O' + (fbi + 2) + '<>"", IFERROR(O' + (fbi + 2) + '/' + lookRef + ', 0), "")'
    );

    // Col Q: FBM Selling Price
    sheet.getRange(fbi + 2, 17).setFormula(
      '=IF(' + fbmSkuCell + '<>"", IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + '), 0), "")'
    );
  }

  // ── FBM Formatting ──
  if (skus.length > 0) {
    sheet.getRange(2, 13, skus.length, 1).setFontWeight('bold'); // FBM SKU
    sheet.getRange(2, 14, skus.length, 2).setNumberFormat('#,##0');  // Available + Sales
    sheet.getRange(2, 16, skus.length, 1).setNumberFormat('#,##0.0'); // Velocity
    sheet.getRange(2, 17, skus.length, 1).setNumberFormat('$#,##0.00'); // Price
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 200);
  for (var c = 2; c <= 11; c++) {
    sheet.setColumnWidth(c, 130);
  }
  sheet.setColumnWidth(12, 20);  // Spacer
  sheet.setColumnWidth(13, 200); // FBM SKU
  for (var fc = 14; fc <= 17; fc++) {
    sheet.setColumnWidth(fc, 130);
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

  // ── FBM Header notes ──
  sheet.getRange(1, 13).setNote(
    'FBM SKU = Your Fulfilled-by-Merchant product ID.\n' +
    'Set the FBM SKU ID in the Settings tab for each product that has\n' +
    'an FBM listing. Leave blank if no FBM counterpart exists.'
  );
  sheet.getRange(1, 14).setNote(
    'FBM Available = Units you have on hand for merchant fulfillment.\n' +
    'This is the FBM equivalent of the FBA "Available" column.\n\n' +
    'Settings field: "FBM on-hand"'
  );
  sheet.getRange(1, 15).setNote(
    'FBM Sales (lookback) = Total FBM units sold in the last N days.\n' +
    'Used to derive FBM daily velocity for the forecast.'
  );
  sheet.getRange(1, 16).setNote(
    'FBM Velocity = How many FBM units you sell per day on average.\n' +
    'Calculated as: FBM Sales (lookback) / Lookback Days.\n\n' +
    'Settings field: "FBM daily sales velocity"'
  );
  sheet.getRange(1, 17).setNote(
    'FBM Price = Your current FBM listing price on Amazon.\n' +
    'May differ from FBA price. Used when the forecast switches\n' +
    'to FBM fulfillment to calculate revenue impact.\n\n' +
    'Settings field: "FBM selling price"'
  );

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}

/**
 * Lightweight linking: sets IFERROR formulas on existing Settings named ranges
 * to point at the Gorilla Data tab. Only touches the fields in GORILLA_LINK_MAP
 * (7 fields × N SKUs), so it runs in seconds instead of minutes.
 *
 * @param {boolean} force  When true, overwrites ALL Gorilla-linkable cells
 *                         (including manual overrides). Use when the user
 *                         explicitly runs "Refresh Gorilla Data".
 *                         When false (default), skips cells that already have
 *                         a user-typed non-zero value to preserve overrides.
 */
function linkSettingsToGorilla(force) {
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

      // If the cell already has a Gorilla formula, re-apply blue indicator but skip formula write
      var existingFormula = cell.getFormula();
      if (existingFormula && existingFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1) {
        cell.setBackground('#E8F0FE');
        cell.setNote('Auto-populated from Gorilla ROI. Type a number to override.');
        continue;
      }

      // If the cell has a non-zero user-entered value (no formula), skip to preserve override
      // — unless force=true (user explicitly ran Refresh Gorilla Data)
      if (!force && !existingFormula) {
        var val = cell.getValue();
        if (val !== '' && val !== 0 && val !== null && val !== undefined) continue;
      }

      var gorillaCol = GORILLA_LINK_MAP[key];
      cell.setFormula("=IFERROR('" + GORILLA_DATA_TAB_NAME + "'!" + gorillaCol + gorillaRow + ", 0)");
      cell.setBackground('#E8F0FE');
      cell.setNote('Auto-populated from Gorilla ROI. Type a number to override.');
    }

    // Link FBA velocity override value formulas (auto-calc from last year's sales)
    for (var ovr = 1; ovr <= 3; ovr++) {
      var ovrKey  = 'VEL_OVERRIDE_' + ovr + '_VALUE';
      var ovrCell = ss.getRangeByName(prefix + '__' + ovrKey);
      if (!ovrCell) continue;

      var ovrNote = 'Auto-calculated from Gorilla ROI historical data.\n' +
        'Pulls last year\'s shipped sales for the same date range\n' +
        'and divides by the number of days to get daily velocity.\n\n' +
        'Just enter the start and end dates — this value fills in automatically.\n' +
        'Type a number to manually override.';

      // If already has a Gorilla-linked formula, re-apply blue indicator but skip formula write
      var ovrFormula = ovrCell.getFormula();
      if (ovrFormula && (ovrFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
                         ovrFormula.indexOf('GORILLA_') > -1)) {
        ovrCell.setBackground('#E8F0FE');
        ovrCell.setNote(ovrNote);
        continue;
      }

      // Skip if the user has a non-empty manual override
      // — unless force=true (user explicitly ran Refresh Gorilla Data)
      if (!force && !ovrFormula) {
        var ovrVal = ovrCell.getValue();
        if (ovrVal !== '' && ovrVal !== 0 && ovrVal !== null && ovrVal !== undefined) continue;
      }

      ovrCell.setFormula(buildVelAutoCalcFormula(prefix, skus[i].id, ovr));
      ovrCell.setBackground('#E8F0FE');
      ovrCell.setNote(ovrNote);
    }

    // Link FBM fields to Gorilla Data tab
    for (var fbmKey in GORILLA_FBM_LINK_MAP) {
      if (!GORILLA_FBM_LINK_MAP.hasOwnProperty(fbmKey)) continue;

      var fbmRangeName = prefix + '__' + fbmKey;
      var fbmCell = ss.getRangeByName(fbmRangeName);
      if (!fbmCell) continue;

      // If already has a Gorilla formula, re-apply blue indicator but skip formula write
      var fbmExistingFormula = fbmCell.getFormula();
      if (fbmExistingFormula && fbmExistingFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1) {
        fbmCell.setBackground('#E8F0FE');
        fbmCell.setNote('Auto-populated from Gorilla ROI (FBM SKU). Type a number to override.');
        continue;
      }

      if (!force && !fbmExistingFormula) {
        var fbmVal = fbmCell.getValue();
        if (fbmVal !== '' && fbmVal !== 0 && fbmVal !== null && fbmVal !== undefined) continue;
      }

      var fbmGorillaCol = GORILLA_FBM_LINK_MAP[fbmKey];
      fbmCell.setFormula("=IFERROR('" + GORILLA_DATA_TAB_NAME + "'!" + fbmGorillaCol + gorillaRow + ", 0)");
      fbmCell.setBackground('#E8F0FE');
      fbmCell.setNote('Auto-populated from Gorilla ROI (FBM SKU). Type a number to override.');
    }

    // Link FBM velocity override value formulas
    for (var fbmOvr = 1; fbmOvr <= 3; fbmOvr++) {
      var fbmOvrKey  = 'FBM_VEL_OVERRIDE_' + fbmOvr + '_VALUE';
      var fbmOvrCell = ss.getRangeByName(prefix + '__' + fbmOvrKey);
      if (!fbmOvrCell) continue;

      var fbmOvrNote = 'Auto-calculated from Gorilla ROI historical data using your FBM SKU ID.\n' +
        'Pulls last year\'s shipped sales for the same date range\n' +
        'and divides by the number of days to get daily FBM velocity.\n\n' +
        'Enter FBM SKU ID and override dates — this value fills in automatically.\n' +
        'Type a number to manually override.';

      // If already has a Gorilla-linked formula, re-apply blue indicator but skip formula write
      var fbmOvrFormula = fbmOvrCell.getFormula();
      if (fbmOvrFormula && (fbmOvrFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
                            fbmOvrFormula.indexOf('GORILLA_') > -1)) {
        fbmOvrCell.setBackground('#E8F0FE');
        fbmOvrCell.setNote(fbmOvrNote);
        continue;
      }

      if (!force && !fbmOvrFormula) {
        var fbmOvrVal = fbmOvrCell.getValue();
        if (fbmOvrVal !== '' && fbmOvrVal !== 0 && fbmOvrVal !== null && fbmOvrVal !== undefined) continue;
      }

      fbmOvrCell.setFormula(buildFbmVelAutoCalcFormula(prefix, fbmOvr));
      fbmOvrCell.setBackground('#E8F0FE');
      fbmOvrCell.setNote(fbmOvrNote);
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
