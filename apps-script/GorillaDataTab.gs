/**
 * GorillaDataTab.gs
 * ---------------------------------------------------------------------------
 * Builds a "Gorilla Data" tab that caches inventory, sales, and pricing data
 * from Amazon Seller Central via Gorilla ROI custom functions.
 *
 * ARCHITECTURE — "Snapshot" pattern:
 *   The Gorilla Data tab stores PLAIN VALUES (cached data), not live Gorilla
 *   formulas. This prevents Google Sheets from firing 100+ API calls every
 *   time the spreadsheet opens, which causes rate-limiting errors.
 *
 *   buildGorillaDataTab()      — Creates tab structure with placeholder 0s
 *   fetchGorillaDataStaged()   — Fetches data one SKU at a time:
 *                                 write formulas → wait → snapshot to values
 *   linkSettingsToGorilla()    — Links Settings cells to the cached values
 *   linkFbmForSku()            — Per-SKU FBM linking after sidebar save
 *
 * Refresh flow:
 *   Menu > Refresh Gorilla Data
 *     → buildGorillaDataTab() (structure)
 *     → fetchGorillaDataStaged() (staged fetch, ~10s per SKU)
 *     → linkSettingsToGorilla(true) (link Settings + velocity overrides)
 *
 * Requires the Gorilla ROI Google Sheets add-on to be installed and connected
 * to a Seller Central account.
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
 * Gorilla inventory categories and their column positions on the data tab.
 */
var GORILLA_INVENTORY_CATS = [
  { col: 2,  category: 'fulfillable' },       // Col B: Available
  { col: 3,  category: 'inbound_working' },    // Col C: Inbound Working
  { col: 4,  category: 'inbound_shipped' },    // Col D: Inbound Shipped
  { col: 5,  category: 'inbound_receiving' },  // Col E: Inbound Receiving
  { col: 6,  category: 'reserved' },           // Col F: Reserved
  { col: 7,  category: 'transfer' },           // Col G: FC Transfer
  { col: 8,  category: 'unsellable' }          // Col H: Unsellable
];

/**
 * Builds (or rebuilds) the Gorilla Data tab STRUCTURE with placeholder values.
 * Does NOT write any Gorilla API formulas — those are added temporarily by
 * fetchGorillaDataStaged() during a refresh, then converted to plain values.
 *
 * This means opening the spreadsheet triggers zero Gorilla API calls.
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

  // 11 FBA cols + 1 spacer + 5 FBM cols = 17 columns
  var requiredCols = 17;
  var requiredRows = skus.length + 1;
  var currentCols  = sheet.getMaxColumns();
  var currentRows  = sheet.getMaxRows();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }

  // ── FBA Headers ──
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

  // ── Col A: SKU IDs (plain values) ──
  var skuValues = [];
  for (var i = 0; i < skus.length; i++) {
    skuValues.push([skus[i].id]);
  }
  if (skus.length > 0) {
    sheet.getRange(2, 1, skus.length, 1).setValues(skuValues);
  }

  // ── Cols B-K: Placeholder values (0) ──
  if (skus.length > 0) {
    var placeholders = [];
    for (var p = 0; p < skus.length; p++) {
      placeholders.push([0, 0, 0, 0, 0, 0, 0, 0, 0, 0]); // B through K
    }
    sheet.getRange(2, 2, skus.length, 10).setValues(placeholders);
  }

  // ── FBA Formatting ──
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
  // FBM SECTION (columns M-Q)
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

  // Col M: FBM SKU IDs (formula referencing Settings named range — NOT an API call)
  for (var fi = 0; fi < skus.length; fi++) {
    var fbmPrefix = namedRangePrefix(skus[fi].id);
    var fbmSkuRef = fbmPrefix + '__FBM_SKU_ID';
    sheet.getRange(fi + 2, 13).setFormula(
      '=IFERROR(' + fbmSkuRef + ', "")'
    );
  }

  // Cols N-Q: Placeholder values
  if (skus.length > 0) {
    var fbmPlaceholders = [];
    for (var fp = 0; fp < skus.length; fp++) {
      fbmPlaceholders.push(['', '', '', '']); // N-Q empty until refresh
    }
    sheet.getRange(2, 14, skus.length, 4).setValues(fbmPlaceholders);
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

  // ── Header notes ──
  sheet.getRange(1, 1).setNote(
    'SKU = Your product ID in Amazon Seller Central.\n' +
    'This tab stores cached data from Gorilla ROI.\n' +
    'Use Menu > Inventory Forecast > Refresh Gorilla Data to fetch latest values.\n\n' +
    'Last Refreshed: never'
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
 * Staged Gorilla data fetch — processes one SKU at a time to avoid
 * Google Sheets rate limits on custom function calls.
 *
 * For each SKU:
 *   1. Writes GORILLA_* formulas to the Gorilla Data tab row
 *   2. Flushes to trigger evaluation
 *   3. Polls until formulas resolve (up to 30s per SKU)
 *   4. Reads computed values and overwrites formulas with plain values
 *   5. Pauses before processing the next SKU
 *
 * After all SKUs: updates "Last Refreshed" timestamp on cell A1 note.
 *
 * Safety: respects a 5-minute runtime limit. If the script is running
 * close to the Apps Script 6-minute timeout, it stops early and tells
 * the user to run again for remaining SKUs.
 */
function fetchGorillaDataStaged() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GORILLA_DATA_TAB_NAME);
  if (!sheet) return;

  var skus = getSkus();
  if (skus.length === 0) return;

  // Named range references used in formulas
  var sellerRef = 'GLOBAL__GORILLA_SELLER_ID';
  var mktRef    = 'GLOBAL__GORILLA_MARKETPLACE';
  var lookRef   = 'GLOBAL__GORILLA_LOOKBACK_DAYS';

  var lookbackRange = ss.getRangeByName('GLOBAL__GORILLA_LOOKBACK_DAYS');
  var lookbackDays  = lookbackRange ? lookbackRange.getValue() : 30;
  if (!lookbackDays || lookbackDays <= 0) lookbackDays = 30;

  var totalStart = new Date().getTime();
  var MAX_RUNTIME = 300000; // 5 minutes (1 min safety margin before Apps Script timeout)
  var skusProcessed = 0;

  for (var s = 0; s < skus.length; s++) {
    // Safety: check if close to timeout
    if (new Date().getTime() - totalStart > MAX_RUNTIME) {
      ss.toast(
        'Processed ' + skusProcessed + '/' + skus.length + ' SKUs before timeout.\n' +
        'Run "Refresh Gorilla Data" again to fetch the remaining SKUs.',
        'Gorilla ROI', 15
      );
      break;
    }

    var rowNum  = s + 2;
    var skuCell = 'A' + rowNum;
    var skuId   = skus[s].id;

    ss.toast(
      'Fetching data for ' + skuId + ' (' + (s + 1) + '/' + skus.length + ')...',
      'Gorilla ROI', 60
    );

    // ── Write FBA Gorilla formulas for this SKU ──
    var formulaCells = [];

    // Cols B-H: GORILLA_INVENTORY per category
    for (var ic = 0; ic < GORILLA_INVENTORY_CATS.length; ic++) {
      var cat  = GORILLA_INVENTORY_CATS[ic];
      var cell = sheet.getRange(rowNum, cat.col);
      cell.setFormula(
        '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ', "' + cat.category + '"), 0)'
      );
      formulaCells.push(cell);
    }

    // Col I: Sales Count (last N days)
    var salesCell = sheet.getRange(rowNum, 9);
    salesCell.setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + skuCell +
      ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );
    formulaCells.push(salesCell);

    // Col K: Selling Price
    var priceCell = sheet.getRange(rowNum, 11);
    priceCell.setFormula(
      '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + skuCell + ', ' + mktRef + '), 0)'
    );
    formulaCells.push(priceCell);

    // ── Write FBM Gorilla formulas if FBM SKU is configured ──
    var fbmSkuVal = sheet.getRange(rowNum, 13).getDisplayValue();
    var hasFbm    = fbmSkuVal && fbmSkuVal !== '';

    if (hasFbm) {
      var fbmSkuCell = 'M' + rowNum;

      // Col N: FBM Available
      var fbmAvailCell = sheet.getRange(rowNum, 14);
      fbmAvailCell.setFormula(
        '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ', "fulfillable"), 0)'
      );
      formulaCells.push(fbmAvailCell);

      // Col O: FBM Sales Count
      var fbmSalesCell = sheet.getRange(rowNum, 15);
      fbmSalesCell.setFormula(
        '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + fbmSkuCell +
        ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
      );
      formulaCells.push(fbmSalesCell);

      // Col Q: FBM Selling Price
      var fbmPriceCell = sheet.getRange(rowNum, 17);
      fbmPriceCell.setFormula(
        '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + '), 0)'
      );
      formulaCells.push(fbmPriceCell);
    }

    // Flush to trigger formula evaluation
    SpreadsheetApp.flush();

    // ── Poll until all formulas resolve (up to 30 seconds) ──
    var maxWait  = 30000;
    var interval = 3000;
    var elapsed  = 0;

    while (elapsed < maxWait) {
      Utilities.sleep(interval);
      SpreadsheetApp.flush();
      elapsed += interval;

      var allDone = true;
      for (var fc = 0; fc < formulaCells.length; fc++) {
        var dispVal = formulaCells[fc].getDisplayValue();
        if (dispVal === '' || dispVal === 'Loading...' || dispVal === null) {
          allDone = false;
          break;
        }
      }
      if (allDone) break;
    }

    // ── Snapshot: read computed values and overwrite formulas ──
    for (var sc = 0; sc < formulaCells.length; sc++) {
      var snapVal = formulaCells[sc].getValue();
      // If the value is an error string or non-numeric, write 0
      if (typeof snapVal !== 'number' || isNaN(snapVal)) {
        snapVal = 0;
      }
      formulaCells[sc].setValue(snapVal);
    }

    // Col J: Daily Velocity (derived — sales / lookback days)
    var salesVal = sheet.getRange(rowNum, 9).getValue();
    sheet.getRange(rowNum, 10).setValue(
      lookbackDays > 0 ? Math.round((salesVal / lookbackDays) * 10) / 10 : 0
    );

    // Col P: FBM Velocity (derived)
    if (hasFbm) {
      var fbmSalesVal = sheet.getRange(rowNum, 15).getValue();
      sheet.getRange(rowNum, 16).setValue(
        lookbackDays > 0 ? Math.round((fbmSalesVal / lookbackDays) * 10) / 10 : 0
      );
    }

    SpreadsheetApp.flush();
    skusProcessed++;

    // Pause before next SKU to stay under Gorilla rate limits
    if (s < skus.length - 1) {
      Utilities.sleep(2000);
    }
  }

  // ── Update "Last Refreshed" timestamp ──
  var refreshTime = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a'
  );
  sheet.getRange(1, 1).setNote(
    'SKU = Your product ID in Amazon Seller Central.\n' +
    'This tab stores cached data from Gorilla ROI.\n' +
    'Use Menu > Inventory Forecast > Refresh Gorilla Data to fetch latest values.\n\n' +
    'Last Refreshed: ' + refreshTime
  );

  if (skusProcessed === skus.length) {
    ss.toast('All ' + skusProcessed + ' SKUs refreshed successfully!', 'Gorilla ROI', 5);
  }
}

/**
 * Fetches Gorilla data for a SINGLE SKU. Used by the save flow when a
 * new FBM SKU is configured and the Gorilla Data tab needs FBM values.
 *
 * Same staged approach as fetchGorillaDataStaged() but for one SKU only.
 *
 * @param {number} skuIndex  The SKU's index in getSkus() (0-based)
 * @param {boolean} fbmOnly  When true, only fetches FBM columns (N, O, Q)
 */
function fetchGorillaDataForSku(skuIndex, fbmOnly) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GORILLA_DATA_TAB_NAME);
  if (!sheet) return;

  var sellerRef = 'GLOBAL__GORILLA_SELLER_ID';
  var mktRef    = 'GLOBAL__GORILLA_MARKETPLACE';
  var lookRef   = 'GLOBAL__GORILLA_LOOKBACK_DAYS';

  var lookbackRange = ss.getRangeByName('GLOBAL__GORILLA_LOOKBACK_DAYS');
  var lookbackDays  = lookbackRange ? lookbackRange.getValue() : 30;
  if (!lookbackDays || lookbackDays <= 0) lookbackDays = 30;

  var rowNum  = skuIndex + 2;
  var skuCell = 'A' + rowNum;
  var formulaCells = [];

  if (!fbmOnly) {
    // FBA formulas
    for (var ic = 0; ic < GORILLA_INVENTORY_CATS.length; ic++) {
      var cat  = GORILLA_INVENTORY_CATS[ic];
      var cell = sheet.getRange(rowNum, cat.col);
      cell.setFormula(
        '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ', "' + cat.category + '"), 0)'
      );
      formulaCells.push(cell);
    }

    var salesCell = sheet.getRange(rowNum, 9);
    salesCell.setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + skuCell +
      ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );
    formulaCells.push(salesCell);

    var priceCell = sheet.getRange(rowNum, 11);
    priceCell.setFormula(
      '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + skuCell + ', ' + mktRef + '), 0)'
    );
    formulaCells.push(priceCell);
  }

  // FBM formulas
  var fbmSkuVal = sheet.getRange(rowNum, 13).getDisplayValue();
  if (fbmSkuVal && fbmSkuVal !== '') {
    var fbmSkuCell = 'M' + rowNum;

    var fbmAvailCell = sheet.getRange(rowNum, 14);
    fbmAvailCell.setFormula(
      '=IFERROR(GORILLA_INVENTORY(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ', "fulfillable"), 0)'
    );
    formulaCells.push(fbmAvailCell);

    var fbmSalesCell = sheet.getRange(rowNum, 15);
    fbmSalesCell.setFormula(
      '=IFERROR(GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + fbmSkuCell +
      ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd")), 0)'
    );
    formulaCells.push(fbmSalesCell);

    var fbmPriceCell = sheet.getRange(rowNum, 17);
    fbmPriceCell.setFormula(
      '=IFERROR(GORILLA_MYPRICE(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + '), 0)'
    );
    formulaCells.push(fbmPriceCell);
  }

  if (formulaCells.length === 0) return;

  SpreadsheetApp.flush();

  // Poll until resolved
  var maxWait = 30000;
  var interval = 3000;
  var elapsed = 0;

  while (elapsed < maxWait) {
    Utilities.sleep(interval);
    SpreadsheetApp.flush();
    elapsed += interval;

    var allDone = true;
    for (var fc = 0; fc < formulaCells.length; fc++) {
      var dispVal = formulaCells[fc].getDisplayValue();
      if (dispVal === '' || dispVal === 'Loading...' || dispVal === null) {
        allDone = false;
        break;
      }
    }
    if (allDone) break;
  }

  // Snapshot to values
  for (var sc = 0; sc < formulaCells.length; sc++) {
    var snapVal = formulaCells[sc].getValue();
    if (typeof snapVal !== 'number' || isNaN(snapVal)) snapVal = 0;
    formulaCells[sc].setValue(snapVal);
  }

  // Derived velocities
  if (!fbmOnly) {
    var salesVal = sheet.getRange(rowNum, 9).getValue();
    sheet.getRange(rowNum, 10).setValue(
      lookbackDays > 0 ? Math.round((salesVal / lookbackDays) * 10) / 10 : 0
    );
  }
  if (fbmSkuVal && fbmSkuVal !== '') {
    var fbmSalesVal = sheet.getRange(rowNum, 15).getValue();
    sheet.getRange(rowNum, 16).setValue(
      lookbackDays > 0 ? Math.round((fbmSalesVal / lookbackDays) * 10) / 10 : 0
    );
  }

  SpreadsheetApp.flush();
}

/**
 * Lightweight linking: sets IFERROR formulas on existing Settings named ranges
 * to point at the Gorilla Data tab. Only touches the fields in GORILLA_LINK_MAP
 * (7 fields x N SKUs), so it runs in seconds instead of minutes.
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

  // Collect velocity override cells for batch snapshotting at the end
  var velocityOverrideCells = [];

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
        // Still collect for snapshotting (has live formula)
        velocityOverrideCells.push(ovrCell);
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
      velocityOverrideCells.push(ovrCell);
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

      var fbmOvrNote = 'Auto-calculated from last year\'s FBA sales \u00d7 FBM conversion rate.\n' +
        'Uses FBA sales (not FBM) because when FBA is in stock, FBA wins\n' +
        'the Buy Box and FBM sales are ~0 — not representative of FBM\n' +
        'demand during an FBA stockout.\n\n' +
        'Enter FBM SKU ID and override dates — this value fills in automatically.\n' +
        'Type a number to manually override.';

      // If already has a Gorilla-linked formula, re-apply blue indicator but skip formula write
      var fbmOvrFormula = fbmOvrCell.getFormula();
      if (fbmOvrFormula && (fbmOvrFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
                            fbmOvrFormula.indexOf('GORILLA_') > -1)) {
        fbmOvrCell.setBackground('#E8F0FE');
        fbmOvrCell.setNote(fbmOvrNote);
        velocityOverrideCells.push(fbmOvrCell);
        continue;
      }

      if (!force && !fbmOvrFormula) {
        var fbmOvrVal = fbmOvrCell.getValue();
        if (fbmOvrVal !== '' && fbmOvrVal !== 0 && fbmOvrVal !== null && fbmOvrVal !== undefined) continue;
      }

      fbmOvrCell.setFormula(buildFbmVelAutoCalcFormula(prefix, skus[i].id, fbmOvr));
      fbmOvrCell.setBackground('#E8F0FE');
      fbmOvrCell.setNote(fbmOvrNote);
      velocityOverrideCells.push(fbmOvrCell);
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

  // ── Snapshot velocity override formulas ──
  // Wait for GORILLA_SALESCOUNT formulas to resolve, then convert to values.
  // This prevents live Gorilla formulas from persisting on the Settings tab.
  if (velocityOverrideCells.length > 0) {
    // Filter to cells that actually have GORILLA_ formulas and whose dates are set
    var liveCells = [];
    for (var lc = 0; lc < velocityOverrideCells.length; lc++) {
      var lcFormula = velocityOverrideCells[lc].getFormula();
      if (lcFormula && lcFormula.indexOf('GORILLA_') > -1) {
        liveCells.push(velocityOverrideCells[lc]);
      }
    }

    if (liveCells.length > 0) {
      ss.toast('Waiting for velocity override formulas...', 'Gorilla ROI', 45);

      var maxWait  = 45000;
      var interval = 3000;
      var elapsed  = 0;

      while (elapsed < maxWait) {
        Utilities.sleep(interval);
        SpreadsheetApp.flush();
        elapsed += interval;

        var allDone = true;
        for (var vc = 0; vc < liveCells.length; vc++) {
          var vcVal = liveCells[vc].getValue();
          if (vcVal === '' || vcVal === null || vcVal === undefined ||
              (typeof vcVal === 'string' && (vcVal === 'Loading...' || vcVal.charAt(0) === '#'))) {
            allDone = false;
            break;
          }
        }
        if (allDone) break;
      }

      // Snapshot: overwrite formulas with computed values
      for (var sv = 0; sv < liveCells.length; sv++) {
        var svVal = liveCells[sv].getValue();
        if (typeof svVal === 'number' && !isNaN(svVal)) {
          liveCells[sv].setValue(svVal);
        }
        // If the formula returned "" (dates not set), leave it as "" (no-op)
      }

      SpreadsheetApp.flush();
    }
  }
}

/**
 * Links just the FBM fields for a single SKU to the Gorilla Data tab.
 * Called automatically when a user saves an FBM SKU ID from the sidebar.
 *
 * Also triggers a targeted Gorilla Data fetch for this SKU's FBM columns
 * so data is available immediately (not just zeros).
 *
 * @param {string} skuId  The SKU whose FBM fields should be linked
 */
function linkFbmForSku(skuId) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var skus = getSkus();

  // Find this SKU's index → Gorilla Data row
  var skuIndex = -1;
  for (var i = 0; i < skus.length; i++) {
    if (skus[i].id === skuId) {
      skuIndex = i;
      break;
    }
  }
  if (skuIndex === -1) return;

  var prefix     = namedRangePrefix(skuId);
  var gorillaRow = skuIndex + 2;

  // Fetch FBM data from Gorilla for this SKU (staged: write formula → wait → snapshot)
  fetchGorillaDataForSku(skuIndex, true);

  // Link FBM data fields (FBM_ONHAND, FBM_DAILY_VELOCITY, FBM_SELLING_PRICE)
  for (var fbmKey in GORILLA_FBM_LINK_MAP) {
    if (!GORILLA_FBM_LINK_MAP.hasOwnProperty(fbmKey)) continue;

    var rangeName = prefix + '__' + fbmKey;
    var cell = ss.getRangeByName(rangeName);
    if (!cell) continue;

    // If already linked to Gorilla Data, skip
    var existingFormula = cell.getFormula();
    if (existingFormula && existingFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1) continue;

    var gorillaCol = GORILLA_FBM_LINK_MAP[fbmKey];
    cell.setFormula("=IFERROR('" + GORILLA_DATA_TAB_NAME + "'!" + gorillaCol + gorillaRow + ", 0)");
    cell.setBackground('#E8F0FE');
    cell.setNote('Auto-populated from Gorilla ROI (FBM SKU). Type a number to override.');
  }

  // Link FBM velocity override auto-calc formulas
  for (var ovr = 1; ovr <= 3; ovr++) {
    var ovrKey  = 'FBM_VEL_OVERRIDE_' + ovr + '_VALUE';
    var ovrCell = ss.getRangeByName(prefix + '__' + ovrKey);
    if (!ovrCell) continue;

    // Skip if already linked
    var ovrFormula = ovrCell.getFormula();
    if (ovrFormula && (ovrFormula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
                       ovrFormula.indexOf('GORILLA_') > -1)) continue;

    // Skip if user has a non-empty manual value
    if (!ovrFormula) {
      var ovrVal = ovrCell.getValue();
      if (ovrVal !== '' && ovrVal !== 0 && ovrVal !== null && ovrVal !== undefined) continue;
    }

    ovrCell.setFormula(buildFbmVelAutoCalcFormula(prefix, skuId, ovr));
    ovrCell.setBackground('#E8F0FE');
    ovrCell.setNote('Auto-calculated from last year\'s FBA sales \u00d7 FBM conversion rate.\nUses FBA sales because FBM sees ~0 sales when FBA has the Buy Box.\nType a number to manually override.');
  }

  SpreadsheetApp.flush();
}
