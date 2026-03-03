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
 * Writes a small batch of GORILLA_* formulas, waits for them to resolve,
 * and snapshots the results to plain values. Returns true if ALL formulas
 * in the batch resolved successfully (no #ERROR!, no timeout).
 *
 * @param {Sheet}   sheet        The Gorilla Data tab
 * @param {Array}   batch        Array of {cell, formula} objects
 * @param {number}  pollTimeout  Max ms to wait for resolution (default 30000)
 * @return {boolean} true if all resolved, false if any errored or timed out
 */
function writeBatchAndSnapshot_(sheet, batch, pollTimeout) {
  if (!batch || batch.length === 0) return true;
  if (!pollTimeout) pollTimeout = 30000;

  // Write formulas
  for (var i = 0; i < batch.length; i++) {
    batch[i].cell.setFormula(batch[i].formula);
  }
  SpreadsheetApp.flush();

  // Poll until all resolve
  var interval = 3000;
  var elapsed  = 0;
  var hasError = false;

  while (elapsed < pollTimeout) {
    Utilities.sleep(interval);
    SpreadsheetApp.flush();
    elapsed += interval;

    var allDone = true;
    hasError = false;
    for (var p = 0; p < batch.length; p++) {
      var dv = batch[p].cell.getDisplayValue();
      if (dv === '' || dv === 'Loading...' || dv === null) {
        allDone = false;
        break;
      }
      if (typeof dv === 'string' && dv.charAt(0) === '#') {
        hasError = true;
      }
    }
    if (allDone) break;
  }

  // Snapshot: overwrite formulas with plain values
  for (var s = 0; s < batch.length; s++) {
    var val = batch[s].cell.getValue();
    if (typeof val !== 'number' || isNaN(val)) val = 0;
    batch[s].cell.setValue(val);
  }
  SpreadsheetApp.flush();

  return !hasError;
}

/**
 * Probes the Gorilla API with a single lightweight formula to check
 * whether the account is currently rate-limited.
 *
 * @param {Sheet}  sheet      The Gorilla Data tab
 * @param {string} sellerRef  Named range for seller ID
 * @param {string} mktRef     Named range for marketplace
 * @param {number} testRow    Row to write the test formula in
 * @return {string} 'ok' if the formula resolved to a number,
 *                  'throttled' if #ERROR!, 'timeout' if it never resolved
 */
function probeGorillaApi_(sheet, sellerRef, mktRef, testRow) {
  var testCell = sheet.getRange(testRow, 2);
  testCell.setFormula(
    '=GORILLA_INVENTORY(' + sellerRef + ', A2, ' + mktRef + ', "fulfillable")'
  );
  SpreadsheetApp.flush();

  var result = 'timeout';
  for (var t = 0; t < 6; t++) {
    Utilities.sleep(3000);
    SpreadsheetApp.flush();
    var dv = testCell.getDisplayValue();
    if (dv === '' || dv === 'Loading...' || dv === null) continue;
    if (typeof dv === 'string' && dv.charAt(0) === '#') {
      result = 'throttled';
    } else {
      result = 'ok';
    }
    break;
  }

  // Clean up — always remove the test formula
  testCell.setValue('');
  SpreadsheetApp.flush();
  return result;
}

/**
 * Staged Gorilla data fetch — processes one SKU at a time, in micro-batches
 * of 3 formulas, to stay well under Google Sheets rate limits.
 *
 * RESILIENCE FEATURES:
 *   - Probe test: checks API availability with a single formula before starting
 *   - Micro-batching: writes 3 formulas at a time (not 12)
 *   - Error detection: if #ERROR! is returned, stops and backs off
 *   - Exponential backoff: delays increase when throttling is detected
 *   - Retry queue: failed SKUs get a second attempt after a cooldown
 *   - 5-minute safety timeout
 *
 * For each SKU:
 *   1. Writes formulas in micro-batches of 3
 *   2. Waits for each batch to resolve before writing the next
 *   3. Snapshots resolved values immediately (removes live formulas)
 *   4. Pauses between SKUs (longer if throttling detected)
 *
 * After all SKUs: updates "Last Refreshed" timestamp on cell A1 note.
 */
function fetchGorillaDataStaged() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GORILLA_DATA_TAB_NAME);
  if (!sheet) return false;

  var skus = getSkus();
  if (skus.length === 0) return false;

  var sellerRef = 'GLOBAL__GORILLA_SELLER_ID';
  var mktRef    = 'GLOBAL__GORILLA_MARKETPLACE';
  var lookRef   = 'GLOBAL__GORILLA_LOOKBACK_DAYS';

  var lookbackRange = ss.getRangeByName('GLOBAL__GORILLA_LOOKBACK_DAYS');
  var lookbackDays  = lookbackRange ? lookbackRange.getValue() : 30;
  if (!lookbackDays || lookbackDays <= 0) lookbackDays = 30;

  // ── Probe: test API availability before starting ──
  ss.toast('Testing Gorilla connection...', 'Gorilla ROI', 20);
  var probeRow = skus.length + 3; // temp row beyond data
  var probeResult = probeGorillaApi_(sheet, sellerRef, mktRef, probeRow);

  if (probeResult === 'throttled') {
    SpreadsheetApp.getUi().alert(
      'Gorilla API Rate Limited',
      'Google is still throttling Gorilla API requests for your account.\n\n' +
      'This is a temporary cooldown imposed by Google — usually resets within\n' +
      '30-60 minutes. The previous burst of formulas used up your quota.\n\n' +
      'What to do:\n' +
      '  1. Wait 30-60 minutes\n' +
      '  2. Try "Refresh Gorilla Data" again\n' +
      '  3. If it still fails, wait a bit longer\n\n' +
      'Your existing cached data (if any) is still intact and the forecast\n' +
      'will continue to work with the last known values.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }

  if (probeResult === 'timeout') {
    SpreadsheetApp.getUi().alert(
      'Gorilla Connection Timeout',
      'The test formula did not resolve within 18 seconds.\n\n' +
      'Possible causes:\n' +
      '  - Gorilla ROI add-on is not installed or not authorized\n' +
      '  - Internet connection issue\n' +
      '  - Gorilla servers are slow\n\n' +
      'Try: Extensions > Add-ons > Manage add-ons > enable Gorilla ROI.\n' +
      'Then try "Refresh Gorilla Data" again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }

  // Probe passed — API is available
  ss.toast('Gorilla connected! Starting staged refresh...', 'Gorilla ROI', 5);

  var totalStart    = new Date().getTime();
  var MAX_RUNTIME   = 300000; // 5 minutes
  var MICRO_BATCH   = 2;      // formulas per batch (kept very low to avoid throttling)
  var basePause     = 5000;   // ms between SKUs (increases on errors)
  var currentPause  = basePause;
  var skusProcessed = 0;
  var failedSkus    = [];     // indices of SKUs that failed (for retry)

  for (var s = 0; s < skus.length; s++) {
    if (new Date().getTime() - totalStart > MAX_RUNTIME) {
      ss.toast(
        'Processed ' + skusProcessed + '/' + skus.length + ' SKUs before timeout.\n' +
        'Run "Refresh Gorilla Data" again for remaining SKUs.',
        'Gorilla ROI', 15
      );
      break;
    }

    var skuOk = fetchSingleSkuMicrobatched_(
      ss, sheet, skus, s, sellerRef, mktRef, lookRef, lookbackDays, MICRO_BATCH
    );

    if (skuOk) {
      skusProcessed++;
      currentPause = basePause; // reset backoff on success
    } else {
      failedSkus.push(s);
      currentPause = Math.min(currentPause * 2, 15000); // exponential backoff, max 15s
      ss.toast(
        'Rate limited on ' + skus[s].id + ' — backing off ' +
        Math.round(currentPause / 1000) + 's...',
        'Gorilla ROI', 10
      );
    }

    // Pause before next SKU
    if (s < skus.length - 1) {
      Utilities.sleep(currentPause);
    }
  }

  // ── Retry failed SKUs once after a longer cooldown ──
  if (failedSkus.length > 0 && new Date().getTime() - totalStart < MAX_RUNTIME) {
    ss.toast(
      'Retrying ' + failedSkus.length + ' failed SKU(s) after cooldown...',
      'Gorilla ROI', 20
    );
    Utilities.sleep(15000); // 15-second cooldown before retry

    for (var r = 0; r < failedSkus.length; r++) {
      if (new Date().getTime() - totalStart > MAX_RUNTIME) break;

      var retryOk = fetchSingleSkuMicrobatched_(
        ss, sheet, skus, failedSkus[r], sellerRef, mktRef, lookRef, lookbackDays, MICRO_BATCH
      );

      if (retryOk) {
        skusProcessed++;
      }

      if (r < failedSkus.length - 1) {
        Utilities.sleep(5000); // longer pause between retries
      }
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
    return true;
  } else {
    var failCount = skus.length - skusProcessed;
    SpreadsheetApp.getUi().alert(
      'Gorilla Refresh Partial',
      skusProcessed + ' of ' + skus.length + ' SKUs refreshed successfully.\n' +
      failCount + ' SKU(s) could not be fetched due to rate limiting.\n\n' +
      'Wait 15-30 minutes, then run "Refresh Gorilla Data" again.\n' +
      'Successfully fetched SKUs will keep their data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return skusProcessed > 0; // partial success — link what we got
  }
}

/**
 * Fetches all Gorilla data for a single SKU using micro-batches.
 * Writes formulas in groups of MICRO_BATCH, waits for each group,
 * then snapshots to plain values before writing the next group.
 *
 * @param {Spreadsheet} ss
 * @param {Sheet}       sheet
 * @param {Array}       skus         Full SKU array
 * @param {number}      skuIndex     Index into skus[]
 * @param {string}      sellerRef    Named range ref
 * @param {string}      mktRef       Named range ref
 * @param {string}      lookRef      Named range ref
 * @param {number}      lookbackDays
 * @param {number}      batchSize    Formulas per micro-batch
 * @return {boolean}    true if all formulas resolved without errors
 */
function fetchSingleSkuMicrobatched_(ss, sheet, skus, skuIndex, sellerRef, mktRef, lookRef, lookbackDays, batchSize) {
  var rowNum  = skuIndex + 2;
  var skuCell = 'A' + rowNum;
  var skuId   = skus[skuIndex].id;

  ss.toast(
    'Fetching data for ' + skuId + ' (' + (skuIndex + 1) + '/' + skus.length + ')...',
    'Gorilla ROI', 60
  );

  // Build all formula descriptors for this SKU.
  // IMPORTANT: NO IFERROR wrapper — we need raw formulas so #ERROR! is
  // visible to our error detection. Errors are caught during snapshot.
  var allFormulas = [];

  // Cols B-H: GORILLA_INVENTORY per category
  for (var ic = 0; ic < GORILLA_INVENTORY_CATS.length; ic++) {
    var cat = GORILLA_INVENTORY_CATS[ic];
    allFormulas.push({
      cell: sheet.getRange(rowNum, cat.col),
      formula: '=GORILLA_INVENTORY(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ', "' + cat.category + '")'
    });
  }

  // Col I: Sales Count
  allFormulas.push({
    cell: sheet.getRange(rowNum, 9),
    formula: '=GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + skuCell +
             ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd"))'
  });

  // Col K: Selling Price
  allFormulas.push({
    cell: sheet.getRange(rowNum, 11),
    formula: '=GORILLA_MYPRICE(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ')'
  });

  // FBM formulas (if configured)
  var fbmSkuVal = sheet.getRange(rowNum, 13).getDisplayValue();
  var hasFbm    = fbmSkuVal && fbmSkuVal !== '';

  if (hasFbm) {
    var fbmSkuCell = 'M' + rowNum;

    allFormulas.push({
      cell: sheet.getRange(rowNum, 14),
      formula: '=GORILLA_INVENTORY(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ', "fulfillable")'
    });
    allFormulas.push({
      cell: sheet.getRange(rowNum, 15),
      formula: '=GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + fbmSkuCell +
               ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd"))'
    });
    allFormulas.push({
      cell: sheet.getRange(rowNum, 17),
      formula: '=GORILLA_MYPRICE(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ')'
    });
  }

  // Process in micro-batches
  var allOk = true;
  for (var b = 0; b < allFormulas.length; b += batchSize) {
    var batch = allFormulas.slice(b, b + batchSize);
    var batchOk = writeBatchAndSnapshot_(sheet, batch, 30000);
    if (!batchOk) {
      allOk = false;
      // Don't break — continue snapshotting remaining batches so we don't
      // leave live formulas on the sheet. But the values will be 0.
    }

    // Pause between micro-batches within the same SKU
    if (b + batchSize < allFormulas.length) {
      Utilities.sleep(3000);
    }
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
  return allOk;
}

/**
 * Fetches Gorilla data for a SINGLE SKU using micro-batches.
 * Used by the save flow when a new FBM SKU is configured.
 *
 * Uses writeBatchAndSnapshot_() for resilient formula handling.
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
  var allFormulas = [];

  // Raw formulas (no IFERROR) — errors caught during snapshot
  if (!fbmOnly) {
    for (var ic = 0; ic < GORILLA_INVENTORY_CATS.length; ic++) {
      var cat = GORILLA_INVENTORY_CATS[ic];
      allFormulas.push({
        cell: sheet.getRange(rowNum, cat.col),
        formula: '=GORILLA_INVENTORY(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ', "' + cat.category + '")'
      });
    }

    allFormulas.push({
      cell: sheet.getRange(rowNum, 9),
      formula: '=GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + skuCell +
               ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd"))'
    });

    allFormulas.push({
      cell: sheet.getRange(rowNum, 11),
      formula: '=GORILLA_MYPRICE(' + sellerRef + ', ' + skuCell + ', ' + mktRef + ')'
    });
  }

  // FBM formulas
  var fbmSkuVal = sheet.getRange(rowNum, 13).getDisplayValue();
  if (fbmSkuVal && fbmSkuVal !== '') {
    var fbmSkuCell = 'M' + rowNum;

    allFormulas.push({
      cell: sheet.getRange(rowNum, 14),
      formula: '=GORILLA_INVENTORY(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ', "fulfillable")'
    });
    allFormulas.push({
      cell: sheet.getRange(rowNum, 15),
      formula: '=GORILLA_SALESCOUNT(' + sellerRef + ', "Custom", ' + mktRef + ', ' + fbmSkuCell +
               ', "Shipped", "Exclude", TEXT(TODAY()-' + lookRef + ', "yyyy-mm-dd"), TEXT(TODAY()-1, "yyyy-mm-dd"))'
    });
    allFormulas.push({
      cell: sheet.getRange(rowNum, 17),
      formula: '=GORILLA_MYPRICE(' + sellerRef + ', ' + fbmSkuCell + ', ' + mktRef + ')'
    });
  }

  if (allFormulas.length === 0) return;

  // Process in micro-batches of 2
  for (var b = 0; b < allFormulas.length; b += 2) {
    var batch = allFormulas.slice(b, b + 2);
    writeBatchAndSnapshot_(sheet, batch, 30000);
    if (b + 2 < allFormulas.length) {
      Utilities.sleep(3000);
    }
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
