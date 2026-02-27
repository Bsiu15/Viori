/**
 * Main.gs
 * ---------------------------------------------------------------------------
 * Orchestration layer — ties together all modules.
 *
 * Entry points:
 *   onOpen()                    — Adds custom menu
 *   openSettingsDialog()        — Opens the SKU settings sidebar
 *   initialSetup()              — First-time setup
 *   recalculateAll()            — Full recalculation (all SKUs + Summary)
 *   recalculateSku(skuId)       — Single SKU recalculation (detail tab only)
 *   getSkuList()                — Returns SKU list for the dialog
 *   getSkuSettingsForDialog()   — Reads one SKU's settings for the dialog
 *   saveSkuSettingsFromDialog() — Writes one SKU's settings from the dialog
 *   addNewSku(skuId, skuName)   — Adds a new SKU to the registry
 *   removeExistingSku(skuId)    — Removes a SKU from the registry
 * ---------------------------------------------------------------------------
 */

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Forecast')
    .addItem('Edit SKU Settings', 'openSettingsDialog')
    .addItem('Set Forecast Date Range', 'openDateRangeDialog')
    .addSeparator()
    .addItem('Recalculate All', 'recalculateAll')
    .addItem('Refresh Gorilla Data', 'refreshGorillaData')
    .addSeparator()
    .addItem('Initial Setup (first time)', 'initialSetup')
    .addItem('Install Auto-Refresh Trigger', 'installTrigger')
    .addToUi();
}

/**
 * Opens the SKU settings sidebar.
 */
function openSettingsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setTitle('SKU Settings')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the date range dialog.
 */
function openDateRangeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DateRangeDialog')
    .setTitle('Forecast Date Range')
    .setWidth(340)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Forecast Date Range');
}

/**
 * Returns the current forecast start and end dates as "yyyy-mm-dd" strings
 * for the date range dialog.
 * @return {Object} {startDate, endDate}
 */
function getDateRangeForDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { startDate: '', endDate: '' };

  var startRange = ss.getRangeByName('GLOBAL__FORECAST_START_DATE');
  if (startRange) {
    var v = startRange.getValue();
    if (v instanceof Date && !isNaN(v.getTime())) {
      result.startDate = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }
  if (!result.startDate) {
    // Default to today
    result.startDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var endRange = ss.getRangeByName('GLOBAL__FORECAST_END_DATE');
  if (endRange) {
    var v = endRange.getValue();
    if (v instanceof Date && !isNaN(v.getTime())) {
      result.endDate = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }
  if (!result.endDate) {
    result.endDate = Utilities.formatDate(END_DATE, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return result;
}

/**
 * Saves the forecast date range from the dialog and recalculates.
 * @param {string} startDateStr  "yyyy-mm-dd"
 * @param {string} endDateStr    "yyyy-mm-dd"
 */
function saveDateRangeFromDialog(startDateStr, endDateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Parse dates
  var startParts = startDateStr.split('-');
  var startDate = new Date(Number(startParts[0]), Number(startParts[1]) - 1, Number(startParts[2]));
  var endParts = endDateStr.split('-');
  var endDate = new Date(Number(endParts[0]), Number(endParts[1]) - 1, Number(endParts[2]));

  // Write to named ranges (create them if they don't exist yet)
  var startRange = ss.getRangeByName('GLOBAL__FORECAST_START_DATE');
  if (startRange) {
    startRange.setValue(startDate);
  }

  var endRange = ss.getRangeByName('GLOBAL__FORECAST_END_DATE');
  if (endRange) {
    endRange.setValue(endDate);
  }

  SpreadsheetApp.flush();

  // Update the global variables so recalculation uses the new dates
  START_DATE = startDate;
  START_DATE.setHours(0, 0, 0, 0);
  END_DATE = endDate;
  END_DATE.setHours(23, 59, 59, 0);

  // Recalculate everything with the new date range
  recalculateAll();
}

/**
 * Returns the SKU list for the dialog dropdown.
 * Called from client-side JS.
 * @return {Object[]} array of {id, name}
 */
function getSkuList() {
  var skus = getSkus();
  var list = [];
  for (var i = 0; i < skus.length; i++) {
    list.push({ id: skus[i].id, name: skus[i].name });
  }
  return list;
}

/**
 * Reads one SKU's current settings and returns them as a flat object
 * for the dialog form. Dates are returned as ISO strings.
 * @param {string} skuId
 * @return {Object}
 */
function getSkuSettingsForDialog(skuId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p  = namedRangePrefix(skuId);
  var result = {};

  var fields = [
    'FBA_OVERRIDE',
    'INBOUND_WORKING', 'INBOUND_SHIPPED', 'INBOUND_RECEIVING',
    'ONHAND_AVAILABLE', 'ONHAND_FC_TRANSFER',
    'RESERVED_CUSTOMER_ORDER', 'RESERVED_FC_PROCESSING',
    'RESEARCHING',
    'UNFULFILLABLE_WAREHOUSE_DAMAGED', 'UNFULFILLABLE_DEFECTIVE', 'UNFULFILLABLE_EXPIRED',
    'UNFULFILLABLE_CUSTOMER_DAMAGED', 'UNFULFILLABLE_CARRIER_DAMAGED', 'UNFULFILLABLE_DISTRIBUTOR_DAMAGED',
    'FBA_CHECKIN_DELAY', 'FBM_ONHAND',
    'DAILY_VELOCITY',
    'VEL_OVERRIDE_1_START', 'VEL_OVERRIDE_1_END', 'VEL_OVERRIDE_1_VALUE',
    'VEL_OVERRIDE_2_START', 'VEL_OVERRIDE_2_END', 'VEL_OVERRIDE_2_VALUE',
    'VEL_OVERRIDE_3_START', 'VEL_OVERRIDE_3_END', 'VEL_OVERRIDE_3_VALUE',
    'CONVERSION_RATE',
    'CR_OVERRIDE_1_START', 'CR_OVERRIDE_1_END', 'CR_OVERRIDE_1_VALUE',
    'CR_OVERRIDE_2_START', 'CR_OVERRIDE_2_END', 'CR_OVERRIDE_2_VALUE',
    'CR_OVERRIDE_3_START', 'CR_OVERRIDE_3_END', 'CR_OVERRIDE_3_VALUE',
    'SPD_UNITS', 'SPD_SEND_DATE', 'SPD_TRANSIT_DAYS',
    'LTL_UNITS', 'LTL_SEND_DATE', 'LTL_TRANSIT_DAYS',
    'DTC_UNITS', 'DTC_START_DATE', 'DTC_END_DATE',
    'ADHOC_UNITS', 'ADHOC_SEND_DATE', 'ADHOC_TRANSIT_DAYS',
    'SELLING_PRICE', 'DPP_MARGIN', 'PAST_OOS_DAYS'
  ];

  for (var i = 0; i < fields.length; i++) {
    var key = fields[i];
    var val = readNamedRange(ss, p + '__' + key);

    // Convert Date objects to ISO string for the HTML form
    if (val instanceof Date) {
      result[key] = val.toISOString();
    } else {
      result[key] = val;
    }
  }

  return result;
}

/**
 * Saves one SKU's settings from the dialog form back to the Settings tab.
 * Dates arrive as "yyyy-mm-dd" strings or null.
 * @param {string} skuId
 * @param {Object} values  flat object of field key -> value
 */
function saveSkuSettingsFromDialog(skuId, values) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p  = namedRangePrefix(skuId);

  var dateFields = [
    'VEL_OVERRIDE_1_START', 'VEL_OVERRIDE_1_END',
    'VEL_OVERRIDE_2_START', 'VEL_OVERRIDE_2_END',
    'VEL_OVERRIDE_3_START', 'VEL_OVERRIDE_3_END',
    'CR_OVERRIDE_1_START', 'CR_OVERRIDE_1_END',
    'CR_OVERRIDE_2_START', 'CR_OVERRIDE_2_END',
    'CR_OVERRIDE_3_START', 'CR_OVERRIDE_3_END',
    'SPD_SEND_DATE', 'LTL_SEND_DATE',
    'DTC_START_DATE', 'DTC_END_DATE',
    'ADHOC_SEND_DATE'
  ];

  for (var key in values) {
    var rangeName = p + '__' + key;
    var range = ss.getRangeByName(rangeName);
    if (!range) continue;

    var val = values[key];

    if (dateFields.indexOf(key) > -1) {
      // Date field: convert "yyyy-mm-dd" string to Date, or clear
      if (val && val !== '') {
        var parts = val.split('-');
        var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        range.setValue(d);
      } else {
        range.setValue('');
      }
    } else {
      // Numeric field: write number or clear
      if (val !== null && val !== undefined && val !== '') {
        range.setValue(Number(val));
      } else {
        range.setValue('');
      }
    }
  }

  SpreadsheetApp.flush();
}

/**
 * First-time setup: initializes the registry, builds the Settings tab
 * with defaults, then runs a full forecast calculation.
 */
function initialSetup() {
  var ui = SpreadsheetApp.getUi();
  var skus = getSkus();

  // Initialize the SKU registry
  initRegistry(skus);

  // Build the Settings tab with all per-SKU input sections
  buildSettingsTab();

  // Run the full forecast
  recalculateAll();

  // Install the onEdit trigger
  installTrigger();

  ui.alert(
    'Setup Complete',
    'Settings tab created with defaults for ' + skus.length + ' SKUs.\n' +
    'Detail tabs and Summary tab generated.\n' +
    'Auto-refresh trigger installed.\n\n' +
    'Use "Inventory Forecast > Edit SKU Settings" to enter your data.',
    ui.ButtonSet.OK
  );
}

/**
 * Rebuilds the Gorilla Data tab and re-links Settings cells.
 * Use after entering or changing Gorilla ROI Seller ID.
 */
function refreshGorillaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if Gorilla is configured
  var sellerRange = ss.getRangeByName('GLOBAL__GORILLA_SELLER_ID');
  if (!sellerRange || !sellerRange.getValue()) {
    SpreadsheetApp.getUi().alert(
      'Gorilla ROI Not Configured',
      'Enter your Gorilla Seller ID in the Settings tab first, then run this again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  ss.toast('Building Gorilla Data tab...', 'Gorilla ROI', 5);

  // Build the Gorilla Data formula sheet
  buildGorillaDataTab();

  // Rebuild Settings tab to link cells to Gorilla Data
  buildSettingsTab();

  // Recalculate with the new data
  recalculateAll();

  ss.toast(
    'Gorilla ROI data linked! Values will populate as Gorilla ROI refreshes.',
    'Done', 5
  );
}

/**
 * Full recalculation: reads all SKU settings, runs the waterfall engine
 * for each SKU, renders all detail tabs, and rebuilds the Summary tab.
 */
function recalculateAll() {
  // Nudge: if Gorilla is configured but Gorilla Data tab doesn't exist yet
  try {
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var gorillaRange = ss2.getRangeByName('GLOBAL__GORILLA_SELLER_ID');
    if (gorillaRange) {
      var gId = gorillaRange.getValue();
      if (gId && gId !== '' && !ss2.getSheetByName(GORILLA_DATA_TAB_NAME)) {
        ss2.toast('Gorilla Seller ID detected! Run "Inventory Forecast > Refresh Gorilla Data" to connect.', 'Gorilla ROI', 10);
      }
    }
  } catch(e) { /* Gorilla config not set up yet */ }

  var skus = getSkus();
  var allResults = [];

  for (var i = 0; i < skus.length; i++) {
    var skuDef    = skus[i];
    var cfg       = readSkuSettings(skuDef.id);
    var snapshots = runWaterfall(cfg);
    var milestones = extractMilestones(snapshots, cfg);

    // Build the detail tab for this SKU
    buildDetailTab(skuDef, snapshots, cfg);

    allResults.push({
      skuDef:     skuDef,
      cfg:        cfg,
      snapshots:  snapshots,
      milestones: milestones
    });
  }

  // Build the Summary tab with all results
  buildSummaryTab(allResults);

  // Move Summary tab to position 2 (after Settings)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) {
    ss.setActiveSheet(summarySheet);
    ss.moveActiveSheet(2);
  }

  // Ensure Settings is first
  var settingsSheet = ss.getSheetByName(SETTINGS_TAB_NAME);
  if (settingsSheet) {
    ss.setActiveSheet(settingsSheet);
    ss.moveActiveSheet(1);
  }

  SpreadsheetApp.flush();
}

/**
 * Single-SKU recalculation: saves and recalculates only the specified
 * SKU's detail tab. Does NOT rebuild the Summary tab.
 * Called from the dialog's "Save This SKU" button.
 *
 * @param {string} skuId
 */
function recalculateSku(skuId) {
  var skus = getSkus();
  var skuDef = null;
  for (var i = 0; i < skus.length; i++) {
    if (skus[i].id === skuId) {
      skuDef = skus[i];
      break;
    }
  }
  if (!skuDef) return;

  var cfg       = readSkuSettings(skuId);
  var snapshots = runWaterfall(cfg);

  // Rebuild only this SKU's detail tab
  buildDetailTab(skuDef, snapshots, cfg);

  SpreadsheetApp.flush();
}

/**
 * Adds a new SKU to the registry and rebuilds the Settings tab.
 * Returns an error string if the SKU can't be added, or null on success.
 *
 * @param {string} skuId    e.g. "SB-NEW-100W-FBA"
 * @param {string} skuName  e.g. "New Product Shampoo Bar"
 * @return {string|null} error message or null
 */
function addNewSku(skuId, skuName) {
  var skus = getSkus();

  // Enforce max limit
  if (skus.length >= MAX_SKUS) {
    return 'Maximum of ' + MAX_SKUS + ' SKUs reached. Google Sheets performance ' +
           'degrades with too many tabs and named ranges. Please remove an existing ' +
           'SKU or create a new spreadsheet for additional SKUs.';
  }

  // Check for duplicate
  for (var i = 0; i < skus.length; i++) {
    if (skus[i].id === skuId) {
      return 'SKU "' + skuId + '" already exists.';
    }
  }

  // Validate input
  if (!skuId || skuId.trim() === '') return 'SKU ID cannot be empty.';
  if (!skuName || skuName.trim() === '') return 'SKU name cannot be empty.';

  // Generate tab name (truncate if needed to fit Sheets' 100-char limit)
  var tabName = 'SKU - ' + skuId;
  if (tabName.length > 100) tabName = tabName.substring(0, 100);

  var newSku = { id: skuId.trim(), name: skuName.trim(), tab: tabName };
  skus.push(newSku);

  // Update the registry
  initRegistry(skus);

  // Rebuild the Settings tab so named ranges exist for the new SKU
  buildSettingsTab();

  return null; // success
}

/**
 * Removes an existing SKU from the registry, deletes its detail tab,
 * and rebuilds the Settings tab.
 *
 * @param {string} skuId
 * @return {string|null} error message or null
 */
function removeExistingSku(skuId) {
  var skus = getSkus();

  // Find the SKU
  var idx = -1;
  var skuDef = null;
  for (var i = 0; i < skus.length; i++) {
    if (skus[i].id === skuId) {
      idx = i;
      skuDef = skus[i];
      break;
    }
  }
  if (idx === -1) return 'SKU "' + skuId + '" not found.';

  // Don't allow removal of the last SKU
  if (skus.length <= 1) {
    return 'Cannot remove the last SKU. At least one SKU must remain.';
  }

  // Remove from array
  skus.splice(idx, 1);

  // Update registry
  initRegistry(skus);

  // Delete the SKU's detail tab
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(skuDef.tab);
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  // Remove named ranges for this SKU
  var prefix = namedRangePrefix(skuId);
  var existingRanges = ss.getNamedRanges();
  for (var r = 0; r < existingRanges.length; r++) {
    if (existingRanges[r].getName().indexOf(prefix) === 0) {
      existingRanges[r].remove();
    }
  }

  // Rebuild Settings tab (without the removed SKU)
  buildSettingsTab();

  return null; // success
}

/**
 * Installable onEdit trigger handler.
 * Fires on every edit; only recalculates if the edit was on the Settings tab.
 *
 * @param {Object} e  Event object from the trigger
 */
function onEditTrigger(e) {
  // Guard: only recalculate if the edit happened on the Settings tab
  if (!e || !e.range) return;

  var editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== SETTINGS_TAB_NAME) return;

  // Only recalculate if the edit was in the values column (column B).
  if (e.range.getColumn() !== SETTINGS_VALUE_COL) return;

  // Debounce: use a lock to prevent overlapping recalculations
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(2000);
  if (!acquired) return;

  try {
    recalculateAll();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Creates an installable onEdit trigger for the active spreadsheet.
 */
function installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Utility: Deletes all SKU detail tabs and the Summary tab.
 */
function deleteGeneratedTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skus = getSkus();

  for (var i = 0; i < skus.length; i++) {
    var sheet = ss.getSheetByName(skus[i].tab);
    if (sheet) ss.deleteSheet(sheet);
  }

  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) ss.deleteSheet(summarySheet);
}
