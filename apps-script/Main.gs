/**
 * Main.gs
 * ---------------------------------------------------------------------------
 * Orchestration layer — ties together all modules.
 *
 * Entry points:
 *   onOpen()                    — Adds custom menu
 *   openSettingsDialog()        — Opens the SKU settings sidebar
 *   initialSetup()              — First-time setup
 *   recalculateAll()            — Full recalculation
 *   getSkuList()                — Returns SKU list for the dialog
 *   getSkuSettingsForDialog()   — Reads one SKU's settings for the dialog
 *   saveSkuSettingsFromDialog() — Writes one SKU's settings from the dialog
 * ---------------------------------------------------------------------------
 */

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Forecast')
    .addItem('Edit SKU Settings', 'openSettingsDialog')
    .addSeparator()
    .addItem('Recalculate All', 'recalculateAll')
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
 * Returns the SKU list for the dialog dropdown.
 * Called from client-side JS.
 * @return {Object[]} array of {id, name}
 */
function getSkuList() {
  var list = [];
  for (var i = 0; i < SKUS.length; i++) {
    list.push({ id: SKUS[i].id, name: SKUS[i].name });
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
    'FBA_AVAILABLE', 'FBA_PROCESSING', 'FBA_CHECKIN_DELAY', 'FBM_ONHAND',
    'DAILY_VELOCITY', 'VEL_OVERRIDE_START', 'VEL_OVERRIDE_END', 'VEL_OVERRIDE_VALUE',
    'CONVERSION_RATE', 'CR_OVERRIDE_START', 'CR_OVERRIDE_END', 'CR_OVERRIDE_VALUE',
    'SPD_UNITS', 'SPD_SEND_DATE', 'SPD_TRANSIT_DAYS',
    'LTL_UNITS', 'LTL_SEND_DATE', 'LTL_TRANSIT_DAYS',
    'DTC_UNITS', 'DTC_START_DATE', 'DTC_END_DATE',
    'ADHOC_UNITS', 'ADHOC_SEND_DATE', 'ADHOC_TRANSIT_DAYS'
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
    'VEL_OVERRIDE_START', 'VEL_OVERRIDE_END',
    'CR_OVERRIDE_START', 'CR_OVERRIDE_END',
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
 * First-time setup: builds the Settings tab with defaults, then runs
 * a full forecast calculation.
 */
function initialSetup() {
  var ui = SpreadsheetApp.getUi();

  // Build the Settings tab with all per-SKU input sections
  buildSettingsTab();

  // Run the full forecast
  recalculateAll();

  // Install the onEdit trigger
  installTrigger();

  ui.alert(
    'Setup Complete',
    'Settings tab created with defaults for all 6 SKUs.\n' +
    'Detail tabs and Summary tab generated.\n' +
    'Auto-refresh trigger installed.\n\n' +
    'Use "Inventory Forecast > Edit SKU Settings" to enter your data.',
    ui.ButtonSet.OK
  );
}

/**
 * Full recalculation: reads all SKU settings, runs the waterfall engine
 * for each SKU, renders all detail tabs, and rebuilds the Summary tab.
 */
function recalculateAll() {
  var allSettings = readAllSkuSettings();
  var allResults  = [];

  for (var i = 0; i < SKUS.length; i++) {
    var skuDef    = SKUS[i];
    var cfg       = allSettings[i];
    var snapshots = runWaterfall(cfg);
    var milestones = extractMilestones(snapshots, cfg);

    // Build the detail tab for this SKU
    buildDetailTab(skuDef, snapshots);

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

  for (var i = 0; i < SKUS.length; i++) {
    var sheet = ss.getSheetByName(SKUS[i].tab);
    if (sheet) ss.deleteSheet(sheet);
  }

  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) ss.deleteSheet(summarySheet);
}
