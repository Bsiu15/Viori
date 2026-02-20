/**
 * Main.gs
 * ---------------------------------------------------------------------------
 * Orchestration layer — ties together all modules.
 *
 * Entry points:
 *   onOpen()          — Adds a custom menu to the spreadsheet
 *   onEditTrigger(e)  — Installable onEdit trigger; recalculates when
 *                        Settings tab values change
 *   initialSetup()    — First-time setup: builds Settings tab, runs forecast
 *   recalculateAll()  — Full recalculation of all SKU detail tabs + Summary
 *   installTrigger()  — Creates the installable onEdit trigger
 * ---------------------------------------------------------------------------
 */

/**
 * Runs when the spreadsheet is opened. Adds a custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Forecast')
    .addItem('Initial Setup (first time)', 'initialSetup')
    .addItem('Recalculate All', 'recalculateAll')
    .addItem('Install Auto-Refresh Trigger', 'installTrigger')
    .addToUi();
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
    'The Settings tab has been created with default values for all 6 SKUs.\n\n' +
    'Detail tabs and the Summary tab have been generated.\n\n' +
    'An onEdit trigger has been installed — any change to the Settings tab ' +
    'will automatically recalculate all forecasts.\n\n' +
    'Enter your actual inventory numbers in the Settings tab to begin forecasting.',
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

  // Use e.range.getSheet() which is authoritative for the edited cell,
  // unlike e.source.getActiveSheet() which can be unreliable if the user
  // switches tabs quickly after editing.
  var editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== SETTINGS_TAB_NAME) return;

  // Only recalculate if the edit was in the values column (column B).
  // Edits to column A (labels) should not trigger a recalculation.
  if (e.range.getColumn() !== SETTINGS_VALUE_COL) return;

  // Debounce: use a lock to prevent overlapping recalculations
  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(2000); // wait up to 2 seconds
  if (!acquired) return; // another recalc is already running

  try {
    recalculateAll();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Creates an installable onEdit trigger for the active spreadsheet.
 * Removes any existing triggers with the same handler name first to
 * avoid duplicates.
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

  // Create new installable onEdit trigger
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * Utility: Deletes all SKU detail tabs and the Summary tab.
 * Useful for a clean rebuild during development.
 */
function deleteGeneratedTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete SKU detail tabs
  for (var i = 0; i < SKUS.length; i++) {
    var sheet = ss.getSheetByName(SKUS[i].tab);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  // Delete Summary tab
  var summarySheet = ss.getSheetByName(SUMMARY_TAB_NAME);
  if (summarySheet) {
    ss.deleteSheet(summarySheet);
  }
}
