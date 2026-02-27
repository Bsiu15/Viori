/**
 * SettingsTab.gs
 * ---------------------------------------------------------------------------
 * Creates and formats the Settings tab with per-SKU input sections.
 * Each SKU gets its own clearly labeled block of input rows.
 * All values are stored in column B; labels are in column A.
 * Named ranges are created for every input so other modules can read them
 * by name (e.g., "SB_HW_100W_FBA__ONHAND_AVAILABLE").
 * ---------------------------------------------------------------------------
 */

/**
 * Master list defining every input row within a SKU block.
 * "key" is appended to the SKU-based named range prefix.
 * "label" is the human-readable text in column A.
 * "default" is the initial value written to column B.
 * "format" is optional: "date" applies date formatting, "percent" applies %.
 */
var SKU_INPUT_ROWS = [
  // ── FBA (auto-calculated or manual override) ──
  { key: 'FBA_OVERRIDE',                label: 'FBA (available for sale)',                       defaultVal: '',   format: 'number'  },
  // ── Starting Inventory: Inbound ──
  { key: 'INBOUND_WORKING',             label: 'Inbound: Working',                              defaultVal: 0,    format: 'number'  },
  { key: 'INBOUND_SHIPPED',             label: 'Inbound: Shipped',                              defaultVal: 0,    format: 'number'  },
  { key: 'INBOUND_RECEIVING',           label: 'Inbound: Receiving',                            defaultVal: 0,    format: 'number'  },
  // ── Starting Inventory: On-hand ──
  { key: 'ONHAND_AVAILABLE',            label: 'On-hand: Available',                             defaultVal: 0,    format: 'number'  },
  { key: 'ONHAND_FC_TRANSFER',          label: 'On-hand: FC transfer',                          defaultVal: 0,    format: 'number'  },
  // ── Starting Inventory: Reserved ──
  { key: 'RESERVED_CUSTOMER_ORDER',     label: 'Reserved: Customer order',                      defaultVal: 0,    format: 'number',
    gorillaWarning: 'LEAVE AT 0 when using Gorilla.\n' +
      'Amazon\'s "Available" already excludes reserved units.\n' +
      'Entering a number here will double-subtract from your forecast.' },
  { key: 'RESERVED_FC_PROCESSING',      label: 'Reserved: FC processing',                       defaultVal: 0,    format: 'number',
    gorillaWarning: 'LEAVE AT 0 when using Gorilla.\n' +
      'Amazon\'s "Available" already excludes reserved units.\n' +
      'FC processing units are not broken out by Gorilla — entering\n' +
      'a value here risks double-counting with the Available number.' },
  // ── Starting Inventory: Researching ──
  { key: 'RESEARCHING',                 label: 'Researching',                                    defaultVal: 0,    format: 'number'  },
  // ── Starting Inventory: Unfulfillable ──
  { key: 'UNFULFILLABLE_WAREHOUSE_DAMAGED', label: 'Unfulfillable: Warehouse damaged',           defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE_DEFECTIVE',         label: 'Unfulfillable: Defective',                   defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE_EXPIRED',           label: 'Unfulfillable: Expired',                     defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE_CUSTOMER_DAMAGED',  label: 'Unfulfillable: Customer damaged',            defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE_CARRIER_DAMAGED',   label: 'Unfulfillable: Carrier damaged',             defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE_DISTRIBUTOR_DAMAGED', label: 'Unfulfillable: Distributor damaged',       defaultVal: 0,    format: 'number'  },
  // ── Forecast config ──
  { key: 'FBA_CHECKIN_DELAY',     label: 'FBA check-in and processing delay (days)',      defaultVal: 7,    format: 'number'  },
  { key: 'FBM_ONHAND',            label: 'FBM on-hand',                                   defaultVal: 0,    format: 'number'  },
  // ── Sales Velocity ──
  { key: 'DAILY_VELOCITY',        label: 'Daily sales velocity (units/day)',              defaultVal: 0,    format: 'number'  },
  { key: 'VEL_OVERRIDE_1_START',  label: 'Velocity override 1: start date',              defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_1_END',    label: 'Velocity override 1: end date',                defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_1_VALUE',  label: 'Velocity override 1: units/day',               defaultVal: '',   format: 'number',  velOverrideNum: 1 },
  { key: 'VEL_OVERRIDE_2_START',  label: 'Velocity override 2: start date',              defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_2_END',    label: 'Velocity override 2: end date',                defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_2_VALUE',  label: 'Velocity override 2: units/day',               defaultVal: '',   format: 'number',  velOverrideNum: 2 },
  { key: 'VEL_OVERRIDE_3_START',  label: 'Velocity override 3: start date',              defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_3_END',    label: 'Velocity override 3: end date',                defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_3_VALUE',  label: 'Velocity override 3: units/day',               defaultVal: '',   format: 'number',  velOverrideNum: 3 },
  // ── Conversion Rate ──
  { key: 'CONVERSION_RATE',       label: 'Conversion rate (%)',                           defaultVal: 100,  format: 'percent' },
  // ── Shipments ──
  { key: 'SPD_UNITS',             label: 'SPD shipment: units',                           defaultVal: 0,    format: 'number'  },
  { key: 'SPD_SEND_DATE',         label: 'SPD shipment: send date',                       defaultVal: '',   format: 'date'    },
  { key: 'SPD_TRANSIT_DAYS',      label: 'SPD shipment: transit time (days)',              defaultVal: 5,    format: 'number'  },
  { key: 'LTL_UNITS',             label: 'LTL shipment: units',                           defaultVal: 0,    format: 'number'  },
  { key: 'LTL_SEND_DATE',         label: 'LTL shipment: send date',                       defaultVal: new Date(2026, 2, 13), format: 'date' },
  { key: 'LTL_TRANSIT_DAYS',      label: 'LTL shipment: transit time (days)',              defaultVal: 14,   format: 'number'  },
  { key: 'DTC_UNITS',             label: 'DTC bridge: units',                              defaultVal: 0,    format: 'number'  },
  { key: 'DTC_START_DATE',        label: 'DTC bridge: start date',                         defaultVal: '',   format: 'date'    },
  { key: 'DTC_END_DATE',          label: 'DTC bridge: end date',                            defaultVal: '',   format: 'date'    },
  { key: 'ADHOC_UNITS',           label: 'Ad-hoc shipment: units',                         defaultVal: 0,    format: 'number'  },
  { key: 'ADHOC_SEND_DATE',       label: 'Ad-hoc shipment: send date',                     defaultVal: '',   format: 'date'    },
  { key: 'ADHOC_TRANSIT_DAYS',    label: 'Ad-hoc shipment: transit time (days)',            defaultVal: 5,    format: 'number'  },
  // ── Financials ──
  { key: 'SELLING_PRICE',         label: 'Selling price ($)',                                defaultVal: '',   format: 'currency' },
  { key: 'DPP_MARGIN',            label: 'DPP margin (%)',                                   defaultVal: '',   format: 'percent' },
  { key: 'PAST_OOS_DAYS',         label: 'Past OOS days (already experienced)',              defaultVal: 0,    format: 'number'  }
];

/**
 * Converts a SKU id into a valid named-range prefix by replacing hyphens with
 * underscores (named ranges cannot contain hyphens).
 * @param {string} skuId  e.g. "SB-HW-100W-FBA"
 * @return {string}       e.g. "SB_HW_100W_FBA"
 */
function namedRangePrefix(skuId) {
  return skuId.replace(/-/g, '_');
}

/**
 * Builds a Gorilla auto-calc formula for a velocity override VALUE field.
 * The formula pulls GORILLA_SALESCOUNT for the same date range shifted back
 * one year and divides by the number of days to produce a daily velocity.
 *
 * Returns "" when the override START/END dates are empty, so the engine
 * skips the override entirely (readNamedRange treats "" as null).
 *
 * @param {string} prefix      Named-range prefix, e.g. "SB_HW_100W_FBA"
 * @param {string} skuId       Original SKU id, e.g. "SB-HW-100W-FBA"
 * @param {number} overrideNum 1, 2, or 3
 * @return {string} A Google Sheets formula string
 */
function buildVelAutoCalcFormula(prefix, skuId, overrideNum) {
  var startRef = prefix + '__VEL_OVERRIDE_' + overrideNum + '_START';
  var endRef   = prefix + '__VEL_OVERRIDE_' + overrideNum + '_END';

  // When both dates are present, pull last year's shipped sales for that range
  // and divide by the number of days to get daily velocity.
  // When dates are empty, return "" so the override is ignored by SettingsReader.
  return '=IF(AND(' + startRef + '<>"", ' + endRef + '<>""), ' +
    'IFERROR(' +
      'GORILLA_SALESCOUNT(GLOBAL__GORILLA_SELLER_ID, "Custom", GLOBAL__GORILLA_MARKETPLACE, "' + skuId + '", ' +
        '"Shipped", "NO", ' +
        'TEXT(' + startRef + ' - 365, "yyyy-mm-dd"), ' +
        'TEXT(' + endRef + ' - 365, "yyyy-mm-dd")) ' +
      '/ (' + endRef + ' - ' + startRef + ' + 1)' +
    ', ""), "")';
}

/**
 * Builds (or rebuilds) the Settings tab from scratch.
 * Preserves existing named range values so data isn't lost during rebuild.
 * Clears existing content, writes all SKU blocks, applies formatting,
 * and creates named ranges for every input cell.
 */
function buildSettingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skus = getSkus();

  // ── Preserve existing values before clearing ──
  var savedValues = {};
  // Preserve global settings
  var savedStartDate = null;
  var globalStartRange = ss.getRangeByName('GLOBAL__FORECAST_START_DATE');
  if (globalStartRange) {
    var gVal = globalStartRange.getValue();
    if (gVal instanceof Date && !isNaN(gVal.getTime())) {
      savedStartDate = gVal;
    }
  }
  var savedEndDate = null;
  var globalEndRange = ss.getRangeByName('GLOBAL__FORECAST_END_DATE');
  if (globalEndRange) {
    var gVal = globalEndRange.getValue();
    if (gVal instanceof Date && !isNaN(gVal.getTime())) {
      savedEndDate = gVal;
    }
  }
  // Preserve Gorilla ROI config
  var savedGorillaSellerId = null;
  var gorillaSellerRange = ss.getRangeByName('GLOBAL__GORILLA_SELLER_ID');
  if (gorillaSellerRange) {
    var gsVal = gorillaSellerRange.getValue();
    if (gsVal !== '' && gsVal !== undefined && gsVal !== null) {
      savedGorillaSellerId = gsVal;
    }
  }
  var savedGorillaMarketplace = null;
  var gorillaMarketRange = ss.getRangeByName('GLOBAL__GORILLA_MARKETPLACE');
  if (gorillaMarketRange) {
    var gmVal = gorillaMarketRange.getValue();
    if (gmVal !== '' && gmVal !== undefined && gmVal !== null) {
      savedGorillaMarketplace = gmVal;
    }
  }
  var savedGorillaLookback = null;
  var gorillaLookbackRange = ss.getRangeByName('GLOBAL__GORILLA_LOOKBACK_DAYS');
  if (gorillaLookbackRange) {
    var glVal = gorillaLookbackRange.getValue();
    if (glVal !== '' && glVal !== undefined && glVal !== null) {
      savedGorillaLookback = glVal;
    }
  }
  for (var sv = 0; sv < skus.length; sv++) {
    var svPrefix = namedRangePrefix(skus[sv].id);
    savedValues[svPrefix] = {};
    for (var sk = 0; sk < SKU_INPUT_ROWS.length; sk++) {
      var svKey = SKU_INPUT_ROWS[sk].key;
      var svRange = ss.getRangeByName(svPrefix + '__' + svKey);
      if (svRange) {
        // Check if the cell has a formula (e.g. FBA auto-calc)
        var svFormula = svRange.getFormula();
        if (svFormula !== '') {
          savedValues[svPrefix][svKey] = { __isFormula: true, formula: svFormula };
        } else {
          var svVal = svRange.getValue();
          if (svVal !== '' && svVal !== undefined && svVal !== null) {
            savedValues[svPrefix][svKey] = svVal;
          }
        }
      }
    }
  }

  // Get or create the Settings sheet
  var sheet = ss.getSheetByName(SETTINGS_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_TAB_NAME, 0);
  }
  sheet.clear();
  sheet.clearFormats();

  // Remove existing named ranges that belong to our SKU inputs or global settings
  var existingRanges = ss.getNamedRanges();
  for (var r = 0; r < existingRanges.length; r++) {
    var rName = existingRanges[r].getName();
    if (rName.indexOf('GLOBAL__') === 0) {
      existingRanges[r].remove();
      continue;
    }
    for (var s = 0; s < skus.length; s++) {
      if (rName.indexOf(namedRangePrefix(skus[s].id)) === 0) {
        existingRanges[r].remove();
        break;
      }
    }
  }

  // ── Column widths ──
  sheet.setColumnWidth(1, 350); // Labels
  sheet.setColumnWidth(2, 200); // Values

  // ── Title row ──
  var row = 1;
  sheet.getRange(row, 1).setValue('Inventory Forecasting Calendar — Settings')
       .setFontSize(14).setFontWeight('bold');
  sheet.getRange(row, 1, 1, 2).setBackground(COLORS.HEADER).setFontColor(COLORS.HEADER_FG);
  row += 1;
  sheet.getRange(row, 1).setValue('All inputs below are per SKU. Change any value to recalculate.')
       .setFontSize(10).setFontStyle('italic').setFontColor('#555555');
  row += 2; // blank spacer

  // ── Global Settings section ──
  sheet.getRange(row, 1, 1, 2).merge()
       .setValue('Global Settings')
       .setBackground(COLORS.SECTION_BG)
       .setFontWeight('bold')
       .setFontSize(11)
       .setBorder(true, true, true, true, false, false);
  row += 1;

  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('Forecast start date')
       .setFontSize(10);
  var startDateCell = sheet.getRange(row, SETTINGS_VALUE_COL);
  if (savedStartDate) {
    startDateCell.setValue(savedStartDate);
  } else {
    startDateCell.setValue(START_DATE);
  }
  startDateCell.setNumberFormat('m/d/yyyy');
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);
  ss.setNamedRange('GLOBAL__FORECAST_START_DATE', startDateCell);
  row += 1;

  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('Forecast end date')
       .setFontSize(10);
  var endDateCell = sheet.getRange(row, SETTINGS_VALUE_COL);
  if (savedEndDate) {
    endDateCell.setValue(savedEndDate);
  } else {
    endDateCell.setValue(END_DATE);
  }
  endDateCell.setNumberFormat('m/d/yyyy');
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, true, true, false, false);
  ss.setNamedRange('GLOBAL__FORECAST_END_DATE', endDateCell);
  row += 2; // spacer

  // ── Gorilla ROI Integration section ──
  sheet.getRange(row, 1, 1, 2).merge()
       .setValue('Gorilla ROI Integration')
       .setBackground(COLORS.SECTION_BG)
       .setFontWeight('bold')
       .setFontSize(11)
       .setBorder(true, true, true, true, false, false);
  row += 1;

  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('Gorilla Seller ID')
       .setFontSize(10);
  var gorillaIdCell = sheet.getRange(row, SETTINGS_VALUE_COL);
  if (savedGorillaSellerId) {
    gorillaIdCell.setValue(savedGorillaSellerId);
  }
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);
  ss.setNamedRange('GLOBAL__GORILLA_SELLER_ID', gorillaIdCell);
  row += 1;

  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('Gorilla Marketplace')
       .setFontSize(10);
  var gorillaMktCell = sheet.getRange(row, SETTINGS_VALUE_COL);
  gorillaMktCell.setValue(savedGorillaMarketplace || 'US');
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);
  ss.setNamedRange('GLOBAL__GORILLA_MARKETPLACE', gorillaMktCell);
  row += 1;

  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('Velocity lookback (days)')
       .setFontSize(10);
  var gorillaLookbackCell = sheet.getRange(row, SETTINGS_VALUE_COL);
  gorillaLookbackCell.setValue(savedGorillaLookback || 30)
                     .setNumberFormat('#,##0');
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);
  ss.setNamedRange('GLOBAL__GORILLA_LOOKBACK_DAYS', gorillaLookbackCell);
  row += 1;

  // Help note
  sheet.getRange(row, SETTINGS_LABEL_COL)
       .setValue('After entering Seller ID, run: Inventory Forecast > Refresh Gorilla Data')
       .setFontSize(9)
       .setFontStyle('italic')
       .setFontColor('#888888');
  sheet.getRange(row, 1, 1, 2).setBorder(null, true, true, true, false, false);
  row += 2; // spacer

  // ── Build each SKU block ──
  for (var i = 0; i < skus.length; i++) {
    var sku = skus[i];
    var prefix = namedRangePrefix(sku.id);
    var skuSaved = savedValues[prefix] || {};

    // SKU header row
    sheet.getRange(row, 1, 1, 2).merge()
         .setValue(sku.id + '  —  ' + sku.name)
         .setBackground(COLORS.SECTION_BG)
         .setFontWeight('bold')
         .setFontSize(11)
         .setBorder(true, true, true, true, false, false);
    row += 1;

    // Input rows
    for (var j = 0; j < SKU_INPUT_ROWS.length; j++) {
      var input = SKU_INPUT_ROWS[j];

      // Label in column A
      sheet.getRange(row, SETTINGS_LABEL_COL)
           .setValue(input.label)
           .setFontSize(10);

      // Restore saved value if it exists, otherwise use default or Gorilla link
      var valueCell = sheet.getRange(row, SETTINGS_VALUE_COL);
      var isGorillaLinked = false;

      if (skuSaved.hasOwnProperty(input.key)) {
        var saved = skuSaved[input.key];
        if (saved && typeof saved === 'object' && saved.__isFormula) {
          valueCell.setFormula(saved.formula);
          // Re-apply Gorilla indicator if it's a Gorilla-linked formula
          if (saved.formula.indexOf(GORILLA_DATA_TAB_NAME) > -1 ||
              saved.formula.indexOf('GORILLA_') > -1) {
            isGorillaLinked = true;
          }
        } else {
          valueCell.setValue(saved);
        }
      } else if (savedGorillaSellerId && GORILLA_LINK_MAP.hasOwnProperty(input.key)) {
        // Auto-link to Gorilla Data tab (no saved value + Gorilla configured + field is linkable)
        var gorillaCol = GORILLA_LINK_MAP[input.key];
        var gorillaRow = i + 2; // SKU index 0 → Gorilla Data row 2
        valueCell.setFormula("=IFERROR('" + GORILLA_DATA_TAB_NAME + "'!" + gorillaCol + gorillaRow + ", 0)");
        isGorillaLinked = true;
      } else if (savedGorillaSellerId && input.velOverrideNum) {
        // Auto-calculate velocity from last year's sales for the override date range
        valueCell.setFormula(buildVelAutoCalcFormula(prefix, sku.id, input.velOverrideNum));
        isGorillaLinked = true;
      } else if (input.defaultVal !== '' && input.defaultVal !== null) {
        valueCell.setValue(input.defaultVal);
      }

      // Visual indicator for Gorilla-linked cells
      if (isGorillaLinked) {
        valueCell.setBackground('#E8F0FE');
        if (input.velOverrideNum) {
          valueCell.setNote(
            'Auto-calculated from Gorilla ROI historical data.\n' +
            'Pulls last year\'s shipped sales for the same date range\n' +
            'and divides by the number of days to get daily velocity.\n\n' +
            'Just enter the start and end dates — this value fills in automatically.\n' +
            'Type a number to manually override.'
          );
        } else {
          valueCell.setNote('Auto-populated from Gorilla ROI. Type a number to override.');
        }
      }

      // Force zero + warning on fields that risk double-counting when Gorilla is active
      if (savedGorillaSellerId && input.gorillaWarning && !isGorillaLinked) {
        valueCell.setValue(0);
        valueCell.setNote(input.gorillaWarning);
        valueCell.setBackground('#FFF3CD'); // amber = caution
      }

      // Apply formatting
      if (input.format === 'date') {
        valueCell.setNumberFormat('m/d/yyyy');
      } else if (input.format === 'percent') {
        valueCell.setNumberFormat('0.0"%"');
      } else if (input.format === 'currency') {
        valueCell.setNumberFormat('$#,##0.00');
      } else {
        valueCell.setNumberFormat('#,##0');
      }

      // Borders
      sheet.getRange(row, 1, 1, 2).setBorder(null, true, null, true, false, false);

      // Create named range: PREFIX__KEY → cell B<row>
      var rangeName = prefix + '__' + input.key;
      ss.setNamedRange(rangeName, valueCell);

      row += 1;
    }

    // Bottom border for block
    sheet.getRange(row - 1, 1, 1, 2).setBorder(null, null, true, null, false, false);

    // Spacer row between SKU blocks
    row += 2;
  }

  // ── Auto-calculate FBA cells ──
  // Set a formula on FBA_OVERRIDE if no manual value was saved.
  // Formula: = ONHAND_AVAILABLE only.
  // FC Transfer and FC Processing are modeled separately in the waterfall
  // engine as delayed batches that clear after the FBA check-in delay.
  // If the user types a number, it replaces the formula (manual override).
  for (var fi = 0; fi < skus.length; fi++) {
    var fbaPrefix = namedRangePrefix(skus[fi].id);
    var fbaSaved = savedValues[fbaPrefix] || {};
    var fbaRange = ss.getRangeByName(fbaPrefix + '__FBA_OVERRIDE');
    if (fbaRange && !fbaSaved.hasOwnProperty('FBA_OVERRIDE')) {
      fbaRange.setFormula(
        '=' + fbaPrefix + '__ONHAND_AVAILABLE'
      );
    }
  }

  // Freeze the title rows and protect structure
  sheet.setFrozenRows(2);

  SpreadsheetApp.flush();
}
