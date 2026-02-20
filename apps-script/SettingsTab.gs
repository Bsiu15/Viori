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
  { key: 'RESERVED_CUSTOMER_ORDER',     label: 'Reserved: Customer order',                      defaultVal: 0,    format: 'number'  },
  { key: 'RESERVED_FC_PROCESSING',      label: 'Reserved: FC processing',                       defaultVal: 0,    format: 'number'  },
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
  { key: 'VEL_OVERRIDE_1_VALUE',  label: 'Velocity override 1: units/day',               defaultVal: '',   format: 'number'  },
  { key: 'VEL_OVERRIDE_2_START',  label: 'Velocity override 2: start date',              defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_2_END',    label: 'Velocity override 2: end date',                defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_2_VALUE',  label: 'Velocity override 2: units/day',               defaultVal: '',   format: 'number'  },
  { key: 'VEL_OVERRIDE_3_START',  label: 'Velocity override 3: start date',              defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_3_END',    label: 'Velocity override 3: end date',                defaultVal: '',   format: 'date'    },
  { key: 'VEL_OVERRIDE_3_VALUE',  label: 'Velocity override 3: units/day',               defaultVal: '',   format: 'number'  },
  // ── Conversion Rate ──
  { key: 'CONVERSION_RATE',       label: 'Conversion rate (%)',                           defaultVal: 100,  format: 'percent' },
  { key: 'CR_OVERRIDE_1_START',   label: 'CR override 1: start date',                    defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_1_END',     label: 'CR override 1: end date',                      defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_1_VALUE',   label: 'CR override 1: value (%)',                     defaultVal: '',   format: 'percent' },
  { key: 'CR_OVERRIDE_2_START',   label: 'CR override 2: start date',                    defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_2_END',     label: 'CR override 2: end date',                      defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_2_VALUE',   label: 'CR override 2: value (%)',                     defaultVal: '',   format: 'percent' },
  { key: 'CR_OVERRIDE_3_START',   label: 'CR override 3: start date',                    defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_3_END',     label: 'CR override 3: end date',                      defaultVal: '',   format: 'date'    },
  { key: 'CR_OVERRIDE_3_VALUE',   label: 'CR override 3: value (%)',                     defaultVal: '',   format: 'percent' },
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
  { key: 'DPP_MARGIN',            label: 'DPP margin (%)',                                   defaultVal: '',   format: 'percent' }
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

  // Remove existing named ranges that belong to our SKU inputs
  var existingRanges = ss.getNamedRanges();
  for (var r = 0; r < existingRanges.length; r++) {
    var rName = existingRanges[r].getName();
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

      // Restore saved value if it exists, otherwise use default
      var valueCell = sheet.getRange(row, SETTINGS_VALUE_COL);
      if (skuSaved.hasOwnProperty(input.key)) {
        var saved = skuSaved[input.key];
        if (saved && typeof saved === 'object' && saved.__isFormula) {
          valueCell.setFormula(saved.formula);
        } else {
          valueCell.setValue(saved);
        }
      } else if (input.defaultVal !== '' && input.defaultVal !== null) {
        valueCell.setValue(input.defaultVal);
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
  // Formula: = ONHAND_AVAILABLE + ONHAND_FC_TRANSFER
  // If the user types a number, it replaces the formula (manual override).
  for (var fi = 0; fi < skus.length; fi++) {
    var fbaPrefix = namedRangePrefix(skus[fi].id);
    var fbaSaved = savedValues[fbaPrefix] || {};
    var fbaRange = ss.getRangeByName(fbaPrefix + '__FBA_OVERRIDE');
    if (fbaRange && !fbaSaved.hasOwnProperty('FBA_OVERRIDE')) {
      fbaRange.setFormula(
        '=' + fbaPrefix + '__ONHAND_AVAILABLE + ' + fbaPrefix + '__ONHAND_FC_TRANSFER'
      );
    }
  }

  // Freeze the title rows and protect structure
  sheet.setFrozenRows(2);

  SpreadsheetApp.flush();
}
