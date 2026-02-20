/**
 * SettingsTab.gs
 * ---------------------------------------------------------------------------
 * Creates and formats the Settings tab with per-SKU input sections.
 * Each SKU gets its own clearly labeled block of input rows.
 * All values are stored in column B; labels are in column A.
 * Named ranges are created for every input so other modules can read them
 * by name (e.g., "SB_HW_100W_FBA__FBA_AVAILABLE").
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
  // ── Starting Inventory ──
  { key: 'FBA_AVAILABLE',         label: 'FBA',                                          defaultVal: 0,    format: 'number'  },
  { key: 'FBA_PROCESSING',        label: 'Inbound',                                      defaultVal: 0,    format: 'number'  },
  { key: 'FBA_CHECKIN_DELAY',     label: 'FBA check-in and processing delay (days)',      defaultVal: 7,    format: 'number'  },
  { key: 'FBM_ONHAND',            label: 'On-hand',                                      defaultVal: 0,    format: 'number'  },
  { key: 'RESERVED',              label: 'Reserved',                                     defaultVal: 0,    format: 'number'  },
  { key: 'RESEARCHING',           label: 'Researching',                                  defaultVal: 0,    format: 'number'  },
  { key: 'UNFULFILLABLE',         label: 'Unfulfillable',                                defaultVal: 0,    format: 'number'  },
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
  { key: 'ADHOC_TRANSIT_DAYS',    label: 'Ad-hoc shipment: transit time (days)',            defaultVal: 5,    format: 'number'  }
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
 * Clears existing content, writes all SKU blocks, applies formatting,
 * and creates named ranges for every input cell.
 */
function buildSettingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    for (var s = 0; s < SKUS.length; s++) {
      if (rName.indexOf(namedRangePrefix(SKUS[s].id)) === 0) {
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
  for (var i = 0; i < SKUS.length; i++) {
    var sku = SKUS[i];
    var prefix = namedRangePrefix(sku.id);

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

      // Default value in column B
      var valueCell = sheet.getRange(row, SETTINGS_VALUE_COL);
      if (input.defaultVal !== '' && input.defaultVal !== null) {
        valueCell.setValue(input.defaultVal);
      }

      // Apply formatting
      if (input.format === 'date') {
        valueCell.setNumberFormat('m/d/yyyy');
      } else if (input.format === 'percent') {
        valueCell.setNumberFormat('0.0"%"');
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

  // Freeze the title rows and protect structure
  sheet.setFrozenRows(2);

  // Add data validation hint: conversion rate 0-100
  // (Applied via named ranges later if needed)

  SpreadsheetApp.flush();
}
