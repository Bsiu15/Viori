# Viori Inventory Forecasting Calendar

A Google Apps Script-based inventory forecasting tool for Amazon sellers. Forecasts day-by-day fulfillment channel allocation (FBA, FBM, DTC) across multiple SKUs based on available inventory, inbound shipments, and daily sales velocity.

## Setup Instructions

### 1. Create a new Google Sheet
Open Google Sheets and create a new blank spreadsheet.

### 2. Open the Apps Script editor
Go to **Extensions > Apps Script**.

### 3. Copy the script files
Create the following `.gs` files in the Apps Script editor (delete the default `Code.gs` first, or rename it):

| File | Purpose |
|------|---------|
| `Config.gs` | SKU definitions, date range, color palette, utility functions |
| `SettingsTab.gs` | Builds the Settings tab with per-SKU input sections and named ranges |
| `SettingsReader.gs` | Reads all per-SKU inputs from named ranges into structured objects |
| `WaterfallEngine.gs` | Day-by-day waterfall simulation engine |
| `DetailTab.gs` | Renders individual SKU detail tabs |
| `SummaryTab.gs` | Builds the Summary tab with milestone table and calendar grid |
| `Main.gs` | Orchestration, menu, triggers, and entry points |

Copy the contents of each file from the `apps-script/` directory in this repository into the corresponding file in the Apps Script editor.

Also copy `appsscript.json`: In the Apps Script editor, go to **Project Settings** (gear icon) and check **Show "appsscript.json" manifest file in editor**. Then replace its contents with the provided `appsscript.json`.

### 4. Run initial setup
1. Save all files in the Apps Script editor (Ctrl+S)
2. Go back to the spreadsheet
3. Reload the page (the custom menu needs a page load to appear)
4. Click **Inventory Forecast > Initial Setup (first time)**
5. Grant the required permissions when prompted

This will:
- Create the Settings tab with default values for all 6 SKUs
- Generate a detail tab for each SKU
- Build the Summary tab with milestone table and calendar grid
- Install an onEdit trigger for automatic recalculation

### 5. Enter your data
Fill in the Settings tab with actual inventory numbers for each SKU. Every change triggers an automatic recalculation of all tabs.

## Architecture

### Settings Tab
All inputs are per-SKU with no global settings. Each SKU has its own labeled section with:
- Starting FBA available / processing / FBM on-hand units
- FBA check-in delay (days)
- Daily sales velocity with optional date-range override
- Conversion rate with optional date-range override
- SPD shipment: units, send date, transit days
- LTL shipment: units, send date (default 3/13/2026), transit days
- DTC bridge: units, start date, end date
- Ad-hoc shipment: units, send date, transit days

### Waterfall Logic
For each SKU, each day is evaluated sequentially:
1. Process inbound shipment arrivals (add to FBA processing)
2. Clear FBA processing batches past their check-in delay (move to FBA available)
3. Determine effective velocity (base velocity * conversion rate, with overrides)
4. Fulfill in priority order: FBA available > FBM on-hand > DTC bridge > OOS

### Color Coding
| Color | Meaning |
|-------|---------|
| Green | Fulfilling from FBA |
| Yellow | Fulfilling from FBM (FBA is zero) |
| Orange | Fulfilling from DTC bridge inventory |
| Red | TRUE OOS -- no inventory on any channel |
| Blue | Shipment arriving / entering FBA processing |
| Gray | Inventory sitting in FBA processing |

### Forecast Window
February 19, 2026 through April 30, 2026 (71 days).

## SKUs Tracked
1. SB-HW-100W-FBA -- Hidden Waterfall Shampoo Bar
2. SB-CY-100W-FBA -- Citrus Yao Clarifying Shampoo Bar
3. SB-TG-100W-FBA -- Terrace Garden Shampoo Bar
4. BUND-1HWS-1HWC-S-FBA -- Hidden Waterfall Shampoo and Conditioner Bar Set
5. BUND-1CYS-1CYC-S-FBA -- Citrus Yao Shampoo and Conditioner Bar Set
6. BUND-1TGS-1TGC-S-FBA -- Terrace Garden Shampoo and Conditioner Bar Set

## Maintenance
- **Recalculate manually**: Inventory Forecast > Recalculate All
- **Re-install trigger**: Inventory Forecast > Install Auto-Refresh Trigger
- All logic is in Apps Script; sheets contain only rendered output (no formulas)
