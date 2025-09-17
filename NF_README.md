# New Flow (NF) - Delivery Lead Time Tracking System

This module provides an end-to-end feature set for computing country-specific delivery lead times and building automated weekly reports for SOK and Kärkkäinen.

## Features

### Core Functionality
- **Daily Flow**: Import nShift Packages reports from Gmail → rebuild tables → refresh statuses
- **Weekly Reports**: Build SOK/Kärkkäinen reports with Sun→Sun windows + status refresh
- **Delivery Analytics**: Per-shipment lead times and weekly country KPIs
- **Automated Scheduling**: Weekday 12:00 daily flow + Monday 02:00 weekly reports

### Key Design Principles
- **"Keikka tehty" Start Time**: Pickup date → Submitted date → Created date (PBI fallback)
- **Tracking-Based Delivery**: Delivered times strictly from tracking events, not PBI timestamps
- **Non-Conflicting**: All functions prefixed with `NF_` to avoid symbol clashes
- **Reuses Existing Patterns**: Leverages Gmail import, status refresh, and weekly logic from current repo

## Installation

### Step 1: Copy Files to Apps Script
Copy these new files into your Google Apps Script project:
- `NF_Main.gs`
- `NF_SOK_KRK_Weekly.gs`
- `NF_Leadtime_Weekly.gs`
- `NF_Menu_Triggers.gs`

### Step 2: Install Menu Trigger
1. In Apps Script, run the function `NF_installMenuTrigger` once
2. Refresh your Google Sheets page
3. You should see "New Flow" menu items under "Tracking" menu

### Step 3: Configure Script Properties
Set these required Script Properties (File → Project properties → Script properties):

#### Required Properties
- `SOK_FREIGHT_ACCOUNT`: SOK freight account number (default: `990719901`)
- `KARKKAINEN_NUMBERS`: Comma-separated Kärkkäinen numbers (default: `615471,802669,7030057`)
- `TARGET_SHEET`: Main packages sheet name (default: `Packages`)
- `ARCHIVE_SHEET`: Archive sheet name (default: `Packages_Archive`)

#### Gmail Configuration
- `GMAIL_QUERY_PACKAGES`: Gmail search for nShift reports (default: `label:"Shipment Report" has:attachment (filename:xlsx OR filename:csv)`)
- `ATTACH_ALLOW_REGEX`: Regex for allowed attachment names (default: `(?:^|\\b)(Packages[ _-]?Report)(?:\\b|$)`)

#### Optional Properties
- `ACTION_SHEET`: Pending actions sheet (default: `Vaatii_toimenpiteitä`)
- `PBI_FOLDER_ID`: Power BI Drive folder ID (if using PBI import)
- `PBI_WEBHOOK_URL`: Power BI webhook URL (if using PBI integration)

### Step 4: Enable Required Services
In Apps Script, ensure these advanced services are enabled:
- **Drive API**: Required for XLSX → Sheets conversion
- **Gmail API**: Required for attachment processing

### Step 5: Install Automation Triggers (Optional)
Use the menu to install automated triggers:
- **"Install weekday 12:00"**: Daily flow on weekdays at noon
- **"Install weekly Mon 02:00"**: Weekly reports every Monday at 2 AM

## Usage

### Manual Operations
Access these from the "New Flow" menu:

#### Daily Operations
- **"Run daily flow now"**: Import latest nShift Packages report and rebuild tables
- **"Build SOK & Kärkkäinen (last week)"**: Create weekly reports for last finished week

#### Analytics
- **"Delivery Times list"**: Generate detailed per-shipment lead time analysis
- **"Country Week Leadtime"**: Create weekly KPI by country and carrier

#### Administration
- **"Install weekday 12:00"**: Setup daily automation
- **"Install weekly Mon 02:00"**: Setup weekly automation  
- **"Remove NF triggers"**: Remove all New Flow triggers
- **"Show NF status"**: Check configuration and system status

### Automated Operations
Once triggers are installed:
- **Weekdays at 12:00**: Automatically runs daily flow (import → rebuild → refresh)
- **Mondays at 02:00**: Automatically builds weekly reports and refreshes their statuses

## Data Sources and Logic

### Start Time ("Keikka Tehty") Priority
1. **Pick up date** (from nShift Packages report) - preferred
2. **Submitted date** (from nShift Packages report) - fallback
3. **Created date** (from Power BI report) - last resort

### Delivered Time Sources
**Only from tracking refresh** - never from Power BI timestamps:
1. **"Delivered date (Confirmed)"** - manually confirmed deliveries
2. **"Delivered Time"** - from tracking API responses
3. **"RefreshTime"** - timestamp of last successful tracking refresh

### Weekly Window Definition
- **Sunday → Sunday**: Matches existing weekly report logic
- **"Last finished week"**: Previous complete week (not current partial week)

### SOK/Kärkkäinen Classification
- **SOK**: Matches `SOK_FREIGHT_ACCOUNT` (digits only)
- **Kärkkäinen**: Matches any number in `KARKKAINEN_NUMBERS` (digits only)
- Based on "Payer", "Freight account", or similar fields

## Data Flow

### Daily Flow (`NF_RunDailyFlow`)
1. Search Gmail for latest nShift "Packages Report" attachment
2. Download and convert XLSX/CSV to data matrix
3. Merge with existing Packages/Archive data (unified headers)
4. Optionally refresh statuses on action sheet

### Weekly Reports (`NF_BuildWeeklyReports`)
1. Calculate last finished Sunday→Sunday window
2. Filter TARGET_SHEET data by date window
3. Split by freight payer into SOK/Kärkkäinen
4. Write Report_SOK and Report_Karkkainen sheets with info headers
5. Refresh tracking statuses for both weekly sheets

### Analytics (`NF_BuildDeliveryTimes` + `NF_MakeCountryWeekLeadtime`)
1. Process all shipments with start time and delivered time
2. Calculate lead times in days
3. Create detailed list (Delivery_Times sheet)
4. Pivot by ISO week + country for KPI view (Leadtime_Weekly_Country sheet)

## Output Sheets

### Weekly Reports
- **Report_SOK**: SOK shipments for last week with info header
- **Report_Karkkainen**: Kärkkäinen shipments for last week with info header

### Analytics Sheets  
- **Delivery_Times**: Per-shipment analysis with lead times and source traceability
- **Leadtime_Weekly_Country**: Weekly KPI by country and carrier

### Info Headers Format
Weekly sheets include metadata rows:
```
Row 1: Column headers
Row 2: Week (SUN→SUN): 2025-01-05 - 2025-01-12 | Rows: 150 | Created: 2025-01-13 08:30:15
Row 3: (empty)
Row 4+: Data rows
```

## Compatibility

### Existing Function Reuse
- `refreshStatuses_Sheet`: For status refresh operations
- `ensureRefreshCols_`: For refresh column management  
- `headerIndexMap_`: For header mapping
- `sanitizeMatrix_`: For data cleanup
- `TRK_trackByCarrierEnhanced`: Preferred tracking engine
- `TRK_trackByCarrier`: Fallback tracking engine

### Script Properties Alignment
Uses same property keys as existing codebase:
- `SOK_FREIGHT_ACCOUNT`, `KARKKAINEN_NUMBERS`
- `TARGET_SHEET`, `ARCHIVE_SHEET`
- `GMAIL_QUERY`, `ATTACH_ALLOW_REGEX`

### Safety Features
- All functions prefixed with `NF_` to avoid naming conflicts
- Trigger management only affects `NF_*` functions
- Fallback implementations for missing helper functions
- Non-destructive operations (no deletion of existing data/functions)

## Troubleshooting

### Common Issues

**Menu not appearing**: 
- Run `NF_installMenuTrigger` in Apps Script
- Refresh the Google Sheets page

**"No nShift Packages report found"**:
- Check `GMAIL_QUERY_PACKAGES` property
- Verify Gmail labels and attachment names
- Check `ATTACH_ALLOW_REGEX` pattern

**Empty weekly reports**:
- Verify `SOK_FREIGHT_ACCOUNT` and `KARKKAINEN_NUMBERS` properties
- Check that payer column exists in data
- Ensure date filtering is working (verify date column names)

**Missing delivered times**:
- Run status refresh on source sheets first
- Check tracking engine availability (`TRK_trackByCarrierEnhanced`)
- Verify tracking codes and carrier names

**Trigger failures**:
- Check Script Properties configuration
- Verify required sheets exist (Packages, Packages_Archive)
- Review execution logs in Apps Script

### Status Check
Run **"Show NF status"** from the menu to check:
- Installed triggers
- Script Properties configuration  
- Required sheets existence
- Tracking engine availability

## Support

For issues specific to this New Flow module:
1. Check the status using "Show NF status" menu item
2. Review Script Properties configuration
3. Verify all required files are copied to Apps Script
4. Check execution logs for specific error messages

The module is designed to work alongside existing functionality without conflicts.