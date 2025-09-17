# NewFlow Apps Script - Installation and Usage Guide

## Overview

NewFlow is an additive Apps Script implementation that provides:
- **Country-level delivery time tracking** from "job done" → "delivered"
- **SOK & Kärkkäinen weekly reporting** with automated scheduling
- **Daily automation** for data imports and metric computation
- **Safe integration** with existing scripts using unique `NF_` prefixes

## Installation

### 1. Copy Script Files

Add these files to your Google Apps Script project:
- `NF_Main.gs` - Menu system and orchestration
- `NF_LeadtimeAndWeekly.gs` - Core analytics and reporting

### 2. Required Script Properties

Go to **Project Settings → Script Properties** and add:

#### Required Properties:
```
SOK_FREIGHT_ACCOUNT = 990719901
KARKKAINEN_NUMBERS = 615471,802669,7030057
```

#### Optional Properties:
```
GMAIL_QUERY_PACKAGES = subject:"package report" OR subject:"Packages Report" newer_than:60d
PBI_FOLDER_ID = <your_google_drive_folder_id>
ERP_FOLDER_ID = <your_erp_drive_folder_id>
ERP_FILE_NAME = <specific_erp_file_pattern>
ATTACH_ALLOW_REGEX = \.(csv|xlsx|xls)$
```

### 3. Enable Advanced Services (if using Drive imports)

1. Go to **Services** in Apps Script editor
2. Add **Drive API** (for XLSX → Sheets conversion)
3. Enable the service

### 4. Add NewFlow Menu

Add this line to your existing `onOpen()` function, or create a new one:

```javascript
function onOpen() {
  // Your existing menu items...
  
  // Add NewFlow menu
  NF_onOpen();
}
```

## Usage

### Menus

After installation, you'll see a **NewFlow** menu with these options:

#### Data Import:
- **Gmail: tuo Packages (nShift) → Rebuild** - Import latest nShift packages from Gmail
- **PBI: tuo Outbound (Drive-kansiosta/staging)** - Import Power BI outbound data
- **ERP: tuo Stock Picking** - Import ERP stock picking data

#### Analytics:
- **Rakenna Maa-kohtainen toimitusaika** - Build country-level leadtime analysis
- **Rakenna SOK & Kärkkäinen -viikkoraportit** - Generate weekly reports

#### Automation:
- **Ajasta päivit. (arkipäivisin 12:00)** - Schedule daily automation at 12:00
- **Ajasta viikko (ma 02:00)** - Schedule weekly reports on Monday at 02:00
- **Poista kaikki NewFlow-ajastukset** - Remove all NewFlow triggers

#### Manual Execution:
- **Aja: Päivittäinen NewFlow** - Run daily import and analysis
- **Aja: Viikkoraportit** - Run weekly SOK & Kärkkäinen reports

### Generated Sheets

NewFlow creates these sheets automatically:

#### Leadtime Analysis:
- **NF_Leadtime_Detail** - Detailed country-level leadtime analysis
- **NF_Leadtime_Weekly_Country** - Weekly KPI by country and carrier

#### Weekly Reports:
- **NF_Report_SOK** - SOK weekly report (Sun→Sun window)
- **NF_Report_Karkkainen** - Kärkkäinen weekly report (Sun→Sun window)

## Data Source Priority

### Job Done Timestamp Priority:
1. **ERP Stock Picking** (Pick Finish, Completed, Picking Completed)
2. **Gmail Packages** (Pick up date, Submitted date)
3. **PowerBI Outbound** (Created, Dispatch date, Shipped date)

### Delivered Timestamp:
- Uses any available: Delivered Time, Delivered At, RefreshTime
- Optionally can enhance missing data using existing tracking functions

### Country Detection:
- Looks for: Country, Dest Country, Destination Country

## Automation Schedule

### Daily Automation (12:00):
1. Import Gmail packages (nShift)
2. Import PowerBI outbound data
3. Import ERP stock picking data
4. Rebuild country leadtime analysis

### Weekly Automation (Monday 02:00):
1. Generate SOK weekly report (last finished Sun→Sun window)
2. Generate Kärkkäinen weekly report (last finished Sun→Sun window)

## Compatibility

NewFlow is designed to be **additive only**:
- ✅ **Safe**: Uses `NF_` prefixes to avoid conflicts
- ✅ **Compatible**: Calls existing functions as wrappers
- ✅ **Non-breaking**: Doesn't modify existing code
- ✅ **Parallel**: Can run alongside existing automation

## Troubleshooting

### Common Issues:

1. **"Gmail-tuonti ei ole käytettävissä"**
   - Ensure `fetchAndRebuild()` or `gmailImportLatestPackagesReport()` functions exist
   - Check Gmail query properties

2. **"Power BI Outbound -tuonti ei ole käytettävissä"**
   - Ensure `pbiImportOutbounds_OldestFirst()` function exists
   - Set `PBI_FOLDER_ID` property

3. **"ERP Stock Picking -tuonti ei ole konfiguroitu"**
   - Ensure `runERPUpdate()` function exists or configure Drive-based import
   - Set `ERP_FOLDER_ID` and `ERP_FILE_NAME` properties

4. **No data in analysis**
   - Check that source sheets (Packages, PowerBI staging, ERP) contain data
   - Verify column headers match expected patterns

### Debug Information:

Check **Apps Script → Executions** for detailed error logs. NewFlow logs errors with context:
- `NF Gmail import failed: ...`
- `NF PBI import failed: ...`
- `NF ERP import failed: ...`
- `NF leadtime build failed: ...`
- `NF weekly failed: ...`

## Advanced Configuration

### Custom Column Mapping:

If your data uses different column names, you can modify the search patterns in `NF_LeadtimeAndWeekly.gs`:

```javascript
// Example: Add custom tracking code patterns
function NF_extractTrackingCode_(row, map) {
  const candidates = [
    'TrackingNumber','Package Number','PackageNumber',
    'YourCustomTrackingField'  // Add your field here
  ];
  // ...
}
```

### Custom Date Patterns:

Modify the `NF_extractJobDone_()` function to handle your specific date column names.

## Support

For issues or questions:
1. Check the execution log in Apps Script
2. Verify Script Properties are set correctly
3. Ensure source sheets contain expected data
4. Review the generated sheets for partial results

Remember: NewFlow is designed to fail gracefully - if one data source is unavailable, it continues with available sources.