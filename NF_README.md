# NF Extension: Drive-based Bulk Imports & Inventory Reconciliation

This extension adds Drive-based bulk imports, inventory reconciliation, and backlog processing capabilities to the Shipment Tracking Toolkit, enabling end-to-end automation without overloading Apps Script quotas.

## New Features

### 🗂️ Drive Folder Auto-Import
- Automatically scans Google Drive folder for latest files by type
- Supports both CSV and XLSX files with Advanced Drive Service conversion
- Case-insensitive filename pattern matching with whitespace tolerance
- Imports to designated sheets based on file content type

### 📊 Inventory Reconciliation
- Cross-references ERP quants, 3PL warehouse balances, and Power BI data
- Builds aggregated views and identifies discrepancies
- Prefers Drive folder imports over OneDrive URLs
- Creates reconciliation reports for inventory management

### 🔄 Bulk Processing & Backfill
- Complete end-to-end daily workflow orchestration
- Batch tracking refresh with conservative limits to avoid quotas
- Historical data backfill and duplicate detection across sources
- Idempotent operations safe to rerun

### 📅 Enhanced Scheduling
- **Daily 00:01**: Inventory/balances update via Drive imports and reconcile
- **Daily 11:00**: Gmail import and tracking refresh, keep weekly SOK/KRK reports current
- **Weekly Mon 02:00**: Build comprehensive weekly SOK/KRK reports

## Installation

### 1. Script Properties Setup
Add the following required Script Property:

```
DRIVE_IMPORT_FOLDER_ID = 1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh
```

**To set Script Properties:**
1. Open Apps Script editor
2. Go to Project Settings (gear icon)
3. Scroll to Script Properties section
4. Add the property above

### 2. Enable Advanced Services
Enable the **Google Drive API** advanced service:
1. In Apps Script editor, go to Services (+ icon)
2. Add "Drive API" service
3. This enables XLSX to Google Sheets conversion

### 3. Deploy Files
All NF_ files are additive and safe to deploy alongside existing code:
- `NF_Drive_Imports.gs` - Drive folder scanning and import automation
- `NF_Bulk_Backfill.gs` - Bulk processing orchestration
- `NF_Inventory_Balance.gs` - Inventory reconciliation logic
- `NF_SOK_KRK_Weekly.gs` - Weekly report management
- `NF_Menu_Triggers.gs` - Menu additions and scheduling

## File Type Recognition

The system auto-detects file types based on filename patterns (case-insensitive):

| File Type | Filename Contains | Target Sheet |
|-----------|-------------------|--------------|
| **ERP Quants** | `quants` | `Import_Quants` |
| **Warehouse Balance** | `warehouse balance` \| `warehouse_balance` \| `3pl balance` | `Import_Warehouse_Balance` |
| **ERP Stock Picking** | `stock picking` \| `erp picking` | `Import_ERP_StockPicking` |
| **PBI Deliveries** | `deliveries` \| `pbi deliveries` \| `pbi_shipment` \| `pbi_outbound` | `Import_Weekly` |
| **PBI Balances** | `pbi balance` \| `pbi stock` \| `pbi inventory` | `Import_PBI_Balance` |

### Example Filenames
✅ **Recognized:**
- `ERP_Quants_2025-01-15.xlsx`
- `Warehouse Balance Report.csv`
- `PBI_Deliveries_Weekly.xlsx`
- `Stock Picking Export (12).csv`

❌ **Not Recognized:**
- `Daily_Report.xlsx` (no matching pattern)
- `inventory.txt` (unsupported format)

## New Menu Items

### Bulk Operations Menu
- **Import from Drive + Rebuild All** → `NF_BulkRebuildAll()` - Complete daily workflow
- **Refresh All Pending (100)** → `NF_RefreshAllPending()` - Batch tracking refresh
- **Find Duplicates (All Sources)** → `NF_BulkFindDuplicates_All()` - Cross-source duplicate detection
- **Import from Drive Only** → `NF_BulkImportFromDrive()` - Drive import without processing
- **Update Inventory Balances** → `NF_UpdateInventoryBalances()` - Inventory reconciliation only
- **Build SOK/Kärkkäinen Always** → `NF_buildSokKarkkainenAlways()` - Historical report merge

### NF Scheduling Menu
- **Setup Daily 00:01 (Inventory)** - Schedule inventory updates
- **Setup Daily 11:00 (Tracking)** - Schedule tracking refresh + reports
- **Setup Weekly Mon 02:00 (Reports)** - Schedule weekly report builds
- **Clear NF Triggers** - Remove all NF-related triggers
- **List Active Triggers** - View currently scheduled triggers

## Key Functions

### Drive Import Functions
```javascript
NF_Drive_ImportLatestAll()                    // Import all file types from Drive folder
NF_Drive_PickLatestFileByPattern_(folderId, patterns)  // Find most recent matching file
NF_Drive_ReadCsvOrXlsxToSheet_(file, sheetName)       // Read file data to sheet
```

### Bulk Processing Functions
```javascript
NF_BulkRebuildAll()                          // Complete daily workflow sequence
NF_RefreshAllPending(limitPerRun)            // Batch refresh tracking statuses
NF_BulkFindDuplicates_All()                  // Cross-source duplicate detection
```

### Inventory Functions
```javascript
NF_UpdateInventoryBalances()                 // Main inventory reconciliation
NF_BuildQuantsAggregate()                    // Aggregate quants by article
NF_ReconcileInventory()                      // Compare ERP vs 3PL vs PBI
```

### Weekly Report Functions
```javascript
NF_buildSokKarkkainenAlways()                // Historical SOK/Kärkkäinen merge
NF_ReconcileWeeklyFromImport()               // Append missing deliveries
NF_getLastFinishedWeekSunWindow_()           // Get last Sunday-Sunday week
```

## Data Flow & Processing

### Daily 00:01 - Inventory Update
1. Scan Drive folder for latest inventory files
2. Import Quants and Warehouse Balance data
3. Build aggregated views and reconciliation reports
4. Update inventory balance sheets

### Daily 11:00 - Tracking & Reports  
1. Run Gmail import and tracking refresh
2. Build/update SOK and Kärkkäinen reports with all historical data
3. Reconcile weekly reports with any new PBI deliveries
4. Batch refresh pending tracking statuses (limited)

### Weekly Monday 02:00 - Weekly Reports
1. Build comprehensive weekly SOK/Kärkkäinen reports
2. Process full historical data merge
3. Generate summary statistics

## Quota Management & Throttling

### Conservative Limits
- **Batch Size**: Max 100 tracking calls per bulk refresh
- **Rate Limiting**: 500ms sleep between API calls
- **Error Handling**: Automatic backoff on rate limits (429 errors)
- **Idempotent**: Safe to rerun operations without duplicating data

### Monitoring
- All operations log to console with detailed progress
- Error logging to `Error_Log` sheet for trigger failures
- Rate limit detection with automatic retry delays

## Data Sources & Truth Hierarchy

### Source Priority (nShift = ground truth)
1. **nShift (Gmail)** - Delivered timestamps and tracking events (never overwritten)
2. **ERP Systems** - Stock picking, quants/balances
3. **3PL Systems** - Warehouse balance verification
4. **Power BI** - Deliveries and balances (reconciled, not overwriting)

### Reconciliation Logic
- ERP vs 3PL vs PBI quantity comparisons
- Duplicate detection across Reference + Article + DeliveryPlace
- Historical data preservation with source tracking

## Troubleshooting

### Common Issues

**"DRIVE_IMPORT_FOLDER_ID not configured"**
- Add the Script Property with the correct folder ID
- Verify folder access permissions

**"Required columns not found"**
- Check file headers match expected patterns
- Files need Article, Location, Quantity columns for inventory
- Case-insensitive matching is enabled

**"Failed to convert XLSX file"**
- Ensure Drive API advanced service is enabled
- Check file format and content validity

**Rate Limiting (429 errors)**
- System automatically handles with backoff
- Reduce batch sizes if persistent issues occur

### Debug Mode
Enable console logging to monitor operations:
```javascript
// All NF functions include detailed console.log statements
// Check Apps Script execution log for detailed progress
```

## Migration Notes

### From OneDrive to Drive Folder
- Drive folder import takes precedence when `DRIVE_IMPORT_FOLDER_ID` is configured
- OneDrive URLs remain as fallback option
- Existing data and sheets are preserved

### Compatibility
- All NF_ functions are additive and non-breaking
- Existing menu items and triggers are preserved
- Legacy function calls remain functional

## Support

For issues or questions:
1. Check Apps Script execution logs for detailed error messages
2. Verify Script Properties configuration
3. Ensure Drive folder contains properly named files
4. Confirm Advanced Drive Service is enabled

---

**Version**: NF Extension 1.0  
**Compatibility**: Apps Script with Drive API advanced service  
**Requirements**: Google Drive folder access, proper Script Properties setup
