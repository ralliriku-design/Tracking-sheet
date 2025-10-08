# Service Level Reporting - Integration Architecture

## Data Flow Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                        Source Data Sheet                         │
│                         (Default: Packages)                       │
│                                                                   │
│  Columns:                                                         │
│  - Date sent / Sent date / Handover date / Submitted date        │
│  - Delivered Time / Delivered At / RefreshTime                   │
│  - Payer / Freight account / Billing / Customer                  │
│  - Created / Timestamp (fallback for window filtering)           │
└─────────────────────────────────────────────────────────────────┘
                                    │
                                    │ Read all data
                                    ▼
┌─────────────────────────────────────────────────────────────────┐
│              NF_buildWeeklyServiceLevels()                        │
│                                                                   │
│  1. Detect columns using configurable hints                      │
│  2. Get last finished Sun→Sun week window                        │
│  3. Filter rows by window date                                   │
│  4. Calculate metrics for each group (ALL, SOK, KARKKAINEN)      │
│  5. Write output to NF_Weekly_ServiceLevels sheet                │
└─────────────────────────────────────────────────────────────────┘
                                    │
                    ┌───────────────┼───────────────┐
                    │               │               │
                    ▼               ▼               ▼
        ┌───────────────┐ ┌──────────────┐ ┌────────────────────┐
        │  ALL Metrics  │ │ SOK Metrics  │ │ KARKKAINEN Metrics │
        └───────────────┘ └──────────────┘ └────────────────────┘
                    │               │               │
                    └───────────────┼───────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────┐
│                  NF_Weekly_ServiceLevels Sheet                    │
│                                                                   │
│  Header Row:  [ISOWeek | Group | OrdersTotal | ... | P90Hours]  │
│  Info Row:    [Week window | Created timestamp | ...]           │
│  Data Rows:   ALL, SOK, KARKKAINEN metrics                       │
└─────────────────────────────────────────────────────────────────┘
```

## Execution Paths

### 1. Manual Execution (via Menu)

```
User clicks menu
      │
      ▼
NF Bulk Operations → Rakenna viikkopalvelutaso (ALL/SOK/KRK)
      │
      ▼
NF_buildWeeklyServiceLevels()
      │
      ▼
Report generated in NF_Weekly_ServiceLevels sheet
```

### 2. Automated Execution (via Trigger)

```
Monday 02:00 trigger fires
      │
      ▼
NF_weeklyReportBuild()
      │
      ├─→ NF_BuildWeeklyReports() or makeWeeklyReportsSunSun()
      │
      └─→ NF_buildWeeklyServiceLevels()  (guarded call)
            │
            ▼
      Report generated automatically
```

### 3. Programmatic Execution

```
Script or external trigger
      │
      ▼
const result = NF_buildWeeklyServiceLevels();
      │
      ▼
Returns: { week, all, sok, karkkainen }
```

## Metrics Calculation Flow

```
For each group (ALL, SOK, KARKKAINEN):
  
  1. Filter rows by payer (if not ALL group)
     │
     ▼
  2. Count OrdersTotal = filtered rows
     │
     ▼
  3. For each row with Delivered timestamp:
     │
     ├─→ Parse Date sent and Delivered Time
     │
     ├─→ Calculate lead time in hours
     │
     ├─→ Categorize into bucket:
     │     - < 24h
     │     - 24-72h
     │     - > 72h
     │
     └─→ Store for statistics calculation
     
  4. Calculate DeliveredTotal and PendingTotal
     │
     ▼
  5. Calculate bucket counts and percentages
     │
     ▼
  6. Calculate statistics (avg, median, P90)
     │
     ▼
  Output: Complete metrics object
```

## Configuration Priority

```
Script Properties
      │
      ├─→ TARGET_SHEET ────────────→ Default: "Packages"
      │
      ├─→ SOK_FREIGHT_ACCOUNT ─────→ Default: "5010"
      │
      ├─→ KARKKAINEN_NUMBERS ──────→ Default: from KARKKAINEN_NUMBERS const
      │
      └─→ NF_DATE_SENT_HINTS ──────→ Default: hardcoded candidates
                                       │
                                       ▼
                              Column detection uses these hints
```

## Integration Points

### Files Modified/Created

```
NF_ServiceLevels.gs (NEW)
├── Main function: NF_buildWeeklyServiceLevels()
├── 12 helper functions (NFSL_*)
└── Test function: NFSL_testServiceLevelCalculations()

NF_Menu_Triggers.gs (MODIFIED)
├── Added menu item in NF_addBulkMenuItems()
└── Added call in NF_weeklyReportBuild()

NF_onOpen_Extension.gs (MODIFIED)
└── Added menu item in NF_enhancedOnOpen()

NF_ServiceLevels_README.md (NEW)
└── Comprehensive documentation
```

### Reused Functions

```
From NF_SOK_KRK_Weekly.gs:
└── NF_getLastFinishedWeekSunWindow_()  (week window calculation)

From Helpers.js (optional, with fallback):
├── parseDateFlexible_()  (date parsing)
├── normalize_()  (text normalization)
└── fmtDateTime_()  (date formatting)

From PropertiesService:
├── getProperty('TARGET_SHEET')
├── getProperty('SOK_FREIGHT_ACCOUNT')
├── getProperty('KARKKAINEN_NUMBERS')
└── getProperty('NF_DATE_SENT_HINTS')
```

## Error Handling

```
Entry Point: NF_buildWeeklyServiceLevels()
      │
      ├─→ Source sheet missing? → Log warning, return early
      │
      ├─→ Required column missing? → Log warning, may return or use fallback
      │
      ├─→ No data in window? → Log info, return with zero metrics
      │
      ├─→ Invalid dates? → Skip in calculations (null safe)
      │
      └─→ General error? → Log to console and Error_Log sheet
                           │
                           ▼
                    Throw error (caught by trigger handler)
```

## Performance Considerations

1. **Batch Operations**: Single `getDataRange().getValues()` read
2. **Batch Write**: Single `setValues()` for all output rows
3. **In-Memory Processing**: All calculations done in JavaScript arrays
4. **Minimal Sheet Access**: Only 2 sheet reads + 3 writes per execution
5. **No External APIs**: All processing local to Apps Script

## Testing Strategy

### Unit Tests (NFSL_testServiceLevelCalculations)

```
✓ Date parsing (multiple formats)
✓ ISO week calculation
✓ Lead time bucket logic
✓ Statistics (avg, median, P90)
✓ Column detection
```

### Integration Testing

```
1. Prepare test data in Packages sheet
   - Add rows with Date sent and Delivered Time
   - Include SOK and Kärkkäinen payer values
   - Span multiple weeks

2. Run NF_buildWeeklyServiceLevels()

3. Verify output:
   - Sheet exists: NF_Weekly_ServiceLevels
   - 3 data rows: ALL, SOK, KARKKAINEN
   - Metrics match expected values
   - Percentages sum to ~100%
```

### Regression Testing

```
Before deployment:
1. Verify existing NF_weeklyReportBuild() still works
2. Verify menu items appear correctly
3. Verify no errors in console logs
4. Verify Error_Log sheet (if exists) has no new errors
```
