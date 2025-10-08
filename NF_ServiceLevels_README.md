# Weekly Service Level Reporting

## Overview

The Weekly Service Level Reporting module (`NF_ServiceLevels.gs`) provides automated calculation and reporting of delivery service level metrics based on the elapsed time between shipment handover (Date sent) and delivery confirmation (Delivered Time).

## Features

- **Automatic Weekly Calculation**: Computes metrics for the last finished Sunday-to-Sunday week
- **Multiple Group Breakdowns**: Reports metrics for ALL, SOK, and Kärkkäinen customer groups
- **Service Level Buckets**: Tracks deliveries in < 24h, 24-72h, and > 72h time windows
- **Statistical Analysis**: Calculates average, median, and P90 lead times
- **Flexible Column Detection**: Adapts to various header naming conventions
- **Integration Ready**: Automatically runs with weekly reports and available via menu

## Output Sheet: NF_Weekly_ServiceLevels

### Headers

| Column | Description |
|--------|-------------|
| ISOWeek | ISO week identifier (e.g., "2025-W03") |
| Group | Customer group (ALL, SOK, or KARKKAINEN) |
| OrdersTotal | Total orders in the week window |
| DeliveredTotal | Orders with delivery confirmation |
| PendingTotal | Orders not yet delivered |
| LTlt24h | Deliveries with lead time < 24 hours |
| LT24_72h | Deliveries with lead time 24-72 hours |
| LTgt72h | Deliveries with lead time > 72 hours |
| Pct_lt24h_of_delivered | Percentage of deliveries < 24h |
| Pct_24_72h_of_delivered | Percentage of deliveries 24-72h |
| Pct_gt72h_of_delivered | Percentage of deliveries > 72h |
| AvgHours | Average lead time in hours |
| MedianHours | Median lead time in hours |
| P90Hours | 90th percentile lead time in hours |

### Info Row

Row 2 contains metadata:
- Column A: Week window in format "Week (SUN→SUN): YYYY-MM-DD - YYYY-MM-DD"
- Column B: Creation timestamp

### Data Rows

Rows 3-5 contain metrics for:
1. Row 3: ALL group (all orders)
2. Row 4: SOK group
3. Row 5: KARKKAINEN group

## Configuration

### Script Properties

Configure these in Script Properties for customization:

| Property | Purpose | Default |
|----------|---------|---------|
| TARGET_SHEET | Source data sheet name | "Packages" |
| SOK_FREIGHT_ACCOUNT | SOK customer account number | "5010" |
| KARKKAINEN_NUMBERS | Comma-separated Kärkkäinen account numbers | "1234,5678,9012" |
| NF_DATE_SENT_HINTS | Custom column name hints for Date sent (comma-separated) | "Date sent,Sent date,Handover date,Submitted date" |

### Column Detection

The module automatically detects required columns using these candidates:

**Date Sent (handover timestamp)**:
- Default candidates: "Date sent", "Sent date", "Handover date", "Submitted date"
- Override with `NF_DATE_SENT_HINTS` Script Property

**Delivered Time (tracking-based)**:
- Candidates: "Delivered Time", "Delivered At", "RefreshTime"
- First match is used

**Payer (for grouping)**:
- Detected by keywords: "payer", "freight account", "billing", "customer", "account"

**Window Date (fallback if Date sent unavailable)**:
- Candidates: "Created", "Created date", "Submitted date", "Booking date", etc.

## Usage

### Manual Execution

Run from the menu:
1. Open Google Sheets
2. Navigate to **NF Bulk Operations** → **Rakenna viikkopalvelutaso (ALL/SOK/KRK)**
3. The report is generated in sheet `NF_Weekly_ServiceLevels`

### Automated Execution

The service level report is automatically included in:
- **Weekly Report Build**: Runs every Monday at 02:00
- Triggered via `NF_weeklyReportBuild()` function

### Programmatic Execution

```javascript
// Build service level report for last finished week
NF_buildWeeklyServiceLevels();

// Returns object with metrics:
{
  week: "2025-W03",
  all: { group: "ALL", ordersTotal: 150, deliveredTotal: 120, ... },
  sok: { group: "SOK", ordersTotal: 80, deliveredTotal: 70, ... },
  karkkainen: { group: "KARKKAINEN", ordersTotal: 30, deliveredTotal: 25, ... }
}
```

## Calculation Logic

### Time Window

- **Week Definition**: Sunday 00:00 → Sunday 00:00 (exclusive)
- **Last Finished Week**: If today is Wednesday, reports the previous Sun→Sun week
- **Filtering**: Uses Date sent when available, otherwise falls back to Created/Submitted

### Lead Time Calculation

1. Parse Date sent timestamp
2. Parse Delivered Time timestamp
3. Calculate difference in milliseconds
4. Convert to hours: `(deliveredDate - sentDate) / (1000 * 60 * 60)`
5. Categorize into buckets:
   - **< 24h**: Lead time < 24 hours
   - **24-72h**: 24 hours ≤ lead time ≤ 72 hours
   - **> 72h**: Lead time > 72 hours

### Statistics

- **Average**: Sum of all lead times / count
- **Median**: Middle value when sorted (or average of two middle values)
- **P90**: Value at 90th percentile position in sorted array

### Group Filtering

- **ALL**: All rows in the time window (no filtering)
- **SOK**: Rows where normalized Payer digits match `SOK_FREIGHT_ACCOUNT`
- **KARKKAINEN**: Rows where normalized Payer digits match any in `KARKKAINEN_NUMBERS`

Digit normalization strips all non-numeric characters for comparison.

## Testing

### Unit Tests

Run the built-in test function:

```javascript
NFSL_testServiceLevelCalculations();
```

Tests verify:
- Date parsing for multiple formats
- ISO week calculation
- Lead time bucket logic
- Statistics calculations (avg, median, P90)
- Column detection with mock headers

### Integration Test

1. Ensure source sheet has required columns
2. Add test data with Date sent and Delivered Time
3. Run `NF_buildWeeklyServiceLevels()`
4. Verify output in `NF_Weekly_ServiceLevels` sheet

## Error Handling

The module includes comprehensive error handling:

- **Missing Source Sheet**: Logs warning and returns early
- **Missing Required Columns**: Logs warning, may return early or use fallback
- **Invalid Dates**: Skipped in calculations (null/empty dates)
- **Zero Delivered**: Shows 0.00 for percentages and empty for statistics
- **Script Property Errors**: Falls back to hardcoded defaults

All errors are logged to console and optionally to Error_Log sheet.

## File Structure

```
NF_ServiceLevels.gs
├── NF_buildWeeklyServiceLevels()          # Main entry point
├── NFSL_findDateSentColumn_()             # Column detection
├── NFSL_findDeliveredColumn_()            # Column detection
├── NFSL_findPayerColumn_()                # Column detection
├── NFSL_findWindowDateColumn_()           # Fallback column detection
├── NFSL_parseKarkkainenNumbers_()         # Config parsing
├── NFSL_parseDate_()                      # Flexible date parser
├── NFSL_formatDate_()                     # Date formatter
├── NFSL_getISOWeek_()                     # ISO week calculator
├── NFSL_normalizeDigits_()                # Digit extractor
├── NFSL_calculateGroupMetrics_()          # Core metrics calculation
├── NFSL_writeServiceLevelSheet_()         # Output sheet writer
└── NFSL_testServiceLevelCalculations()    # Test function
```

## Integration with Existing Modules

### NF_Menu_Triggers.gs

- Menu item added to "NF Bulk Operations" submenu
- Integrated into `NF_weeklyReportBuild()` trigger function
- Guarded with `typeof` check to prevent errors if module not loaded

### NF_onOpen_Extension.gs

- Menu item added to enhanced onOpen menu
- Available in "NF Bulk Operations" section

### NF_SOK_KRK_Weekly.gs

- Reuses `NF_getLastFinishedWeekSunWindow_()` for week calculation
- Compatible with existing SOK/Kärkkäinen grouping logic
- Uses same Script Properties for customer account numbers

## Troubleshooting

### No Data in Report

1. Check that source sheet (default: "Packages") exists and has data
2. Verify Date sent or fallback date column exists
3. Verify Delivered Time column exists
4. Confirm data exists for the last finished Sun→Sun week

### Incorrect Grouping

1. Verify `SOK_FREIGHT_ACCOUNT` Script Property is set correctly
2. Verify `KARKKAINEN_NUMBERS` Script Property contains comma-separated values
3. Check that Payer column is detected correctly (view console logs)
4. Verify Payer values match configured account numbers (digits only)

### Column Not Found

1. Review console logs to see which columns were detected
2. Use `NF_DATE_SENT_HINTS` to provide custom column name hints
3. Ensure column headers match expected candidates (exact match, trimmed)

### Percentages Don't Sum to 100%

- This is expected when there are rows with Date sent but no Delivered Time
- Only delivered rows are counted in bucket percentages
- Check PendingTotal to see undelivered count

## Future Enhancements

Potential improvements:
- Historical trend analysis (multiple weeks)
- Carrier-specific breakdowns
- Custom SLA threshold configuration
- Email/notification alerts for SLA breaches
- Export to CSV/PDF for reporting
- Dashboard visualization integration
