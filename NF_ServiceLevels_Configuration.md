# Service Level Reporting - Configuration Guide

## Quick Start

### 1. Set Script Properties

Open Script Editor → Project Settings → Script Properties, then add:

| Property Key | Example Value | Purpose |
|-------------|---------------|---------|
| `TARGET_SHEET` | `Packages` | Source data sheet name |
| `SOK_FREIGHT_ACCOUNT` | `5010` | SOK customer account (digits only) |
| `KARKKAINEN_NUMBERS` | `1234,5678,9012` | Kärkkäinen accounts (comma-separated) |
| `NF_DATE_SENT_HINTS` | `Date sent,Sent date` | (Optional) Custom column hints |

### 2. Verify Column Names

Check your source sheet (default: "Packages") has these columns:

**Required**:
- ✅ One of: `Delivered Time`, `Delivered At`, or `RefreshTime`
- ✅ One of: `Date sent`, `Sent date`, `Handover date`, or `Submitted date`

**Recommended**:
- ⭐ Payer column (any name containing: payer, freight, billing, customer, account)
- ⭐ Date column for filtering (Created, Timestamp, etc.)

### 3. Run First Time

1. Open your Google Sheets spreadsheet
2. Navigate to menu: **NF Bulk Operations** → **Rakenna viikkopalvelutaso (ALL/SOK/KRK)**
3. Wait for processing (usually < 10 seconds)
4. Check new sheet: `NF_Weekly_ServiceLevels`

## Configuration Examples

### Example 1: Standard Configuration

```javascript
// Script Properties setup
const properties = {
  'TARGET_SHEET': 'Packages',
  'SOK_FREIGHT_ACCOUNT': '5010',
  'KARKKAINEN_NUMBERS': '1234,5678,9012'
};

// Headers in Packages sheet:
// Package Number | Date sent | Payer | Delivered Time | Status | Created
```

### Example 2: Custom Column Names

If your sheet uses different column names:

```javascript
// Script Properties
const properties = {
  'TARGET_SHEET': 'Packages',
  'SOK_FREIGHT_ACCOUNT': '5010',
  'KARKKAINEN_NUMBERS': '1234,5678,9012',
  'NF_DATE_SENT_HINTS': 'Handover Date,Dispatch Date,Sent Time'
};

// Headers in Packages sheet:
// Tracking | Handover Date | Customer | RefreshTime | Current Status
```

### Example 3: Multiple Kärkkäinen Accounts

```javascript
// Script Properties
const properties = {
  'TARGET_SHEET': 'Shipments',
  'SOK_FREIGHT_ACCOUNT': '990719901',
  'KARKKAINEN_NUMBERS': '615471,802669,7030057,1112223'
};
```

## Programmatic Configuration

### Set Properties via Script

```javascript
function setupServiceLevelConfig() {
  const props = PropertiesService.getScriptProperties();
  
  props.setProperties({
    'TARGET_SHEET': 'Packages',
    'SOK_FREIGHT_ACCOUNT': '5010',
    'KARKKAINEN_NUMBERS': '1234,5678,9012'
  });
  
  console.log('Configuration saved successfully');
}
```

### Read Current Configuration

```javascript
function showCurrentConfig() {
  const props = PropertiesService.getScriptProperties();
  
  const config = {
    targetSheet: props.getProperty('TARGET_SHEET') || 'Packages',
    sokAccount: props.getProperty('SOK_FREIGHT_ACCOUNT') || '5010',
    krkNumbers: props.getProperty('KARKKAINEN_NUMBERS') || '1234,5678,9012',
    dateSentHints: props.getProperty('NF_DATE_SENT_HINTS') || '(default)'
  };
  
  console.log('Current configuration:', config);
  return config;
}
```

## Validation

### Check Configuration

```javascript
function validateServiceLevelConfig() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check TARGET_SHEET exists
  const targetSheet = props.getProperty('TARGET_SHEET') || 'Packages';
  const sheet = ss.getSheetByName(targetSheet);
  if (!sheet) {
    console.error(`❌ Target sheet "${targetSheet}" not found`);
    return false;
  }
  console.log(`✅ Target sheet "${targetSheet}" exists`);
  
  // Check required columns
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    console.error('❌ Target sheet is empty');
    return false;
  }
  
  const headers = data[0];
  const headerStr = headers.join('|').toLowerCase();
  
  // Check for Date sent
  const hasDateSent = 
    headerStr.includes('date sent') ||
    headerStr.includes('sent date') ||
    headerStr.includes('handover') ||
    headerStr.includes('submitted');
  
  if (!hasDateSent) {
    console.warn('⚠️ Date sent column not found (service levels may be limited)');
  } else {
    console.log('✅ Date sent column found');
  }
  
  // Check for Delivered Time
  const hasDelivered = 
    headerStr.includes('delivered time') ||
    headerStr.includes('delivered at') ||
    headerStr.includes('refreshtime');
  
  if (!hasDelivered) {
    console.error('❌ Delivered timestamp column not found');
    return false;
  }
  console.log('✅ Delivered timestamp column found');
  
  // Check for Payer
  const hasPayer = 
    headerStr.includes('payer') ||
    headerStr.includes('freight') ||
    headerStr.includes('billing') ||
    headerStr.includes('customer') ||
    headerStr.includes('account');
  
  if (!hasPayer) {
    console.warn('⚠️ Payer column not found (grouping may not work)');
  } else {
    console.log('✅ Payer column found');
  }
  
  console.log('\n✅ Configuration validation complete');
  return true;
}
```

## Sample Data Structure

### Minimal Working Example

```
| Package Number | Date sent        | Payer | Delivered Time      | Status    |
|----------------|------------------|-------|---------------------|-----------|
| PKG001         | 2025-01-13 09:00 | 5010  | 2025-01-13 18:30   | Delivered |
| PKG002         | 2025-01-13 10:00 | 1234  | 2025-01-14 11:15   | Delivered |
| PKG003         | 2025-01-13 11:00 | 5678  |                     | In Transit|
| PKG004         | 2025-01-14 08:00 | 5010  | 2025-01-15 14:00   | Delivered |
```

### Expected Output

```
NF_Weekly_ServiceLevels sheet:

| ISOWeek  | Group      | OrdersTotal | DeliveredTotal | PendingTotal | LTlt24h | LT24_72h | LTgt72h | ... |
|----------|------------|-------------|----------------|--------------|---------|----------|---------|-----|
| 2025-W02 | ALL        | 4           | 3              | 1            | 2       | 1        | 0       | ... |
| 2025-W02 | SOK        | 2           | 2              | 0            | 1       | 1        | 0       | ... |
| 2025-W02 | KARKKAINEN | 2           | 1              | 1            | 1       | 0        | 0       | ... |
```

## Troubleshooting

### Issue: "Target sheet not found"

**Solution**: Set `TARGET_SHEET` property to match your sheet name exactly (case-sensitive)

```javascript
PropertiesService.getScriptProperties()
  .setProperty('TARGET_SHEET', 'YourActualSheetName');
```

### Issue: "No SOK/Kärkkäinen data in report"

**Solution**: Verify account numbers match exactly (digits only)

```javascript
// Check what values are in your Payer column
function checkPayerValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Packages');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find Payer column
  const payerIdx = headers.findIndex(h => 
    String(h).toLowerCase().includes('payer') ||
    String(h).toLowerCase().includes('freight') ||
    String(h).toLowerCase().includes('account')
  );
  
  if (payerIdx < 0) {
    console.log('Payer column not found');
    return;
  }
  
  // Get unique payer values
  const payers = new Set();
  for (let i = 1; i < data.length; i++) {
    const value = String(data[i][payerIdx] || '').replace(/\D/g, '');
    if (value) payers.add(value);
  }
  
  console.log('Unique payer values (digits only):');
  Array.from(payers).sort().forEach(p => console.log(`  ${p}`));
}
```

### Issue: "Column not detected"

**Solution**: Use `NF_DATE_SENT_HINTS` to provide custom column names

```javascript
// List all column names
function showColumnHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Packages');
  const headers = sheet.getDataRange().getValues()[0];
  
  console.log('Column headers:');
  headers.forEach((h, i) => console.log(`  ${i}: "${h}"`));
}

// Then set custom hints
PropertiesService.getScriptProperties()
  .setProperty('NF_DATE_SENT_HINTS', 'Your Column Name,Alternative Name');
```

## Advanced Configuration

### Multiple Source Sheets

To process multiple sheets, modify the function:

```javascript
function NF_buildServiceLevelsMultiSheet() {
  const sheets = ['Packages', 'Packages_Archive', 'Import_Weekly'];
  
  // Combine data from all sheets
  let allData = [];
  let unionHeaders = [];
  
  for (const sheetName of sheets) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      // Merge and process...
    }
  }
  
  // Continue with combined data...
}
```

### Custom Grouping Logic

To add more customer groups:

```javascript
// In Script Properties
const properties = {
  'CUSTOMER_A_NUMBERS': '1111,2222',
  'CUSTOMER_B_NUMBERS': '3333,4444'
};

// Modify NFSL_calculateGroupMetrics_ to support additional groups
```

### Historical Reporting

To keep historical data:

```javascript
function NF_appendWeeklyServiceLevels() {
  // Run normal calculation
  const result = NF_buildWeeklyServiceLevels();
  
  // Append to history sheet instead of overwriting
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('ServiceLevel_History') || 
                       ss.insertSheet('ServiceLevel_History');
  
  // Append rows with timestamp...
}
```
