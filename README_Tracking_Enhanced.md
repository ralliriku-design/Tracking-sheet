# Enhanced Tracking System

This document describes the enhanced tracking system capabilities including the ad hoc tracker functionality.

## Overview

The enhanced tracking system provides comprehensive tracking data processing with support for:
- Multiple carrier APIs (Posti, GLS, DHL, Bring, Matkahuolto)
- Enhanced data extraction and KPI generation
- Automated delivery time calculations
- Country and weekly performance metrics

## Core Features

### Carrier Support
- **Posti**: Full API integration with OAuth
- **GLS**: Finnish API and OAuth fallback
- **DHL**: API key-based tracking
- **Bring**: Full API integration with credentials
- **Matkahuolto**: Basic API integration

### Data Processing
- Automatic date extraction for created/delivered events
- Delivery time calculations in days
- ISO week numbering for time-based analytics
- Country-based performance grouping

## Ad hoc tracker

The ad hoc tracker provides flexible import and tracking capabilities for one-off data processing needs.

### Features

- **Enhanced Header Detection**: Automatically detects tracking columns with support for Finnish and English column names
- **Fallback Content Scanning**: When header detection fails, analyzes column content to identify tracking codes
- **KPI Generation**: Automatically calculates delivery performance metrics by country and week
- **Flexible Import**: Supports Excel and CSV files from Google Drive

### Supported Column Names

The tracker automatically detects these column header variations:

**Tracking Columns:**
- `package id (seurantakoodi)` (Finnish with English)
- `package id`, `seurantakoodi`
- `tracking number`, `tracking`, `barcode`
- `waybill`, `waybill no`, `awb`
- `package number`, `shipment id`
- `consignment number`, `parcel id`

**Carrier Columns:**
- `carrier`, `carrier name`, `delivery carrier`
- `courier`, `logistics provider`, `shipper`
- `kuljetusliike`, `kuljetusyhtiö`, `kuljetus`
- `toimitustapa`, `delivery method`

**Country Columns:**
- `country`, `destination country`, `dest country`
- `delivery country`, `country code`
- `maa`, `kohdamaa`, `destination`

### Installation and Setup

1. **Install Menu Trigger** (one-time setup):
   ```javascript
   Adhoc_installMenuTrigger()
   ```
   
   Or run from the Scripts editor, or use the menu item once manually installed.

2. **Initialize System** (optional):
   ```javascript
   Adhoc_initialize()
   ```
   
   This creates the necessary sheets and sets up the trigger system.

### Usage

#### From Menu
1. Install the menu trigger using `Adhoc_installMenuTrigger()`
2. Use **Tracking → Ad hoc -tracker → Aja tuonti & päivitys**
3. Paste Google Drive URL or file ID when prompted
4. The system will automatically:
   - Detect column headers
   - Extract tracking data
   - Process all tracking codes
   - Generate KPI metrics
   - Create results in `Adhoc_Results` and `Adhoc_KPI` sheets

#### Programmatic Usage
```javascript
// Basic import
ADHOC_RunFromUrl('drive_url_or_file_id');

// With options
ADHOC_RunFromUrl('drive_url_or_file_id', {
  headerRow: 2,  // If headers are not on row 1
  trackingHeaderHints: ['custom_tracking_col'],
  carrierHeaderHints: ['custom_carrier_col'],
  countryHeaderHints: ['custom_country_col']
});

// Refresh existing results
ADHOC_RefreshResults();
```

#### Advanced Options
```javascript
const options = {
  sourceSheetName: 'Sheet1',      // Specific sheet name (for Excel files)
  headerRow: 2,                   // Header row number (default: 1)
  trackingHeaderHints: [          // Additional tracking column hints
    'my_tracking_column',
    'package_reference'
  ],
  carrierHeaderHints: [           // Additional carrier column hints
    'shipping_company',
    'logistics_partner'
  ],
  countryHeaderHints: [           // Additional country column hints
    'destination_country',
    'delivery_location'
  ]
};

ADHOC_buildFromValues(spreadsheetData, 'label', options);
```

### Output Sheets

#### Adhoc_Results
Contains detailed tracking results for each processed code:
- `Carrier`: Detected or provided carrier name
- `Tracking`: Tracking code
- `Country`: Destination country (if available)
- `Status`: Current delivery status
- `CreatedISO`: Package creation/acceptance timestamp
- `DeliveredISO`: Package delivery timestamp
- `DaysToDeliver`: Calculated delivery time in days
- `WeekISO`: ISO week number for delivery date
- `RefreshAt`: Last update timestamp
- `Location`: Last known location
- `Raw`: Raw API response (truncated)

#### Adhoc_KPI
Contains aggregated performance metrics:
- `Country`: Destination country
- `ISO Week`: Week in YYYY-Wxx format
- `Deliveries`: Number of delivered packages
- `Avg Days`: Average delivery time
- `Median Days`: Median delivery time
- `Min Days`: Fastest delivery time
- `Max Days`: Slowest delivery time

### Error Handling

The tracker includes comprehensive error handling:

- **Missing Headers**: Clear error messages indicating which columns couldn't be found
- **Content Fallback**: Automatic content analysis when header detection fails
- **Tracking Errors**: Individual tracking failures don't stop batch processing
- **API Limits**: Respects carrier API rate limits and retry policies

### Troubleshooting

**Column Detection Issues:**
- Verify column headers match supported patterns
- Use `trackingHeaderHints` for custom column names
- Check that tracking data is in the expected format

**Menu Not Appearing:**
```javascript
// Check trigger status
Adhoc_checkTriggerStatus();

// Manually add menu for current session
Adhoc_installMenuManual();

// Reinstall trigger
Adhoc_removeMenuTrigger();
Adhoc_installMenuTrigger();
```

**Import Failures:**
- Ensure Google Drive file is accessible
- Verify file format (Excel .xlsx or CSV)
- Check that file contains actual tracking data
- Review error messages in the toast notifications

### Integration Notes

The ad hoc tracker integrates with the existing tracking infrastructure:
- Uses the same carrier API functions (`TRK_trackByCarrier_`)
- Respects the same Script Properties for API credentials
- Follows the same rate limiting and caching patterns
- Compatible with existing refresh and bulk operations

This ensures consistency across all tracking operations in the system.