/**
 * NF_Bulk_Backfill.gs — Bulk processing orchestration for end-to-end workflows
 * 
 * Provides bulk import, rebuild, and refresh functions to process historical data
 * and keep tracking statuses up to date without overloading Apps Script quotas.
 * 
 * Key Functions:
 *   - NF_BulkImportFromDrive(): Import all files from Drive folder
 *   - NF_BulkRebuildAll(): Complete daily workflow sequence
 *   - NF_RefreshAllPending(): Batch refresh missing tracking statuses
 *   - NF_BulkFindDuplicates_All(): Cross-source duplicate detection
 */

/**
 * Wrapper function to import all files from the configured Drive folder.
 * Calls NF_Drive_ImportLatestAll() and returns results.
 */
function NF_BulkImportFromDrive() {
  console.log('NF_BulkImportFromDrive: Starting Drive import...');
  
  try {
    const results = NF_Drive_ImportLatestAll();
    
    // Log summary
    const successful = results.filter(r => r.status === 'success');
    const errors = results.filter(r => r.status === 'error');
    
    console.log(`Drive import completed: ${successful.length} successful, ${errors.length} errors`);
    
    if (errors.length > 0) {
      console.warn('Import errors:', errors);
    }
    
    return results;
  } catch (error) {
    console.error('NF_BulkImportFromDrive failed:', error);
    throw error;
  }
}

/**
 * Complete daily rebuild sequence without overloading Apps Script.
 * Performs all steps needed for a full daily run in the correct order.
 */
function NF_BulkRebuildAll() {
  console.log('NF_BulkRebuildAll: Starting complete rebuild sequence...');
  const startTime = Date.now();
  
  try {
    // Step 1: Import from Drive folder
    console.log('Step 1: Importing from Drive...');
    NF_BulkImportFromDrive();
    
    // Step 2: Update inventory balances
    console.log('Step 2: Updating inventory balances...');
    if (typeof NF_UpdateInventoryBalances === 'function') {
      NF_UpdateInventoryBalances();
    } else {
      console.warn('NF_UpdateInventoryBalances not available, skipping...');
    }
    
    // Step 3: Reconcile weekly reports from import
    console.log('Step 3: Reconciling weekly reports...');
    if (typeof NF_OpenImportTemplate === 'function') {
      NF_OpenImportTemplate();
    }
    if (typeof NF_ReconcileWeeklyFromImport === 'function') {
      NF_ReconcileWeeklyFromImport();
    } else {
      console.warn('Weekly reconciliation functions not available, skipping...');
    }
    
    // Step 4: Validate and reconcile ERP data
    console.log('Step 4: Processing ERP data...');
    if (typeof NF_ValidateERPImport === 'function') {
      NF_ValidateERPImport();
    }
    if (typeof NF_ReconcileERPToWeekly === 'function') {
      NF_ReconcileERPToWeekly();
    } else {
      console.warn('ERP reconciliation functions not available, skipping...');
    }
    
    // Step 5: Run daily Gmail import and tracking refresh
    console.log('Step 5: Running daily flow (Gmail + tracking)...');
    if (typeof NF_RunDailyFlow === 'function') {
      NF_RunDailyFlow();
    } else if (typeof runDailyFlowOnce === 'function') {
      runDailyFlowOnce();
    } else {
      console.warn('Daily flow function not available, skipping...');
    }
    
    // Step 6: Build/merge SOK and Kärkkäinen reports
    console.log('Step 6: Building SOK/Kärkkäinen reports...');
    NF_buildSokKarkkainenAlways();
    
    // Step 7: Refresh pending tracking statuses (limited batch)
    console.log('Step 7: Refreshing pending statuses...');
    NF_RefreshAllPending(100);
    
    const duration = (Date.now() - startTime) / 1000;
    console.log(`NF_BulkRebuildAll completed successfully in ${duration}s`);
    
    return {
      status: 'success',
      duration: duration,
      steps: 7
    };
    
  } catch (error) {
    const duration = (Date.now() - startTime) / 1000;
    console.error(`NF_BulkRebuildAll failed after ${duration}s:`, error);
    throw error;
  }
}

/**
 * Scans key sheets for rows missing tracking data and refreshes them in batches.
 * Uses existing tracking functions with conservative throttling.
 * 
 * @param {number} limitPerRun - Maximum number of tracking calls per execution (default: 100)
 */
function NF_RefreshAllPending(limitPerRun = 100) {
  console.log(`NF_RefreshAllPending: Starting with limit ${limitPerRun}`);
  
  const sheetsToCheck = ['Packages', 'Report_SOK', 'Report_Karkkainen', 'Delivery_Times'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let totalCalls = 0;
  let totalUpdated = 0;
  
  for (const sheetName of sheetsToCheck) {
    if (totalCalls >= limitPerRun) {
      console.log(`Reached call limit (${limitPerRun}), stopping...`);
      break;
    }
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      console.log(`Sheet ${sheetName} not found or empty, skipping...`);
      continue;
    }
    
    console.log(`Processing sheet: ${sheetName}`);
    
    try {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Ensure refresh columns exist
      NF_ensureRefreshCols_(sheet, headers);
      
      // Find columns
      const colMap = {};
      headers.forEach((header, index) => {
        const h = String(header || '').toLowerCase();
        if (h.includes('carrier')) colMap.carrier = index;
        if (h.includes('tracking') && (h.includes('number') || h.includes('code'))) colMap.tracking = index;
        if (h.includes('delivered') && !h.includes('source')) colMap.delivered = index;
        if (h.includes('refresh') && h.includes('time')) colMap.refreshTime = index;
      });
      
      if (colMap.carrier === undefined || colMap.tracking === undefined) {
        console.log(`Required columns not found in ${sheetName}, skipping...`);
        continue;
      }
      
      // Process rows that need refresh
      let sheetUpdated = 0;
      
      for (let rowIndex = 1; rowIndex < data.length && totalCalls < limitPerRun; rowIndex++) {
        const row = data[rowIndex];
        
        // Skip if already has recent refresh
        if (colMap.refreshTime !== undefined && row[colMap.refreshTime]) {
          const refreshTime = new Date(row[colMap.refreshTime]);
          const sixHoursAgo = new Date(Date.now() - 6 * 60 * 60 * 1000);
          if (refreshTime > sixHoursAgo) {
            continue; // Recently refreshed
          }
        }
        
        // Skip if already delivered
        if (colMap.delivered !== undefined && row[colMap.delivered]) {
          continue;
        }
        
        const carrier = String(row[colMap.carrier] || '').trim();
        const trackingCode = String(row[colMap.tracking] || '').trim();
        
        if (!carrier || !trackingCode) {
          continue;
        }
        
        // Perform tracking call
        try {
          let result;
          
          // Try enhanced version first, fallback to basic
          if (typeof TRK_trackByCarrierEnhanced === 'function') {
            result = TRK_trackByCarrierEnhanced(carrier, trackingCode);
          } else if (typeof TRK_trackByCarrier_ === 'function') {
            result = TRK_trackByCarrier_(carrier, trackingCode);
          } else {
            console.warn('No tracking function available');
            break;
          }
          
          // Update row with results
          if (result && typeof result === 'object') {
            NF_updateRowWithTrackingResult_(sheet, rowIndex + 1, result);
            sheetUpdated++;
          }
          
          totalCalls++;
          
          // Rate limiting
          Utilities.sleep(500); // 500ms between calls
          
        } catch (error) {
          console.error(`Error tracking ${carrier} ${trackingCode}:`, error);
          
          // If rate limited, wait longer
          if (String(error).includes('429') || String(error).includes('rate limit')) {
            console.log('Rate limit detected, sleeping 5 seconds...');
            Utilities.sleep(5000);
          }
        }
      }
      
      console.log(`Sheet ${sheetName}: ${sheetUpdated} rows updated`);
      totalUpdated += sheetUpdated;
      
    } catch (error) {
      console.error(`Error processing sheet ${sheetName}:`, error);
    }
  }
  
  console.log(`NF_RefreshAllPending completed: ${totalUpdated} rows updated, ${totalCalls} API calls`);
  
  return {
    totalUpdated: totalUpdated,
    totalCalls: totalCalls,
    sheetsProcessed: sheetsToCheck.length
  };
}

/**
 * Ensures that refresh columns exist in the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {any[]} headers
 */
function NF_ensureRefreshCols_(sheet, headers) {
  const requiredColumns = ['Delivered', 'Delivered_Source', 'RefreshTime', 'RefreshStatus', 'RefreshLocation'];
  const existingHeaders = headers.map(h => String(h || ''));
  
  const missingColumns = requiredColumns.filter(col => 
    !existingHeaders.some(existing => 
      existing.toLowerCase().includes(col.toLowerCase())
    )
  );
  
  if (missingColumns.length > 0) {
    console.log(`Adding missing columns to sheet: ${missingColumns.join(', ')}`);
    
    const lastCol = sheet.getLastColumn();
    for (let i = 0; i < missingColumns.length; i++) {
      sheet.getRange(1, lastCol + 1 + i).setValue(missingColumns[i]);
    }
  }
}

/**
 * Updates a sheet row with tracking result data.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum - 1-based row number
 * @param {Object} result - Tracking result object
 */
function NF_updateRowWithTrackingResult_(sheet, rowNum, result) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const now = new Date();
  
  // Find column positions
  const updates = [];
  
  headers.forEach((header, colIndex) => {
    const h = String(header || '').toLowerCase();
    
    if (h.includes('delivered') && !h.includes('source') && result.delivered) {
      updates.push({ col: colIndex + 1, value: result.delivered });
    } else if (h.includes('delivered_source') && result.deliveredSource) {
      updates.push({ col: colIndex + 1, value: result.deliveredSource });
    } else if (h.includes('refreshtime')) {
      updates.push({ col: colIndex + 1, value: now });
    } else if (h.includes('refreshstatus') && result.status) {
      updates.push({ col: colIndex + 1, value: result.status });
    } else if (h.includes('refreshlocation') && result.location) {
      updates.push({ col: colIndex + 1, value: result.location });
    }
  });
  
  // Apply updates
  updates.forEach(update => {
    sheet.getRange(rowNum, update.col).setValue(update.value);
  });
}

/**
 * Extended duplicate detection across PBI import and ERP_Validated sheets.
 * Groups by Reference orderid + Article + DeliveryPlace and flags distinct Orderid > 1.
 */
function NF_BulkFindDuplicates_All() {
  console.log('NF_BulkFindDuplicates_All: Starting cross-source duplicate detection...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['Import_Weekly', 'Import_PBI_Balance', 'Import_ERP_StockPicking', 'ERP_Validated'];
  
  // Get or create duplicates sheet
  let duplicatesSheet = ss.getSheetByName('Reconcile_Duplicates');
  if (!duplicatesSheet) {
    duplicatesSheet = ss.insertSheet('Reconcile_Duplicates');
    duplicatesSheet.getRange(1, 1, 1, 7).setValues([
      ['Source', 'Reference', 'Article', 'DeliveryPlace', 'Orderid', 'DuplicateCount', 'DetectedAt']
    ]);
    duplicatesSheet.setFrozenRows(1);
  }
  
  const duplicateGroups = new Map();
  
  // Process each sheet
  for (const sheetName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      console.log(`Sheet ${sheetName} not found or empty, skipping...`);
      continue;
    }
    
    console.log(`Processing ${sheetName} for duplicates...`);
    
    try {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Find relevant columns
      const colMap = {};
      headers.forEach((header, index) => {
        const h = String(header || '').toLowerCase();
        if (h.includes('reference') || h.includes('orderid')) colMap.reference = index;
        if (h.includes('article') || h.includes('product')) colMap.article = index;
        if (h.includes('delivery') && h.includes('place')) colMap.deliveryPlace = index;
        if (h.includes('orderid') && !h.includes('reference')) colMap.orderid = index;
      });
      
      // Process data rows
      for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        
        const reference = String(row[colMap.reference] || '').trim();
        const article = String(row[colMap.article] || '').trim();
        const deliveryPlace = String(row[colMap.deliveryPlace] || '').trim();
        const orderid = String(row[colMap.orderid] || row[colMap.reference] || '').trim();
        
        if (!reference || !article) continue;
        
        // Create composite key
        const key = `${reference}|${article}|${deliveryPlace}`;
        
        if (!duplicateGroups.has(key)) {
          duplicateGroups.set(key, {
            reference: reference,
            article: article,
            deliveryPlace: deliveryPlace,
            orderids: new Set(),
            sources: new Set()
          });
        }
        
        const group = duplicateGroups.get(key);
        group.orderids.add(orderid);
        group.sources.add(sheetName);
      }
      
    } catch (error) {
      console.error(`Error processing sheet ${sheetName}:`, error);
    }
  }
  
  // Find and write duplicates
  const duplicates = [];
  const now = new Date();
  
  duplicateGroups.forEach((group, key) => {
    if (group.orderids.size > 1) {
      duplicates.push([
        Array.from(group.sources).join(', '),
        group.reference,
        group.article,
        group.deliveryPlace,
        Array.from(group.orderids).join(', '),
        group.orderids.size,
        now
      ]);
    }
  });
  
  // Write results
  if (duplicates.length > 0) {
    const startRow = duplicatesSheet.getLastRow() + 1;
    duplicatesSheet.getRange(startRow, 1, duplicates.length, 7).setValues(duplicates);
    console.log(`Found ${duplicates.length} duplicate groups, written to Reconcile_Duplicates`);
  } else {
    console.log('No duplicates found');
  }
  
  return {
    duplicatesFound: duplicates.length,
    totalGroups: duplicateGroups.size
  };
}

/**
 * Placeholder for SOK/Kärkkäinen always build function.
 * This should merge all historical data into reports.
 */
function NF_buildSokKarkkainenAlways() {
  console.log('NF_buildSokKarkkainenAlways: Building historical SOK/Kärkkäinen reports...');
  
  // Try to call existing function if available
  if (typeof buildSokKarkkainenAlways === 'function') {
    return buildSokKarkkainenAlways();
  } else if (typeof makeWeeklyReportsSunSun_Only === 'function') {
    return makeWeeklyReportsSunSun_Only();
  } else {
    console.warn('SOK/Kärkkäinen build function not available');
    return null;
  }
}