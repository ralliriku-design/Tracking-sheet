/**
 * NF_Inventory_Balance.gs — Inventory reconciliation with Drive folder preference
 * 
 * Updates inventory balances by preferring Drive folder auto-import over OneDrive URLs.
 * Builds aggregates and reconciliation reports from imported data.
 * 
 * Key Functions:
 *   - NF_UpdateInventoryBalances(): Main inventory update orchestrator
 *   - NF_BuildQuantsAggregate(): Aggregate quants data by location
 *   - NF_ReconcileInventory(): Compare ERP vs 3PL vs PBI data
 */

/**
 * Main function to update inventory balances.
 * Prefers Drive folder imports, falls back to OneDrive if configured.
 */
function NF_UpdateInventoryBalances() {
  console.log('NF_UpdateInventoryBalances: Starting inventory balance update...');
  
  const driveImportFolderId = PropertiesService.getScriptProperties().getProperty('DRIVE_IMPORT_FOLDER_ID');
  
  try {
    // Step 1: Try Drive import first if configured
    if (driveImportFolderId) {
      console.log('Using Drive folder import for inventory data...');
      
      // Import latest Quants and Warehouse Balance from Drive
      const quantsFile = NF_Drive_PickLatestFileByPattern_(driveImportFolderId, ['quants']);
      const warehouseFile = NF_Drive_PickLatestFileByPattern_(driveImportFolderId, ['warehouse balance', 'warehouse_balance', '3pl balance']);
      
      if (quantsFile) {
        console.log(`Importing quants from: ${quantsFile.getName()}`);
        NF_Drive_ReadCsvOrXlsxToSheet_(quantsFile, 'Import_Quants');
      }
      
      if (warehouseFile) {
        console.log(`Importing warehouse balance from: ${warehouseFile.getName()}`);
        NF_Drive_ReadCsvOrXlsxToSheet_(warehouseFile, 'Import_Warehouse_Balance');
      }
    } else {
      console.log('DRIVE_IMPORT_FOLDER_ID not configured, trying OneDrive fallback...');
      
      // Try OneDrive import if available
      if (typeof importFromOneDriveUrls === 'function') {
        importFromOneDriveUrls();
      } else {
        console.warn('No OneDrive import function available and no Drive folder configured');
      }
    }
    
    // Step 2: Build aggregates from imported data
    console.log('Building quants aggregates...');
    NF_BuildQuantsAggregate();
    
    // Step 3: Build location-based summary
    console.log('Building quants by location...');
    NF_BuildQuantsByLocation();
    
    // Step 4: Reconcile inventory data
    console.log('Reconciling inventory...');
    NF_ReconcileInventory();
    
    console.log('NF_UpdateInventoryBalances completed successfully');
    
    return {
      status: 'success',
      driveUsed: !!driveImportFolderId,
      timestamp: new Date()
    };
    
  } catch (error) {
    console.error('NF_UpdateInventoryBalances failed:', error);
    throw error;
  }
}

/**
 * Builds aggregated quants data from Import_Quants sheet.
 * Groups by Article and sums quantities across locations.
 */
function NF_BuildQuantsAggregate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Import_Quants');
  
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    console.log('Import_Quants sheet not found or empty, skipping aggregate...');
    return;
  }
  
  console.log('Building Quants_Aggregate from Import_Quants...');
  
  try {
    const data = sourceSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const colMap = {};
    headers.forEach((header, index) => {
      const h = String(header || '').toLowerCase().trim();
      if (h.includes('article') || h.includes('product')) colMap.article = index;
      if (h.includes('location') || h.includes('warehouse')) colMap.location = index;
      if (h.includes('quantity') || h.includes('qty')) colMap.quantity = index;
    });
    
    if (colMap.article === undefined || colMap.quantity === undefined) {
      console.warn('Required columns (Article, Quantity) not found in Import_Quants');
      return;
    }
    
    // Aggregate data by article
    const aggregates = new Map();
    
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      const article = String(row[colMap.article] || '').trim();
      const quantity = parseFloat(row[colMap.quantity]) || 0;
      
      if (!article) continue;
      
      if (aggregates.has(article)) {
        aggregates.set(article, aggregates.get(article) + quantity);
      } else {
        aggregates.set(article, quantity);
      }
    }
    
    // Create or update Quants_Aggregate sheet
    let aggregateSheet = ss.getSheetByName('Quants_Aggregate');
    if (!aggregateSheet) {
      aggregateSheet = ss.insertSheet('Quants_Aggregate');
    }
    
    // Clear and write data
    aggregateSheet.clear();
    
    const outputData = [['Article', 'Total_Quantity', 'LastUpdated']];
    const now = new Date();
    
    aggregates.forEach((quantity, article) => {
      outputData.push([article, quantity, now]);
    });
    
    if (outputData.length > 1) {
      aggregateSheet.getRange(1, 1, outputData.length, 3).setValues(outputData);
      aggregateSheet.setFrozenRows(1);
      aggregateSheet.autoResizeColumns(1, 3);
    }
    
    console.log(`Quants_Aggregate updated with ${aggregates.size} articles`);
    
  } catch (error) {
    console.error('Error building quants aggregate:', error);
    throw error;
  }
}

/**
 * Builds location-based quants summary from Import_Quants sheet.
 */
function NF_BuildQuantsByLocation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Import_Quants');
  
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    console.log('Import_Quants sheet not found or empty, skipping location summary...');
    return;
  }
  
  console.log('Building Quants_By_Location from Import_Quants...');
  
  try {
    const data = sourceSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const colMap = {};
    headers.forEach((header, index) => {
      const h = String(header || '').toLowerCase().trim();
      if (h.includes('article') || h.includes('product')) colMap.article = index;
      if (h.includes('location') || h.includes('warehouse')) colMap.location = index;
      if (h.includes('quantity') || h.includes('qty')) colMap.quantity = index;
    });
    
    if (colMap.article === undefined || colMap.location === undefined || colMap.quantity === undefined) {
      console.warn('Required columns (Article, Location, Quantity) not found in Import_Quants');
      return;
    }
    
    // Group by location and article
    const locationData = new Map();
    
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      const article = String(row[colMap.article] || '').trim();
      const location = String(row[colMap.location] || '').trim();
      const quantity = parseFloat(row[colMap.quantity]) || 0;
      
      if (!article || !location) continue;
      
      const key = `${location}|${article}`;
      locationData.set(key, (locationData.get(key) || 0) + quantity);
    }
    
    // Create or update Quants_By_Location sheet
    let locationSheet = ss.getSheetByName('Quants_By_Location');
    if (!locationSheet) {
      locationSheet = ss.insertSheet('Quants_By_Location');
    }
    
    // Clear and write data
    locationSheet.clear();
    
    const outputData = [['Location', 'Article', 'Quantity', 'LastUpdated']];
    const now = new Date();
    
    locationData.forEach((quantity, key) => {
      const [location, article] = key.split('|');
      outputData.push([location, article, quantity, now]);
    });
    
    if (outputData.length > 1) {
      locationSheet.getRange(1, 1, outputData.length, 4).setValues(outputData);
      locationSheet.setFrozenRows(1);
      locationSheet.autoResizeColumns(1, 4);
    }
    
    console.log(`Quants_By_Location updated with ${locationData.size} location-article combinations`);
    
  } catch (error) {
    console.error('Error building quants by location:', error);
    throw error;
  }
}

/**
 * Reconciles inventory data between ERP (Quants), 3PL (Warehouse Balance), and PBI data.
 * Creates a reconciliation report showing discrepancies.
 */
function NF_ReconcileInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('Reconciling inventory between ERP, 3PL, and PBI...');
  
  try {
    // Get data from different sources
    const erpData = NF_getSheetData_(ss, 'Quants_Aggregate', ['article', 'total_quantity']);
    const warehouseData = NF_getSheetData_(ss, 'Import_Warehouse_Balance', ['article', 'quantity']);
    const pbiData = NF_getSheetData_(ss, 'Import_PBI_Balance', ['article', 'quantity']);
    
    // Combine all articles
    const allArticles = new Set();
    erpData.forEach(item => allArticles.add(item.article));
    warehouseData.forEach(item => allArticles.add(item.article));
    pbiData.forEach(item => allArticles.add(item.article));
    
    // Create reconciliation data
    const reconcileData = [['Article', 'ERP_Quantity', 'Warehouse_Quantity', 'PBI_Quantity', 'ERP_vs_Warehouse', 'ERP_vs_PBI', 'Status', 'LastUpdated']];
    const now = new Date();
    
    allArticles.forEach(article => {
      const erp = erpData.find(item => item.article === article);
      const warehouse = warehouseData.find(item => item.article === article);
      const pbi = pbiData.find(item => item.article === article);
      
      const erpQty = erp ? erp.quantity : 0;
      const warehouseQty = warehouse ? warehouse.quantity : 0;
      const pbiQty = pbi ? pbi.quantity : 0;
      
      const erpVsWarehouse = erpQty - warehouseQty;
      const erpVsPbi = erpQty - pbiQty;
      
      // Determine status
      let status = 'OK';
      if (Math.abs(erpVsWarehouse) > 0.01 || Math.abs(erpVsPbi) > 0.01) {
        status = 'DISCREPANCY';
      }
      if (!erp && !warehouse && !pbi) {
        status = 'NO_DATA';
      }
      
      reconcileData.push([
        article,
        erpQty,
        warehouseQty,
        pbiQty,
        erpVsWarehouse,
        erpVsPbi,
        status,
        now
      ]);
    });
    
    // Create or update reconciliation sheet
    let reconcileSheet = ss.getSheetByName('Inventory_Reconcile');
    if (!reconcileSheet) {
      reconcileSheet = ss.insertSheet('Inventory_Reconcile');
    }
    
    // Clear and write data
    reconcileSheet.clear();
    
    if (reconcileData.length > 1) {
      reconcileSheet.getRange(1, 1, reconcileData.length, 8).setValues(reconcileData);
      reconcileSheet.setFrozenRows(1);
      reconcileSheet.autoResizeColumns(1, 8);
      
      // Add conditional formatting for discrepancies
      const statusRange = reconcileSheet.getRange(2, 7, reconcileData.length - 1, 1);
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('DISCREPANCY')
        .setBackground('#ffcccb')
        .setRanges([statusRange])
        .build();
      reconcileSheet.setConditionalFormatRules([rule]);
    }
    
    console.log(`Inventory reconciliation completed with ${allArticles.size} articles`);
    
  } catch (error) {
    console.error('Error reconciling inventory:', error);
    throw error;
  }
}

/**
 * Helper function to extract data from a sheet with flexible column mapping.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @param {string} sheetName - Sheet name to read from
 * @param {string[]} requiredFields - Array of required field names (lowercase)
 * @return {Object[]} Array of objects with article and quantity fields
 */
function NF_getSheetData_(ss, sheetName, requiredFields) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices for required fields
  const colMap = {};
  requiredFields.forEach(field => {
    headers.forEach((header, index) => {
      const h = String(header || '').toLowerCase().trim();
      if (h.includes(field.replace('_', '')) || h === field) {
        colMap[field] = index;
      }
    });
  });
  
  // Check if we have required columns
  const articleCol = colMap[requiredFields[0]];
  const quantityCol = colMap[requiredFields[1]];
  
  if (articleCol === undefined || quantityCol === undefined) {
    console.warn(`Required columns not found in ${sheetName}`);
    return [];
  }
  
  // Extract data
  const result = [];
  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    const article = String(row[articleCol] || '').trim();
    const quantity = parseFloat(row[quantityCol]) || 0;
    
    if (article) {
      result.push({ article, quantity });
    }
  }
  
  return result;
}