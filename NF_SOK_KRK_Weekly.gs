/**
 * NF_SOK_KRK_Weekly.gs — SOK and Kärkkäinen weekly report management
 * 
 * Provides functions for building and maintaining SOK and Kärkkäinen weekly reports
 * with historical data merge capabilities and date window utilities.
 * 
 * Key Functions:
 *   - NF_buildSokKarkkainenAlways(): Merge all historical rows into reports
 *   - NF_getLastFinishedWeekSunWindow_(): Get last completed Sunday-to-Sunday week
 *   - NF_ReconcileWeeklyFromImport(): Append missing deliveries from PBI import
 */

// Constants for payer identification
const SOK_FREIGHT_ACCOUNT = '5010';
const KARKKAINEN_NUMBERS = ['1234', '5678', '9012']; // Update with actual Kärkkäinen account numbers

/**
 * Merges all rows historically into Report_SOK and Report_Karkkainen using payer split rules.
 * This function processes all available data to ensure reports are fully up to date.
 */
function NF_buildSokKarkkainenAlways() {
  console.log('NF_buildSokKarkkainenAlways: Building historical SOK/Kärkkäinen reports...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get all source sheets that might contain deliveries
    const sourceSheets = ['Packages', 'Import_Weekly', 'PowerBI_New', 'Packages_Archive'];
    
    let allData = [];
    let unionHeaders = [];
    
    // Collect data from all source sheets
    for (const sheetName of sourceSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) {
        console.log(`Sheet ${sheetName} not found or empty, skipping...`);
        continue;
      }
      
      console.log(`Processing data from ${sheetName}...`);
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      // Merge headers
      unionHeaders = NF_mergeHeaders_(unionHeaders, headers);
      
      // Add rows with source identification
      for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        // Extend row to match union headers and add source info
        const extendedRow = NF_extendRowToHeaders_(row, headers, unionHeaders);
        extendedRow.push(sheetName); // Add source sheet name
        allData.push(extendedRow);
      }
    }
    
    if (allData.length === 0) {
      console.log('No data found in source sheets');
      return;
    }
    
    // Add source column to union headers
    unionHeaders.push('SourceSheet');
    
    // Find payer column
    const payerIndex = NF_findPayerIndex_(unionHeaders);
    if (payerIndex < 0) {
      console.warn('Payer column not found, cannot split SOK/Kärkkäinen');
      return;
    }
    
    // Split data by payer
    const sokData = [];
    const krkData = [];
    
    for (const row of allData) {
      const payerValue = String(row[payerIndex] || '').trim();
      const payerDigits = NF_normalizeDigits_(payerValue);
      
      if (!payerDigits) continue;
      
      if (payerDigits === NF_normalizeDigits_(SOK_FREIGHT_ACCOUNT)) {
        sokData.push(row);
      } else if (KARKKAINEN_NUMBERS.some(n => NF_normalizeDigits_(n) === payerDigits)) {
        krkData.push(row);
      }
    }
    
    // Write to report sheets
    NF_writeReportSheet_(ss, 'Report_SOK', unionHeaders, sokData);
    NF_writeReportSheet_(ss, 'Report_Karkkainen', unionHeaders, krkData);
    
    // Add info row with build timestamp
    const now = new Date();
    const infoRow = new Array(unionHeaders.length).fill('');
    infoRow[0] = `Built: ${now.toISOString()}`;
    infoRow[1] = `SOK rows: ${sokData.length}`;
    infoRow[2] = `Kärkkäinen rows: ${krkData.length}`;
    
    const sokSheet = ss.getSheetByName('Report_SOK');
    const krkSheet = ss.getSheetByName('Report_Karkkainen');
    
    if (sokSheet) {
      sokSheet.insertRowsAfter(1, 1);
      sokSheet.getRange(2, 1, 1, unionHeaders.length).setValues([infoRow]);
    }
    
    if (krkSheet) {
      krkSheet.insertRowsAfter(1, 1);
      krkSheet.getRange(2, 1, 1, unionHeaders.length).setValues([infoRow]);
    }
    
    console.log(`SOK/Kärkkäinen reports built: SOK=${sokData.length} rows, Kärkkäinen=${krkData.length} rows`);
    
    return {
      sokRows: sokData.length,
      karkkainenRows: krkData.length,
      totalProcessed: allData.length
    };
    
  } catch (error) {
    console.error('Error building SOK/Kärkkäinen reports:', error);
    throw error;
  }
}

/**
 * Returns the last finished week's Sunday-to-Sunday time window.
 * @return {Object} Object with start and end Date objects
 */
function NF_getLastFinishedWeekSunWindow_() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  
  const dayOfWeek = now.getDay(); // 0 = Sunday
  
  // Calculate the most recent Sunday
  const thisSunday = new Date(now);
  thisSunday.setDate(now.getDate() - dayOfWeek);
  
  // Last finished week ends at this Sunday, starts 7 days before
  const end = new Date(thisSunday);
  const start = new Date(thisSunday);
  start.setDate(thisSunday.getDate() - 7);
  
  console.log(`Last finished week: ${start.toISOString().split('T')[0]} to ${end.toISOString().split('T')[0]}`);
  
  return { start, end };
}

/**
 * Appends missing deliveries from PBI import to SOK/KRK weekly reports.
 * This ensures that any last-minute PBI deliveries are included in the reports.
 */
function NF_ReconcileWeeklyFromImport() {
  console.log('NF_ReconcileWeeklyFromImport: Reconciling weekly reports with PBI import...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Import_Weekly');
  
  if (!importSheet || importSheet.getLastRow() < 2) {
    console.log('Import_Weekly sheet not found or empty, skipping reconciliation...');
    return;
  }
  
  try {
    const importData = importSheet.getDataRange().getValues();
    const importHeaders = importData[0];
    
    // Find relevant columns
    const payerIndex = NF_findPayerIndex_(importHeaders);
    const dateIndex = NF_findDateIndex_(importHeaders);
    
    if (payerIndex < 0) {
      console.warn('Payer column not found in Import_Weekly');
      return;
    }
    
    // Get last week's window
    const { start, end } = NF_getLastFinishedWeekSunWindow_();
    
    // Filter import data for last week
    const lastWeekData = [];
    for (let rowIndex = 1; rowIndex < importData.length; rowIndex++) {
      const row = importData[rowIndex];
      
      // Check date if available
      if (dateIndex >= 0) {
        const dateValue = row[dateIndex];
        if (dateValue) {
          const rowDate = new Date(dateValue);
          if (rowDate < start || rowDate >= end) {
            continue; // Outside last week window
          }
        }
      }
      
      lastWeekData.push(row);
    }
    
    if (lastWeekData.length === 0) {
      console.log('No data found for last week in Import_Weekly');
      return;
    }
    
    // Split by payer and check for missing entries
    const sokEntries = [];
    const krkEntries = [];
    
    for (const row of lastWeekData) {
      const payerValue = String(row[payerIndex] || '').trim();
      const payerDigits = NF_normalizeDigits_(payerValue);
      
      if (!payerDigits) continue;
      
      if (payerDigits === NF_normalizeDigits_(SOK_FREIGHT_ACCOUNT)) {
        sokEntries.push(row);
      } else if (KARKKAINEN_NUMBERS.some(n => NF_normalizeDigits_(n) === payerDigits)) {
        krkEntries.push(row);
      }
    }
    
    // Append to existing reports if there are new entries
    if (sokEntries.length > 0) {
      NF_appendToReportSheet_(ss, 'Report_SOK', importHeaders, sokEntries);
    }
    
    if (krkEntries.length > 0) {
      NF_appendToReportSheet_(ss, 'Report_Karkkainen', importHeaders, krkEntries);
    }
    
    console.log(`Reconciliation completed: ${sokEntries.length} SOK entries, ${krkEntries.length} Kärkkäinen entries appended`);
    
  } catch (error) {
    console.error('Error in weekly reconciliation:', error);
    throw error;
  }
}

/**
 * Helper function to merge headers from multiple sources.
 * @param {string[]} oldHeaders - Existing headers
 * @param {string[]} newHeaders - New headers to merge
 * @return {string[]} Merged headers array
 */
function NF_mergeHeaders_(oldHeaders, newHeaders) {
  const result = oldHeaders.slice(); // Copy existing headers
  const existing = new Set(oldHeaders.map(h => String(h || '').toLowerCase().trim()));
  
  for (const header of newHeaders) {
    const normalized = String(header || '').toLowerCase().trim();
    if (!existing.has(normalized)) {
      result.push(header);
      existing.add(normalized);
    }
  }
  
  return result;
}

/**
 * Helper function to extend a row to match union headers.
 * @param {any[]} row - Original row data
 * @param {string[]} sourceHeaders - Headers from source sheet
 * @param {string[]} unionHeaders - Target union headers
 * @return {any[]} Extended row matching union headers
 */
function NF_extendRowToHeaders_(row, sourceHeaders, unionHeaders) {
  const result = new Array(unionHeaders.length).fill('');
  
  // Map source data to union header positions
  for (let i = 0; i < sourceHeaders.length && i < row.length; i++) {
    const sourceHeader = String(sourceHeaders[i] || '').toLowerCase().trim();
    const unionIndex = unionHeaders.findIndex(h => 
      String(h || '').toLowerCase().trim() === sourceHeader
    );
    
    if (unionIndex >= 0) {
      result[unionIndex] = row[i];
    }
  }
  
  return result;
}

/**
 * Helper function to find payer column index.
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NF_findPayerIndex_(headers) {
  const payerCandidates = ['payer', 'freight account', 'billing', 'customer', 'account'];
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();
    if (payerCandidates.some(candidate => header.includes(candidate))) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Helper function to find date column index.
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NF_findDateIndex_(headers) {
  const dateCandidates = ['date', 'created', 'submitted', 'booking', 'dispatch', 'shipped', 'timestamp'];
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();
    if (dateCandidates.some(candidate => header.includes(candidate))) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Helper function to normalize digits from a string.
 * @param {string} value - String to normalize
 * @return {string} Normalized digits only
 */
function NF_normalizeDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

/**
 * Helper function to write data to a report sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @param {string} sheetName - Target sheet name
 * @param {string[]} headers - Headers array
 * @param {any[][]} data - Data rows
 */
function NF_writeReportSheet_(ss, sheetName, headers, data) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Clear existing data
  sheet.clear();
  
  // Write headers
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  
  // Write data
  if (data.length > 0 && headers.length > 0) {
    // Ensure all data rows match header length
    const normalizedData = data.map(row => {
      const normalizedRow = new Array(headers.length).fill('');
      for (let i = 0; i < Math.min(row.length, headers.length); i++) {
        normalizedRow[i] = row[i];
      }
      return normalizedRow;
    });
    
    sheet.getRange(2, 1, normalizedData.length, headers.length).setValues(normalizedData);
  }
  
  // Auto-resize columns
  if (headers.length > 0) {
    sheet.autoResizeColumns(1, Math.min(headers.length, 20));
  }
}

/**
 * Helper function to append data to an existing report sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @param {string} sheetName - Target sheet name
 * @param {string[]} sourceHeaders - Headers from source data
 * @param {any[][]} data - Data rows to append
 */
function NF_appendToReportSheet_(ss, sheetName, sourceHeaders, data) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.warn(`Report sheet ${sheetName} not found for appending`);
    return;
  }
  
  if (data.length === 0) {
    return;
  }
  
  // Get existing headers
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Map source data to existing header structure
  const mappedData = data.map(row => 
    NF_extendRowToHeaders_(row, sourceHeaders, existingHeaders)
  );
  
  // Append data
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, mappedData.length, existingHeaders.length).setValues(mappedData);
  
  console.log(`Appended ${mappedData.length} rows to ${sheetName}`);
}