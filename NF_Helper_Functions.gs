/**
 * NF_Helper_Functions.gs — Helper functions for NF extension
 * 
 * Provides fallback implementations of helper functions that might be missing
 * or ensures compatibility with the existing codebase.
 */

/**
 * Safe implementation of headerIndexMap_ if not available elsewhere.
 * Creates a map from header names to column indices.
 */
function NF_safeHeaderIndexMap_(headers) {
  // Try to use existing function first
  if (typeof headerIndexMap_ === 'function') {
    return headerIndexMap_(headers);
  }
  
  // Fallback implementation
  const map = {};
  (headers || []).forEach((header, index) => {
    map[header] = index;
  });
  return map;
}

/**
 * Safe implementation of firstCode_ if not available elsewhere.
 * Extracts the first tracking code from a cell value.
 */
function NF_safeFirstCode_(value) {
  // Try to use existing function first
  if (typeof firstCode_ === 'function') {
    return firstCode_(value);
  }
  
  // Fallback implementation
  let s = String(value || '').trim();
  if (!s) return '';
  
  // If multiple codes in one cell, split and take first
  const parts = s.split(/[,\n;]/).map(x => String(x || '').trim()).filter(x => x);
  return parts.length ? parts[0] : s;
}

/**
 * Safe implementation of fmtDateTime_ if not available elsewhere.
 * Formats a date for display.
 */
function NF_safeFmtDateTime_(date) {
  // Try to use existing function first
  if (typeof fmtDateTime_ === 'function') {
    return fmtDateTime_(date);
  }
  
  // Fallback implementation
  if (!date) return '';
  
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    
    return d.toISOString().replace('T', ' ').substring(0, 19);
  } catch (error) {
    return '';
  }
}

/**
 * Safe way to get or create a sheet.
 * @param {string} sheetName - Name of the sheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Sheet object
 */
function NF_safeGetOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Safe way to check if a value indicates delivery.
 * @param {any} value - Value to check
 * @return {boolean} True if value indicates delivery
 */
function NF_safeIsDelivered_(value) {
  const str = String(value || '').toLowerCase();
  const deliveredKeywords = [
    'delivered', 'toimitettu', 'luovutettu', 'delivered to pickup point',
    'delivered to parcel locker', 'delivered to recipient', 'delivered - picked up'
  ];
  
  return deliveredKeywords.some(keyword => str.includes(keyword));
}

/**
 * Safe way to parse a flexible date format.
 * @param {any} value - Date value to parse
 * @return {Date|null} Parsed date or null
 */
function NF_safeParseDateFlexible_(value) {
  if (!value) return null;
  
  try {
    // Try direct Date parsing first
    const direct = new Date(value);
    if (!isNaN(direct.getTime())) {
      return direct;
    }
    
    // Try to handle common Finnish date formats
    const str = String(value).trim();
    
    // dd.MM.yyyy HH:mm format
    const ddmmyyyy = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s*(\d{1,2}):(\d{2})?/);
    if (ddmmyyyy) {
      const [, day, month, year, hour, minute] = ddmmyyyy;
      return new Date(parseInt(year), parseInt(month) - 1, parseInt(day), 
                     parseInt(hour) || 0, parseInt(minute) || 0);
    }
    
    // yyyy-MM-dd format
    const yyyymmdd = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (yyyymmdd) {
      const [, year, month, day] = yyyymmdd;
      return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * Safe way to get configuration integer from Script Properties.
 * @param {string} key - Property key
 * @param {number} fallback - Fallback value
 * @return {number} Configuration value
 */
function NF_safeCfgInt_(key, fallback) {
  // Try to use existing function first
  if (typeof getCfgInt_ === 'function') {
    return getCfgInt_(key, fallback);
  }
  
  // Fallback implementation
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    const parsed = parseInt(value);
    return isNaN(parsed) ? fallback : parsed;
  } catch (error) {
    return fallback;
  }
}

/**
 * Safe error logging function.
 * @param {string} context - Context or function name where error occurred
 * @param {Error} error - Error object
 */
function NF_safeLogError_(context, error) {
  // Try to use existing function first
  if (typeof logError_ === 'function') {
    return logError_(context, error);
  }
  
  // Fallback implementation
  console.error(`[${context}] Error:`, error);
  
  try {
    // Try to log to Error_Log sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('Error_Log');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('Error_Log');
      errorSheet.getRange(1, 1, 1, 4).setValues([
        ['Timestamp', 'Context', 'Message', 'Stack']
      ]);
    }
    
    errorSheet.appendRow([
      new Date(),
      context,
      error.message || String(error),
      error.stack || ''
    ]);
    
  } catch (logError) {
    console.error('Failed to log error to sheet:', logError);
  }
}

/**
 * Checks if all required helper functions are available.
 * @return {Object} Status of helper function availability
 */
function NF_checkHelperFunctions() {
  const functions = [
    { name: 'headerIndexMap_', required: true },
    { name: 'firstCode_', required: true },
    { name: 'fmtDateTime_', required: false },
    { name: 'getCfgInt_', required: false },
    { name: 'logError_', required: false },
    { name: 'TRK_trackByCarrierEnhanced', required: false },
    { name: 'TRK_trackByCarrier_', required: false }
  ];
  
  const status = { available: [], missing: [], critical: [] };
  
  for (const func of functions) {
    const isAvailable = typeof eval(`typeof ${func.name}`) === 'function';
    
    if (isAvailable) {
      status.available.push(func.name);
    } else {
      status.missing.push(func.name);
      if (func.required) {
        status.critical.push(func.name);
      }
    }
  }
  
  return status;
}