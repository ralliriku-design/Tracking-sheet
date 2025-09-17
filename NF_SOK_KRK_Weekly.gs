/**
 * NF_SOK_KRK_Weekly.gs - SOK/Kärkkäinen Weekly Reports
 * 
 * Builds weekly reports for SOK and Kärkkäinen based on freight payer information.
 * Uses Sunday→Sunday week windows and Script Properties for payer identification.
 */

/********************* WEEKLY WINDOW LOGIC ***************************/

/**
 * Get last finished week (Sunday → Sunday) window
 * @returns {Object} {start: Date, end: Date}
 */
function NF_getLastFinishedWeekSunWindow_() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const dayOfWeek = now.getDay(); // 0 = Sunday
  
  // Calculate this Sunday
  const thisSunday = new Date(now);
  thisSunday.setDate(now.getDate() - dayOfWeek);
  
  // Last finished week is the previous week
  const end = thisSunday;
  const start = new Date(end);
  start.setDate(end.getDate() - 7);
  
  return { start, end };
}

/********************* SOK/KARKKAINEN BUILDERS ***************************/

/**
 * Build SOK and Kärkkäinen reports for specific week window
 * @param {Date} start - Week start (Sunday)
 * @param {Date} end - Week end (Sunday)
 */
function NF_buildSokKarkkainenForWindow_(start, end) {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages';
  const sheet = ss.getSheetByName(targetSheet);
  
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(`"${targetSheet}" is empty or missing.`);
  }
  
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(h => String(h || '').trim());
  const rows = data.slice(1);
  
  // Filter rows by date window
  const dateIndex = NF_pickDateIndex_(headers);
  const inWindow = rows.filter(row => {
    if (dateIndex < 0) return true; // Include all if no date column
    const dateValue = NF_parseDateFlexible_(row[dateIndex]);
    return dateValue && dateValue >= start && dateValue < end;
  });
  
  // Split by payer
  const payerIndex = NF_pickPayerIndex_(headers);
  const sokRows = [];
  const karkkainenRows = [];
  
  if (payerIndex >= 0) {
    const sokAccount = PropertiesService.getScriptProperties().getProperty('SOK_FREIGHT_ACCOUNT') || '990719901';
    const karkkainenNumbers = (PropertiesService.getScriptProperties().getProperty('KARKKAINEN_NUMBERS') || '615471,802669,7030057').split(',');
    
    const sokNormalized = NF_normalizeDigits_(sokAccount);
    const karkkainenSet = new Set(karkkainenNumbers.map(NF_normalizeDigits_));
    
    for (const row of inWindow) {
      const payerValue = NF_normalizeDigits_(row[payerIndex]);
      if (!payerValue) continue;
      
      if (payerValue === sokNormalized) {
        sokRows.push(row);
      } else if (karkkainenSet.has(payerValue)) {
        karkkainenRows.push(row);
      }
    }
  }
  
  // Write sheets
  NF_writeWeeklySheet_(ss, 'Report_SOK', headers, sokRows, start, end);
  NF_writeWeeklySheet_(ss, 'Report_Karkkainen', karkkainenRows, headers, start, end);
  
  ss.toast(`SOK=${sokRows.length} | Kärkkäinen=${karkkainenRows.length}`, 'Weekly Reports Built');
}

/**
 * Build SOK and Kärkkäinen reports from all historical data
 */
function NF_buildSokKarkkainenAlways() {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages';
  const archiveSheet = PropertiesService.getScriptProperties().getProperty('ARCHIVE_SHEET') || 'Packages_Archive';
  
  // Collect all tables (similar to existing pattern)
  const tables = [];
  
  const addTable = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getDisplayValues();
      tables.push({
        headers: data[0].map(String),
        rows: data.slice(1)
      });
    }
  };
  
  addTable(targetSheet);
  addTable(archiveSheet);
  
  if (!tables.length) {
    throw new Error('No data found in source tables.');
  }
  
  // Merge headers from all tables
  let unionHeaders = [];
  tables.forEach(table => {
    unionHeaders = NF_mergeHeaders_(unionHeaders, table.headers);
  });
  
  // Collect and classify all rows
  const sokRows = [];
  const karkkainenRows = [];
  const payerIndex = NF_pickPayerIndex_(unionHeaders);
  
  if (payerIndex >= 0) {
    const sokAccount = PropertiesService.getScriptProperties().getProperty('SOK_FREIGHT_ACCOUNT') || '990719901';
    const karkkainenNumbers = (PropertiesService.getScriptProperties().getProperty('KARKKAINEN_NUMBERS') || '615471,802669,7030057').split(',');
    
    const sokNormalized = NF_normalizeDigits_(sokAccount);
    const karkkainenSet = new Set(karkkainenNumbers.map(NF_normalizeDigits_));
    const headerMap = NF_headerIndexMap_(unionHeaders);
    
    for (const table of tables) {
      const sourceMap = NF_headerIndexMap_(table.headers);
      
      for (const sourceRow of table.rows) {
        // Skip completely empty rows
        if (!sourceRow.some(cell => String(cell || '').trim() !== '')) continue;
        
        // Map to union header structure
        const unifiedRow = new Array(unionHeaders.length).fill('');
        for (const [headerName, sourceIndex] of Object.entries(sourceMap)) {
          const targetIndex = headerMap[headerName];
          if (typeof targetIndex === 'number') {
            unifiedRow[targetIndex] = sourceRow[sourceIndex];
          }
        }
        
        // Classify by payer
        const payerValue = NF_normalizeDigits_(unifiedRow[payerIndex]);
        if (!payerValue) continue;
        
        if (payerValue === sokNormalized) {
          sokRows.push(unifiedRow);
        } else if (karkkainenSet.has(payerValue)) {
          karkkainenRows.push(unifiedRow);
        }
      }
    }
  }
  
  // Write flat reports (no week window info)
  NF_writeFlat_(ss, 'Report_SOK', unionHeaders, sokRows);
  NF_writeFlat_(ss, 'Report_Karkkainen', unionHeaders, karkkainenRows);
  
  ss.toast(`SOK=${sokRows.length} | Kärkkäinen=${karkkainenRows.length}`, 'Full Reports Built');
}

/********************* HELPER FUNCTIONS ***************************/

/**
 * Write weekly sheet with info header rows
 */
function NF_writeWeeklySheet_(spreadsheet, sheetName, headers, rows, start, end) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  sheet.clear();
  
  // Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Info row
  const infoRow = new Array(headers.length).fill('');
  infoRow[0] = `Week (SUN→SUN): ${NF_dateToYMD_(start)} - ${NF_dateToYMD_(end)}`;
  if (headers.length > 1) infoRow[1] = `Rows: ${rows.length}`;
  if (headers.length > 2) infoRow[2] = `Created: ${NF_formatDateTime_(new Date())}`;
  
  sheet.getRange(2, 1, 1, headers.length).setValues([infoRow]);
  sheet.getRange(2, 1, 1, Math.min(headers.length, 3)).setFontStyle('italic');
  
  // Data rows
  if (rows && rows.length) {
    const normalizedRows = rows.map(row => {
      const normalizedRow = row.slice(0, headers.length);
      while (normalizedRow.length < headers.length) {
        normalizedRow.push('');
      }
      return normalizedRow;
    });
    sheet.getRange(4, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }
  
  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, Math.min(headers.length, 20));
}

/**
 * Write flat sheet (no week info, for full rebuild)
 */
function NF_writeFlat_(spreadsheet, sheetName, headers, rows) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  sheet.clear();
  
  // Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Data rows
  if (rows && rows.length) {
    const normalizedRows = rows.map(row => {
      const normalizedRow = row.slice(0, headers.length);
      while (normalizedRow.length < headers.length) {
        normalizedRow.push('');
      }
      return normalizedRow;
    });
    sheet.getRange(2, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }
  
  sheet.autoResizeColumns(1, Math.min(headers.length, 20));
}

/**
 * Pick date index for filtering window
 */
function NF_pickDateIndex_(headers) {
  const dateCandidates = [
    'Submitted date', 'Created', 'Created date', 'Booking date', 'Booked time',
    'Dispatch date', 'Shipped date', 'Timestamp', 'Date'
  ];
  
  return NF_pickAnyIndex_(headers, dateCandidates);
}

/**
 * Pick payer index for SOK/Kärkkäinen classification
 */
function NF_pickPayerIndex_(headers) {
  const payerCandidates = [
    'Payer', 'Freight account', 'Billing account', 'Customer number', 'Customer ID', 'Customer #'
  ];
  
  return NF_pickAnyIndex_(headers, payerCandidates);
}

/**
 * Pick any index from candidates (similar to existing pickAnyIndex_)
 */
function NF_pickAnyIndex_(headers, candidates) {
  if (typeof pickAnyIndex_ === 'function') {
    return pickAnyIndex_(headers, candidates);
  }
  
  const normalized = headers.map(h => NF_normalize_(h));
  for (const candidate of candidates) {
    const index = normalized.indexOf(NF_normalize_(candidate));
    if (index >= 0) return index;
  }
  return -1;
}

/**
 * Normalize text (similar to existing normalize_)
 */
function NF_normalize_(text) {
  return String(text || '').toLowerCase().trim().replace(/\s+/g, ' ').replace(/[^\p{L}\p{N}]+/gu, ' ');
}

/**
 * Normalize digits only (for payer matching)
 */
function NF_normalizeDigits_(text) {
  return String(text || '').replace(/\D/g, '');
}

/**
 * Parse date flexibly
 */
function NF_parseDateFlexible_(value) {
  if (!value) return null;
  try {
    const date = new Date(value);
    return isNaN(date.getTime()) ? null : date;
  } catch (e) {
    return null;
  }
}

/**
 * Format date as YYYY-MM-DD
 */
function NF_dateToYMD_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return date.toISOString().split('T')[0];
  }
}

/**
 * Format datetime
 */
function NF_formatDateTime_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return date.toISOString().replace('T', ' ').substring(0, 19);
  }
}

/**
 * Merge header arrays
 */
function NF_mergeHeaders_(arr1, arr2) {
  if (typeof mergeHeaders_ === 'function') {
    return mergeHeaders_(arr1, arr2);
  }
  
  return Array.from(new Set([...(arr1 || []), ...(arr2 || [])]));
}

/**
 * Create header index map
 */
function NF_headerIndexMap_(headers) {
  if (typeof headerIndexMap_ === 'function') {
    return headerIndexMap_(headers);
  }
  
  const map = {};
  (headers || []).forEach((header, index) => {
    map[header] = index;
  });
  return map;
}