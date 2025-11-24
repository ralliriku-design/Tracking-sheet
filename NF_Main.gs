/**
 * NF_Main.gs - New Flow Core Functionality
 * 
 * Provides the main entry points for the new delivery lead time tracking system:
 * - Daily flow: Gmail import → rebuild → status refresh
 * - Weekly reports: SOK/Kärkkäinen split with tracking refresh
 * - Delivery time analytics: detailed list and country KPI
 * 
 * Uses "keikka tehty" logic: pickup date → submitted date → created date (PBI fallback)
 * Delivered time strictly from tracking events, not PBI timestamps.
 */

/********************* PUBLIC FLOWS ***************************/

/**
 * Daily flow: Import latest nShift Packages report → rebuild → refresh statuses
 */
function NF_RunDailyFlow() {
  const ss = SpreadsheetApp.getActive();
  const startTime = Date.now();
  
  try {
    // Import Gmail attachment
    const attachment = NF_findLatestPackagesAttachment_();
    if (!attachment) {
      throw new Error('No nShift Packages report found in Gmail');
    }
    
    const values = NF_readAttachmentToValues_(attachment.blob, attachment.filename);
    const matrix = NF_sanitizeMatrix_(values);
    
    // Rebuild Packages/Archive with merged headers
    NF_rebuildWithArchive_(matrix);
    
    // Optionally refresh statuses on ACTION_SHEET
    const actionSheet = PropertiesService.getScriptProperties().getProperty('ACTION_SHEET') || 'Vaatii_toimenpiteitä';
    if (ss.getSheetByName(actionSheet)) {
      try {
        if (typeof refreshStatuses_Sheet === 'function') {
          refreshStatuses_Sheet(actionSheet);
        }
      } catch (e) {
        console.log('Warning: Could not refresh action sheet:', e.message);
      }
    }
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    ss.toast(`Daily flow completed in ${duration}s`, 'New Flow', 5);
    
  } catch (error) {
    ss.toast(`Daily flow failed: ${error.message}`, 'New Flow Error', 10);
    throw error;
  }
}

/**
 * Build weekly reports for SOK and Kärkkäinen, then refresh their statuses
 */
function NF_BuildWeeklyReports() {
  const ss = SpreadsheetApp.getActive();
  const startTime = Date.now();
  
  try {
    // Build SOK/Kärkkäinen reports for last finished week
    const { start, end } = NF_getLastFinishedWeekSunWindow_();
    NF_buildSokKarkkainenForWindow_(start, end);
    
    // Refresh statuses for the weekly reports
    NF_RefreshWeeklyStatuses();
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    ss.toast(`Weekly reports built and refreshed in ${duration}s`, 'New Flow', 5);
    
  } catch (error) {
    ss.toast(`Weekly reports failed: ${error.message}`, 'New Flow Error', 10);
    throw error;
  }
}

/**
 * Refresh tracking statuses for weekly reports (Report_SOK, Report_Karkkainen)
 */
function NF_RefreshWeeklyStatuses() {
  const reports = ['Report_SOK', 'Report_Karkkainen'];
  
  for (const sheetName of reports) {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 3) { // Skip if only headers + info rows
      try {
        if (typeof refreshStatuses_Sheet === 'function') {
          refreshStatuses_Sheet(sheetName);
        } else {
          // Fallback: use ensureRefreshCols pattern
          NF_refreshStatusesFallback_(sheet);
        }
      } catch (e) {
        console.log(`Warning: Could not refresh ${sheetName}:`, e.message);
      }
    }
  }
}

/**
 * Build detailed delivery times list
 */
function NF_BuildDeliveryTimes() {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages';
  const sheet = ss.getSheetByName(targetSheet);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Missing sheet: ${targetSheet}`);
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('No data found.');
    return;
  }
  
  const map = NF_headerIndexMap_(data[0]);
  const rows = [['Carrier', 'DestCountry', 'Start(keikka tehty)', 'Delivered', 'LeadTime(days)', 'Source', 'Submitted', 'Pickup', 'Created']];
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const carrier = row[NF_colIndexOf_(map, ['Carrier'])] || '';
    const destCountry = row[NF_colIndexOf_(map, ['Country', 'Dest Country', 'Destination Country'])] || '';
    
    const startInfo = NF_parseStartTime_(row, map);
    const deliveredInfo = NF_pickDeliveredTime_(row, map);
    
    let leadTime = '';
    if (startInfo.time && deliveredInfo.time) {
      const startMs = new Date(startInfo.time).getTime();
      const deliveredMs = new Date(deliveredInfo.time).getTime();
      if (isFinite(startMs) && isFinite(deliveredMs) && deliveredMs > startMs) {
        leadTime = ((deliveredMs - startMs) / 86400000).toFixed(2);
      }
    }
    
    rows.push([
      carrier,
      destCountry,
      startInfo.time || '',
      deliveredInfo.time || '',
      leadTime,
      startInfo.source || '',
      startInfo.submitted || '',
      startInfo.pickup || '',
      startInfo.created || ''
    ]);
  }
  
  const reportSheet = ss.getSheetByName('Delivery_Times') || ss.insertSheet('Delivery_Times');
  reportSheet.clear();
  reportSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  
  SpreadsheetApp.getUi().alert('Delivery times report ready: Delivery_Times');
}

/**
 * Build country/week lead time KPI
 */
function NF_MakeCountryWeekLeadtime() {
  const ss = SpreadsheetApp.getActive();
  const deliverySheet = ss.getSheetByName('Delivery_Times');
  
  if (!deliverySheet) {
    SpreadsheetApp.getUi().alert('Run "Delivery Times list" first.');
    return;
  }
  
  const data = deliverySheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('Delivery_Times is empty.');
    return;
  }
  
  const header = data[0];
  const carrierIdx = header.indexOf('Carrier');
  const countryIdx = header.indexOf('DestCountry');
  const deliveredIdx = header.indexOf('Delivered');
  const leadTimeIdx = header.indexOf('LeadTime(days)');
  
  const aggregated = {};
  
  for (let r = 1; r < data.length; r++) {
    const carrier = String(data[r][carrierIdx] || '').trim();
    const country = String(data[r][countryIdx] || '').trim();
    const delivered = data[r][deliveredIdx];
    const leadTime = parseFloat(data[r][leadTimeIdx]);
    
    if (!carrier || !country || !delivered || !isFinite(leadTime)) continue;
    
    const weekKey = NF_yearIsoWeek_(new Date(delivered));
    const key = `${weekKey}|${country}|${carrier}`;
    
    if (!aggregated[key]) {
      aggregated[key] = { sum: 0, count: 0 };
    }
    aggregated[key].sum += leadTime;
    aggregated[key].count += 1;
  }
  
  const rows = [['ISOWeek', 'Country', 'Carrier', 'AvgLeadTime(days)', 'Count']];
  Object.keys(aggregated).sort().forEach(key => {
    const [weekKey, country, carrier] = key.split('|');
    const agg = aggregated[key];
    rows.push([weekKey, country, carrier, (agg.sum / agg.count).toFixed(2), agg.count]);
  });
  
  const reportSheet = ss.getSheetByName('Leadtime_Weekly_Country') || ss.insertSheet('Leadtime_Weekly_Country');
  reportSheet.clear();
  reportSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  
  SpreadsheetApp.getUi().alert('Country week report ready: Leadtime_Weekly_Country');
}

/********************* HELPER FUNCTIONS ***************************/

/**
 * Parse start time using "keikka tehty" priority: pickup → submitted → created
 */
function NF_parseStartTime_(row, headerMap) {
  const pickupIdx = NF_colIndexOf_(headerMap, ['Pick up date', 'Pick up', 'Pickup date']);
  const submittedIdx = NF_colIndexOf_(headerMap, ['Submitted date', 'Submitted', 'B']);
  const createdIdx = NF_colIndexOf_(headerMap, ['Created Date', 'Created']);
  
  const pickup = pickupIdx >= 0 ? row[pickupIdx] : '';
  const submitted = submittedIdx >= 0 ? row[submittedIdx] : '';
  const created = createdIdx >= 0 ? row[createdIdx] : '';
  
  let time = '';
  let source = '';
  
  if (pickup) {
    time = pickup;
    source = 'pickup';
  } else if (submitted) {
    time = submitted;
    source = 'submitted';
  } else if (created) {
    time = created;
    source = 'created';
  }
  
  return {
    time: time,
    source: source,
    pickup: pickup,
    submitted: submitted,
    created: created
  };
}

/**
 * Pick delivered time from tracking refresh.
 * 
 * IMPORTANT: RefreshTime should NOT be used as delivery time by itself!
 * RefreshTime = last time we checked tracking status (not when package delivered)
 * 
 * Priority logic (FIXED):
 * 1. "Delivered date (Confirmed)" - PBI confirmed delivery (most reliable)
 * 2. "Delivered Time" / "Delivered At" - from tracking events when delivered
 * 3. "RefreshTime" ONLY if RefreshStatus indicates delivery (e.g., "Delivered", "Toimitettu")
 * 
 * This fixes the issue where RefreshTime was incorrectly used as delivery date
 * even when package was still in transit.
 * 
 * @param {Array} row - Data row
 * @param {Object} headerMap - Header index map
 * @return {Object} {time: date string, source: source name}
 */
function NF_pickDeliveredTime_(row, headerMap) {
  const deliveredIdx = NF_colIndexOf_(headerMap, ['Delivered Time', 'Delivered At', 'Delivered']);
  const refreshIdx = NF_colIndexOf_(headerMap, ['RefreshTime']);
  const refreshStatusIdx = NF_colIndexOf_(headerMap, ['RefreshStatus', 'Refresh Status']);
  const confirmedIdx = NF_colIndexOf_(headerMap, ['Delivered date (Confirmed)']);
  
  let time = '';
  let source = 'none';
  
  // Priority 1: Confirmed delivery from PBI
  if (confirmedIdx >= 0 && row[confirmedIdx]) {
    time = row[confirmedIdx];
    source = 'confirmed';
    return { time: time, source: source };
  }
  
  // Priority 2: Delivered Time from tracking events
  if (deliveredIdx >= 0 && row[deliveredIdx]) {
    time = row[deliveredIdx];
    source = 'delivered';
    return { time: time, source: source };
  }
  
  // Priority 3: RefreshTime ONLY if status indicates delivered
  // Check RefreshStatus to ensure package is actually delivered
  if (refreshIdx >= 0 && row[refreshIdx] && refreshStatusIdx >= 0) {
    const status = String(row[refreshStatusIdx] || '').toLowerCase();
    // Only use RefreshTime if status clearly indicates delivery
    if (status.includes('delivered') || 
        status.includes('toimitettu') || 
        status.includes('luovutettu') ||
        status.includes('utlevert')) {
      time = row[refreshIdx];
      source = 'refresh_delivered';
      return { time: time, source: source };
    }
  }
  
  // No delivery date found
  return { time: '', source: 'none' };
}

/**
 * Convert date to ISO week format YYYY-'W'ww
 */
function NF_yearIsoWeek_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "YYYY-'W'ww");
  } catch (e) {
    return '';
  }
}

/**
 * Find latest nShift Packages report attachment from Gmail
 */
function NF_findLatestPackagesAttachment_() {
  const gmailQuery = PropertiesService.getScriptProperties().getProperty('GMAIL_QUERY_PACKAGES') || 
                    PropertiesService.getScriptProperties().getProperty('GMAIL_QUERY') ||
                    'label:"Shipment Report" has:attachment (filename:xlsx OR filename:csv)';
  
  const attachRegex = new RegExp(
    PropertiesService.getScriptProperties().getProperty('ATTACH_ALLOW_REGEX') || 
    '(?:^|\\b)(Packages[ _-]?Report)(?:\\b|$)', 'i'
  );
  
  const threads = GmailApp.search(gmailQuery, 0, 50);
  let best = null;
  
  for (const thread of threads) {
    for (const message of thread.getMessages().reverse()) {
      const attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true }) || [];
      for (const attachment of attachments) {
        const name = (attachment.getName() || '').toLowerCase();
        if (!(name.endsWith('.xlsx') || name.endsWith('.csv'))) continue;
        if (!attachRegex.test(attachment.getName())) continue;
        
        if (!best || message.getDate() > best.date) {
          best = {
            blob: attachment.copyBlob(),
            filename: attachment.getName(),
            date: message.getDate()
          };
        }
      }
    }
  }
  
  return best;
}

/**
 * Read attachment (XLSX/CSV) to values matrix
 */
function NF_readAttachmentToValues_(blob, filename) {
  if (typeof readAttachmentToValues_ === 'function') {
    // Use existing implementation if available
    return readAttachmentToValues_(blob, filename).values || readAttachmentToValues_(blob, filename);
  }
  
  const name = String(filename || '').toLowerCase();
  if (name.endsWith('.csv')) {
    const csv = Utilities.parseCsv(blob.getDataAsString());
    return csv;
  }
  
  if (name.endsWith('.xlsx')) {
    // Convert to Google Sheet and read first sheet
    let tempId = null;
    try {
      const file = Drive.Files.insert(
        { title: filename, mimeType: 'application/vnd.google-apps.spreadsheet' },
        blob,
        { convert: true }
      );
      tempId = file.id;
      const tmpSs = SpreadsheetApp.openById(tempId);
      const sheet = tmpSs.getSheets()[0];
      const values = sheet.getDataRange().getValues();
      return values;
    } finally {
      if (tempId) {
        try {
          Drive.Files.remove(tempId);
        } catch (e) {
          // Ignore cleanup errors
        }
      }
    }
  }
  
  throw new Error('Unsupported file type: ' + filename);
}

/**
 * Sanitize data matrix
 */
function NF_sanitizeMatrix_(values) {
  if (typeof sanitizeMatrix_ === 'function') {
    return sanitizeMatrix_(values);
  }
  
  if (!values || !values.length) return [];
  const matrix = values.map(row => row.map(cell => (cell === null || cell === undefined) ? '' : cell));
  
  // Remove BOM and trim headers
  if (typeof matrix[0][0] === 'string') {
    matrix[0][0] = matrix[0][0].replace(/^\uFEFF/, '');
  }
  matrix[0] = matrix[0].map(header => String(header || '').trim());
  
  return matrix;
}

/**
 * Rebuild Packages/Archive with merged headers
 */
function NF_rebuildWithArchive_(matrix) {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages';
  const archiveSheet = PropertiesService.getScriptProperties().getProperty('ARCHIVE_SHEET') || 'Packages_Archive';
  
  let packages = ss.getSheetByName(targetSheet);
  if (!packages) packages = ss.insertSheet(targetSheet);
  
  let archive = ss.getSheetByName(archiveSheet);
  if (!archive) archive = ss.insertSheet(archiveSheet);
  
  // Write to target sheet with merge logic if existing data
  if (packages.getLastRow() < 2) {
    // Empty sheet, just write the matrix
    packages.clear();
    packages.getRange(1, 1, matrix.length, matrix[0].length).setValues(matrix);
  } else {
    // Merge with existing data using existing patterns
    if (typeof writeMerged_ === 'function') {
      const existing = packages.getDataRange().getValues();
      writeMerged_(packages, existing, matrix);
    } else {
      // Simple fallback: append new data
      const existingData = packages.getDataRange().getValues();
      const existingHeaders = existingData[0];
      const newHeaders = matrix[0];
      
      // Merge headers
      const allHeaders = [...existingHeaders];
      newHeaders.forEach(header => {
        if (!allHeaders.includes(header)) {
          allHeaders.push(header);
        }
      });
      
      // Rebuild sheet with merged headers
      packages.clear();
      packages.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
      
      // Add existing data
      for (let r = 1; r < existingData.length; r++) {
        const row = new Array(allHeaders.length).fill('');
        for (let c = 0; c < existingHeaders.length; c++) {
          const headerIdx = allHeaders.indexOf(existingHeaders[c]);
          if (headerIdx >= 0) {
            row[headerIdx] = existingData[r][c];
          }
        }
        packages.appendRow(row);
      }
      
      // Add new data (skip header)
      for (let r = 1; r < matrix.length; r++) {
        const row = new Array(allHeaders.length).fill('');
        for (let c = 0; c < newHeaders.length; c++) {
          const headerIdx = allHeaders.indexOf(newHeaders[c]);
          if (headerIdx >= 0) {
            row[headerIdx] = matrix[r][c];
          }
        }
        packages.appendRow(row);
      }
    }
  }
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

/**
 * Find column index from candidates
 */
function NF_colIndexOf_(headerMap, candidates) {
  if (typeof colIndexOf_ === 'function') {
    return colIndexOf_(headerMap, candidates);
  }
  
  for (const candidate of candidates || []) {
    if (typeof headerMap[candidate] === 'number') {
      return headerMap[candidate];
    }
  }
  return -1;
}

/**
 * Fallback status refresh implementation
 */
function NF_refreshStatusesFallback_(sheet) {
  // Basic implementation - just ensure refresh columns exist
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const refreshCols = [
    'RefreshCarrier', 'RefreshStatus', 'RefreshTime', 'RefreshLocation', 'RefreshRaw', 'RefreshAt',
    'RefreshAttempts', 'RefreshNextAt', 'Delivered date (Confirmed)', 'Delivered_Source'
  ];
  
  let needsUpdate = false;
  const newHeaders = [...headers];
  
  refreshCols.forEach(col => {
    if (!newHeaders.includes(col)) {
      newHeaders.push(col);
      needsUpdate = true;
    }
  });
  
  if (needsUpdate) {
    // Expand sheet with new columns
    const existingData = sheet.getDataRange().getValues();
    sheet.clear();
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    
    for (let r = 1; r < existingData.length; r++) {
      const row = new Array(newHeaders.length).fill('');
      for (let c = 0; c < headers.length; c++) {
        row[c] = existingData[r][c];
      }
      sheet.getRange(r + 1, 1, 1, newHeaders.length).setValues([row]);
    }
  }
}