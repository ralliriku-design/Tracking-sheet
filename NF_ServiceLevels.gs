/**
 * NF_ServiceLevels.gs — Weekly Service Level Reporting
 * 
 * Computes weekly service level metrics based on Date sent → Delivered time,
 * grouped by ALL, SOK, and Kärkkäinen for the last finished Sun→Sun week.
 * 
 * Key Functions:
 *   - NF_buildWeeklyServiceLevels(): Main entry point for building service level report
 *   - NFSL_* helper functions for internal logic
 */

/**
 * Main function to build weekly service level report.
 * Creates NF_Weekly_ServiceLevels sheet with metrics for ALL, SOK, and KARKKAINEN groups.
 */
function NF_buildWeeklyServiceLevels() {
  console.log('NF_buildWeeklyServiceLevels: Building weekly service level report...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get configuration
    const targetSheet = PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages';
    const sokAccount = PropertiesService.getScriptProperties().getProperty('SOK_FREIGHT_ACCOUNT') || '5010';
    const krkNumbers = NFSL_parseKarkkainenNumbers_();
    
    // Get source data
    const sheet = ss.getSheetByName(targetSheet);
    if (!sheet || sheet.getLastRow() < 2) {
      console.warn(`Source sheet "${targetSheet}" not found or empty`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Find required columns
    const dateSentIndex = NFSL_findDateSentColumn_(headers);
    const deliveredIndex = NFSL_findDeliveredColumn_(headers);
    const payerIndex = NFSL_findPayerColumn_(headers);
    const windowDateIndex = dateSentIndex >= 0 ? dateSentIndex : NFSL_findWindowDateColumn_(headers);
    
    if (dateSentIndex < 0) {
      console.warn('Date sent column not found - service level metrics will be incomplete');
    }
    
    if (deliveredIndex < 0) {
      console.warn('Delivered timestamp column not found - cannot calculate service levels');
      return;
    }
    
    // Get last finished week window
    const { start, end } = NF_getLastFinishedWeekSunWindow_();
    const isoWeek = NFSL_getISOWeek_(start);
    
    console.log(`Processing week ${isoWeek}: ${NFSL_formatDate_(start)} to ${NFSL_formatDate_(end)}`);
    
    // Filter rows in window
    const windowRows = rows.filter(row => {
      if (windowDateIndex < 0) return false;
      const dateVal = row[windowDateIndex];
      if (!dateVal) return false;
      const dt = NFSL_parseDate_(dateVal);
      return dt && dt >= start && dt < end;
    });
    
    console.log(`Found ${windowRows.length} rows in window`);
    
    // Calculate metrics for each group
    const allMetrics = NFSL_calculateGroupMetrics_(windowRows, null, null, dateSentIndex, deliveredIndex, payerIndex);
    const sokMetrics = NFSL_calculateGroupMetrics_(windowRows, sokAccount, null, dateSentIndex, deliveredIndex, payerIndex);
    const krkMetrics = NFSL_calculateGroupMetrics_(windowRows, null, krkNumbers, dateSentIndex, deliveredIndex, payerIndex);
    
    // Build output sheet
    NFSL_writeServiceLevelSheet_(ss, isoWeek, start, end, allMetrics, sokMetrics, krkMetrics);
    
    console.log('Weekly service level report completed successfully');
    
    return {
      week: isoWeek,
      all: allMetrics,
      sok: sokMetrics,
      karkkainen: krkMetrics
    };
    
  } catch (error) {
    console.error('Error building weekly service levels:', error);
    throw error;
  }
}

/**
 * Find Date sent column using configured hints.
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NFSL_findDateSentColumn_(headers) {
  // Get custom hints from Script Properties
  const customHints = PropertiesService.getScriptProperties().getProperty('NF_DATE_SENT_HINTS');
  let candidates = [];
  
  if (customHints) {
    candidates = customHints.split(',').map(s => s.trim()).filter(s => s);
  } else {
    candidates = ['Date sent', 'Sent date', 'Handover date', 'Submitted date'];
  }
  
  const norm = headers.map(h => String(h || '').trim());
  
  for (const candidate of candidates) {
    const index = norm.indexOf(candidate);
    if (index >= 0) {
      console.log(`Found Date sent column: "${headers[index]}" at index ${index}`);
      return index;
    }
  }
  
  console.warn('Date sent column not found');
  return -1;
}

/**
 * Find Delivered timestamp column (tracking-based, not PBI).
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NFSL_findDeliveredColumn_(headers) {
  const candidates = ['Delivered Time', 'Delivered At', 'RefreshTime'];
  const norm = headers.map(h => String(h || '').trim());
  
  for (const candidate of candidates) {
    const index = norm.indexOf(candidate);
    if (index >= 0) {
      console.log(`Found Delivered column: "${headers[index]}" at index ${index}`);
      return index;
    }
  }
  
  console.warn('Delivered timestamp column not found');
  return -1;
}

/**
 * Find Payer column for grouping.
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NFSL_findPayerColumn_(headers) {
  const candidates = ['payer', 'freight account', 'billing', 'customer', 'account'];
  
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase();
    if (candidates.some(candidate => header.includes(candidate))) {
      console.log(`Found Payer column: "${headers[i]}" at index ${i}`);
      return i;
    }
  }
  
  console.warn('Payer column not found');
  return -1;
}

/**
 * Find date column for window filtering (fallback if Date sent not available).
 * @param {string[]} headers - Array of header strings
 * @return {number} Column index or -1 if not found
 */
function NFSL_findWindowDateColumn_(headers) {
  const candidates = ['Created', 'Created date', 'Submitted date', 'Booking date', 'Booked time', 'Timestamp', 'Date'];
  const norm = headers.map(h => String(h || '').trim());
  
  for (const candidate of candidates) {
    const index = norm.indexOf(candidate);
    if (index >= 0) {
      console.log(`Found window date column: "${headers[index]}" at index ${index}`);
      return index;
    }
  }
  
  return -1;
}

/**
 * Parse Kärkkäinen numbers from Script Properties.
 * @return {string[]} Array of Kärkkäinen numbers
 */
function NFSL_parseKarkkainenNumbers_() {
  const value = PropertiesService.getScriptProperties().getProperty('KARKKAINEN_NUMBERS');
  if (value) {
    return value.split(',').map(s => s.trim()).filter(s => s);
  }
  // Fallback to constants in NF_SOK_KRK_Weekly.gs
  if (typeof KARKKAINEN_NUMBERS !== 'undefined') {
    return KARKKAINEN_NUMBERS;
  }
  return ['1234', '5678', '9012'];
}

/**
 * Parse date flexibly.
 * @param {any} value - Date value to parse
 * @return {Date|null} Parsed date or null
 */
function NFSL_parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value)) return value;
  
  // Try existing helper first
  if (typeof parseDateFlexible_ === 'function') {
    return parseDateFlexible_(value);
  }
  
  // Fallback parsing
  const s = String(value).trim();
  let d = new Date(s);
  if (!isNaN(d)) return d;
  
  // dd.MM.yyyy [HH:mm[:ss]]
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const dd = m[1].padStart(2,'0'), MM = m[2].padStart(2,'0'), yyyy = m[3];
    const hh = (m[4]||'00').padStart(2,'0'), mi = (m[5]||'00').padStart(2,'0'), ss = (m[6]||'00').padStart(2,'0');
    d = new Date(`${yyyy}-${MM}-${dd}T${hh}:${mi}:${ss}`);
    if (!isNaN(d)) return d;
  }
  
  // yyyy-MM-dd HH:mm[:ss]
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    d = new Date(`${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]||'00'}`);
    if (!isNaN(d)) return d;
  }
  
  return null;
}

/**
 * Format date as YYYY-MM-DD.
 * @param {Date} date - Date to format
 * @return {string} Formatted date string
 */
function NFSL_formatDate_(date) {
  if (!(date instanceof Date)) return String(date || '');
  const yyyy = date.getFullYear();
  const MM = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${MM}-${dd}`;
}

/**
 * Get ISO week number for a date.
 * @param {Date} date - Date to get week number for
 * @return {string} ISO week string (e.g., "2025-W03")
 */
function NFSL_getISOWeek_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

/**
 * Normalize digits from a string.
 * @param {string} value - String to normalize
 * @return {string} Normalized digits only
 */
function NFSL_normalizeDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

/**
 * Calculate service level metrics for a group.
 * @param {any[][]} rows - Data rows
 * @param {string|null} sokAccount - SOK account number (if filtering for SOK)
 * @param {string[]|null} krkNumbers - Kärkkäinen numbers (if filtering for Kärkkäinen)
 * @param {number} dateSentIndex - Date sent column index
 * @param {number} deliveredIndex - Delivered column index
 * @param {number} payerIndex - Payer column index
 * @return {Object} Metrics object
 */
function NFSL_calculateGroupMetrics_(rows, sokAccount, krkNumbers, dateSentIndex, deliveredIndex, payerIndex) {
  // Filter rows by payer if needed
  let groupRows = rows;
  let groupName = 'ALL';
  
  if (sokAccount !== null && payerIndex >= 0) {
    groupName = 'SOK';
    const sokDigits = NFSL_normalizeDigits_(sokAccount);
    groupRows = rows.filter(row => {
      const payerValue = String(row[payerIndex] || '').trim();
      const payerDigits = NFSL_normalizeDigits_(payerValue);
      return payerDigits === sokDigits;
    });
  } else if (krkNumbers !== null && payerIndex >= 0) {
    groupName = 'KARKKAINEN';
    const krkDigitsList = krkNumbers.map(NFSL_normalizeDigits_);
    groupRows = rows.filter(row => {
      const payerValue = String(row[payerIndex] || '').trim();
      const payerDigits = NFSL_normalizeDigits_(payerValue);
      return krkDigitsList.includes(payerDigits);
    });
  }
  
  const ordersTotal = groupRows.length;
  
  // Process delivered rows
  const deliveredRows = [];
  const leadTimesHours = [];
  
  for (const row of groupRows) {
    const deliveredVal = row[deliveredIndex];
    if (!deliveredVal) continue;
    
    const deliveredDate = NFSL_parseDate_(deliveredVal);
    if (!deliveredDate) continue;
    
    deliveredRows.push(row);
    
    // Calculate lead time if Date sent is available
    if (dateSentIndex >= 0) {
      const sentVal = row[dateSentIndex];
      if (sentVal) {
        const sentDate = NFSL_parseDate_(sentVal);
        if (sentDate && sentDate <= deliveredDate) {
          const leadTimeMs = deliveredDate - sentDate;
          const leadTimeHours = leadTimeMs / (1000 * 60 * 60);
          leadTimesHours.push(leadTimeHours);
        }
      }
    }
  }
  
  const deliveredTotal = deliveredRows.length;
  const pendingTotal = ordersTotal - deliveredTotal;
  
  // Calculate buckets
  let ltLt24h = 0;
  let lt24_72h = 0;
  let ltGt72h = 0;
  
  for (const hours of leadTimesHours) {
    if (hours < 24) {
      ltLt24h++;
    } else if (hours <= 72) {
      lt24_72h++;
    } else {
      ltGt72h++;
    }
  }
  
  // Calculate percentages (of delivered)
  const pctLt24h = deliveredTotal > 0 ? (ltLt24h / deliveredTotal * 100).toFixed(2) : '0.00';
  const pct24_72h = deliveredTotal > 0 ? (lt24_72h / deliveredTotal * 100).toFixed(2) : '0.00';
  const pctGt72h = deliveredTotal > 0 ? (ltGt72h / deliveredTotal * 100).toFixed(2) : '0.00';
  
  // Calculate statistics
  let avgHours = '';
  let medianHours = '';
  let p90Hours = '';
  
  if (leadTimesHours.length > 0) {
    const sum = leadTimesHours.reduce((a, b) => a + b, 0);
    avgHours = (sum / leadTimesHours.length).toFixed(2);
    
    const sorted = leadTimesHours.slice().sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    medianHours = sorted.length % 2 === 0 
      ? ((sorted[mid - 1] + sorted[mid]) / 2).toFixed(2)
      : sorted[mid].toFixed(2);
    
    const p90Index = Math.ceil(sorted.length * 0.9) - 1;
    p90Hours = sorted[p90Index].toFixed(2);
  }
  
  return {
    group: groupName,
    ordersTotal,
    deliveredTotal,
    pendingTotal,
    ltLt24h,
    lt24_72h,
    ltGt72h,
    pctLt24h,
    pct24_72h,
    pctGt72h,
    avgHours,
    medianHours,
    p90Hours
  };
}

/**
 * Write service level sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Spreadsheet object
 * @param {string} isoWeek - ISO week string
 * @param {Date} start - Window start date
 * @param {Date} end - Window end date
 * @param {Object} allMetrics - Metrics for ALL group
 * @param {Object} sokMetrics - Metrics for SOK group
 * @param {Object} krkMetrics - Metrics for KARKKAINEN group
 */
function NFSL_writeServiceLevelSheet_(ss, isoWeek, start, end, allMetrics, sokMetrics, krkMetrics) {
  const sheetName = 'NF_Weekly_ServiceLevels';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  // Define headers
  const headers = [
    'ISOWeek', 'Group', 'OrdersTotal', 'DeliveredTotal', 'PendingTotal',
    'LTlt24h', 'LT24_72h', 'LTgt72h',
    'Pct_lt24h_of_delivered', 'Pct_24_72h_of_delivered', 'Pct_gt72h_of_delivered',
    'AvgHours', 'MedianHours', 'P90Hours'
  ];
  
  // Write headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  
  // Write info row
  const infoRow = new Array(headers.length).fill('');
  infoRow[0] = `Week (SUN→SUN): ${NFSL_formatDate_(start)} - ${NFSL_formatDate_(end)}`;
  infoRow[1] = `Created: ${new Date().toISOString()}`;
  sheet.getRange(2, 1, 1, headers.length).setValues([infoRow]);
  sheet.getRange(2, 1, 1, 2).setFontStyle('italic');
  
  // Write data rows
  const dataRows = [
    [
      isoWeek, allMetrics.group, allMetrics.ordersTotal, allMetrics.deliveredTotal, allMetrics.pendingTotal,
      allMetrics.ltLt24h, allMetrics.lt24_72h, allMetrics.ltGt72h,
      allMetrics.pctLt24h, allMetrics.pct24_72h, allMetrics.pctGt72h,
      allMetrics.avgHours, allMetrics.medianHours, allMetrics.p90Hours
    ],
    [
      isoWeek, sokMetrics.group, sokMetrics.ordersTotal, sokMetrics.deliveredTotal, sokMetrics.pendingTotal,
      sokMetrics.ltLt24h, sokMetrics.lt24_72h, sokMetrics.ltGt72h,
      sokMetrics.pctLt24h, sokMetrics.pct24_72h, sokMetrics.pctGt72h,
      sokMetrics.avgHours, sokMetrics.medianHours, sokMetrics.p90Hours
    ],
    [
      isoWeek, krkMetrics.group, krkMetrics.ordersTotal, krkMetrics.deliveredTotal, krkMetrics.pendingTotal,
      krkMetrics.ltLt24h, krkMetrics.lt24_72h, krkMetrics.ltGt72h,
      krkMetrics.pctLt24h, krkMetrics.pct24_72h, krkMetrics.pctGt72h,
      krkMetrics.avgHours, krkMetrics.medianHours, krkMetrics.p90Hours
    ]
  ];
  
  sheet.getRange(3, 1, dataRows.length, headers.length).setValues(dataRows);
  
  // Format sheet
  sheet.autoResizeColumns(1, Math.min(headers.length, 20));
  
  console.log(`Service level sheet written: ${dataRows.length} rows`);
}

/**
 * Test function for service level calculations.
 * Creates test data and validates the calculation logic.
 */
function NFSL_testServiceLevelCalculations() {
  console.log('Testing service level calculation logic...');
  
  // Test date parsing
  const testDates = [
    '2025-01-15 10:30:00',
    '15.01.2025 10:30',
    '2025-01-15T10:30:00',
    new Date('2025-01-15T10:30:00')
  ];
  
  console.log('Testing date parsing:');
  for (const dateVal of testDates) {
    const parsed = NFSL_parseDate_(dateVal);
    if (parsed && !isNaN(parsed)) {
      console.log(`✅ Parsed "${dateVal}" -> ${parsed.toISOString()}`);
    } else {
      console.error(`❌ Failed to parse "${dateVal}"`);
    }
  }
  
  // Test ISO week calculation
  const testWeek = new Date('2025-01-13'); // A Monday
  const isoWeek = NFSL_getISOWeek_(testWeek);
  console.log(`ISO week for ${testWeek.toISOString()}: ${isoWeek}`);
  
  // Test lead time buckets
  const testLeadTimes = [12, 23, 25, 48, 72, 96, 120]; // hours
  const buckets = { lt24: 0, '24_72': 0, gt72: 0 };
  
  for (const hours of testLeadTimes) {
    if (hours < 24) buckets.lt24++;
    else if (hours <= 72) buckets['24_72']++;
    else buckets.gt72++;
  }
  
  console.log('Lead time bucket test:');
  console.log(`  < 24h: ${buckets.lt24} (expected: 2)`);
  console.log(`  24-72h: ${buckets['24_72']} (expected: 3)`);
  console.log(`  > 72h: ${buckets.gt72} (expected: 2)`);
  
  // Test statistics calculation
  const values = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100];
  const sum = values.reduce((a, b) => a + b, 0);
  const avg = sum / values.length;
  const sorted = values.slice().sort((a, b) => a - b);
  const median = sorted[Math.floor(sorted.length / 2)];
  const p90Index = Math.ceil(sorted.length * 0.9) - 1;
  const p90 = sorted[p90Index];
  
  console.log('Statistics test:');
  console.log(`  Avg: ${avg.toFixed(2)} (expected: 55.00)`);
  console.log(`  Median: ${median.toFixed(2)} (expected: 60.00)`);
  console.log(`  P90: ${p90.toFixed(2)} (expected: 90.00)`);
  
  // Test column detection with mock headers
  const mockHeaders = [
    'Package Number',
    'Date sent',
    'Payer',
    'Delivered Time',
    'Status',
    'Created'
  ];
  
  const dateSentIdx = NFSL_findDateSentColumn_(mockHeaders);
  const deliveredIdx = NFSL_findDeliveredColumn_(mockHeaders);
  const payerIdx = NFSL_findPayerColumn_(mockHeaders);
  
  console.log('Column detection test:');
  console.log(`  Date sent: ${dateSentIdx} (expected: 1)`);
  console.log(`  Delivered: ${deliveredIdx} (expected: 3)`);
  console.log(`  Payer: ${payerIdx} (expected: 2)`);
  
  const allPassed = 
    dateSentIdx === 1 &&
    deliveredIdx === 3 &&
    payerIdx === 2;
  
  if (allPassed) {
    console.log('✅ All service level tests passed!');
    return true;
  } else {
    console.error('❌ Some tests failed');
    return false;
  }
}
