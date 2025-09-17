// Adhoc_Tracker.gs — Enhanced ad hoc tracker with improved header detection and tracking integration

/**
 * Enhanced ad hoc tracker that reliably detects tracking columns including Finnish/English aliases
 * and integrates with the enhanced tracking system for KPI generation.
 */

// Enhanced header candidates with Finnish/English aliases
const ADHOC_TRACKING_CANDIDATES = [
  'package id (seurantakoodi)', 'package id', 'seurantakoodi', 'tracking number', 'tracking',
  'barcode', 'waybill', 'waybill no', 'awb', 'package number', 'packagenumber',
  'shipment id', 'consignment number', 'consignment', 'parcel id', 'parcel number'
];

const ADHOC_CARRIER_CANDIDATES = [
  'carrier', 'carrier name', 'carriername', 'delivery carrier', 'courier',
  'logistics provider', 'logisticsprovider', 'shipper', 'service provider',
  'forwarder', 'transporter', 'kuljetusliike', 'kuljetusyhtiö', 'kuljetus',
  'toimitustapa', 'delivery method', 'service family', 'refreshcarrier'
];

const ADHOC_COUNTRY_CANDIDATES = [
  'country', 'destination country', 'dest country', 'delivery country',
  'country code', 'maa', 'kohdamaa', 'destination', 'dest'
];

/**
 * Normalize header text for better matching
 */
function adhocNormalize_(s) {
  if (!s) return '';
  return String(s)
    .toLowerCase()
    .replace(/\([^)]*\)/g, '') // Remove parentheses and their content
    .replace(/[^\p{L}\p{N}]+/gu, ' ') // Replace punctuation with spaces
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Enhanced header detection with fallback content scanning
 */
function adhocFindColumn_(headers, data, candidates, contentValidator) {
  // First try exact header matching
  const normalizedHeaders = headers.map(h => adhocNormalize_(h));
  const normalizedCandidates = candidates.map(c => adhocNormalize_(c));
  
  for (let i = 0; i < normalizedHeaders.length; i++) {
    for (const candidate of normalizedCandidates) {
      if (normalizedHeaders[i].includes(candidate) || candidate.includes(normalizedHeaders[i])) {
        return i;
      }
    }
  }
  
  // Fallback: content-based detection if validator provided
  if (contentValidator && data && data.length > 1) {
    let bestCol = -1;
    let bestScore = 0;
    
    for (let col = 0; col < headers.length; col++) {
      let validCount = 0;
      let totalCount = 0;
      
      for (let row = 1; row < Math.min(data.length, 50); row++) { // Check up to 50 rows
        const value = String(data[row][col] || '').trim();
        if (value) {
          totalCount++;
          if (contentValidator(value)) {
            validCount++;
          }
        }
      }
      
      if (totalCount > 0) {
        const score = validCount / totalCount;
        if (score > bestScore && score > 0.3) { // At least 30% tracking-like
          bestScore = score;
          bestCol = col;
        }
      }
    }
    
    return bestCol;
  }
  
  return -1;
}

/**
 * Validate if a value looks like a tracking code
 */
function adhocIsTrackingLike_(value) {
  const s = String(value || '').trim();
  if (s.length < 4) return false;
  
  // Should contain letters and/or numbers, not just special characters
  if (!/[a-zA-Z0-9]/.test(s)) return false;
  
  // Common tracking code patterns
  const patterns = [
    /^[A-Z0-9]{8,}$/i,     // Alphanumeric 8+ chars
    /^\d{10,}$/,           // Numeric 10+ digits
    /^[A-Z]{2}\d{9}[A-Z]{2}$/i, // Postal format
    /^1Z[A-Z0-9]{16}$/i,   // UPS format
    /^\d{4}\s?\d{4}\s?\d{4}$/  // Grouped digits
  ];
  
  return patterns.some(p => p.test(s.replace(/\s/g, '')));
}

/**
 * Validate if a value looks like a carrier name
 */
function adhocIsCarrierLike_(value) {
  const s = String(value || '').trim().toLowerCase();
  if (s.length < 2) return false;
  
  const knownCarriers = [
    'posti', 'gls', 'dhl', 'fedex', 'ups', 'bring', 'matkahuolto', 
    'schenker', 'dsv', 'db', 'kaukokiito', 'mh'
  ];
  
  return knownCarriers.some(carrier => s.includes(carrier)) || 
         /^[a-zA-Z\s]{2,}$/.test(s); // Letters and spaces only
}

/**
 * Main enhanced tracker function that replaces buildAdhocFromValues_
 */
function ADHOC_buildFromValues(values, label, options = {}) {
  if (!values || !values.length) {
    throw new Error('Tuotu tiedosto on tyhjä.');
  }
  
  const {
    sourceSheetName = null,
    headerRow = 1,
    trackingHeaderHints = [],
    carrierHeaderHints = [],
    countryHeaderHints = []
  } = options;
  
  // Get headers from specified row (default row 1)
  const headerRowIndex = Math.max(0, headerRow - 1);
  if (headerRowIndex >= values.length) {
    throw new Error(`Header row ${headerRow} not found in data.`);
  }
  
  const headers = values[headerRowIndex].map(v => String(v || '').trim());
  const dataRows = values.slice(headerRowIndex + 1).filter(r => r.some(x => String(x || '').trim() !== ''));
  
  // Enhanced column detection
  const trackingCandidates = [...ADHOC_TRACKING_CANDIDATES, ...trackingHeaderHints];
  const carrierCandidates = [...ADHOC_CARRIER_CANDIDATES, ...carrierHeaderHints];
  const countryCandidates = [...ADHOC_COUNTRY_CANDIDATES, ...countryHeaderHints];
  
  const trackingCol = adhocFindColumn_(headers, values, trackingCandidates, adhocIsTrackingLike_);
  const carrierCol = adhocFindColumn_(headers, values, carrierCandidates, adhocIsCarrierLike_);
  const countryCol = adhocFindColumn_(headers, values, countryCandidates, null);
  
  if (trackingCol < 0) {
    throw new Error(`Adhoc: ei löytynyt seurantakoodin saraketta. Tarkistetut sarakkeet: ${headers.join(', ')}`);
  }
  
  if (carrierCol < 0) {
    Logger.log('Carrier column not found, will use empty values');
  }
  
  // Extract data
  const extractedData = dataRows.map(row => {
    const carrier = carrierCol >= 0 ? String(row[carrierCol] || '').trim() : '';
    const tracking = String(row[trackingCol] || '').trim();
    const country = countryCol >= 0 ? String(row[countryCol] || '').trim() : '';
    
    return {
      carrier: carrier || 'Unknown',
      tracking: tracking,
      country: country || 'Unknown',
      originalRow: row
    };
  }).filter(item => item.tracking); // Only keep rows with tracking codes
  
  if (extractedData.length === 0) {
    throw new Error('Ei löytynyt yhtään riviä seurantakoodilla.');
  }
  
  // Create or update Adhoc_Results sheet
  const ss = SpreadsheetApp.getActive();
  const resultsSheet = ss.getSheetByName('Adhoc_Results') || ss.insertSheet('Adhoc_Results');
  
  // Setup results headers
  const resultsHeaders = [
    'Carrier', 'Tracking', 'Country', 'Status', 'CreatedISO', 'DeliveredISO', 
    'DaysToDeliver', 'WeekISO', 'RefreshAt', 'Location', 'Raw'
  ];
  
  resultsSheet.clear();
  resultsSheet.getRange(1, 1, 1, resultsHeaders.length).setValues([resultsHeaders]);
  
  // Process tracking with enhanced system
  const results = [];
  const now = new Date();
  
  for (const item of extractedData) {
    try {
      // Use existing tracking system
      const trackResult = TRK_trackByCarrier_(item.carrier, item.tracking);
      
      // Extract key dates for KPI calculation
      let createdISO = '';
      let deliveredISO = '';
      let daysToDeliver = '';
      
      if (trackResult.status && /accepted|received|picked/i.test(trackResult.status)) {
        createdISO = trackResult.time || '';
      }
      
      if (trackResult.status && /delivered|toimitettu|luovutettu/i.test(trackResult.status)) {
        deliveredISO = trackResult.time || '';
      }
      
      // Calculate delivery time if both dates available
      if (createdISO && deliveredISO) {
        try {
          const created = new Date(createdISO);
          const delivered = new Date(deliveredISO);
          if (!isNaN(created) && !isNaN(delivered)) {
            daysToDeliver = Math.round((delivered - created) / (1000 * 60 * 60 * 24));
          }
        } catch (e) {
          Logger.log('Date calculation error: ' + e.message);
        }
      }
      
      // Calculate ISO week
      const weekISO = deliveredISO ? getISOWeek_(new Date(deliveredISO)) : '';
      
      results.push([
        item.carrier,
        item.tracking,
        item.country,
        trackResult.status || '',
        createdISO,
        deliveredISO,
        daysToDeliver,
        weekISO,
        Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
        trackResult.location || '',
        (trackResult.raw || '').substring(0, 500)
      ]);
    } catch (e) {
      Logger.log('Error processing tracking ' + item.tracking + ': ' + e.message);
      results.push([
        item.carrier,
        item.tracking,
        item.country,
        'ERROR: ' + e.message,
        '', '', '', '',
        Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
        '', ''
      ]);
    }
  }
  
  // Write results
  if (results.length > 0) {
    resultsSheet.getRange(2, 1, results.length, resultsHeaders.length).setValues(results);
  }
  
  resultsSheet.setFrozenRows(1);
  resultsSheet.autoResizeColumns(1, Math.min(resultsHeaders.length, 10));
  
  // Build KPI sheet
  ADHOC_buildKPI();
  
  ss.toast(`Adhoc: käsitelty ${results.length} seurantakoodia (${label})`);
  return results.length;
}

/**
 * Build KPI summary with average delivery delays per country per ISO week
 */
function ADHOC_buildKPI() {
  const ss = SpreadsheetApp.getActive();
  const resultsSheet = ss.getSheetByName('Adhoc_Results');
  
  if (!resultsSheet || resultsSheet.getLastRow() < 2) {
    return;
  }
  
  const data = resultsSheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  // Find column indices
  const countryCol = headers.indexOf('Country');
  const weekCol = headers.indexOf('WeekISO');
  const daysCol = headers.indexOf('DaysToDeliver');
  const deliveredCol = headers.indexOf('DeliveredISO');
  
  if (countryCol < 0 || weekCol < 0 || daysCol < 0) {
    Logger.log('Required columns not found for KPI calculation');
    return;
  }
  
  // Group by country and week
  const kpiData = {};
  
  for (const row of rows) {
    const country = row[countryCol] || 'Unknown';
    const week = row[weekCol];
    const days = row[daysCol];
    const delivered = row[deliveredCol];
    
    if (!week || !days || !delivered) continue; // Only delivered items
    
    const key = `${country}|${week}`;
    if (!kpiData[key]) {
      kpiData[key] = {
        country: country,
        week: week,
        deliveries: [],
        totalDays: 0,
        count: 0
      };
    }
    
    const daysNum = parseFloat(days);
    if (!isNaN(daysNum)) {
      kpiData[key].deliveries.push(daysNum);
      kpiData[key].totalDays += daysNum;
      kpiData[key].count += 1;
    }
  }
  
  // Calculate averages and create KPI data
  const kpiRows = [];
  
  for (const [key, data] of Object.entries(kpiData)) {
    if (data.count > 0) {
      const avgDays = (data.totalDays / data.count).toFixed(1);
      const medianDays = calculateMedian_(data.deliveries).toFixed(1);
      
      kpiRows.push([
        data.country,
        data.week,
        data.count,
        avgDays,
        medianDays,
        Math.min(...data.deliveries),
        Math.max(...data.deliveries)
      ]);
    }
  }
  
  // Sort by country, then week
  kpiRows.sort((a, b) => {
    if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
    return a[1].localeCompare(b[1]);
  });
  
  // Create or update KPI sheet
  const kpiSheet = ss.getSheetByName('Adhoc_KPI') || ss.insertSheet('Adhoc_KPI');
  const kpiHeaders = [
    'Country', 'ISO Week', 'Deliveries', 'Avg Days', 'Median Days', 'Min Days', 'Max Days'
  ];
  
  kpiSheet.clear();
  kpiSheet.getRange(1, 1, 1, kpiHeaders.length).setValues([kpiHeaders]);
  
  if (kpiRows.length > 0) {
    kpiSheet.getRange(2, 1, kpiRows.length, kpiHeaders.length).setValues(kpiRows);
  }
  
  kpiSheet.setFrozenRows(1);
  kpiSheet.autoResizeColumns(1, kpiHeaders.length);
  
  Logger.log('KPI: Generated ' + kpiRows.length + ' country/week combinations');
}

/**
 * Calculate median of an array of numbers
 */
function calculateMedian_(numbers) {
  if (numbers.length === 0) return 0;
  
  const sorted = numbers.slice().sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  
  if (sorted.length % 2 === 0) {
    return (sorted[middle - 1] + sorted[middle]) / 2;
  } else {
    return sorted[middle];
  }
}

/**
 * Get ISO week number for a date
 */
function getISOWeek_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  const week1 = new Date(d.getFullYear(), 0, 4);
  const weekNum = 1 + Math.round(((d.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
  return `${d.getFullYear()}-W${String(weekNum).padStart(2, '0')}`;
}

/**
 * Main entry point for URL-based import (replaces buildAdhocFromValues_)
 */
function ADHOC_RunFromUrl(url, options = {}) {
  if (!url) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Ad hoc seuranta',
      'Liitä Drive-URL tai -ID (.xlsx/.csv):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    url = response.getResponseText();
  }
  
  const fileId = extractId_(url);
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  
  let values;
  const mimeType = blob.getContentType();
  
  if (mimeType.includes('sheet') || mimeType.includes('excel')) {
    // Excel file
    const xlsxData = Utilities.parseCsv(Utilities.newBlob(blob.getBytes(), 'text/csv').getDataAsString());
    values = xlsxData;
  } else {
    // CSV file
    values = Utilities.parseCsv(blob.getDataAsString());
  }
  
  return ADHOC_buildFromValues(values, file.getName(), options);
}

/**
 * Helper to extract file ID from Drive URL
 */
function extractId_(urlOrId) {
  const s = String(urlOrId || '').trim();
  const match = s.match(/[-\w]{25,}/);
  return match ? match[0] : s;
}

/**
 * Refresh existing Adhoc_Results data
 */
function ADHOC_RefreshResults() {
  const ss = SpreadsheetApp.getActive();
  const resultsSheet = ss.getSheetByName('Adhoc_Results');
  
  if (!resultsSheet || resultsSheet.getLastRow() < 2) {
    ss.toast('Ei Adhoc_Results dataa päivitettäväksi');
    return;
  }
  
  const data = resultsSheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const carrierCol = headers.indexOf('Carrier');
  const trackingCol = headers.indexOf('Tracking');
  
  if (carrierCol < 0 || trackingCol < 0) {
    throw new Error('Required columns not found in Adhoc_Results');
  }
  
  const now = new Date();
  const updatedRows = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].slice();
    const carrier = row[carrierCol];
    const tracking = row[trackingCol];
    
    if (carrier && tracking) {
      try {
        const trackResult = TRK_trackByCarrier_(carrier, tracking);
        
        // Update status and related fields
        row[headers.indexOf('Status')] = trackResult.status || '';
        row[headers.indexOf('RefreshAt')] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        row[headers.indexOf('Location')] = trackResult.location || '';
        row[headers.indexOf('Raw')] = (trackResult.raw || '').substring(0, 500);
        
        // Update dates if needed
        if (trackResult.status && /accepted|received|picked/i.test(trackResult.status) && !row[headers.indexOf('CreatedISO')]) {
          row[headers.indexOf('CreatedISO')] = trackResult.time || '';
        }
        
        if (trackResult.status && /delivered|toimitettu|luovutettu/i.test(trackResult.status)) {
          row[headers.indexOf('DeliveredISO')] = trackResult.time || '';
          row[headers.indexOf('WeekISO')] = trackResult.time ? getISOWeek_(new Date(trackResult.time)) : '';
          
          // Recalculate delivery days
          const createdISO = row[headers.indexOf('CreatedISO')];
          if (createdISO && trackResult.time) {
            try {
              const created = new Date(createdISO);
              const delivered = new Date(trackResult.time);
              if (!isNaN(created) && !isNaN(delivered)) {
                row[headers.indexOf('DaysToDeliver')] = Math.round((delivered - created) / (1000 * 60 * 60 * 24));
              }
            } catch (e) {
              Logger.log('Date calculation error: ' + e.message);
            }
          }
        }
      } catch (e) {
        Logger.log('Error refreshing tracking ' + tracking + ': ' + e.message);
        row[headers.indexOf('Status')] = 'ERROR: ' + e.message;
        row[headers.indexOf('RefreshAt')] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      }
    }
    
    updatedRows.push(row);
  }
  
  // Write updated data
  resultsSheet.getRange(2, 1, updatedRows.length, headers.length).setValues(updatedRows);
  
  // Rebuild KPI
  ADHOC_buildKPI();
  
  ss.toast(`Adhoc: päivitetty ${updatedRows.length} seurantakoodia`);
}