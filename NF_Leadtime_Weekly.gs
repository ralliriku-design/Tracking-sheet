/**
 * NF_Leadtime_Weekly.gs - Lead Time Analytics
 * 
 * Provides delivery time analysis and weekly country KPIs.
 * Based on existing 08_delivery_reports.txt patterns but with "keikka tehty" start time logic.
 */

/********************* DELIVERY TIME REPORTS ***************************/

/**
 * Create detailed delivery times report
 * Uses "keikka tehty" priority: pickup → submitted → created
 * Delivered time from tracking events only
 */
function NF_makeDeliveryTimeReport() {
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
  
  const headerMap = NF_headerIndexMap_(data[0]);
  
  // Column indices
  const carrierIdx = NF_colIndexOf_(headerMap, ['Carrier']);
  const countryIdx = NF_colIndexOf_(headerMap, ['Country', 'Dest Country', 'Destination Country']);
  const submittedIdx = NF_colIndexOf_(headerMap, ['Submitted date', 'Submitted', 'B']);
  const pickupIdx = NF_colIndexOf_(headerMap, ['Pick up date', 'Pick up', 'Pickup date']);
  const createdIdx = NF_colIndexOf_(headerMap, ['Created Date', 'Created']);
  const deliveredIdx = NF_colIndexOf_(headerMap, ['Delivered Time', 'Delivered At', 'Delivered']);
  const refreshTimeIdx = NF_colIndexOf_(headerMap, ['RefreshTime']);
  const refreshStatusIdx = NF_colIndexOf_(headerMap, ['RefreshStatus', 'Refresh Status']);
  const confirmedDeliveredIdx = NF_colIndexOf_(headerMap, ['Delivered date (Confirmed)']);
  
  const resultRows = [['Carrier', 'DestCountry', 'Start(keikka tehty)', 'Delivered', 'LeadTime(days)', 'Source', 'Submitted', 'Pickup', 'Created']];
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    
    const carrier = row[carrierIdx] || '';
    const country = row[countryIdx] || '';
    const submitted = row[submittedIdx] || '';
    const pickup = row[pickupIdx] || '';
    const created = row[createdIdx] || '';
    
    // Pick delivered time (FIXED: RefreshTime only if status indicates delivered)
    // Priority: confirmed → delivered → refresh time (only if status = delivered)
    let delivered = '';
    if (confirmedDeliveredIdx >= 0 && row[confirmedDeliveredIdx]) {
      delivered = row[confirmedDeliveredIdx];
    } else if (deliveredIdx >= 0 && row[deliveredIdx]) {
      delivered = row[deliveredIdx];
    } else if (refreshTimeIdx >= 0 && row[refreshTimeIdx] && refreshStatusIdx >= 0) {
      // Only use RefreshTime if RefreshStatus indicates delivered
      const status = String(row[refreshStatusIdx] || '').toLowerCase();
      if (status.includes('delivered') || 
          status.includes('toimitettu') || 
          status.includes('luovutettu') ||
          status.includes('utlevert')) {
        delivered = row[refreshTimeIdx];
      }
    }
    
    // "Keikka tehty" start time priority: pickup → submitted → created
    let startTime = '';
    let source = '';
    if (pickup) {
      startTime = pickup;
      source = 'pickup';
    } else if (submitted) {
      startTime = submitted;
      source = 'submitted';
    } else if (created) {
      startTime = created;
      source = 'created';
    }
    
    // Calculate lead time
    let leadTime = '';
    if (startTime && delivered) {
      const startMs = new Date(startTime).getTime();
      const deliveredMs = new Date(delivered).getTime();
      if (isFinite(startMs) && isFinite(deliveredMs) && deliveredMs > startMs) {
        leadTime = ((deliveredMs - startMs) / 86400000).toFixed(2);
      }
    }
    
    resultRows.push([
      carrier,
      country,
      startTime,
      delivered,
      leadTime,
      source,
      submitted,
      pickup,
      created
    ]);
  }
  
  const reportSheet = ss.getSheetByName('Delivery_Times') || ss.insertSheet('Delivery_Times');
  reportSheet.clear();
  reportSheet.getRange(1, 1, resultRows.length, resultRows[0].length).setValues(resultRows);
  
  SpreadsheetApp.getUi().alert('Delivery times report ready: Delivery_Times');
}

/**
 * Create country/week lead time KPI report
 * Pivots Delivery_Times data by ISO week and country
 */
function NF_makeCountryWeekLeadtimeReport() {
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
  
  const headers = data[0];
  const carrierIdx = headers.indexOf('Carrier');
  const countryIdx = headers.indexOf('DestCountry');
  const deliveredIdx = headers.indexOf('Delivered');
  const leadTimeIdx = headers.indexOf('LeadTime(days)');
  
  // Aggregate by week|country|carrier
  const aggregationMap = {};
  
  for (let r = 1; r < data.length; r++) {
    const carrier = String(data[r][carrierIdx] || '').trim();
    const country = String(data[r][countryIdx] || '').trim();
    const delivered = data[r][deliveredIdx];
    const leadTime = parseFloat(data[r][leadTimeIdx]);
    
    if (!carrier || !country || !delivered || !isFinite(leadTime)) continue;
    
    // Get year and ISO week
    const deliveredDate = new Date(delivered);
    const year = deliveredDate.getFullYear();
    const isoWeek = NF_getISOWeek_(deliveredDate);
    const weekKey = `${year}-W${isoWeek.toString().padStart(2, '0')}`;
    
    const aggregationKey = `${year}|${weekKey}|${country}|${carrier}`;
    
    if (!aggregationMap[aggregationKey]) {
      aggregationMap[aggregationKey] = { sum: 0, count: 0 };
    }
    aggregationMap[aggregationKey].sum += leadTime;
    aggregationMap[aggregationKey].count += 1;
  }
  
  // Build result rows
  const resultRows = [['Year', 'ISOWeek', 'Country', 'Carrier', 'AvgLeadTime(days)', 'Count']];
  
  Object.keys(aggregationMap).sort().forEach(key => {
    const [year, weekKey, country, carrier] = key.split('|');
    const aggregation = aggregationMap[key];
    const avgLeadTime = (aggregation.sum / aggregation.count).toFixed(2);
    
    resultRows.push([year, weekKey, country, carrier, avgLeadTime, aggregation.count]);
  });
  
  const reportSheet = ss.getSheetByName('Leadtime_Weekly_Country') || ss.insertSheet('Leadtime_Weekly_Country');
  reportSheet.clear();
  reportSheet.getRange(1, 1, resultRows.length, resultRows[0].length).setValues(resultRows);
  
  SpreadsheetApp.getUi().alert('Country week report ready: Leadtime_Weekly_Country');
}

/********************* HELPER FUNCTIONS ***************************/

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
 * Get ISO week number for a date
 * Based on existing Utilities.formatDate pattern but with fallback
 */
function NF_getISOWeek_(date) {
  try {
    // Try using Utilities.formatDate first
    const weekString = Utilities.formatDate(date, Session.getScriptTimeZone(), "ww");
    return parseInt(weekString, 10);
  } catch (e) {
    // Fallback: manual ISO week calculation
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    // Set to nearest Thursday: current date + 4 - current day number
    // Make Sunday's day number 7
    d.setDate(d.getDate() + 4 - (d.getDay() || 7));
    // Get first day of year
    const yearStart = new Date(d.getFullYear(), 0, 1);
    // Calculate full weeks to nearest Thursday
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return weekNo;
  }
}

/**
 * Alternative year-week format (YYYY-'W'ww) using Utilities or fallback
 */
function NF_yearIsoWeek_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "YYYY-'W'ww");
  } catch (e) {
    const year = date.getFullYear();
    const week = NF_getISOWeek_(date);
    return `${year}-W${week.toString().padStart(2, '0')}`;
  }
}