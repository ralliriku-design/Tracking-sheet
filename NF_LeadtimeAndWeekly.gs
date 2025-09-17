/** NF_LeadtimeAndWeekly.gs — Country leadtime (job-done → delivered) and SOK/Kärkkäinen weekly (Sun→Sun) */

// Helper: safe property get
function NF_Prop_(k, d) { try { const v = PropertiesService.getScriptProperties().getProperty(String(k)); return (v==null||v==='')?d:v; } catch(e){ return d; } }

// Resolve commonly used sheets
function NF_getSheet_(name) { const ss=SpreadsheetApp.getActive(); return ss.getSheetByName(name) || ss.insertSheet(name); }

// Week window (last finished Sun→Sun)
function NF_lastFinishedWeek_() {
  const now = new Date(); now.setHours(0,0,0,0);
  const day = now.getDay();
  const end = new Date(now); end.setDate(now.getDate() - day); // this Sunday 00:00
  const start = new Date(end); start.setDate(end.getDate() - 7);
  return { start, end };
}

// Normalize header → index map
function NF_mapHdr_(hdr) { const m={}; for (let i=0;i<hdr.length;i++){ const h=String(hdr[i]||''); m[h]=i; m[h.toLowerCase().trim()]=i; } return m; }

// Find column by candidates
function NF_findCol_(map, candidates) { for (const c of candidates) { if (c in map) return map[c]; const lc = c.toLowerCase().trim(); if (lc in map) return map[lc]; } return -1; }

// Parse date flexibly
function NF_parseDate_(v) { if (!v) return null; try { const d = new Date(v); return isNaN(d.getTime()) ? null : d; } catch(e) { return null; } }

// Format date-time
function NF_fmtDT_(d) { try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone()||'Europe/Helsinki', 'yyyy-MM-dd HH:mm:ss'); } catch(e) { return String(d||''); } }

// Normalize digits (for payer matching)
function NF_digits_(v) { return String(v||'').replace(/\D+/g,''); }

/**
 * Main function: Build country-level leadtime analysis
 */
function NF_buildCountryLeadtime() {
  const ss = SpreadsheetApp.getActive();
  
  // Collect data from all sources
  const packages = NF_getPackagesData_(ss);
  const powerbi = NF_getPowerBIData_(ss);
  const erp = NF_getERPData_(ss);
  
  if (!packages.length && !powerbi.length && !erp.length) {
    SpreadsheetApp.getUi().alert('NewFlow: Ei dataa lähteistä (Packages/PowerBI/ERP)');
    return;
  }
  
  // Merge and analyze
  const merged = NF_mergeSourcesByTrackingCode_(packages, powerbi, erp);
  const analyzed = NF_analyzeLeadtimes_(merged);
  
  // Write detail sheet
  const detailSh = NF_getSheet_('NF_Leadtime_Detail');
  NF_writeDetailSheet_(detailSh, analyzed);
  
  // Build weekly country KPI
  const weeklyKPI = NF_buildWeeklyCountryKPI_(analyzed);
  const kpiSh = NF_getSheet_('NF_Leadtime_Weekly_Country');
  NF_writeWeeklyKPI_(kpiSh, weeklyKPI);
  
  ss.toast('NewFlow: Maa-kohtainen toimitusaika-analyysi valmis');
}

function NF_getPackagesData_(ss) {
  const sh = ss.getSheetByName('Packages');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getDisplayValues();
  const hdr = data[0]; const rows = data.slice(1);
  const map = NF_mapHdr_(hdr);
  
  return rows.map(r => ({
    source: 'Packages',
    trackingCode: NF_extractTrackingCode_(r, map),
    country: NF_extractCountry_(r, map),
    carrier: NF_extractCarrier_(r, map),
    jobDoneTs: NF_extractJobDone_(r, map, 'packages'),
    deliveredTs: NF_extractDelivered_(r, map),
    rawRow: r,
    headers: hdr
  })).filter(x => x.trackingCode);
}

function NF_getPowerBIData_(ss) {
  const sh = ss.getSheetByName('PBI_Outbound_Staging') || ss.getSheetByName('PowerBI_Import') || ss.getSheetByName('PowerBI_New');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getDisplayValues();
  const hdr = data[0]; const rows = data.slice(1);
  const map = NF_mapHdr_(hdr);
  
  return rows.map(r => ({
    source: 'PowerBI',
    trackingCode: NF_extractTrackingCode_(r, map),
    country: NF_extractCountry_(r, map),
    carrier: NF_extractCarrier_(r, map),
    jobDoneTs: NF_extractJobDone_(r, map, 'powerbi'),
    deliveredTs: NF_extractDelivered_(r, map),
    rawRow: r,
    headers: hdr
  })).filter(x => x.trackingCode);
}

function NF_getERPData_(ss) {
  // Try to find ERP sheet - look for common patterns
  const erpSheets = ss.getSheets().filter(s => {
    const name = s.getName().toLowerCase();
    return name.includes('stock') || name.includes('picking') || name.includes('erp');
  });
  
  if (!erpSheets.length) return [];
  
  const sh = erpSheets[0]; // Use first found
  if (sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getDisplayValues();
  const hdr = data[0]; const rows = data.slice(1);
  const map = NF_mapHdr_(hdr);
  
  return rows.map(r => ({
    source: 'ERP',
    trackingCode: NF_extractTrackingCode_(r, map),
    country: NF_extractCountry_(r, map),
    carrier: NF_extractCarrier_(r, map),
    jobDoneTs: NF_extractJobDone_(r, map, 'erp'),
    deliveredTs: NF_extractDelivered_(r, map),
    rawRow: r,
    headers: hdr
  })).filter(x => x.trackingCode);
}

function NF_extractTrackingCode_(row, map) {
  const candidates = ['TrackingNumber','Package Number','PackageNumber','Package Id','PackageId','Tracking Code','Code','Barcode'];
  const idx = NF_findCol_(map, candidates);
  return idx >= 0 ? String(row[idx]||'').trim() : '';
}

function NF_extractCountry_(row, map) {
  const candidates = ['Country','Dest Country','Destination Country','Country Code','Destination'];
  const idx = NF_findCol_(map, candidates);
  return idx >= 0 ? String(row[idx]||'').trim() : '';
}

function NF_extractCarrier_(row, map) {
  const candidates = ['Carrier','LogisticsProvider','Shipper','Kuljetusyritys','Shipping Company'];
  const idx = NF_findCol_(map, candidates);
  return idx >= 0 ? String(row[idx]||'').trim() : '';
}

function NF_extractJobDone_(row, map, source) {
  // Priority: ERP > Gmail > PowerBI
  let candidates = [];
  
  if (source === 'erp') {
    candidates = ['Pick Finish','Completed','Picking Completed','Pick Complete','Done','Finish Time'];
  } else if (source === 'packages') {
    candidates = ['Pick up date','Submitted date','Submitted','Pick up','Booking date','Booked time'];
  } else if (source === 'powerbi') {
    candidates = ['Created','Created date','Dispatch date','Shipped date','Created Date'];
  }
  
  const idx = NF_findCol_(map, candidates);
  if (idx >= 0) {
    const val = row[idx];
    if (val) return NF_parseDate_(val);
  }
  return null;
}

function NF_extractDelivered_(row, map) {
  const candidates = ['Delivered Time','Delivered At','Delivered','RefreshTime','Delivered Date','Delivery Date'];
  const idx = NF_findCol_(map, candidates);
  if (idx >= 0) {
    const val = row[idx];
    if (val) return NF_parseDate_(val);
  }
  return null;
}

function NF_mergeSourcesByTrackingCode_(packages, powerbi, erp) {
  const byCode = {};
  
  // Add all sources by tracking code
  [...packages, ...powerbi, ...erp].forEach(item => {
    const code = item.trackingCode;
    if (!code) return;
    
    if (!byCode[code]) byCode[code] = { sources: [], finalItem: null };
    byCode[code].sources.push(item);
  });
  
  // For each code, pick best jobDone and delivered timestamps
  Object.values(byCode).forEach(group => {
    const sources = group.sources;
    
    // Priority for jobDone: ERP > Packages > PowerBI
    let bestJobDone = null;
    let bestJobSource = '';
    for (const s of sources) {
      if (s.jobDoneTs && (!bestJobDone || 
          (s.source === 'ERP' && bestJobSource !== 'ERP') ||
          (s.source === 'Packages' && bestJobSource === 'PowerBI'))) {
        bestJobDone = s.jobDoneTs;
        bestJobSource = s.source;
      }
    }
    
    // For delivered: use any available (prefer most recent non-null)
    let bestDelivered = null;
    for (const s of sources) {
      if (s.deliveredTs && (!bestDelivered || s.deliveredTs > bestDelivered)) {
        bestDelivered = s.deliveredTs;
      }
    }
    
    // Pick primary source (prefer ERP > Packages > PowerBI)
    let primary = sources.find(s => s.source === 'ERP') || 
                 sources.find(s => s.source === 'Packages') || 
                 sources[0];
    
    group.finalItem = {
      ...primary,
      jobDoneTs: bestJobDone,
      deliveredTs: bestDelivered,
      jobDoneSource: bestJobSource,
      allSources: sources.map(s => s.source).join(',')
    };
  });
  
  return Object.values(byCode).map(g => g.finalItem);
}

function NF_analyzeLeadtimes_(merged) {
  return merged.map(item => {
    let leadTimeDays = '';
    if (item.jobDoneTs && item.deliveredTs) {
      const diffMs = item.deliveredTs.getTime() - item.jobDoneTs.getTime();
      if (diffMs > 0) {
        leadTimeDays = (diffMs / (24 * 60 * 60 * 1000)).toFixed(2);
      }
    }
    
    return {
      ...item,
      leadTimeDays,
      jobDoneFmt: item.jobDoneTs ? NF_fmtDT_(item.jobDoneTs) : '',
      deliveredFmt: item.deliveredTs ? NF_fmtDT_(item.deliveredTs) : ''
    };
  });
}

function NF_writeDetailSheet_(sheet, analyzed) {
  sheet.clear();
  
  const headers = [
    'TrackingCode','Country','Carrier','JobDone','Delivered','LeadTime(days)',
    'JobDoneSource','AllSources','Source'
  ];
  
  const rows = analyzed.map(item => [
    item.trackingCode || '',
    item.country || '',
    item.carrier || '',
    item.jobDoneFmt || '',
    item.deliveredFmt || '',
    item.leadTimeDays || '',
    item.jobDoneSource || '',
    item.allSources || '',
    item.source || ''
  ]);
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

function NF_buildWeeklyCountryKPI_(analyzed) {
  const weeklyData = {};
  
  analyzed.forEach(item => {
    if (!item.deliveredTs || !item.leadTimeDays) return;
    
    const deliveredDate = item.deliveredTs;
    const year = deliveredDate.getFullYear();
    const week = Utilities.formatDate(deliveredDate, Session.getScriptTimeZone()||'Europe/Helsinki', "YYYY-'W'ww");
    const country = item.country || 'Unknown';
    const carrier = item.carrier || 'Unknown';
    
    const key = `${week}|${country}|${carrier}`;
    
    if (!weeklyData[key]) {
      weeklyData[key] = {
        week, year, country, carrier,
        leadTimes: [],
        count: 0
      };
    }
    
    weeklyData[key].leadTimes.push(parseFloat(item.leadTimeDays));
    weeklyData[key].count++;
  });
  
  // Calculate averages
  return Object.values(weeklyData).map(group => {
    const avgLeadTime = group.leadTimes.length > 0 
      ? (group.leadTimes.reduce((a,b) => a+b, 0) / group.leadTimes.length).toFixed(2)
      : '';
    
    return {
      ...group,
      avgLeadTime
    };
  }).sort((a,b) => {
    // Sort by week desc, then country, then carrier
    if (a.week !== b.week) return b.week.localeCompare(a.week);
    if (a.country !== b.country) return a.country.localeCompare(b.country);
    return a.carrier.localeCompare(b.carrier);
  });
}

function NF_writeWeeklyKPI_(sheet, weeklyKPI) {
  sheet.clear();
  
  const headers = ['Week','Country','Carrier','Count','AvgLeadTime(days)'];
  const rows = weeklyKPI.map(kpi => [
    kpi.week || '',
    kpi.country || '',
    kpi.carrier || '',
    kpi.count || 0,
    kpi.avgLeadTime || ''
  ]);
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

/**
 * SOK & Kärkkäinen weekly reports (Sun→Sun)
 */
function NF_buildSokKarkkainenWeekly() {
  const ss = SpreadsheetApp.getActive();
  const sourceSheet = ss.getSheetByName('Packages');
  if (!sourceSheet || sourceSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('NewFlow: Packages-taulukko on tyhjä');
    return;
  }
  
  const { start, end } = NF_lastFinishedWeek_();
  const data = sourceSheet.getDataRange().getDisplayValues();
  const hdr = data[0]; const rows = data.slice(1);
  const map = NF_mapHdr_(hdr);
  
  // Find date and payer columns
  const dateIdx = NF_findCol_(map, ['Submitted date','Created','Created date','Booking date','Booked time','Dispatch date','Shipped date','Timestamp','Date']);
  const payerIdx = NF_findCol_(map, ['Payer','Freight account','Billing account','Customer number','Customer ID','Customer #']);
  
  // Filter rows by date window
  const inWindow = rows.filter(r => {
    if (dateIdx < 0) return true;
    const dt = NF_parseDate_(r[dateIdx]);
    return dt && dt >= start && dt < end;
  });
  
  // Split by payer
  const sokAccount = NF_Prop_('SOK_FREIGHT_ACCOUNT', '990719901');
  const karkkainenNumbers = NF_Prop_('KARKKAINEN_NUMBERS', '615471,802669,7030057').split(',').map(s => s.trim());
  
  const sokRows = [];
  const karkkainenRows = [];
  
  if (payerIdx >= 0) {
    const sokDigits = NF_digits_(sokAccount);
    const karkkainenDigitsSet = new Set(karkkainenNumbers.map(NF_digits_));
    
    for (const r of inWindow) {
      const payerDigits = NF_digits_(r[payerIdx] || '');
      if (!payerDigits) continue;
      
      if (payerDigits === sokDigits) {
        sokRows.push(r);
      } else if (karkkainenDigitsSet.has(payerDigits)) {
        karkkainenRows.push(r);
      }
    }
  }
  
  // Write reports
  NF_writeWeeklyReport_(ss, 'NF_Report_SOK', hdr, sokRows, start, end);
  NF_writeWeeklyReport_(ss, 'NF_Report_Karkkainen', hdr, karkkainenRows, start, end);
  
  ss.toast(`NewFlow: Viikkoraportit valmis | SOK=${sokRows.length} | Kärkkäinen=${karkkainenRows.length}`);
}

function NF_writeWeeklyReport_(ss, sheetName, headers, rows, start, end) {
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear();
  
  // Headers
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Info row
  const infoRow = new Array(headers.length).fill('');
  infoRow[0] = `Viikko (SUN→SUN): ${NF_fmtDT_(start).substring(0,10)} - ${NF_fmtDT_(end).substring(0,10)}`;
  if (headers.length > 1) infoRow[1] = `Rivejä: ${rows.length}`;
  if (headers.length > 2) infoRow[2] = `Luotu: ${NF_fmtDT_(new Date())}`;
  sh.getRange(2, 1, 1, headers.length).setValues([infoRow]);
  
  // Style info row
  sh.getRange(2, 1, 1, Math.min(headers.length, 3)).setFontStyle('italic');
  
  // Data rows
  if (rows && rows.length > 0) {
    const normalizedRows = rows.map(r => {
      const rr = r.slice(0, headers.length);
      while (rr.length < headers.length) rr.push('');
      return rr;
    });
    sh.getRange(4, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }
  
  sh.setFrozenRows(3);
  sh.autoResizeColumns(1, Math.min(headers.length, 20));
}