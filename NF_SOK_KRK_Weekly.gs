/**
 * NF_SOK_KRK_Weekly.gs — SOK and Kärkkäinen weekly report management
 *
 * Provides functions for building and maintaining SOK and Kärkkäinen weekly reports
 * with historical data merge capabilities and date window utilities.
 *
 * Key Functions:
 *   - NF_buildSokKarkkainenForWindow_(start, end): Build last-week windowed reports
 *   - NF_buildSokKarkkainenAlways(): Merge all historical rows into reports
 *   - NF_getLastFinishedWeekSunWindow_(): Get last completed Sunday-to-Sunday week
 *   - NF_ReconcileWeeklyFromImport(): Append missing deliveries from PBI import
 */

// Defaults (used only as fallback if Script Properties are not set)
const SOK_FREIGHT_ACCOUNT = '5010';
const KARKKAINEN_NUMBERS = ['1234', '5678', '9012']; // Update with actual Kärkkäinen accounts

/********************* WINDOW ***************************/

/**
 * Returns the last finished week's Sunday-to-Sunday time window.
 * Start = previous Sunday 00:00, End = this Sunday 00:00 (exclusive).
 * @return {{start: Date, end: Date}}
 */
function NF_getLastFinishedWeekSunWindow_() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  const dayOfWeek = now.getDay(); // 0 = Sunday

  // Most recent Sunday (today if Sunday)
  const thisSunday = new Date(now);
  thisSunday.setDate(now.getDate() - dayOfWeek);

  // Last finished week ends at this Sunday, starts 7 days before
  const end = new Date(thisSunday);
  const start = new Date(thisSunday);
  start.setDate(thisSunday.getDate() - 7);

  try {
    Logger.log(`Last finished week: ${start.toISOString().split('T')[0]} to ${end.toISOString().split('T')[0]}`);
  } catch (_) {}

  return { start, end };
}

/********************* BUILDERS ***************************/

/**
 * Build SOK and Kärkkäinen reports for specific week window.
 * @param {Date} start - Week start (Sunday)
 * @param {Date} end   - Week end (Sunday, exclusive)
 */
function NF_buildSokKarkkainenForWindow_(start, end) {
  const ss = SpreadsheetApp.getActive();
  const targetSheet = (PropertiesService.getScriptProperties().getProperty('TARGET_SHEET') || 'Packages');
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
    const { sok, krkSet } = NF_getAccounts_();

    for (const row of inWindow) {
      const payerValue = NF_normalizeDigits_(row[payerIndex]);
      if (!payerValue) continue;

      if (sok && payerValue === sok) {
        sokRows.push(row);
      } else if (krkSet.has(payerValue)) {
        karkkainenRows.push(row);
      }
    }
  }

  // Write sheets (with info header rows)
  NF_writeWeeklySheet_(ss, 'Report_SOK', headers, sokRows, start, end);
  NF_writeWeeklySheet_(ss, 'Report_Karkkainen', headers, karkkainenRows, start, end);

  ss.toast(`SOK=${sokRows.length} | Kärkkäinen=${karkkainenRows.length}`, 'Weekly Reports Built');
}

/**
 * Merges all rows historically into Report_SOK and Report_Karkkainen using payer split rules.
 * Aggregates from multiple likely sources and writes flat reports (no week window).
 */
function NF_buildSokKarkkainenAlways() {
  const ss = SpreadsheetApp.getActive();

  // Candidate sources (can be extended via Script Properties CSV: NF_ALL_SOURCES)
  const propSources = (PropertiesService.getScriptProperties().getProperty('NF_ALL_SOURCES') || '').split(',').map(s => s.trim()).filter(Boolean);
  const defaultSources = ['Packages', 'Import_Weekly', 'PowerBI_New', 'Packages_Archive'];
  const sourceSheets = (propSources.length ? propSources : defaultSources);

  const tables = [];

  const addTable = (sheetName) => {
    const sh = ss.getSheetByName(sheetName);
    if (sh && sh.getLastRow() > 1) {
      const data = sh.getDataRange().getDisplayValues();
      tables.push({
        source: sheetName,
        headers: data[0].map(String),
        rows: data.slice(1)
      });
    } else {
      Logger.log(`NF_buildSokKarkkainenAlways: ${sheetName} missing or empty, skipped`);
    }
  };

  sourceSheets.forEach(addTable);

  if (!tables.length) {
    throw new Error('No data found in source tables.');
  }

  // Merge headers from all tables
  let unionHeaders = [];
  tables.forEach(table => {
    unionHeaders = NF_mergeHeaders_(unionHeaders, table.headers);
  });

  // Ensure we have a payer index after union
  const payerIndex = NF_pickPayerIndex_(unionHeaders);
  if (payerIndex < 0) {
    Logger.log('NF_buildSokKarkkainenAlways: payer column not found in union headers, aborting.');
    return;
  }

  const { sok, krkSet } = NF_getAccounts_();

  // Collect and classify all rows into union shape
  const headerMap = NF_headerIndexMap_(unionHeaders);
  const sokRows = [];
  const karkkainenRows = [];

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

      if (sok && payerValue === sok) {
        sokRows.push(unifiedRow);
      } else if (krkSet.has(payerValue)) {
        karkkainenRows.push(unifiedRow);
      }
    }
  }

  // Write flat reports
  NF_writeFlat_(ss, 'Report_SOK', unionHeaders, sokRows);
  NF_writeFlat_(ss, 'Report_Karkkainen', unionHeaders, karkkainenRows);

  ss.toast(`SOK=${sokRows.length} | Kärkkäinen=${karkkainenRows.length}`, 'Full Reports Built');
}

/**
 * Appends missing deliveries from PBI import to SOK/KRK weekly reports for the last finished week.
 */
function NF_ReconcileWeeklyFromImport() {
  Logger.log('NF_ReconcileWeeklyFromImport: Reconciling weekly reports with Import_Weekly...');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('Import_Weekly');

  if (!importSheet || importSheet.getLastRow() < 2) {
    Logger.log('Import_Weekly sheet not found or empty, skipping reconciliation...');
    return;
  }

  const importData = importSheet.getDataRange().getValues();
  const importHeaders = importData[0];

  // Find relevant columns
  const payerIndex = NF_findPayerIndex_(importHeaders);
  const dateIndex = NF_findDateIndex_(importHeaders);

  if (payerIndex < 0) {
    Logger.log('Payer column not found in Import_Weekly');
    return;
  }

  // Get last week's window
  const { start, end } = NF_getLastFinishedWeekSunWindow_();

  // Filter import data for last week
  const lastWeekData = [];
  for (let rowIndex = 1; rowIndex < importData.length; rowIndex++) {
    const row = importData[rowIndex];

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

  if (!lastWeekData.length) {
    Logger.log('No data found for last week in Import_Weekly');
    return;
  }

  // Split by payer and append
  const { sok, krkSet } = NF_getAccounts_();
  const sokEntries = [];
  const krkEntries = [];

  for (const row of lastWeekData) {
    const payerValue = NF_normalizeDigits_(row[payerIndex]);
    if (!payerValue) continue;

    if (sok && payerValue === sok) {
      sokEntries.push(row);
    } else if (krkSet.has(payerValue)) {
      krkEntries.push(row);
    }
  }

  if (sokEntries.length > 0) {
    NF_appendToReportSheet_(ss, 'Report_SOK', importHeaders, sokEntries);
  }
  if (krkEntries.length > 0) {
    NF_appendToReportSheet_(ss, 'Report_Karkkainen', importHeaders, krkEntries);
  }

  Logger.log(`Reconciliation completed: SOK=${sokEntries.length}, Kärkkäinen=${krkEntries.length} appended`);
}

/********************* HELPERS ***************************/

/**
 * Read accounts from Script Properties, fallback to file-level defaults.
 * @return {{sok: string, krkSet: Set<string>}}
 */
function NF_getAccounts_() {
  const props = PropertiesService.getScriptProperties();
  const sokProp = (props.getProperty('SOK_FREIGHT_ACCOUNT') || '').trim();
  const krkProp = (props.getProperty('KARKKAINEN_NUMBERS') || '').trim();

  const sok = NF_normalizeDigits_(sokProp || SOK_FREIGHT_ACCOUNT || '');
  const krkList = (krkProp || (Array.isArray(KARKKAINEN_NUMBERS) ? KARKKAINEN_NUMBERS.join(',') : '') || '')
    .split(',')
    .map(NF_normalizeDigits_)
    .filter(Boolean);

  return { sok, krkSet: new Set(krkList) };
}

/**
 * Merge header arrays (case-insensitive), preserving original casing of first occurrence.
 * Falls back to existing mergeHeaders_ if available.
 */
function NF_mergeHeaders_(oldHeaders, newHeaders) {
  if (typeof mergeHeaders_ === 'function') {
    try { return mergeHeaders_(oldHeaders, newHeaders); } catch (_) {}
  }
  const result = (oldHeaders || []).slice();
  const seen = new Set(result.map(h => String(h || '').toLowerCase().trim()));
  for (const h of (newHeaders || [])) {
    const key = String(h || '').toLowerCase().trim();
    if (!seen.has(key)) { result.push(h); seen.add(key); }
  }
  return result;
}

/**
 * Extend a row to match union headers by header text (case-insensitive).
 */
function NF_extendRowToHeaders_(row, sourceHeaders, unionHeaders) {
  const out = new Array(unionHeaders.length).fill('');
  const u = (unionHeaders || []).map(h => String(h || '').toLowerCase().trim());
  for (let i = 0; i < (sourceHeaders || []).length && i < (row || []).length; i++) {
    const key = String(sourceHeaders[i] || '').toLowerCase().trim();
    const idx = u.indexOf(key);
    if (idx >= 0) out[idx] = row[i];
  }
  return out;
}

/**
 * Create header index map (exact name match).
 * Falls back to headerIndexMap_ if present.
 */
function NF_headerIndexMap_(headers) {
  if (typeof headerIndexMap_ === 'function') {
    try { return headerIndexMap_(headers); } catch (_) {}
  }
  const map = {};
  (headers || []).forEach((h, i) => map[h] = i);
  return map;
}

/**
 * Find a 'payer' column index by loose matching.
 */
function NF_findPayerIndex_(headers) {
  const cands = ['payer', 'freight account', 'freight', 'billing', 'customer', 'account', 'invoice account'];
  for (let i = 0; i < (headers || []).length; i++) {
    const h = String(headers[i] || '').toLowerCase();
    if (cands.some(k => h.includes(k))) return i;
  }
  return -1;
}

/**
 * Find any date-like column index by loose matching.
 */
function NF_findDateIndex_(headers) {
  const cands = ['date', 'created', 'submitted', 'booking', 'dispatch', 'shipped', 'timestamp'];
  for (let i = 0; i < (headers || []).length; i++) {
    const h = String(headers[i] || '').toLowerCase();
    if (cands.some(k => h.includes(k))) return i;
  }
  return -1;
}

/**
 * Pick date index for filtering window (exact-ish candidates).
 */
function NF_pickDateIndex_(headers) {
  const dateCandidates = [
    'Submitted date', 'Created', 'Created date', 'Booking date', 'Booked time',
    'Dispatch date', 'Shipped date', 'Timestamp', 'Date'
  ];
  return NF_pickAnyIndex_(headers, dateCandidates);
}

/**
 * Pick payer index for SOK/Kärkkäinen classification (exact-ish candidates).
 */
function NF_pickPayerIndex_(headers) {
  const payerCandidates = [
    'Invoice account', 'Payer', 'Freight account', 'Billing account',
    'Customer number', 'Customer ID', 'Customer #', 'Account'
  ];
  return NF_pickAnyIndex_(headers, payerCandidates);
}

/**
 * Try to pick the first existing column index from a candidate list.
 * If a global pickAnyIndex_ exists, use it.
 */
function NF_pickAnyIndex_(headers, candidates) {
  if (typeof pickAnyIndex_ === 'function') {
    try { return pickAnyIndex_(headers, candidates); } catch (_) {}
  }
  const normalized = (headers || []).map(h => NF_normalize_(h));
  for (const c of (candidates || [])) {
    const idx = normalized.indexOf(NF_normalize_(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

/**
 * Normalize headline text (lowercase, trim, collapse spaces, strip punctuation).
 */
function NF_normalize_(text) {
  return String(text || '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[^\p{L}\p{N}]+/gu, ' ');
}

/**
 * Keep only digits (for payer/account matching).
 */
function NF_normalizeDigits_(text) {
  return String(text || '').replace(/\D/g, '');
}

/**
 * Parse date flexibly (supports common formats).
 */
function NF_parseDateFlexible_(value) {
  if (!value) return null;
  try {
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
  } catch (_) {
    return null;
  }
}

/**
 * Format date as YYYY-MM-DD (sheet timezone if possible).
 */
function NF_dateToYMD_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (_) {
    return date.toISOString().split('T')[0];
  }
}

/**
 * Format datetime as YYYY-MM-DD HH:mm:ss (sheet timezone if possible).
 */
function NF_formatDateTime_(date) {
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (_) {
    return date.toISOString().replace('T', ' ').substring(0, 19);
  }
}

/********************* WRITERS ***************************/

/**
 * Write weekly sheet with info header rows.
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
      const r = row.slice(0, headers.length);
      while (r.length < headers.length) r.push('');
      return r;
    });
    sheet.getRange(4, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }

  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, Math.min(headers.length, 20));
}

/**
 * Write flat sheet (no week info, for full rebuild).
 */
function NF_writeFlat_(spreadsheet, sheetName, headers, rows) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  sheet.clear();

  // Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Data rows
  if (rows && rows.length) {
    const normalizedRows = rows.map(row => {
      const r = row.slice(0, headers.length);
      while (r.length < headers.length) r.push('');
      return r;
    });
    sheet.getRange(2, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }

  sheet.autoResizeColumns(1, Math.min(headers.length, 20));
}

/**
 * Write data to a report sheet (simple writer).
 */
function NF_writeReportSheet_(ss, sheetName, headers, data) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();

  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  if (data.length > 0 && headers.length > 0) {
    const normalizedData = data.map(row => {
      const out = new Array(headers.length).fill('');
      for (let i = 0; i < Math.min(row.length, headers.length); i++) {
        out[i] = row[i];
      }
      return out;
    });
    sheet.getRange(2, 1, normalizedData.length, headers.length).setValues(normalizedData);
  }

  if (headers.length > 0) {
    sheet.autoResizeColumns(1, Math.min(headers.length, 20));
  }
}

/**
 * Append data to an existing report sheet, aligned to existing headers.
 */
function NF_appendToReportSheet_(ss, sheetName, sourceHeaders, data) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Report sheet ${sheetName} not found for appending`);
    return;
  }
  if (!data.length) return;

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mapped = data.map(row => NF_extendRowToHeaders_(row, sourceHeaders, existingHeaders));

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, mapped.length, existingHeaders.length).setValues(mapped);

  Logger.log(`Appended ${mapped.length} rows to ${sheetName}`);
}