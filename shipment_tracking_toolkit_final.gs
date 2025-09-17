// Shipment Tracking Toolkit – Final unified script

// This script integrates all features for automated shipment tracking and reporting
// in Google Sheets via Apps Script. It includes Gmail imports for shipment reports
// (CSV and XLSX via Drive API conversion), building of a pending shipments sheet,
// status updates across multiple carriers (Posti, GLS Finland, DHL, Bring,
// Matkahuolto) with proper rate limiting, dynamic throttling, and fallback
// endpoints. It also implements bulk refresh logic, weekly report generation,
// Quick actions for confirming and archiving delivered packages, error logging,
// scheduling functions, and a credentials hub panel for storing API keys.

// The code below is a cohesive, ready-to-deploy Apps Script file. Before
// deploying, enable the Google Drive advanced service for XLSX conversions
// and set your API keys and URLs in the Script Properties using the provided
// Credentials Hub.

/* Configuration constants */
const ACTION_SHEET = 'Vaatii_toimenpiteitä';
const TARGET_SHEET = 'Packages';
const ARCHIVE_SHEET = 'Packages_Archive';
const PBI_IMPORT_SHEET = 'PowerBI_Import';
const PBI_NEW_SHEET = 'PowerBI_New';

// Bulk refresh parameters
const BULK_MAX_API_CALLS_PER_RUN = 300;
const BULK_TIME_LIMIT_MS = 5 * 60 * 1000 - 15 * 1000;
const BULK_BACKOFF_MINUTES_BASE = 30;
const BULK_BACKOFF_MINUTES_MAX = 24 * 60;
const BULK_COOLDOWN_NO_CODE_H = 24;
const SLA_OPEN_DAYS = 5;
const SLA_PD_DAYS = 5;

// Helper to fetch script properties
function TRK_props_(k) {
  return PropertiesService.getScriptProperties().getProperty(k) || '';
}

// Credential panel to manage API keys
function openCredsPanel() {
  const t = HtmlService.createTemplateFromFile('Creds');
  t.props = credGetProps();
  const html = t.evaluate()
    .setTitle('Integraatioavaimet')
    .setWidth(760)
    .setHeight(760);
  SpreadsheetApp.getUi().showSidebar(html);
}

function credGetProps() {
  const sp = PropertiesService.getScriptProperties();
  const keys = [
    'MH_TRACK_URL','MH_BASIC',
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL','POSTI_TRK_URL','POSTI_TRK_BASIC','POSTI_TRK_USER','POSTI_TRK_PASS',
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY','GLS_FI_SENDER_ID','GLS_FI_SENDER_IDS','GLS_FI_METHOD',
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    'DHL_TRACK_URL','DHL_API_KEY',
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    'BULK_BACKOFF_MINUTES_BASE',
    'PBI_WEBHOOK_URL'
  ];
  const props = {};
  keys.forEach(k => props[k] = sp.getProperty(k) || '');
  return props;
}

function credSaveProps(data) {
  const sp = PropertiesService.getScriptProperties();
  try {
    Object.keys(data).forEach(k => {
      if (data[k] !== '') sp.setProperty(k, String(data[k]));
      else sp.deleteProperty(k);
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}

function credTestTrack(carrier, code) {
  if (!carrier || !code) return { error: 'Anna kuljetusyhtiö ja koodi' };
  try {
    return TRK_trackByCarrier_(carrier, code);
  } catch (e) {
    return { error: String(e) };
  }
}

function showCredentialsHub() {
  openCredsPanel();
}

// Normalisation helpers
function canonicalCarrier_(s) {
  const c = String(s || '').toLowerCase();
  if (/posti/.test(c)) return 'posti';
  if (/gls/.test(c)) return 'gls';
  if (/dhl/.test(c)) return 'dhl';
  if (/bring/.test(c)) return 'bring';
  if (/matkahuolto|mh/.test(c)) return 'matkahuolto';
  return 'other';
}

function dateToYMD_(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function fmtDateTime_(d) {
  return `${dateToYMD_(d)} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
}
function parseDateFlexible_(v) {
  if (!v) return null;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5));
  const s = String(v).trim().replace(/[.;]/g,'-').replace(/T/,' ');
  const iso = s.match(/^\d{4}-\d{2}-\d{2}/) ? s : s.replace(/^(\d{1,2})[-.](\d{1,2})[-.](\d{2,4})/, '$3-$2-$1');
  const dt = new Date(iso);
  return isNaN(dt) ? null : dt;
}
function normalize_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/[^\p{L}\p{N}]+/gu,' ');
}
function headerIndexMap_(hdr) {
  const m = {}; (hdr || []).forEach((h,i) => m[h] = i);
  return m;
}
function mergeHeaders_(a,b) {
  return Array.from(new Set([...(a || []), ...(b || [])]));
}
function firstCode_(cell) {
  return String(cell || '').split(/[\n,;]/)[0].trim();
}

// Rate limiting functions
function trkRateLimitWait_(carrier) {
  const sp = PropertiesService.getScriptProperties();
  const keyLast = `RATE_LAST_${carrier.toUpperCase()}`;
  const keyMinMs = `RATE_MINMS_${carrier.toUpperCase()}`;
  const last = Number(sp.getProperty(keyLast) || '0');
  const minMs = Number(sp.getProperty(keyMinMs) || '0');
  const now = Date.now();
  const elapsed = now - last;
  if (minMs && elapsed < minMs) Utilities.sleep(minMs - elapsed);
  sp.setProperty(keyLast, String(Date.now()));
}

function autoTuneRateLimitOn429_(carrier, addMs, capMs) {
  const sp = PropertiesService.getScriptProperties();
  const key = `RATE_MINMS_${carrier.toUpperCase()}`;
  const current = Number(sp.getProperty(key) || '0');
  const updated = Math.min(current + addMs, capMs || current + addMs);
  sp.setProperty(key, String(updated));
}

function getCfgInt_(key, fallback) {
  const sp = PropertiesService.getScriptProperties();
  const v = sp.getProperty(key);
  return v ? Number(v) : fallback;
}

// Menu and UI
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipment')
    .addItem('Credentials Hub', 'showCredentialsHub')
    .addItem('Hae Gmail → Päivitä taulut nyt', 'runDailyFlowOnce')
    .addItem('Rakenna Vaatii_toimenpiteitä (ei-toimitetut)', 'buildPendingFromPackagesAndArchive')
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
    .addSeparator()
    .addItem('Asetuspaneeli (pop-in)', 'openControlPanel')
    .addSeparator()
    .addItem('Rakenna SOK & Kärkkäinen (aina)', 'buildSokKarkkainenAlways')
    .addItem('Tee viikkoraportit (SUN→SUN) + status', 'makeWeeklyReportsSunSun')
    .addItem('Arkistoi toimitetut (Vaatii_toimenpiteitä)', 'archiveDeliveredFromVaatii')
    .addItem('Aseta valinnaiset oletukset', 'seedOptionalDefaults')
    .addItem('Täytä API-avaimet & asiakasnumerot', 'seedKnownAccountsAndKeys')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Päivitä STATUS (valittu yhtiö)')
        .addItem('Matkahuolto', 'menuRefreshCarrier_MH')
        .addItem('Posti', 'menuRefreshCarrier_POSTI')
        .addItem('Bring', 'menuRefreshCarrier_BRING')
        .addItem('GLS', 'menuRefreshCarrier_GLS')
        .addItem('DHL', 'menuRefreshCarrier_DHL')
        .addItem('Kaikki (ei suodatusta)', 'menuRefreshCarrier_ALL')
    )
    .addSubMenu(
      ui.createMenu('Iso statusajo (turvallinen)')
        .addItem('Aloita (aktiivinen välilehti)', 'bulkStartForActiveSheet')
        .addItem('Aloita: Vaatii_toimenpiteitä', 'bulkStart_Vaatii')
        .addItem('Aloita: Packages', 'bulkStart_Packages')
        .addItem('Aloita: Packages_Archive', 'bulkStart_Archive')
        .addItem('Pysäytä bulk-ajo', 'bulkStop')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Adhoc (Power BI / XLSX)')
        .addItem('Tuo Drive-URL/ID → Adhoc_Tracking', 'adhocImportFromDriveFile')
        .addItem('Tuo Gmailin uusin "Outbound order" -liite', 'adhocImportLatestOutboundFromGmail')
        .addItem('Päivitä statukset (Adhoc_Tracking)', 'adhocRefresh')
    )
    .addSubMenu(
      ui.createMenu('Power BI (historia)')
        .addItem('Tuo Power BI -tiedosto → suodata uudet', 'powerBiImportAndFilter')
        .addItem('Arkistoi PBI-uudet → Packages_Archive', 'powerBiArchiveNew')
        .addItem('Käynnistä Power BI -päivitys (webhook)', 'triggerPowerBI')
    )
    .addSeparator()
    .addItem('Gmail: tuo KAIKKI (vanhin → uusin)', 'fetchHistoryFromGmailOldToNew')
    .addSeparator()
    .addItem('Toimitusajat – listaus', 'makeDeliveryTimeReport')
    .addItem('Toimitusajat viikkotasolla maittain', 'makeCountryWeekLeadtimeReport')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Tarkistimet')
        .addItem('Puuttuvat sarakeotsikot (aktiivinen välilehti)', 'showMissingColumns_Active')
        .addItem('Integraatioavaimet (Script Properties)', 'showMissingProperties')
    )
    .addSeparator()
    .addItem('Ajasta: arkipäivät klo 12:00', 'setupWeekday1200_DailyFlow')
    .addItem('Ajasta: viikkoraportti maanantai 02:00', 'setupWeeklyMon0200')
    .addSeparator()
    .addItem('Tarkista arkiston tuplat', 'checkArchiveDuplicates')
    .addItem('Poista kaikki ajastukset', 'clearAllTriggers')
    .addItem('Tyhjennä seurannan välimuisti', 'trkClearCache')
    .addItem('Näytä edistyminen', 'showProgressSidebar')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Pikatoiminnot / Vaatii')
        .addItem('Vahvista toimitukset (Vaatii)', 'confirmMissingDeliveredTimes')
        .addItem('Vahvista + Arkistoi (Vaatii)', 'confirmAndArchiveDelivered')
    )
    .addToUi();
}

function onInstall() {
  onOpen();
}

// Gmail attachment handling
function readAttachmentToValues_(blob, filename) {
  if (!blob) return null;
  filename = filename || blob.getName();
  try {
    if (/\.xlsx$/i.test(filename) && typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.create === 'function') {
      const file = Drive.Files.create({ title: 'AdhocUploadTemp', mimeType: MimeType.GOOGLE_SHEETS }, blob);
      const tempFileId = file.id;
      const ss = SpreadsheetApp.openById(tempFileId);
      const values = ss.getSheets()[0].getDataRange().getValues();
      DriveApp.getFileById(tempFileId).setTrashed(true);
      return values;
    }
    const csv = Utilities.parseCsv(blob.getDataAsString('UTF-8'));
    return csv;
  } catch (e) {
    SpreadsheetApp.getUi().alert('Virhe liitteen lukemisessa: ' + e);
    return null;
  }
}

function sanitizeMatrix_(matrix) {
  if (!matrix || !matrix.length) return matrix;
  let lastNonEmpty = matrix.length;
  for (let r = matrix.length - 1; r >= 0; r--) {
    if (!matrix[r].join('')) lastNonEmpty = r;
    else break;
  }
  return matrix.slice(0, lastNonEmpty);
}

function chooseKeyIndex_(headers) {
  const hdr = headers.map(h => normalize_(h));
  const keys = ['tilaus','order','toimitus','shipment','delivery'];
  for (let i=0; i<hdr.length; i++) if (keys.some(k => hdr[i].includes(k))) return i;
  return 0;
}

function findLatestAttachment_() {
  const threads = GmailApp.search('filename:(Outbound) OR subject:(Outbound)');
  if (!threads.length) return null;
  const msgs = GmailApp.getMessagesForThreads([threads[0]])[0];
  for (let i = msgs.length - 1; i >= 0; i--) {
    const atts = msgs[i].getAttachments();
    for (const att of atts) {
      if (/Outbound/i.test(att.getName())) return { blob: att, name: att.getName() };
    }
  }
  return null;
}

function fetchAndRebuild() {
  const attachment = findLatestAttachment_();
  if (!attachment) return;
  const values = readAttachmentToValues_(attachment.blob, attachment.name);
  if (!values) return;
  const matrix = sanitizeMatrix_(values);
  if (!matrix || matrix.length < 2) return;
  const ss = SpreadsheetApp.getActive();
  const sheetName = attachment.name.toLowerCase().includes('archive') ? ARCHIVE_SHEET : TARGET_SHEET;
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clear();
  sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
}

function fetchHistoryFromGmailOldToNew() {
  const threads = GmailApp.search('filename:(Outbound) OR subject:(Outbound)');
  if (!threads.length) {
    SpreadsheetApp.getUi().alert('Ei löytynyt yhtään Outbound-viestiä Gmailista.');
    return;
  }
  const ss = SpreadsheetApp.getActive();
  let sheetPack = ss.getSheetByName(TARGET_SHEET);
  if (!sheetPack) sheetPack = ss.insertSheet(TARGET_SHEET);
  let sheetArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!sheetArch) sheetArch = ss.insertSheet(ARCHIVE_SHEET);
  sheetPack.clear(); sheetArch.clear();
  let headersSet = false;
  for (let ti = threads.length-1; ti>=0; ti--) {
    const msgs = GmailApp.getMessagesForThread(threads[ti]);
    for (let mi = msgs.length-1; mi>=0; mi--) {
      const atts = msgs[mi].getAttachments();
      for (const att of atts) {
        const name = att.getName();
        if (/Outbound.*\.xlsx|Outbound.*\.csv/i.test(name)) {
          const values = readAttachmentToValues_(att, name);
          if (!values) continue;
          const matrix = sanitizeMatrix_(values);
          if (!matrix || matrix.length < 2) continue;
          const isArchive = name.toLowerCase().includes('archive');
          const sh = isArchive ? sheetArch : sheetPack;
          if (!headersSet) {
            sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
            headersSet = true;
          } else {
            const oldData = sh.getDataRange().getValues();
            const headersOld = oldData.shift();
            const headersNew = matrix.shift();
            const merged = mergeHeaders_(headersOld, headersNew);
            const oldMap = headerIndexMap_(headersOld);
            const newMap = headerIndexMap_(headersNew);
            const allData = [merged];
            oldData.forEach(row => {
              const newRow = merged.map(col => (col in oldMap ? row[oldMap[col]] : ''));
              allData.push(newRow);
            });
            matrix.forEach(row => {
              const newRow = merged.map(col => (col in newMap ? row[newMap[col]] : ''));
              allData.push(newRow);
            });
            sh.clear();
            sh.getRange(1,1,allData.length,allData[0].length).setValues(allData);
          }
        }
      }
    }
  }
}

// Build pending shipments sheet
function buildPendingFromPackagesAndArchive() {
  const ss = SpreadsheetApp.getActive();
  const shPack = ss.getSheetByName(TARGET_SHEET);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shPack || !shArch) {
    SpreadsheetApp.getUi().alert('Packages tai Packages_Archive -taulua ei löydy.');
    return;
  }
  const dataPack = shPack.getDataRange().getValues();
  const dataArch = shArch.getDataRange().getValues();
  if (dataPack.length < 2) {
    SpreadsheetApp.getUi().alert('Packages-taulu on tyhjä.');
    return;
  }
  const hdr = dataPack[0];
  const keyIdx = chooseKeyIndex_(hdr);
  const archMap = {};
  const archHdr = dataArch.length ? dataArch[0] : [];
  const archKeyIdx = archHdr.length ? chooseKeyIndex_(archHdr) : -1;
  for (let i=1; i<dataArch.length; i++) {
    const key = dataArch[i][archKeyIdx];
    if (key) archMap[String(key)] = true;
  }
  const pending = [hdr];
  for (let j=1; j<dataPack.length; j++) {
    const key = dataPack[j][keyIdx];
    if (key && !archMap[String(key)]) pending.push(dataPack[j]);
  }
  let shPending = ss.getSheetByName(ACTION_SHEET);
  if (!shPending) shPending = ss.insertSheet(ACTION_SHEET);
  shPending.clear();
  shPending.getRange(1,1,pending.length,pending[0].length).setValues(pending);
}

// Update statuses for a given sheet
function refreshStatuses_Vaatii() {
  refreshStatuses_Sheet(ACTION_SHEET, true);
}

function refreshStatuses_Sheet(sheetName, removeDelivered) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const hdr = data[0];
  const outHdr = ['RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw'];
  let writeHdr = false;
  outHdr.forEach(col => {
    if (!hdr.includes(col)) {
      hdr.push(col);
      writeHdr = true;
    }
  });
  if (writeHdr) {
    sh.getRange(1,1,1,hdr.length).setValues([hdr]);
    data[0] = hdr;
  }
  const colIndex = headerIndexMap_(hdr);
  const now = Date.now();
  const cutoff = 6 * 3600 * 1000;
  const output = [];
  for (let i=1; i<data.length; i++) {
    const row = data[i];
    if (!row.join('')) {
      output.push(['','','','','']);
      continue;
    }
    const carrier = row[colIndex['Carrier']] || row[colIndex['LogisticsProvider']] || row[colIndex['Shipper']] || '';
    const code = firstCode_(row[colIndex['TrackingNumber']]);
    const lastTime = row[colIndex['RefreshTime']] || '';
    const lastTs = lastTime ? new Date(lastTime).getTime() : 0;
    const expired = !lastTs || (now - lastTs > cutoff);
    if (!code || !carrier || !expired) {
      output.push([
        row[colIndex['RefreshCarrier']] || carrier,
        row[colIndex['RefreshStatus']] || '',
        row[colIndex['RefreshTime']] || '',
        row[colIndex['RefreshLocation']] || '',
        row[colIndex['RefreshRaw']] || ''
      ]);
    } else {
      const res = TRK_trackByCarrier_(carrier, code);
      output.push([
        res.carrier || carrier || '',
        res.status || '',
        res.time || fmtDateTime_(new Date()),
        res.location || '',
        res.raw || ''
      ]);
      if (removeDelivered && (String(res.status).toLowerCase().includes('delivered') || String(res.status).toLowerCase().includes('toimitettu'))) {
        archiveRow_(row, sheetName);
      }
    }
  }
  const firstCol = hdr.indexOf('RefreshCarrier') + 1;
  sh.getRange(2, firstCol, output.length, outHdr.length).setValues(output);
}

function archiveRow_(row, sourceSheet) {
  const ss = SpreadsheetApp.getActive();
  const shSource = ss.getSheetByName(sourceSheet);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shSource || !shArch) return;
  shArch.appendRow(row);
  const data = shSource.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i].join() === row.join()) {
      shSource.deleteRow(i+1);
      break;
    }
  }
}

// Bulk refresh helpers
const EXTRA_PRIORITY_CALLS = 0;

function menuRefreshCarrier_MH() { bulkStartForSheetFiltered_(ACTION_SHEET, 'matkahuolto'); }
function menuRefreshCarrier_POSTI() { bulkStartForSheetFiltered_(ACTION_SHEET, 'posti'); }
function menuRefreshCarrier_BRING() { bulkStartForSheetFiltered_(ACTION_SHEET, 'bring'); }
function menuRefreshCarrier_GLS() { bulkStartForSheetFiltered_(ACTION_SHEET, 'gls'); }
function menuRefreshCarrier_DHL() { bulkStartForSheetFiltered_(ACTION_SHEET, 'dhl'); }
function menuRefreshCarrier_ALL() { bulkStartForSheetFiltered_(ACTION_SHEET, null); }

function bulkStartForSheetFiltered_(sheetName, carrierFilter) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const job = { sheet: sheetName, row: 2, calls: 0, started: Date.now(), total: sh.getLastRow()-1, carrierFilter: carrierFilter };
  PropertiesService.getScriptProperties().setProperty('BULKJOB|' + sheetName, JSON.stringify(job));
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('bulkTick').timeBased().everyMinutes(1).create();
  SpreadsheetApp.getUi().alert(`Bulk päivitys aloitettu taululle "${sheetName}"${carrierFilter ? ' (suodatus: '+carrierFilter+')' : ''}.`);
}

function bulkStartForActiveSheet() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh) return;
  bulkStartForSheetFiltered_(sh.getName(), null);
}
function bulkStart_Vaatii() { bulkStartForSheetFiltered_(ACTION_SHEET, null); }
function bulkStart_Packages() { bulkStartForSheetFiltered_(TARGET_SHEET, null); }
function bulkStart_Archive() { bulkStartForSheetFiltered_(ARCHIVE_SHEET, null); }

function bulkStop() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  PropertiesService.getScriptProperties().deleteProperty('BULKJOB|' + ACTION_SHEET);
  SpreadsheetApp.getUi().alert('Bulk-ajo pysäytetty.');
}

function bulkTick() {
  const sp = PropertiesService.getScriptProperties();
  const jobJSON = sp.getProperty('BULKJOB|' + ACTION_SHEET);
  if (!jobJSON) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
    return;
  }
  const job = JSON.parse(jobJSON);
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(job.sheet);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  const hdr = data[0];
  const colIndex = headerIndexMap_(hdr);
  const output = [];
  const maxCalls = Math.min(job.total, job.calls + EXTRA_PRIORITY_CALLS + BULK_MAX_API_CALLS_PER_RUN);
  for (let r=job.row; r<data.length && job.calls < maxCalls; r++) {
    const row = data[r];
    if (!row.join('')) continue;
    const carrier = String(row[colIndex['RefreshCarrier']] || row[colIndex['Carrier']] || '').toLowerCase();
    if (job.carrierFilter && carrier.indexOf(job.carrierFilter) === -1) continue;
    const code = firstCode_(row[colIndex['TrackingNumber']] || '');
    const lastTime = row[colIndex['RefreshTime']] || '';
    const expired = !lastTime || (Date.now() - new Date(lastTime).getTime() > 6*3600*1000);
    if (!code || !carrier || !expired) continue;
    trkRateLimitWait_(carrier);
    const res = TRK_trackByCarrier_(carrier, code);
    output.push({ rowIndex: r, data: [res.carrier || carrier, res.status || '', res.time || fmtDateTime_(new Date()), res.location || '', res.raw || ''] });
    job.calls++;
    job.row = r + 1;
    if (res.status === 'RATE_LIMIT_429') {
      const backoff = getCfgInt_('BULK_BACKOFF_MINUTES_BASE', BULK_BACKOFF_MINUTES_BASE);
      sp.setProperty('PAUSE_UNTIL_' + carrier.toUpperCase(), String(Date.now() + backoff * 60000));
      break;
    }
  }
  if (output.length) {
    const firstCol = hdr.indexOf('RefreshCarrier') !== -1 ? hdr.indexOf('RefreshCarrier') + 1 : hdr.length + 1;
    output.forEach(e => {
      sh.getRange(e.rowIndex + 1, firstCol, 1, e.data.length).setValues([e.data]);
    });
  }
  if (job.row >= data.length || job.calls >= job.total) {
    sp.deleteProperty('BULKJOB|' + job.sheet);
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
    SpreadsheetApp.getUi().alert('Bulk-päivitys valmis: ' + job.calls + ' kutsua suoritettu.');
  } else {
    sp.setProperty('BULKJOB|' + job.sheet, JSON.stringify(job));
  }
}

// Quick actions
function confirmMissingDeliveredTimes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ACTION_SHEET);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const hdr = data[0];
  const statusCol = hdr.findIndex(h => normalize_(h).includes('status'));
  const dateCol = hdr.findIndex(h => normalize_(h).includes('delivered') || normalize_(h).includes('received'));
  if (statusCol === -1 || dateCol === -1) {
    SpreadsheetApp.getUi().alert('Ei löytynyt "Status" tai "Delivered at" -saraketta.');
    return;
  }
  const now = new Date();
  let filled = 0;
  for (let r=1; r<data.length; r++) {
    const status = String(data[r][statusCol] || '').toLowerCase();
    const dateVal = data[r][dateCol];
    if (status.includes('delivered') && !dateVal) {
      sh.getRange(r+1, dateCol+1).setValue(now);
      filled++;
    }
  }
  SpreadsheetApp.getUi().alert(`Lisätty toimitusaika ${filled} riville.`);
}

function confirmAndArchiveDelivered() {
  confirmMissingDeliveredTimes();
  const ss = SpreadsheetApp.getActive();
  const shVaatii = ss.getSheetByName(ACTION_SHEET);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shVaatii || !shArch) return;
  const data = shVaatii.getDataRange().getValues();
  if (data.length < 2) return;
  const hdr = data[0];
  const statusIdx = hdr.findIndex(h => normalize_(h).includes('status'));
  const toMove = [];
  for (let r=1; r<data.length; r++) {
    const status = String(data[r][statusIdx] || '').toLowerCase();
    if (status.includes('delivered')) toMove.push(data[r]);
  }
  toMove.forEach(row => archiveRow_(row, ACTION_SHEET));
}

// Error logging
function logError_(context, error) {
  const ss = SpreadsheetApp.getActive();
  const shLog = ss.getSheetByName('ErrorLog') || ss.insertSheet('ErrorLog');
  shLog.appendRow([new Date(), context, String(error)]);
  if (shLog.getLastRow() > 1000) shLog.deleteRows(1, shLog.getLastRow() - 1000);
}

function openErrorLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ErrorLog');
  if (sh) ss.setActiveSheet(sh);
}

function clearErrorLog() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('ErrorLog');
  if (sh) sh.clear();
}

// Scheduling functions
function setupWeekday1200_DailyFlow() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'runDailyFlowOnce') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('runDailyFlowOnce').timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: päivittäinen ajo klo 12:00 (ma-pe)');
}

function setupWeeklyMon0200() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'makeWeeklyReportsSunSun') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('makeWeeklyReportsSunSun').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: viikkoraportti maanantaisin 02:00.');
}

function clearAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(tr => ScriptApp.deleteTrigger(tr));
  SpreadsheetApp.getUi().alert('Kaikki ajastukset poistettu.');
}

// Carrier API calls
function pickLatestEvent_(events) {
  if (!Array.isArray(events) || !events.length) return null;
  const ts = e => {
    const cand = e.eventDateTime || e.eventTime || e.timestamp || e.dateTime || e.dateIso || e.date || e.time || '';
    const d = new Date(cand);
    return isNaN(d) ? 0 : d.getTime();
  };
  return events.slice().sort((a,b) => ts(a)-ts(b))[events.length-1];
}

function TRK_safeFetch_(url, opt) {
  try { return UrlFetchApp.fetch(url, opt); }
  catch(e) {
    return {
      getResponseCode: () => 0,
      getContentText: () => String(e),
      getAllHeaders: () => ({})
    };
  }
}

// Posti API with OAuth and fallback
function TRK_trackPosti(code) {
  const carrierName = 'Posti';
  trkRateLimitWait_(carrierName);
  const tokenUrl = TRK_props_('POSTI_TOKEN_URL');
  const basic = TRK_props_('POSTI_BASIC');
  const trackUrl = (TRK_props_('POSTI_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (tokenUrl && basic && trackUrl) {
    const tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    let access='';
    try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(e) {}
    if (access) {
      const res = TRK_safeFetch_(trackUrl, {
        method:'get',
        headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
        muteHttpExceptions:true
      });
      const http = res.getResponseCode ? res.getResponseCode() : 0;
      let body='';
      try { body = res.getContentText() || ''; } catch(e) {}
      if (http === 429) {
        autoTuneRateLimitOn429_('POSTI', 100, 1500);
        return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
      }
      if (http >= 200 && http < 300) {
        try {
          const j = JSON.parse(body);
          const pickText = v => (v && typeof v === 'object') ? (v.fi || v.en || v.sv || v.name || v.description || '') : (v || '');
          let status='', time='', location='';
          if (Array.isArray(j?.parcelShipments) && j.parcelShipments.length) {
            const p = j.parcelShipments[0];
            const last = pickLatestEvent_(p.events || []) || p.latestEvent || {};
            status = pickText(last.eventDescription) || pickText(last.eventShortName) || last.status || '';
            time = last.eventDateTime || last.eventTime || last.timestamp || last.dateTime || last.dateIso || last.date || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (!status && (j?.shipments?.length || j?.items?.length)) {
            const ship = j.shipments?.[0] || j.items?.[0];
            const last = (ship?.events && ship.events.length) ? (pickLatestEvent_(ship.events) || {}) : {};
            status = pickText(last.eventDescription) || last.description || last.status || '';
            time = last.timestamp || last.dateTime || last.dateIso || last.date || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (!status && Array.isArray(j?.freightShipments) && j.freightShipments.length) {
            const f = j.freightShipments[0];
            const last = (Array.isArray(f.events) && f.events.length) ? (pickLatestEvent_(f.events) || {}) : (f.latestEvent || f.latestStatus || {});
            status = pickText(last.eventDescription) || last.statusText || last.status || pickText(f.latestStatus?.description) || '';
            time = last.timestamp || last.dateTime || last.dateIso || last.date || f.latestStatus?.timestamp || '';
            location = (last.location && (pickText(last.location.displayName) || pickText(last.location.name))) || last.location || last.city || '';
          }
          if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
        } catch(e) {}
      }
    }
  }
  // fallback Basic
  const fbUrl = (TRK_props_('POSTI_TRK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  let fbBasic = TRK_props_('POSTI_TRK_BASIC');
  const fbUser = TRK_props_('POSTI_TRK_USER');
  const fbPass = TRK_props_('POSTI_TRK_PASS');
  if (fbUrl && (fbBasic || (fbUser && fbPass))) {
    if (!fbBasic) fbBasic = Utilities.base64Encode(fbUser + ':' + fbPass);
    const res = TRK_safeFetch_(fbUrl, {
      method:'get',
      headers:{ 'Authorization':'Basic ' + fbBasic, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) {
      autoTuneRateLimitOn429_('POSTI', 100, 1500);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    }
    if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
    try {
      const j = JSON.parse(body);
      let ev = j.events || j.trackingEvents || j.parcelShipments?.[0]?.events || [];
      ev = Array.isArray(ev) ? ev : [];
      const parseTs = e => Date.parse(e?.eventDateTime || e?.timestamp || e?.dateTime || e?.date || '');
      ev.sort((a,b) => (parseTs(a) || 0) - (parseTs(b) || 0));
      const last = ev[ev.length-1] || {};
      const status = last.eventDescription || last.description || last.status || '';
      const time = last.eventDateTime || last.timestamp || last.dateTime || last.date || '';
      const location = last.location?.name || last.location || last.city || '';
      if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
    } catch(e) {}
    return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
  }
  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

// GLS tracking with FI v2 and fallback
function TRK_trackGLS(code) {
  const carrierName = 'GLS';
  trkRateLimitWait_(carrierName);
  // primary: FI Customer API v2 with x-api-key
  const base = TRK_props_('GLS_FI_TRACK_URL');
  const key = TRK_props_('GLS_FI_API_KEY');
  const senderSingle = TRK_props_('GLS_FI_SENDER_ID');
  const senderList = TRK_props_('GLS_FI_SENDER_IDS');
  const sender = senderSingle || (senderList ? senderList.split(',')[0].trim() : '');
  const method = (TRK_props_('GLS_FI_METHOD') || 'POST').toUpperCase();
  if (base && key) {
    let url = base;
    const headers = { 'x-api-key': key, 'accept':'application/json' };
    let opt = { muteHttpExceptions:true, headers: headers };
    if (method === 'POST') {
      opt.method = 'post';
      opt.contentType = 'application/json';
      opt.payload = JSON.stringify(sender ? { senderId: sender, parcelNo: code } : { parcelNo: code });
    } else {
      url = url + (url.includes('?') ? '&' : '?') + (sender ? ('senderId='+encodeURIComponent(sender)+'&') : '') + 'parcelNo=' + encodeURIComponent(code);
      opt.method = 'get';
    }
    const res = TRK_safeFetch_(url, opt);
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) {
      autoTuneRateLimitOn429_('GLS', 200, 2000);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    }
    if (http >= 200 && http < 300) {
      try {
        const j = JSON.parse(body);
        const p = j?.parcels?.[0] || j?.parcel || j;
        if (p) {
          const ev = Array.isArray(p.events) ? p.events : [];
          const last = ev.length ? (pickLatestEvent_(ev) || {}) : {};
          const status = last.description || p.status || last.code || '';
          const time = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
          const location = [last.city || '', last.postalCode || '', last.country || ''].filter(Boolean).join(' ');
          if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
        }
        // general fallback within FI response
        let ev = j.events || j.trackingEvents || j.parcel?.events || j.parcelShipments?.[0]?.events || j.parcels?.[0]?.events || j.items?.[0]?.events || j.eventList || j.tracking || [];
        ev = Array.isArray(ev) ? ev : [];
        const last = pickLatestEvent_(ev) || {};
        const status = last.eventDescription || last.description || last.eventShortName?.fi || last.eventShortName?.en || last.name || last.status || last.code || '';
        const time = last.eventTime || last.dateTime || last.timestamp || last.dateIso || last.date || last.time || '';
        const location = last.locationName || last.location || last.depot || last.city || '';
        if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
      } catch(e) {}
    }
    // fallback to global
  }
  const tokenUrl = TRK_props_('GLS_TOKEN_URL');
  const basic = TRK_props_('GLS_BASIC');
  const trackUrl = (TRK_props_('GLS_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  if (tokenUrl && basic && trackUrl) {
    const tok = TRK_safeFetch_(tokenUrl, {
      method:'post',
      headers:{ 'Authorization':'Basic ' + basic, 'Content-Type':'application/x-www-form-urlencoded' },
      payload:{ grant_type:'client_credentials' },
      muteHttpExceptions:true
    });
    let access=''; try { access = JSON.parse(tok.getContentText()).access_token || ''; } catch(e) {}
    if (!access) return { carrier: carrierName, status:'TOKEN_FAIL', raw: tok.getContentText().slice(0,500) };
    const res = TRK_safeFetch_(trackUrl, {
      method:'get',
      headers:{ 'Authorization':'Bearer ' + access, 'Accept':'application/json' },
      muteHttpExceptions:true
    });
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) {
      autoTuneRateLimitOn429_('GLS', 200, 2000);
      return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
    }
    if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
    try {
      const j = JSON.parse(body);
      const it = j?.parcels?.[0] || j?.items?.[0] || j?.shipment || j?.consignment?.[0] || j;
      let ev = it?.events || it?.milestones || it?.statusHistory || it?.trackAndTraceInfo?.events || it?.trackings || [];
      ev = Array.isArray(ev) ? ev : [];
      const last = pickLatestEvent_(ev) || {};
      const status = last.description || last.eventDescription || last.name || last.status || last.code || '';
      const time = last.dateTime || last.timestamp || last.dateIso || last.date || last.time || '';
      const location = (last.location && (last.location.name || last.location.displayName)) || last.location || last.depot || last.city || '';
      return { carrier: carrierName, found:!!status, status, time, location, raw: body.slice(0,2000) };
    } catch(e) {}
  }
  return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
}

// DHL API
function TRK_trackDHL(code) {
  const carrierName = 'DHL';
  trkRateLimitWait_(carrierName);
  const url = (TRK_props_('DHL_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const key = TRK_props_('DHL_API_KEY');
  if (!url || !key) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'DHL-API-Key': key, 'Accept':'application/json' }, muteHttpExceptions:true });
  const http = res.getResponseCode ? res.getResponseCode() : 0;
  const body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter: 15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
  try {
    const j = JSON.parse(body);
    if (Array.isArray(j?.parcels) && j.parcels.length) {
      const p = j.parcels[0];
      const ev = Array.isArray(p.events) ? p.events : [];
      const last = ev.length ? (pickLatestEvent_(ev) || {}) : {};
      const status = last.description || p.status || last.code || '';
      const time = p.statusDateTime || last.eventDateTime || last.timestamp || last.dateTime || '';
      const location = [last.city || '', last.postalCode || '', last.country || ''].filter(Boolean).join(', ');
      if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
    }
    if (Array.isArray(j?.shipments) && j.shipments.length) {
      const s = j.shipments[0];
      let ev = Array.isArray(s.events) ? s.events : [];
      if (!ev.length && s.status) ev = [s.status];
      const last = ev.length ? (pickLatestEvent_(ev) || {}) : s.status || {};
      const status = last.description || last.status || s.status?.status || '';
      const time = last.timestamp || last.dateTime || last.date || s.status?.timestamp || '';
      const addr = last.location?.address;
      let location='';
      if (addr) location = [addr.addressLocality, (addr.countryCode || addr.addressCountryCode)].filter(Boolean).join(', ');
      return { carrier: carrierName, found: !!status, status, time, location, raw: body.slice(0,2000) };
    }
  } catch(e) {}
  return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
}

// Bring API
function TRK_trackBring(code) {
  const carrierName = 'Bring';
  trkRateLimitWait_(carrierName);
  const url = (TRK_props_('BRING_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const uid = TRK_props_('BRING_UID'), key = TRK_props_('BRING_KEY'), cli = TRK_props_('BRING_CLIENT_URL');
  if (!url) return { carrier: carrierName, status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'X-MyBring-API-Uid': uid, 'X-MyBring-API-Key': key, 'X-Bring-Client-URL': cli, 'Accept':'application/json' }, muteHttpExceptions:true });
  const http = res.getResponseCode ? res.getResponseCode() : 0;
  const body = res.getContentText ? (res.getContentText() || '') : '';
  if (http === 429) return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 };
  if (http >= 400) return { carrier: carrierName, status:'HTTP_'+http, raw: body.slice(0,1000) };
  try {
    const j = JSON.parse(body);
    const c = j?.consignments?.[0] || j?.items?.[0] || j;
    const ev = Array.isArray(c?.events) ? c.events : [];
    const last = pickLatestEvent_(ev) || {};
    const status = last.description || last.eventDescription || last.status || '';
    const time = last.dateIso || last.timestamp || last.dateTime || last.date || '';
    const location = last.location || last.city || '';
    return { carrier: carrierName, found: !!status, status, time, location, raw: body.slice(0,2000) };
  } catch(e) {
    return { carrier: carrierName, status:'NO_DATA', raw: body.slice(0,2000) };
  }
}

// Matkahuolto API
function TRK_trackMH(code) {
  trkRateLimitWait_('Matkahuolto');
  const url = (TRK_props_('MH_TRACK_URL') || '').replace('{{code}}', encodeURIComponent(code));
  const basic = TRK_props_('MH_BASIC');
  if (!url || !basic) return { carrier:'Matkahuolto', status:'MISSING_CREDENTIALS' };
  const res = TRK_safeFetch_(url, { method:'get', headers:{ 'Authorization':'Basic '+basic, 'Accept':'application/json' }, muteHttpExceptions:true });
  const body = res.getContentText() || '';
  let status='', time='', location='';
  try {
    const j = JSON.parse(body);
    const first = j?.[0] || j?.items?.[0] || j?.consignment || j?.shipment || j;
    let ev = first?.events || first?.history || [];
    ev = Array.isArray(ev) ? ev : [];
    const last = pickLatestEvent_(ev) || {};
    status = last.status || last.description || last.eventDescription || last.name || '';
    time = last.time || last.timestamp || last.dateTime || last.dateIso || last.date || '';
    location = last.location || last.place || last.depot || last.city || '';
  } catch(e) {}
  return { carrier:'Matkahuolto', found: !!status, status, time, location, raw: body.slice(0,2000) };
}

// Dispatcher
function TRK_trackByCarrier_(carrier, code) {
  const c = String(canonicalCarrier_(carrier));
  if (c === 'gls') return TRK_trackGLS(code);
  if (c === 'posti') return TRK_trackPosti(code);
  if (c === 'dhl') return TRK_trackDHL(code);
  if (c === 'bring') return TRK_trackBring(code);
  if (c === 'matkahuolto') return TRK_trackMH(code);
  return { carrier: carrier, status:'UNKNOWN_CARRIER' };
}

// Power BI functions
function powerBiImportAndFilter() {
  SpreadsheetApp.getUi().alert('Tämä toiminto ei ole käytössä tässä versiossa.');
}
function powerBiArchiveNew() {
  const ss = SpreadsheetApp.getActive();
  const shNew = ss.getSheetByName(PBI_NEW_SHEET);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shNew || !shArch) {
    SpreadsheetApp.getUi().alert('PBI_New tai Packages_Archive -taulua ei löydy.');
    return;
  }
  const data = shNew.getDataRange().getValues();
  if (data.length > 1) {
    for (let r=1; r<data.length; r++) shArch.appendRow(data[r]);
  }
  shNew.clearContents();
  SpreadsheetApp.getUi().alert('PBI-uudet toimitukset arkistoitu.');
}
function triggerPowerBI() {
  const url = PropertiesService.getScriptProperties().getProperty('PBI_WEBHOOK_URL');
  if (!url) {
    SpreadsheetApp.getUi().alert('Power BI -webhook URL puuttuu asetuksista (PBI_WEBHOOK_URL)');
    return;
  }
  try {
    UrlFetchApp.fetch(url, { muteHttpExceptions:true });
    SpreadsheetApp.getUi().alert('Power BI -päivitys kutsuttu.');
  } catch(err) {
    SpreadsheetApp.getUi().alert('Virhe Power BI -webhook-kutsussa: ' + err);
  }
}

// Script property seeds
function seedOptionalDefaults() {
  SpreadsheetApp.getUi().alert('Ei valinnaisia oletuksia asetettavaksi.');
}
function seedKnownAccountsAndKeys() {
  const sp = PropertiesService.getScriptProperties();
  const setIfEmpty = (k,v) => { if (!sp.getProperty(k)) sp.setProperty(k, v); };
  setIfEmpty('GLS_FI_API_KEY','<GLS_FI_API_KEY>');
  setIfEmpty('GLS_FI_TRACK_URL','https://<gls-fi-host>/customerapi/v2/get_tracking_events');
  setIfEmpty('GLS_FI_SENDER_ID','<GLS_SENDER_ID>');
  setIfEmpty('GLS_FI_SENDER_IDS','');
  setIfEmpty('GLS_FI_METHOD','POST');
  setIfEmpty('GLS_TOKEN_URL','https://api.gls-group.net/oauth2/v1/token');
  setIfEmpty('GLS_BASIC','<base64_clientId:secret>');
  setIfEmpty('GLS_TRACK_URL','https://api.gls-group.net/track-and-trace-v1/tracking/simple/references/{{code}}');
  setIfEmpty('POSTI_TOKEN_URL','https://oauth2.posti.com/oauth/token');
  setIfEmpty('POSTI_BASIC','<base64_clientId:secret>');
  setIfEmpty('POSTI_TRACK_URL','https://api.posti.fi/tracking/7/shipments/trackingnumbers/{{code}}');
  setIfEmpty('POSTI_TRK_URL','https://atlas.posti.fi/track-shipment-json?ShipmentId={{code}}');
  setIfEmpty('POSTI_TRK_BASIC','<base64_user:pass>');
  setIfEmpty('POSTI_TRK_USER','<ma_contract_user>');
  setIfEmpty('POSTI_TRK_PASS','<ma_contract_pass>');
  setIfEmpty('MH_TRACK_URL','https://extservices.matkahuolto.fi/mpaketti/public/tracking?ids={{code}}');
  setIfEmpty('MH_BASIC','<base64_user:pass>');
  setIfEmpty('DHL_TRACK_URL','https://api-eu.dhl.com/track/shipments?trackingNumber={{code}}');
  setIfEmpty('DHL_API_KEY','<DHL_API_KEY>');
  setIfEmpty('BRING_TRACK_URL','https://api.bring.com/tracking/api/v2/tracking.json?q={{code}}');
  setIfEmpty('BRING_UID','<BRING_API_UID>');
  setIfEmpty('BRING_KEY','<BRING_API_KEY>');
  setIfEmpty('BRING_CLIENT_URL','https://<yourapp>.domain.com');
  setIfEmpty('BULK_BACKOFF_MINUTES_BASE','30');
  setIfEmpty('PBI_WEBHOOK_URL','https://<oma-webhook-url>');
  SpreadsheetApp.getUi().alert('Tunnetut oletusavaimet asetettu, tarkista Script Properties.');
}

// Old control panel stub to maintain compatibility
function openControlPanel() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html><meta charset="utf-8"><title>Shipment – Asetuspaneeli</title><style>body{font:14px system-ui;margin:16px;color:#111}h1{font-size:18px;margin:0 0 10px}.row{display:grid;grid-template-columns:170px 1fr;gap:8px;align-items:center;margin:6px 0}input{width:100%}button{padding:6px 12px;margin:8px 8px 0 0}.info{font-size:12px;color:#555}</style><body><h1>Asetuspaneeli</h1><div class="row"><span>MH_BASIC</span><input id="mhb" type="text"></div><div class="row"><span>MH_TRACK_URL</span><input id="mhu" type="text"></div><div class="row"><span>POSTI_BASIC</span><input id="pob" type="text"></div><div class="row"><span>POSTI_TOKEN_URL</span><input id="pot" type="text"></div><div class="row"><span>POSTI_TRACK_URL</span><input id="pou" type="text"></div><div class="row"><span>POSTI_TRK_USER</span><input id="ptu" type="text"></div><div class="row"><span>POSTI_TRK_PASS</span><input id="ptp" type="password"></div><div class="row"><span>POSTI_TRK_URL</span><input id="ptr" type="text"></div><div class="row"><span>POSTI_TRK_BASIC</span><input id="ptb" type="text"></div><div class="row"><span>GLS_FI_API_KEY</span><input id="gak" type="text"></div><div class="row"><span>GLS_FI_TRACK_URL</span><input id="gfu" type="text"></div><div class="row"><span>GLS_FI_SENDER_ID</span><input id="gsi" type="text"></div><div class="row"><span>GLS_FI_METHOD</span><input id="gfm" type="text"></div><div class="row"><span>GLS_TOKEN_URL</span><input id="gtu" type="text"></div><div class="row"><span>GLS_BASIC</span><input id="gba" type="text"></div><div class="row"><span>GLS_TRACK_URL</span><input id="gtr" type="text"></div><div class="row"><span>DHL_TRACK_URL</span><input id="dtu" type="text"></div><div class="row"><span>DHL_API_KEY</span><input id="dak" type="text"></div><div class="row"><span>BRING_TRACK_URL</span><input id="btu" type="text"></div><div class="row"><span>BRING_UID</span><input id="bui" type="text"></div><div class="row"><span>BRING_KEY</span><input id="bk" type="text"></div><div class="row"><span>BRING_CLIENT_URL</span><input id="bcu" type="text"></div><div class="row"><span>BULK_BACKOFF_MINUTES_BASE</span><input id="bbm" type="text" placeholder="esim. 5"></div><p class="info">Jätä kenttä tyhjäksi poistaaksesi asetuksen.</p><button onclick="seed()">Täytä oletukset</button><button onclick="flush()">Tyhjennä välimuisti</button><script>function seed(){google.script.run.seedKnownAccountsAndKeys();}function flush(){google.script.run.trkClearCache();}</script></body>`)
    .setWidth(640).setHeight(480).setTitle('Asetuspaneeli');
  SpreadsheetApp.getUi().showModalDialog(html, 'Asetuspaneeli');
}