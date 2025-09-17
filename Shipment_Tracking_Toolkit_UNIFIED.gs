// Shipment Tracking Toolkit – Unified & Patched (2025-09-15)
// Author: ChatGPT (for Riku / Ralli Logistics)
// Notes:
// - Includes: Gmail imports (CSV/XLSX), Pending builder, Status refresh (Posti, GLS, DHL, Bring, Matkahuolto),
//   bulk updater with cooldown + rate limiting auto-tune, archive handling, triggers, error log, credentials panel.
// - Safe header detection, safe archiving after writes, robust XLSX conversion via Drive Advanced (v2/v3) or CSV fallback.
// - Added cache-buster for GETs, label-configurable Gmail query, weekday-safe daily schedule, and sanity check menus.

/* ========================= Configuration ========================= */
const ACTION_SHEET = 'Vaatii_toimenpiteitä';
const TARGET_SHEET = 'Packages';
const ARCHIVE_SHEET = 'Packages_Archive';
const PBI_IMPORT_SHEET = 'PowerBI_Import';
const PBI_NEW_SHEET = 'PowerBI_New';

// Bulk refresh parameters
const BULK_MAX_API_CALLS_PER_RUN = 300;
const BULK_BACKOFF_MINUTES_BASE = 30;
const BULK_BACKOFF_MINUTES_MAX = 24 * 60;

// SLA examples (not enforced in this version but reserved)
const SLA_OPEN_DAYS = 5;
const SLA_PD_DAYS = 5;

/* ========================= Properties Helpers ========================= */
function TRK_props_(k) {
  return PropertiesService.getScriptProperties().getProperty(k) || '';
}
function getCfgInt_(key, fallback) {
  const sp = PropertiesService.getScriptProperties();
  const v = sp.getProperty(key);
  return v ? Number(v) : fallback;
}

/* ========================= Normalization & Headers ========================= */
function normalize_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/[^\p{L}\p{N}]+/gu,' ');
}
function headerIndexMap_(hdr) {
  const m = {};
  (hdr || []).forEach((h,i) => {
    try {
      const n = normalize_(h || '');
      m[h] = i;        // original
      m[n] = i;        // normalized
    } catch (e) {
      m[h] = i;
      m[String(h).toLowerCase()] = i;
    }
  });
  return m;
}
function colIndexOf_(hdrMap, candidates) {
  for (const label of candidates) {
    if (label in hdrMap) return hdrMap[label];
    const n = normalize_(label);
    if (n in hdrMap) return hdrMap[n];
  }
  return -1;
}
const CARRIER_CANDIDATES = [
  'RefreshCarrier','Carrier','Carrier name','LogisticsProvider','Shipper','Service provider','Forwarder','Transporter'
];
const TRACKING_CANDIDATES = [
  'Tracking number','TrackingNumber','Barcode','Waybill','Waybill No','AWB','Shipment ID','Consignment number'
];
const KEY_CANDIDATES = [
  'Package Number','PackageNumber','Tracking number','TrackingNumber','Consignment number','Shipment ID','Outbound order','Orderid','Order id'
];
function chooseKeyIndex_(headers) {
  const m = headerIndexMap_(headers);
  const idx = colIndexOf_(m, KEY_CANDIDATES);
  return idx >= 0 ? idx : 0;
}
function mergeHeaders_(a,b) {
  return Array.from(new Set([...(a || []), ...(b || [])]));
}
function firstCode_(cell) {
  return String(cell || '').split(/[\n,;]/)[0].trim();
}

/* ========================= Dates & Formatting ========================= */
function fmtDateTime_(d) {
  if (!d) return '';
  try {
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  } catch (e) {
    try { return new Date(d).toISOString(); } catch (err) { return String(d); }
  }
}

/* ========================= Rate Limiting & Caching ========================= */
function trkRateLimitWait_(carrier) {
  const sp = PropertiesService.getScriptProperties();
  const C  = String(carrier||'').toUpperCase();

  // Cooldown
  const pauseUntil = parseInt(sp.getProperty('PAUSE_UNTIL_'+C) || '0', 10);
  if (pauseUntil && Date.now() < pauseUntil) {
    Utilities.sleep(Math.min(pauseUntil - Date.now(), 30000));
  }

  // Per-carrier min interval (default 700ms)
  const minMs = parseInt(sp.getProperty('RATE_MINMS_'+C), 10) || 700;
  const key   = 'RATE_LAST_'+C;
  const last  = parseInt(sp.getProperty(key)||'0',10) || 0;
  const now   = Date.now();
  const wait  = Math.max(0, (last + minMs) - now);
  if (wait > 0) Utilities.sleep(Math.min(wait, 5000));
  sp.setProperty(key, String(Date.now()));
}
function autoTuneRateLimitOn429_(carrier, addMs, capMs) {
  const sp = PropertiesService.getScriptProperties();
  const key = `RATE_MINMS_${String(carrier||'').toUpperCase()}`;
  const current = Number(sp.getProperty(key) || '0');
  const updated = Math.min(current + addMs, capMs || current + addMs);
  sp.setProperty(key, String(updated));

  const pauseKey = 'PAUSE_UNTIL_' + String(carrier||'').toUpperCase();
  const baseMin = getCfgInt_('BULK_BACKOFF_MINUTES_BASE', BULK_BACKOFF_MINUTES_BASE);
  const pauseMs = Math.min(baseMin, BULK_BACKOFF_MINUTES_MAX) * 60000;
  sp.setProperty(pauseKey, String(Date.now() + pauseMs));
}
function trkClearCache(){
  PropertiesService.getScriptProperties().setProperty('TRK_CACHE_BUSTER', String(Date.now()));
  SpreadsheetApp.getUi().alert('Seurannan välimuisti nollattu (cache-buster päivitetty).');
}
function withBust_(url){
  const b = PropertiesService.getScriptProperties().getProperty('TRK_CACHE_BUSTER');
  return b ? url + (url.indexOf('?')>=0?'&':'?') + '_=' + encodeURIComponent(b) : url;
}

/* ========================= UI: Menu ========================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipment')
    .addItem('Credentials Hub', 'showCredentialsHub')
    .addItem('Hae Gmail → Päivitä taulut nyt', 'runDailyFlowOnce')
    .addItem('Rakenna Vaatii_toimenpiteitä (ei-toimitetut)', 'buildPendingFromPackagesAndArchive')
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
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
        .addItem('Näytä edistyminen', 'showProgressSidebar')
    )
    .addSeparator()
    .addItem('Gmail: tuo KAIKKI (vanhin → uusin)', 'fetchHistoryFromGmailOldToNew')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Power BI')
        .addItem('Arkistoi PBI-uudet → Packages_Archive', 'powerBiArchiveNew')
        .addItem('Käynnistä Power BI -päivitys (webhook)', 'triggerPowerBI')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Tarkistimet')
        .addItem('Puuttuvat sarakeotsikot (aktiivinen välilehti)', 'showMissingColumns_Active')
        .addItem('Integraatioavaimet (Script Properties)', 'showMissingProperties')
        .addItem('Avaa ErrorLog', 'openErrorLog')
        .addItem('Tyhjennä ErrorLog', 'clearErrorLog')
    )
    .addSeparator()
    .addItem('Ajasta: arkipäivät klo 12:00', 'setupWeekday1200_DailyFlow')
    .addItem('Ajasta: viikkoraportti maanantai 02:00', 'setupWeeklyMon0200')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Pikatoiminnot / Vaatii')
        .addItem('Vahvista toimitukset (Vaatii)', 'confirmMissingDeliveredTimes')
        .addItem('Vahvista + Arkistoi (Vaatii)', 'confirmAndArchiveDelivered')
    )
    .addSeparator()
    .addItem('Tyhjennä seurannan välimuisti', 'trkClearCache')
    .addItem('Poista kaikki ajastukset', 'clearAllTriggers')
    .addToUi();
}
function onInstall() { onOpen(); }

/* ========================= Credentials Panel ========================= */
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
    'PBI_WEBHOOK_URL',
    'GMAIL_QUERY'
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
function showCredentialsHub() {
  try {
    openCredsPanel(); // try external template "Creds"
  } catch (e) {
    openControlPanel(); // fallback inline
  }
}
function openCredsPanel() {
  const t = HtmlService.createTemplateFromFile('Creds');
  t.props = credGetProps();
  const html = t.evaluate()
    .setTitle('Integraatioavaimet')
    .setWidth(760)
    .setHeight(760);
  SpreadsheetApp.getUi().showSidebar(html);
}
function openControlPanel() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html><meta charset="utf-8"><title>Shipment – Asetuspaneeli</title>
<style>
body{font:14px system-ui;margin:16px;color:#111}
h1{font-size:18px;margin:0 0 10px}
.row{display:grid;grid-template-columns:190px 1fr;gap:8px;align-items:center;margin:6px 0}
input{width:100%}button{padding:6px 12px;margin:8px 8px 0 0}
.info{font-size:12px;color:#555}
</style>
<body>
  <h1>Asetuspaneeli</h1>
  <div class="row"><span>GMAIL_QUERY</span><input id="gq" type="text" placeholder='esim: label:ShipmentReports OR subject:(Outbound)'></div>
  <div class="row"><span>MH_BASIC</span><input id="mhb" type="text"></div>
  <div class="row"><span>MH_TRACK_URL</span><input id="mhu" type="text"></div>
  <div class="row"><span>POSTI_BASIC</span><input id="pob" type="text"></div>
  <div class="row"><span>POSTI_TOKEN_URL</span><input id="pot" type="text"></div>
  <div class="row"><span>POSTI_TRACK_URL</span><input id="pou" type="text"></div>
  <div class="row"><span>POSTI_TRK_USER</span><input id="ptu" type="text"></div>
  <div class="row"><span>POSTI_TRK_PASS</span><input id="ptp" type="password"></div>
  <div class="row"><span>POSTI_TRK_URL</span><input id="ptr" type="text"></div>
  <div class="row"><span>POSTI_TRK_BASIC</span><input id="ptb" type="text"></div>
  <div class="row"><span>GLS_FI_API_KEY</span><input id="gak" type="text"></div>
  <div class="row"><span>GLS_FI_TRACK_URL</span><input id="gfu" type="text"></div>
  <div class="row"><span>GLS_FI_SENDER_ID</span><input id="gsi" type="text"></div>
  <div class="row"><span>GLS_FI_SENDER_IDS</span><input id="gsl" type="text"></div>
  <div class="row"><span>GLS_FI_METHOD</span><input id="gfm" type="text" placeholder="POST tai GET"></div>
  <div class="row"><span>GLS_TOKEN_URL</span><input id="gtu" type="text"></div>
  <div class="row"><span>GLS_BASIC</span><input id="gba" type="text"></div>
  <div class="row"><span>GLS_TRACK_URL</span><input id="gtr" type="text"></div>
  <div class="row"><span>DHL_TRACK_URL</span><input id="dtu" type="text"></div>
  <div class="row"><span>DHL_API_KEY</span><input id="dak" type="text"></div>
  <div class="row"><span>BRING_TRACK_URL</span><input id="btu" type="text"></div>
  <div class="row"><span>BRING_UID</span><input id="bui" type="text"></div>
  <div class="row"><span>BRING_KEY</span><input id="bk" type="text"></div>
  <div class="row"><span>BRING_CLIENT_URL</span><input id="bcu" type="text"></div>
  <div class="row"><span>BULK_BACKOFF_MINUTES_BASE</span><input id="bbm" type="text" placeholder="esim. 30"></div>
  <p class="info">Jätä kenttä tyhjäksi poistaaksesi asetuksen.</p>
  <button onclick="seed()">Täytä oletukset</button>
  <button onclick="flush()">Tyhjennä välimuisti</button>
  <script>
    google.script.run.withSuccessHandler(props => {
      document.getElementById('gq').value = props.GMAIL_QUERY || '';
      document.getElementById('mhb').value = props.MH_BASIC || '';
      document.getElementById('mhu').value = props.MH_TRACK_URL || '';
      document.getElementById('pob').value = props.POSTI_BASIC || '';
      document.getElementById('pot').value = props.POSTI_TOKEN_URL || '';
      document.getElementById('pou').value = props.POSTI_TRACK_URL || '';
      document.getElementById('ptu').value = props.POSTI_TRK_USER || '';
      document.getElementById('ptp').value = props.POSTI_TRK_PASS || '';
      document.getElementById('ptr').value = props.POSTI_TRK_URL || '';
      document.getElementById('ptb').value = props.POSTI_TRK_BASIC || '';
      document.getElementById('gak').value = props.GLS_FI_API_KEY || '';
      document.getElementById('gfu').value = props.GLS_FI_TRACK_URL || '';
      document.getElementById('gsi').value = props.GLS_FI_SENDER_ID || '';
      document.getElementById('gsl').value = props.GLS_FI_SENDER_IDS || '';
      document.getElementById('gfm').value = props.GLS_FI_METHOD || '';
      document.getElementById('gtu').value = props.GLS_TOKEN_URL || '';
      document.getElementById('gba').value = props.GLS_BASIC || '';
      document.getElementById('gtr').value = props.GLS_TRACK_URL || '';
      document.getElementById('dtu').value = props.DHL_TRACK_URL || '';
      document.getElementById('dak').value = props.DHL_API_KEY || '';
      document.getElementById('btu').value = props.BRING_TRACK_URL || '';
      document.getElementById('bui').value = props.BRING_UID || '';
      document.getElementById('bk').value = props.BRING_KEY || '';
      document.getElementById('bcu').value = props.BRING_CLIENT_URL || '';
      document.getElementById('bbm').value = props.BULK_BACKOFF_MINUTES_BASE || '';
    }).credGetProps();

    function collect() {
      return {
        GMAIL_QUERY: document.getElementById('gq').value,
        MH_BASIC: document.getElementById('mhb').value,
        MH_TRACK_URL: document.getElementById('mhu').value,
        POSTI_BASIC: document.getElementById('pob').value,
        POSTI_TOKEN_URL: document.getElementById('pot').value,
        POSTI_TRACK_URL: document.getElementById('pou').value,
        POSTI_TRK_USER: document.getElementById('ptu').value,
        POSTI_TRK_PASS: document.getElementById('ptp').value,
        POSTI_TRK_URL: document.getElementById('ptr').value,
        POSTI_TRK_BASIC: document.getElementById('ptb').value,
        GLS_FI_API_KEY: document.getElementById('gak').value,
        GLS_FI_TRACK_URL: document.getElementById('gfu').value,
        GLS_FI_SENDER_ID: document.getElementById('gsi').value,
        GLS_FI_SENDER_IDS: document.getElementById('gsl').value,
        GLS_FI_METHOD: document.getElementById('gfm').value,
        GLS_TOKEN_URL: document.getElementById('gtu').value,
        GLS_BASIC: document.getElementById('gba').value,
        GLS_TRACK_URL: document.getElementById('gtr').value,
        DHL_TRACK_URL: document.getElementById('dtu').value,
        DHL_API_KEY: document.getElementById('dak').value,
        BRING_TRACK_URL: document.getElementById('btu').value,
        BRING_UID: document.getElementById('bui').value,
        BRING_KEY: document.getElementById('bk').value,
        BRING_CLIENT_URL: document.getElementById('bcu').value,
        BULK_BACKOFF_MINUTES_BASE: document.getElementById('bbm').value
      };
    }
    function seed(){ google.script.run.seedKnownAccountsAndKeys(); }
    function flush(){ google.script.run.trkClearCache(); }
    window.onbeforeunload = () => google.script.run.credSaveProps(collect());
  </script>
</body>`).setWidth(720).setHeight(560).setTitle('Asetuspaneeli');
  SpreadsheetApp.getUi().showModalDialog(html, 'Asetuspaneeli');
}

/* ========================= Error logging & checks ========================= */
function logError_(context, error) {
  try {
    const ss = SpreadsheetApp.getActive();
    const shLog = ss.getSheetByName('ErrorLog') || ss.insertSheet('ErrorLog');
    shLog.appendRow([new Date(), context, String(error)]);
    if (shLog.getLastRow() > 2000) shLog.deleteRows(1, shLog.getLastRow() - 2000);
  } catch(e) {}
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
function showMissingColumns_Active(){
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh) return;
  const hdr = sh.getDataRange().getValues()[0] || [];
  const m = headerIndexMap_(hdr);
  const need = ['Carrier','Tracking number','RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw'];
  const missing = need.filter(n => !(n in m) && !(normalize_(n) in m));
  SpreadsheetApp.getUi().alert(missing.length ? ('Puuttuvat: ' + missing.join(', ')) : 'Kaikki OK (vaaditut sarakkeet löytyivät).');
}
function showMissingProperties(){
  const props = credGetProps();
  const must = ['POSTI_TRACK_URL','GLS_FI_TRACK_URL','DHL_TRACK_URL','BRING_TRACK_URL','MH_TRACK_URL','GMAIL_QUERY'];
  const missing = must.filter(k => !(props[k]||'').trim());
  SpreadsheetApp.getUi().alert(missing.length ? ('Puuttuvat asetukset: ' + missing.join(', ')) : 'Asetukset näyttävät hyviltä (ydinavaimet löytyvät).');
}

/* ========================= Gmail & Attachments ========================= */
function sanitizeMatrix_(matrix) {
  if (!matrix || !matrix.length) return matrix;
  let lastNonEmpty = matrix.length;
  for (let r = matrix.length - 1; r >= 0; r--) {
    if (!matrix[r].join('')) lastNonEmpty = r;
    else break;
  }
  return matrix.slice(0, lastNonEmpty);
}
function readAttachmentToValues_(blob, filename) {
  if (!blob) return null;
  filename = filename || blob.getName();
  try {
    // Preferred: Drive Advanced v2 insert with convert:true
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function') {
      const file = Drive.Files.insert({ title: filename.replace(/\.(xlsx|csv)$/i,''), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true });
      const ss = SpreadsheetApp.openById(file.id);
      const values = ss.getSheets()[0].getDataRange().getDisplayValues();
      DriveApp.getFileById(file.id).setTrashed(true);
      return values;
    }
    // Fallback: Upload then copy/convert
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.copy === 'function') {
      const up = DriveApp.createFile(blob);
      const copied = Drive.Files.copy({ title: filename.replace(/\.(xlsx|csv)$/i,''), mimeType: MimeType.GOOGLE_SHEETS }, up.getId());
      const ss = SpreadsheetApp.openById(copied.id);
      const values = ss.getSheets()[0].getDataRange().getDisplayValues();
      DriveApp.getFileById(up.getId()).setTrashed(true);
      DriveApp.getFileById(copied.id).setTrashed(true);
      return values;
    }
    // CSV fallback
    return Utilities.parseCsv(blob.getDataAsString('UTF-8'));
  } catch (e) {
    SpreadsheetApp.getUi().alert('Virhe liitteen lukemisessa: ' + e);
    return null;
  }
}
function gmailQuery_() {
  return TRK_props_('GMAIL_QUERY') || 'subject:(Outbound) OR filename:(Outbound)';
}
function findLatestAttachment_() {
  const threads = GmailApp.search(gmailQuery_());
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
function writeMerged_(sheet, existing, incoming){
  // Merge two matrices by headers
  const old = existing.slice();
  const newM = incoming.slice();
  const hdrOld = old.shift() || [];
  const hdrNew = newM.shift() || [];
  const mergedHdr = mergeHeaders_(hdrOld, hdrNew);
  const mapOld = headerIndexMap_(hdrOld);
  const mapNew = headerIndexMap_(hdrNew);
  const out = [mergedHdr];
  old.forEach(row => {
    const nr = mergedHdr.map(col => (col in mapOld ? row[mapOld[col]] : ''));
    out.push(nr);
  });
  newM.forEach(row => {
    const nr = mergedHdr.map(col => (col in mapNew ? row[mapNew[col]] : ''));
    out.push(nr);
  });
  sheet.clear();
  sheet.getRange(1,1,out.length,out[0].length).setValues(out);
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
  if (sh.getLastRow() < 2) {
    sh.clear();
    sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
  } else {
    // merge-in
    const existing = sh.getDataRange().getValues();
    writeMerged_(sh, existing, matrix);
  }
}
function fetchHistoryFromGmailOldToNew() {
  const threads = GmailApp.search(gmailQuery_());
  if (!threads.length) {
    SpreadsheetApp.getUi().alert('Ei löytynyt viestejä annetulla Gmail-haulla.');
    return;
  }
  const ss = SpreadsheetApp.getActive();
  let sheetPack = ss.getSheetByName(TARGET_SHEET) || ss.insertSheet(TARGET_SHEET);
  let sheetArch = ss.getSheetByName(ARCHIVE_SHEET) || ss.insertSheet(ARCHIVE_SHEET);
  sheetPack.clear(); sheetArch.clear();
  let firstPack = true, firstArch = true;
  for (let ti = 0; ti < threads.length; ti++) {
    const msgs = GmailApp.getMessagesForThread(threads[ti]);
    for (let mi = 0; mi < msgs.length; mi++) {
      const atts = msgs[mi].getAttachments();
      for (const att of atts) {
        const name = att.getName();
        if (/Outbound.*\.(xlsx|csv)/i.test(name)) {
          const values = readAttachmentToValues_(att, name);
          if (!values) continue;
          const matrix = sanitizeMatrix_(values);
          if (!matrix || matrix.length < 2) continue;
          const isArchive = name.toLowerCase().includes('archive');
          const sh = isArchive ? sheetArch : sheetPack;
          if ((isArchive && firstArch) || (!isArchive && firstPack)) {
            sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
            if (isArchive) firstArch = false; else firstPack = false;
          } else {
            const existing = sh.getDataRange().getValues();
            writeMerged_(sh, existing, matrix);
          }
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('Gmail-historia tuotu ja yhdistetty (Packages & Packages_Archive).');
}

/* ========================= Pending Builder ========================= */
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
  let shPending = ss.getSheetByName(ACTION_SHEET) || ss.insertSheet(ACTION_SHEET);
  shPending.clear();
  shPending.getRange(1,1,pending.length,pending[0].length).setValues(pending);
}

/* ========================= Status Refresh & Archiving ========================= */
function refreshStatuses_Vaatii() {
  refreshStatuses_Sheet(ACTION_SHEET, true);
}
function refreshStatuses_Sheet(sheetName, removeDelivered) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  // ensure Refresh columns
  const hdr = data[0].slice();
  const extra = ['RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw'];
  let writeHdr = false;
  extra.forEach(c => { if (!hdr.includes(c)) { hdr.push(c); writeHdr = true; } });
  if (writeHdr) sh.getRange(1,1,1,hdr.length).setValues([hdr]);

  const col = headerIndexMap_(hdr);
  const carrierI  = colIndexOf_(col, CARRIER_CANDIDATES);
  const trackI    = colIndexOf_(col, TRACKING_CANDIDATES);
  const rCarrierI = col['RefreshCarrier'];
  const rStatusI  = col['RefreshStatus'];
  const rTimeI    = col['RefreshTime'];
  const rLocI     = col['RefreshLocation'];
  const rRawI     = col['RefreshRaw'];

  const now = Date.now();
  const cutoffMs = 6 * 3600 * 1000;
  const writeRows = [];
  const toArchive = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.join('')) continue;

    const carrier = String(row[carrierI] || row[rCarrierI] || '').toLowerCase();
    const code    = firstCode_(row[trackI] || '');
    const last    = row[rTimeI] ? new Date(row[rTimeI]).getTime() : 0;
    const expired = !last || (now - last > cutoffMs);

    let res = null;
    if (code && carrier && expired) {
      res = TRK_trackByCarrier_(carrier, code);
    }

    const values = [
      (res && res.carrier) || (carrier || ''),
      (res && res.status)  || row[rStatusI] || '',
      (res && res.time)    || row[rTimeI]  || '',
      (res && res.location)|| row[rLocI]   || '',
      (res && res.raw)     || row[rRawI]   || ''
    ];
    const delivered = !!(res && String(res.status).toLowerCase().match(/delivered|toimitettu|luovutettu/));

    writeRows.push({ idx: i, values, delivered, rowSnapshot: row.slice() });

    if (removeDelivered && delivered) {
      toArchive.push({ idx: i, row: row.slice() });
    }
  }

  // Write updates
  const firstCol = hdr.indexOf('RefreshCarrier') + 1;
  writeRows.forEach(w => {
    try { sh.getRange(w.idx + 1, firstCol, 1, 5).setValues([w.values]); } catch(e){ logError_('refreshStatuses_Sheet setValues', e); }
  });

  // Archive after writes (bottom-up)
  if (removeDelivered && toArchive.length) {
    const shArch = ss.getSheetByName(ARCHIVE_SHEET) || ss.insertSheet(ARCHIVE_SHEET);
    toArchive.sort((a,b) => b.idx - a.idx).forEach(item => {
      try {
        shArch.appendRow(item.row);
        sh.deleteRow(item.idx + 1);
      } catch(e){ logError_('refreshStatuses_Sheet archive', e); }
    });
  }
}

/* ========================= Quick Actions ========================= */
function confirmMissingDeliveredTimes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ACTION_SHEET);
  if (!sh) return;
  const dataRange = sh.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) return;

  const hdr = (data[0] || []).map(h => normalize_(h));
  const statusCol = hdr.findIndex(h => h.includes('status'));
  const dateCol = hdr.findIndex(h => h.includes('delivered') || h.includes('received') || h.includes('toimitettu'));
  if (statusCol === -1 || dateCol === -1) {
    SpreadsheetApp.getUi().alert('Ei löytynyt "Status" tai "Delivered at" -saraketta.');
    return;
  }

  const now = new Date();
  let filled = 0;
  for (let r = 1; r < data.length; r++) {
    const status = String(data[r][statusCol] || '').toLowerCase();
    const dateVal = data[r][dateCol];
    if (status.match(/delivered|toimitettu|luovutettu/) && !dateVal) {
      try { sh.getRange(r+1, dateCol+1).setValue(now); filled++; } catch(e){ logError_('confirmMissingDeliveredTimes setValue', e); }
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
    if (status.match(/delivered|toimitettu|luovutettu/)) toMove.push({idx:r, row:data[r]});
  }
  toMove.sort((a,b)=>b.idx-a.idx).forEach(item => {
    try {
      shArch.appendRow(item.row);
      shVaatii.deleteRow(item.idx+1);
    } catch(e){ logError_('confirmAndArchiveDelivered', e); }
  });
  SpreadsheetApp.getUi().alert(`Arkistoitu ${toMove.length} toimitettua riviä.`);
}

/* ========================= Scheduling & Orchestration ========================= */
function runDailyFlowOnce(){
  // Skip weekends for the daily run
  const day = (new Date()).getDay(); // 0=Sun,6=Sat
  if (day === 0 || day === 6) return;
  try { fetchAndRebuild(); } catch(e){ logError_('runDailyFlowOnce.fetchAndRebuild', e); }
  try { buildPendingFromPackagesAndArchive(); } catch(e){ logError_('runDailyFlowOnce.buildPending', e); }
  try { refreshStatuses_Vaatii(); } catch(e){ logError_('runDailyFlowOnce.refreshVaatii', e); }
}
function setupWeekday1200_DailyFlow() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'runDailyFlowOnce') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('runDailyFlowOnce').timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: päivittäinen ajo klo 12:00 (weekdays only, skripti ohittaa viikonloput).');
}
function setupWeeklyMon0200() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'makeWeeklyReportsSunSun') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('makeWeeklyReportsSunSun').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: viikkoraportti maanantaisin 02:00 (stub).');
}
function clearAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(tr => ScriptApp.deleteTrigger(tr));
  SpreadsheetApp.getUi().alert('Kaikki ajastukset poistettu.');
}
function showProgressSidebar(){
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const jobs = Object.keys(props).filter(k => k.indexOf('BULKJOB|') === 0).map(k => JSON.parse(props[k]));
  const html = HtmlService.createHtmlOutput(
    '<h3 style="font-family:system-ui;margin:10px 0">Bulk-tila</h3>' +
    (jobs.length ? jobs.map(j => {
      const done = Math.max(0, Math.min(j.row-2, j.total));
      const pct = j.total ? Math.round((done / j.total) * 100) : 0;
      return `<div style="font:13px system-ui;margin:6px 0"><b>${j.sheet}</b> — ${done}/${j.total} (${pct}%)</div>`;
    }).join('') : '<div style="font:13px system-ui">Ei käynnissä olevia ajoja.</div>')
  ).setTitle('Bulk-tila').setWidth(320).setHeight(180);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ========================= Bulk Processing ========================= */
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
  const sp = PropertiesService.getScriptProperties();
  Object.keys(sp.getProperties()).forEach(k => { if (k.indexOf('BULKJOB|')===0) sp.deleteProperty(k); });
  SpreadsheetApp.getUi().alert('Bulk-ajo pysäytetty.');
}
function bulkTick() {
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const jobKeys = Object.keys(props).filter(k => k.indexOf('BULKJOB|') === 0);
  if (!jobKeys.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
    return;
  }

  const ss = SpreadsheetApp.getActive();
  function paused_(carrier){
    const until = parseInt(sp.getProperty('PAUSE_UNTIL_' + String(carrier||'').toUpperCase())||'0',10);
    return until && Date.now() < until;
  }

  jobKeys.forEach(jobKey => {
    try {
      const job = JSON.parse(sp.getProperty(jobKey));
      if (!job || !job.sheet) { sp.deleteProperty(jobKey); return; }
      const sh = ss.getSheetByName(job.sheet);
      if (!sh) { sp.deleteProperty(jobKey); return; }

      const data = sh.getDataRange().getValues();
      if (!data || data.length < 2) { sp.deleteProperty(jobKey); return; }

      const hdr = data[0];
      const colIndex = headerIndexMap_(hdr);
      const carrierI = colIndexOf_(colIndex, CARRIER_CANDIDATES);
      const trackI   = colIndexOf_(colIndex, TRACKING_CANDIDATES);
      const rCarrierI = colIndex['RefreshCarrier'];
      const rTimeI    = colIndex['RefreshTime'];

      const maxCalls = BULK_MAX_API_CALLS_PER_RUN;
      let callsThisTick = 0;

      for (let r = job.row || 2; r < data.length && callsThisTick < maxCalls; r++) {
        const row = data[r];
        if (!row.join('')) continue;
        const carrier = String((row[rCarrierI] || row[carrierI] || '')).toLowerCase();
        if (job.carrierFilter && carrier.indexOf(job.carrierFilter) === -1) continue;
        if (paused_(carrier)) continue;

        const code = firstCode_(row[trackI] || '');
        const lastTime = row[rTimeI] || '';
        const expired = !lastTime || (Date.now() - new Date(lastTime).getTime() > 6*3600*1000);
        if (!code || !carrier || !expired) continue;

        trkRateLimitWait_(carrier);
        const res = TRK_trackByCarrier_(carrier, code);

        const out = [
          res.carrier || carrier,
          res.status || '',
          res.time || fmtDateTime_(new Date()),
          res.location || '',
          res.raw || ''
        ];
        const firstCol = hdr.indexOf('RefreshCarrier') !== -1 ? hdr.indexOf('RefreshCarrier') + 1 : hdr.length + 1;
        try {
          sh.getRange(r + 1, firstCol, 1, out.length).setValues([out]);
        } catch(err) { logError_('bulkTick setValues', err); }

        job.calls = (job.calls || 0) + 1;
        job.row = r + 1;
        callsThisTick++;

        if (res.status === 'RATE_LIMIT_429') {
          autoTuneRateLimitOn429_(carrier, 200, 2000);
          break;
        }
      }

      if ((job.row && job.row >= data.length) || (job.calls && job.calls >= (job.total || data.length-1))) {
        sp.deleteProperty(jobKey);
      } else {
        sp.setProperty(jobKey, JSON.stringify(job));
      }
    } catch (e) {
      logError_('bulkTick processing ' + jobKey, e);
    }
  });

  const remaining = Object.keys(PropertiesService.getScriptProperties().getProperties()).filter(k => k.indexOf('BULKJOB|')===0);
  if (!remaining.length) {
    ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'bulkTick') ScriptApp.deleteTrigger(tr); });
  }
}

/* ========================= Carrier API Helpers ========================= */
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
function pickLatestEvent_(events) {
  if (!Array.isArray(events) || !events.length) return null;
  const ts = e => {
    const cand = e.eventDateTime || e.eventTime || e.timestamp || e.dateTime || e.dateIso || e.date || e.time || '';
    const d = new Date(cand);
    return isNaN(d) ? 0 : d.getTime();
  };
  return events.slice().sort((a,b) => ts(a)-ts(b))[events.length-1];
}

/* ========================= Carrier API Implementations ========================= */
function TRK_trackPosti(code) {
  const carrierName = 'Posti';
  trkRateLimitWait_('POSTI');
  const tokenUrl = TRK_props_('POSTI_TOKEN_URL');
  const basic = TRK_props_('POSTI_BASIC');
  const trackUrlRaw = TRK_props_('POSTI_TRACK_URL');
  const trackUrl = trackUrlRaw ? withBust_(trackUrlRaw.replace('{{code}}', encodeURIComponent(code))) : '';

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
      if (http === 429) { autoTuneRateLimitOn429_('POSTI', 100, 1500); return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 }; }
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
  const fbUrlRaw = TRK_props_('POSTI_TRK_URL');
  const fbUrl = fbUrlRaw ? withBust_(fbUrlRaw.replace('{{code}}', encodeURIComponent(code))) : '';
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
    if (http === 429) { autoTuneRateLimitOn429_('POSTI', 100, 1500); return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 }; }
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

function TRK_trackGLS(code) {
  const carrierName = 'GLS';
  trkRateLimitWait_('GLS');
  // primary: FI Customer API v2 with x-api-key
  const baseRaw = TRK_props_('GLS_FI_TRACK_URL');
  const base = baseRaw || '';
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
      url = withBust_(url);
    }
    const res = TRK_safeFetch_(url, opt);
    const http = res.getResponseCode ? res.getResponseCode() : 0;
    let body=''; try { body = res.getContentText() || ''; } catch(e) {}
    if (http === 429) { autoTuneRateLimitOn429_('GLS', 200, 2000); return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 }; }
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
        // general fallback
        let ev = j.events || j.trackingEvents || j.parcel?.events || j.parcelShipments?.[0]?.events || j.parcels?.[0]?.events || j.items?.[0]?.events || j.eventList || j.tracking || [];
        ev = Array.isArray(ev) ? ev : [];
        const last = pickLatestEvent_(ev) || {};
        const status = last.eventDescription || last.description || last.eventShortName?.fi || last.eventShortName?.en || last.name || last.status || last.code || '';
        const time = last.eventTime || last.dateTime || last.timestamp || last.dateIso || last.date || last.time || '';
        const location = last.locationName || last.location || last.depot || last.city || '';
        if (status) return { carrier: carrierName, found:true, status, time, location, raw: body.slice(0,2000) };
      } catch(e) {}
    }
  }
  // GLS OAuth global fallback
  const tokenUrl = TRK_props_('GLS_TOKEN_URL');
  const basic = TRK_props_('GLS_BASIC');
  const trackUrlRaw = TRK_props_('GLS_TRACK_URL');
  const trackUrl = trackUrlRaw ? withBust_(trackUrlRaw.replace('{{code}}', encodeURIComponent(code))) : '';
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
    if (http === 429) { autoTuneRateLimitOn429_('GLS', 200, 2000); return { carrier: carrierName, status:'RATE_LIMIT_429', raw: body.slice(0,1000), retryAfter:15 }; }
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

function TRK_trackDHL(code) {
  const carrierName = 'DHL';
  trkRateLimitWait_('DHL');
  const urlRaw = TRK_props_('DHL_TRACK_URL');
  const url = urlRaw ? withBust_(urlRaw.replace('{{code}}', encodeURIComponent(code))) : '';
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

function TRK_trackBring(code) {
  const carrierName = 'Bring';
  trkRateLimitWait_('BRING');
  const urlRaw = TRK_props_('BRING_TRACK_URL');
  const url = urlRaw ? withBust_(urlRaw.replace('{{code}}', encodeURIComponent(code))) : '';
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

function TRK_trackMH(code) {
  trkRateLimitWait_('MATKAHUOLTO');
  const urlRaw = TRK_props_('MH_TRACK_URL');
  const url = urlRaw ? withBust_(urlRaw.replace('{{code}}', encodeURIComponent(code))) : '';
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

/* ========================= Dispatcher ========================= */
function TRK_trackByCarrier_(carrier, code) {
  const c = String((carrier||'')).toLowerCase();
  if (c.includes('gls')) return TRK_trackGLS(code);
  if (c.includes('posti')) return TRK_trackPosti(code);
  if (c.includes('dhl')) return TRK_trackDHL(code);
  if (c.includes('bring')) return TRK_trackBring(code);
  if (c.includes('matkahuolto') || c === 'mh') return TRK_trackMH(code);
  return { carrier: carrier, status:'UNKNOWN_CARRIER' };
}

/* ========================= Power BI ========================= */
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

/* ========================= Seeds ========================= */
function seedKnownAccountsAndKeys() {
  const sp = PropertiesService.getScriptProperties();
  const setIfEmpty = (k,v) => { if (!sp.getProperty(k)) sp.setProperty(k, v); };
  setIfEmpty('GMAIL_QUERY','subject:(Outbound) OR filename:(Outbound)');

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