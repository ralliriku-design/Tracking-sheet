// utils.gs — helpers, logging, headers, rate limiting, cache-buster, stats

/* ===== Config ===== */
const ACTION_SHEET = 'Vaatii_toimenpiteitä';
const TARGET_SHEET = 'Packages';
const ARCHIVE_SHEET = 'Packages_Archive';
const PBI_IMPORT_SHEET = 'PowerBI_Import';
const PBI_NEW_SHEET = 'PowerBI_New';

// Bulk parameters
const BULK_MAX_API_CALLS_PER_RUN = 300;
const BULK_BACKOFF_MINUTES_BASE = 30;
const BULK_BACKOFF_MINUTES_MAX = 24 * 60;

/* ===== Properties helpers ===== */
function TRK_props_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function getCfgInt_(key, fallback) {
  const sp = PropertiesService.getScriptProperties();
  const v = sp.getProperty(key);
  return v ? Number(v) : fallback;
}

/* ===== Normalization & headers ===== */
function normalize_(s) { return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/[^\p{L}\p{N}]+/gu,' '); }
function headerIndexMap_(hdr) {
  const m = {};
  (hdr || []).forEach((h,i) => {
    try { const n = normalize_(h || ''); m[h] = i; m[n] = i; }
    catch(e){ m[h] = i; m[String(h).toLowerCase()] = i; }
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
  'RefreshCarrier','Carrier','Carrier name','CarrierName','Delivery Carrier','Courier','Courier name',
  'LogisticsProvider','Logistics Provider','Shipper','Service provider','Forwarder','Forwarder name',
  'Transporter','Kuljetusliike','Kuljetusyhtiö','Kuljetus','Toimitustapa','Delivery method'
];
const TRACKING_CANDIDATES = [
  'Tracking number','TrackingNumber','Tracking','Tracking code','Tracking Code','Tracking No','Tracking ID',
  'Barcode','Bar code','Barcode ID','Waybill','Waybill No','AWB',
  'Parcel number','Parcel No','Parcel ID','ParcelID',
  'Shipment ID','ShipmentID','Shipment Number','Shipment no','ShipmentNo',
  'Consignment number','Consignment no','ConsignmentNo','Consignment ID','ConsignmentID',
  'Package number','Package Number','PackageNumber','Package no','Package No','PackageNo','Package ID','PackageID'
];
const KEY_CANDIDATES = [
  'Package Number','PackageNumber','Tracking number','TrackingNumber','Barcode','Waybill','Waybill No','AWB',
  'Shipment ID','ShipmentID','Shipment Number','Shipment no','ShipmentNo',
  'Consignment number','Consignment no','ConsignmentNo','ConsignmentID',
  'Parcel number','ParcelNo','Parcel ID','ParcelID','Collo ID','ColloID',
  'Orderid','Order id','Order number','OrderNumber','Reference','Customer Reference'
];
function chooseKeyIndex_(headers) {
  const m = headerIndexMap_(headers);
  const idx = colIndexOf_(m, KEY_CANDIDATES);
  return idx >= 0 ? idx : 0;
}
function mergeHeaders_(a,b) { return Array.from(new Set([...(a || []), ...(b || [])])); }
function firstCode_(cell) { return String(cell || '').split(/[\n,;]/)[0].trim(); }

/* ===== Dates & formatting ===== */
function fmtDateTime_(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"); }
  catch(e){ try { return new Date(d).toISOString(); } catch(err){ return String(d); } }
}

/* ===== Logging ===== */
function logError_(context, error) {
  try {
    const ss = SpreadsheetApp.getActive();
    const shLog = ss.getSheetByName('ErrorLog') || ss.insertSheet('ErrorLog');
    shLog.appendRow([new Date(), context, String(error)]);
    if (shLog.getLastRow() > 2000) shLog.deleteRows(1, shLog.getLastRow() - 2000);
  } catch(e) {}
}
function openErrorLog() { const ss = SpreadsheetApp.getActive(); const sh = ss.getSheetByName('ErrorLog'); if (sh) ss.setActiveSheet(sh); }
function clearErrorLog() { const ss = SpreadsheetApp.getActive(); const sh = ss.getSheetByName('ErrorLog'); if (sh) sh.clear(); }

/* ===== Rate limiting & cache ===== */
function trkRateLimitWait_(carrier) {
  const sp = PropertiesService.getScriptProperties();
  const C  = String(carrier||'').toUpperCase();
  const pauseUntil = parseInt(sp.getProperty('PAUSE_UNTIL_'+C) || '0', 10);
  if (pauseUntil && Date.now() < pauseUntil) { Utilities.sleep(Math.min(pauseUntil - Date.now(), 30000)); }
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
function trkClearCache(){ PropertiesService.getScriptProperties().setProperty('TRK_CACHE_BUSTER', String(Date.now())); SpreadsheetApp.getUi().alert('Cache-buster päivitetty.'); }
function withBust_(url){ const b = PropertiesService.getScriptProperties().getProperty('TRK_CACHE_BUSTER'); return b ? url + (url.indexOf('?')>=0?'&':'?') + '_=' + encodeURIComponent(b) : url; }

/* ===== Fetch helpers ===== */
function TRK_safeFetch_(url, opt) {
  try { return UrlFetchApp.fetch(url, opt); }
  catch(e) {
    return { getResponseCode: () => 0, getContentText: () => String(e), getAllHeaders: () => ({}) };
  }
}
function pickLatestEvent_(events) {
  if (!Array.isArray(events) || !events.length) return null;
  const ts = e => { const cand = e.eventDateTime || e.eventTime || e.timestamp || e.dateTime || e.dateIso || e.date || e.time || ''; const d = new Date(cand); return isNaN(d) ? 0 : d.getTime(); };
  return events.slice().sort((a,b) => ts(a)-ts(b))[events.length-1];
}


/* ===== Date parsing ===== */
function parseDateFlexible_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date(v);
  const s = String(v).trim().replace(/[.;]/g,'-').replace(/T/,' ');
  const iso = /^\d{4}-\d{2}-\d{2}/.test(s) ? s : s.replace(/^(\d{1,2})[-.](\d{1,2})[-.](\d{2,4})/, '$3-$2-$1');
  const dt = new Date(iso);
  return isNaN(dt) ? null : dt;
}

/* ===== Success meter ("mittari") ===== */
function isSuccessResult_(res){
  if (!res) return false;
  if (res.found) return true;
  const s = String(res.status||'').toUpperCase();
  if (!s) return false;
  if (s === 'MISSING_CREDENTIALS' || s === 'NO_DATA' || s === 'UNKNOWN_CARRIER' || s === 'TOKEN_FAIL' || s.indexOf('HTTP_')===0 || s === 'RATE_LIMIT_429') return false;
  return true;
}
function statsAppend_(mode, sheetName, tried, success, errors, rate429){
  try{
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('RunStats') || ss.insertSheet('RunStats');
    if (sh.getLastRow() === 0){
      sh.appendRow(['Timestamp','Mode','Sheet','Tried','Success','Errors','Rate429','Success%']);
    }
    const pct = tried ? (success / tried) : 0;
    sh.appendRow([new Date(), mode, sheetName || '', tried||0, success||0, errors||0, rate429||0, pct]);
  }catch(e){ logError_('statsAppend_', e); }
}


/* Heuristic fallback: choose a column that looks like an identifier (many unique, code-like) */
function chooseKeyIndexHeuristic_(headers, data){
  try{
    if (!headers || !headers.length || !data || data.length < 2) return -1;
    const rxCode = /^[A-Z0-9\-\s]{8,}$/i;
    let best = {idx:-1, score: -1};
    for (let c=0;c<headers.length;c++){
      let nonEmpty=0, unique=0;
      const seen = {};
      for (let r=1; r<Math.min(data.length, 200); r++){
        const v = String(data[r][c]||'').trim();
        if (!v) continue;
        nonEmpty++;
        if (!(v in seen)) { seen[v]=1; unique++; }
      }
      if (!nonEmpty) continue;
      const matchRate = (()=>{
        let m=0;
        for (let r=1; r<Math.min(data.length, 200); r++){ const v = String(data[r][c]||'').trim(); if (!v) continue; if (rxCode.test(v)) m++; }
        return nonEmpty ? (m/nonEmpty) : 0;
      })();
      const uniqRate = unique / Math.max(nonEmpty,1);
      const score = (uniqRate*0.6) + (matchRate*0.4);
      if (score > best.score) best = {idx:c, score};
    }
    return (best.score >= 0.5) ? best.idx : -1;
  }catch(e){ return -1; }
}


function findHeaderIndexIncludes_(headers, substrs){
  const n = headers.map(h => normalize_(h));
  for (let i=0;i<n.length;i++){
    for (const s of substrs){
      if (n[i].includes(normalize_(s))) return i;
    }
  }
  return -1;
}


function pickTrackingIndex_(hdr, data){
  const m = headerIndexMap_(hdr);
  // 0) property hint(s)
  const hint = TRK_props_('TRACKING_HEADER_HINT');
  if (hint){
    const names = String(hint).split(',').map(s=>s.trim()).filter(Boolean);
    for (const h of names){
      const direct = colIndexOf_(m, [h]);
      if (direct >= 0) return direct;
      const inc = findHeaderIndexIncludes_(hdr, [h]);
      if (inc >= 0) return inc;
    }
  }
  // 1) standard candidates
  let idx = colIndexOf_(m, TRACKING_CANDIDATES);
  // 2) broader includes (now also 'package')
  if (idx < 0) idx = findHeaderIndexIncludes_(hdr, ['track','barcode','waybill','awb','parcel','package']);
  // 3) heuristic (code-like, many unique)
  if (idx < 0 && typeof chooseKeyIndexHeuristic_==='function'){
    const hIdx = chooseKeyIndexHeuristic_(hdr, data);
    if (hIdx >= 0) idx = hIdx;
  }
  // 4) default to configured column (1-based), fallback to column A
  if (idx < 0){
    const def = Number(TRK_props_('TRACKING_DEFAULT_COL') || '1'); // default A
    if (def >= 1 && def <= hdr.length) idx = def - 1;
  }
  return idx;
}


function pickCarrierIndex_(hdr){
  const m = headerIndexMap_(hdr);
  // property hints (comma-separated)
  const hint = TRK_props_('CARRIER_HEADER_HINT');
  if (hint){
    const names = String(hint).split(',').map(s=>s.trim()).filter(Boolean);
    for (const h of names){
      const direct = colIndexOf_(m, [h]);
      if (direct >= 0) return direct;
      const inc = findHeaderIndexIncludes_(hdr, [h]);
      if (inc >= 0) return inc;
    }
  }
  // standard candidates
  let idx = colIndexOf_(m, CARRIER_CANDIDATES);
  // includes fallback
  if (idx < 0) idx = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper','delivery']);
  return idx;
}
