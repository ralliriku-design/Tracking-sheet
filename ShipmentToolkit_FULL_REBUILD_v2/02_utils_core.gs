
// 02_utils_core.gs — helpers, props, headers, sanitizer
const ACTION_SHEET = 'Vaatii_toimenpiteitä';
const ARCHIVE_SHEET = 'Packages_Archive';

function TRK_props_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function normalize_(s){ return String(s||'').toLowerCase().replace(/\s+/g,' ').trim(); }

function headerIndexMap_(hdr){
  const m = {}; (hdr||[]).forEach((h,i)=>{ m[String(h)] = i; m[normalize_(h)] = i; });
  return m;
}
function colIndexOf_(map, names){
  for (const n of (names||[])){
    if (n in map) return map[n];
    const low = normalize_(n);
    if (low in map) return map[low];
  }
  return -1;
}

function sanitizeTrackingCode_(val){
  if (val === null || typeof val === 'undefined') return '';
  let s = String(val).trim();
  if (/^[0-9]+(?:\.[0-9]+)?e\+[0-9]+$/i.test(s)){
    const m = s.toLowerCase().match(/^([0-9]+)(?:\.([0-9]+))?e\+([0-9]+)$/);
    if (m){
      let int = m[1]||'', frac = m[2]||'', exp = Number(m[3]||'0');
      s = exp <= frac.length ? int + frac.slice(0,exp) + (frac.slice(exp)||'')
                             : int + frac + '0'.repeat(exp - frac.length);
    }
  }
  if (/^[0-9]+\.0$/.test(s)) s = s.replace(/\.0$/, '');
  return s;
}
function firstCode_(cell){
  const raw = String(cell||''); const first = raw.split(/[\n,;]/)[0].trim();
  return sanitizeTrackingCode_(first);
}

function pickTrackingIndex_(hdr, data){
  const m = headerIndexMap_(hdr);
  const hint = TRK_props_('TRACKING_HEADER_HINT');
  if (hint){
    const names = String(hint).split(',').map(s=>s.trim()).filter(Boolean);
    const idx = colIndexOf_(m, names);
    if (idx >= 0) return idx;
  }
  const cands = ['Package Number','Tracking Number','Tracking','Barcode','Waybill','AWB','Consignment','Parcel'];
  let idx = colIndexOf_(m, cands);
  if (idx >= 0) return idx;
  const def = Number(TRK_props_('TRACKING_DEFAULT_COL') || '1'); // default A
  return Math.max(0, def-1);
}

function findHeaderIndexIncludes_(headers, substrs){
  const n = (headers || []).map(h => normalize_(h));
  for (let i=0;i<n.length;i++){
    for (const s of (substrs || [])){
      if (n[i].includes(normalize_(s))) return i;
    }
  }
  return -1;
}
const CARRIER_CANDIDATES = [
  'RefreshCarrier','Carrier','Carrier name','CarrierName','Delivery Carrier','Courier',
  'LogisticsProvider','Logistics Provider','Shipper','Service provider','Forwarder',
  'Transporter','Kuljetusliike','Kuljetusyhtiö','Kuljetus','Toimitustapa','Delivery method'
];
function pickCarrierIndex_(hdr){
  const m = headerIndexMap_(hdr);
  const hint = TRK_props_('CARRIER_HEADER_HINT');
  if (hint){
    const names = String(hint).split(',').map(s=>s.trim()).filter(Boolean);
    const idx = colIndexOf_(m, names);
    if (idx >= 0) return idx;
    const inc = findHeaderIndexIncludes_(hdr, names);
    if (inc >= 0) return inc;
  }
  let idx = colIndexOf_(m, CARRIER_CANDIDATES);
  if (idx < 0) idx = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper','delivery']);
  if (idx < 0){
    const def = Number(TRK_props_('CARRIER_DEFAULT_COL') || '25'); // default Y
    if (def >= 1 && def <= hdr.length) idx = def - 1;
  }
  return idx;
}

function fmtDateTime_(d) {
  if (!d) return '';
  try {
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  } catch (e) {
    try { return new Date(d).toISOString(); } catch (err) { return String(d); }
  }
}

function logError_(context, error) {
  const ss = SpreadsheetApp.getActive();
  const shLog = ss.getSheetByName('ErrorLog') || ss.insertSheet('ErrorLog');
  shLog.appendRow([new Date(), context, String(error)]);
  if (shLog.getLastRow() > 2000) shLog.deleteRows(1, shLog.getLastRow() - 2000);
}
function openErrorLog(){
  const ss = SpreadsheetApp.getActive(); const sh = ss.getSheetByName('ErrorLog'); if (sh) ss.setActiveSheet(sh);
}
