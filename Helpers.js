/******************************************************
 * Helpers, Logging, Formatting, Headers, Cache
 ******************************************************/

const DELIVERED_KEYWORDS = [
  'delivered','toimitettu','luovutettu',
  'delivered to pickup point','delivered to recipient','delivered - picked up'
];

const KEY_CANDIDATES = [
  ['Package Number','PackageNumber','Package No'],
  ['Orderid','Order id','Outbound order','Outbound order id'],
  ['Consignment ID','Consignment number','Shipment ID','Shipment No'],
  ['Waybill','Waybill No','AWB'],
  ['Tracking number','Tracking','Barcode']
];
const TRACKING_CODE_CANDIDATES = [
  'Tracking number','Tracking','Barcode','Waybill','Waybill No','AWB',
  'Package Number','PackageNumber','Shipment ID','Consignment number'
];
const CARRIER_CANDIDATES = [
  'Carrier','Carrier name','Service provider','Forwarder','Transporter','Kuljetusliike','Service family'
];
const DATE_FIELDS = [
  'Submitted date','Created','Created date','Booked time','Booking date',
  'Dispatch date','Shipped date','Timestamp','Date','Delivered date','Delivery date'
];

function getOrCreateSheet_(name){
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function writeHeaderOnce_(sh, headers){
  if (!sh.getLastRow()) sh.getRange(1,1,1,headers.length).setValues([headers]);
}

function normalize_(s){
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/[^\p{L}\p{N}]+/gu,' ');
}
function sanitizeMatrix_(values){
  if (!values || !values.length) return [];
  const m = values.map(r => r.map(v => (v === null || v === undefined) ? '' : v));
  if (typeof m[0][0] === 'string') m[0][0] = m[0][0].replace(/^\uFEFF/, '');
  m[0] = m[0].map(h => String(h||'').trim());
  let lastCol = m[0].length - 1;
  while (lastCol > 0 && m.every(row => String(row[lastCol]||'').trim() === '')) lastCol--;
  m.forEach(row => row.splice(lastCol+1));
  return m;
}
function mergeHeaders_(oldHdr, newHdr){
  const out = (oldHdr || []).slice();
  const have = new Set((oldHdr || []).map(normalize_));
  for (const h of (newHdr || [])){
    const n = normalize_(h);
    if (!have.has(n)){ out.push(h); have.add(n); }
  }
  return out;
}
function headerIndexMap_(hdr){
  const m = {}; (hdr || []).forEach((h,i)=>m[h]=i); return m;
}
function pickAnyIndex_(headers, candidates){
  const norm = headers.map(normalize_);
  for (const name of candidates){
    const i = norm.indexOf(normalize_(name));
    if (i >= 0) return i;
  }
  return -1;
}
function chooseKeyIndex_(headers){
  const hdr = headers.map(h => String(h||'').trim());
  for (const alts of KEY_CANDIDATES){
    const norm = hdr.map(normalize_), cand = alts.map(normalize_);
    for (const a of cand){ const i = norm.indexOf(a); if (i >= 0) return i; }
  }
  return -1;
}
function pickDeliveredIndex_(headers){
  const norm = headers.map(normalize_);
  const list = ['Delivered date','Delivery date','Delivered','Delivered on','Toimitettu','Luovutettu'];
  for (const l of list){ const i = norm.indexOf(normalize_(l)); if (i >= 0) return i; }
  return -1;
}

function isDeliveredByText_(s){
  const L = String(s||'').toLowerCase();
  return DELIVERED_KEYWORDS.some(w => L.includes(String(w).toLowerCase()));
}
function isDelivered_(row, statusI, dateI){
  if (dateI >= 0 && String(row[dateI]||'').trim()) return true;
  if (statusI >= 0 && isDeliveredByText_(row[statusI])) return true;
  return false;
}

function firstCode_(v){
  let s = String(v||'').trim();
  if (!s) return '';
  const parts = s.split(/[,\n;]/).map(x => String(x||'').trim()).filter(x => x);
  return parts.length ? parts[0] : s;
}
function isCodeLikely_(code){
  const s = String(code||'').trim(); return s.length >= 4;
}

function fmtDateTime_(d){
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function dateToYMD_(d){
  if (!(d instanceof Date)) return String(d||'');
  const yyyy = d.getFullYear();
  const MM = String(d.getMonth()+1).padStart(2,'0');
  const dd = String(d.getDate()).padStart(2,'0');
  return `${yyyy}-${MM}-${dd}`;
}
function parseDateFlexible_(val){
  if (!val) return null;
  if (val instanceof Date && !isNaN(val)) return val;
  const s = String(val).trim();

  // ISO-like
  let d = new Date(s);
  if (!isNaN(d)) return d;

  // dd.MM.yyyy [HH:mm[:ss]]
  let m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m){
    const dd = m[1].padStart(2,'0'), MM = m[2].padStart(2,'0'), yyyy = m[3];
    const hh = (m[4]||'00').padStart(2,'0'), mi = (m[5]||'00').padStart(2,'0'), ss = (m[6]||'00').padStart(2,'0');
    d = new Date(`${yyyy}-${MM}-${dd}T${hh}:${mi}:${ss}`);
    if (!isNaN(d)) return d;
  }

  // yyyy-MM-dd HH:mm[:ss]
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (m){
    d = new Date(`${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]||'00'}`);
    if (!isNaN(d)) return d;
  }
  return null;
}

function canonicalCarrier_(s){
  const c = String(s||'').toLowerCase();
  if (/posti/.test(c)) return 'posti';
  if (/gls/.test(c)) return 'gls';
  if (/dhl/.test(c)) return 'dhl';
  if (/bring/.test(c)) return 'bring';
  if (/matkahuolto|mh/.test(c)) return 'matkahuolto';
  return 'other';
}

/** Script Properties helpers */
function spGet_(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function spSet_(k,v){ PropertiesService.getScriptProperties().setProperty(k, String(v)); }

/** Logging HTTP errors into RUN_LOG_SHEET */
function logHttpError_(carrier, codeOrTrack, where, url, httpCode, tag, body, retryAfter){
  try {
    const sh = getOrCreateSheet_(RUN_LOG_SHEET);
    writeHeaderOnce_(sh, ['Time','Carrier','Fn','HTTP','Tag','Code/Track','Retry-After','Snippet']);
    sh.insertRowsAfter(1, 1);
    sh.getRange(2,1,1,8).setValues([[
      fmtDateTime_(new Date()), carrier, where, httpCode || '', tag || '',
      String(codeOrTrack || ''), retryAfter || '', String(body||'').slice(0,250)
    ]]);
  } catch(e) {}
}

/** Script Cache */
function trkCacheKey_(carrier, code){
  const b = spGet_('TRK_CACHE_BUSTER') || '0';
  return `TRK|${b}|${(carrier||'').toLowerCase()}|${code}`;
}
function trkGetCached_(carrier, code){
  try { const c = CacheService.getScriptCache(); const raw = c.get(trkCacheKey_(carrier,code)); return raw ? JSON.parse(raw) : null; } catch(e){ return null; }
}
function trkSetCached_(carrier, code, res){
  try { CacheService.getScriptCache().put(trkCacheKey_(carrier,code), JSON.stringify(res||{}), 6*3600); } catch(e){}
}
function trkClearCache(){ spSet_('TRK_CACHE_BUSTER', String(Date.now())); }

/** Per-carrier throttle (min ms between calls) */
function trkThrottle_(carrierTag){
  const minMs = Number(spGet_(`RATE_MINMS_${carrierTag}`) || '0') || 0;
  if (!minMs) return;
  const key = `LAST_CALL_${carrierTag}`;
  const now = Date.now();
  const last = Number(spGet_(key) || '0') || 0;
  const sleep = last + minMs - now;
  if (sleep > 0 && sleep < 30000) Utilities.sleep(sleep);
  spSet_(key, String(Date.now()));
}

/** Suggest new limits on 429 */
function autoTuneRateLimitOn429_(carrierTag, minMs, maxMs){
  if (minMs) spSet_(`RATE_MINMS_${carrierTag}`, String(minMs));
  if (maxMs) spSet_(`RATE_MAXMS_${carrierTag}`, String(maxMs));
}

/** Lock helpers to prevent overlapping bulk runs */
function withLock_(name, fn){
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try { return fn(); } finally { lock.releaseLock(); }
}