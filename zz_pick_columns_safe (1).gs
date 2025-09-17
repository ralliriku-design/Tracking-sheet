/**
 * zz_pick_columns_safe.gs
 * Collision-safe helpers for picking Tracking/Carrier columns.
 * - Uses unique UX_* function names to avoid "already declared" errors.
 * - Guards globals so this file can be added multiple times without conflicts.
 *
 * How to use:
 *  1) Add this file to your Apps Script project as "zz_pick_columns_safe.gs".
 *  2) Run debugActiveSheetConfig_UX() once to verify indices.
 *  3) In your code, get indices like:
 *       const hdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
 *       const trkIdx = UX_pickTrackingIdx(hdr, Number(TRK_props_('TRACKING_DEFAULT_COL')||1)-1);
 *       const carIdx = UX_pickCarrierIdx(hdr, Number(TRK_props_('CARRIER_DEFAULT_COL')||25)-1);
 */

// ---- Guards for globals (avoid "already declared") ----
if (typeof UX_TRACKING_CANDIDATES === 'undefined') {
  var UX_TRACKING_CANDIDATES = [
    'Package Number','Packagenumber','Tracking Number','Trackingnumber',
    'Tracking','Tracking Code','Seuranta','Seurantanumero','Waybill','Waybill Number',
    'AWB','TrackingReference','Transfer/Tracking Reference'
  ];
}
if (typeof UX_CARRIER_CANDIDATES === 'undefined') {
  var UX_CARRIER_CANDIDATES = [
    'Carrier','Logistics Provider','LogisticsProvider','Shipper',
    'Kuljetusliike','Kuljetusyhtiö'
  ];
}

// ---- Local normalize and header map (namespaced to avoid clashes) ----
function UX_normalize_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim();
}
function UX_headerIndexMap_(hdr) {
  const m = {};
  (hdr || []).forEach((h,i) => {
    const n = UX_normalize_(h || '');
    m[h] = i;        // original
    m[n] = i;        // normalized
  });
  return m;
}

// ---- Collision-safe version of your helper ----
function UX_colIndexOf_(map, names){
  for (const n of (names||[])){
    if (n in map) return map[n];
    const low = UX_normalize_(n);
    if (low in map) return map[low];
  }
  return -1;
}

// ---- Public pickers (with fallback index support) ----
function UX_pickTrackingIdx(headers, fallbackZeroBased){
  const map = UX_headerIndexMap_(headers);
  const idx = UX_colIndexOf_(map, UX_TRACKING_CANDIDATES);
  return idx >= 0 ? idx : (typeof fallbackZeroBased === 'number' ? fallbackZeroBased : -1);
}
function UX_pickCarrierIdx(headers, fallbackZeroBased){
  const map = UX_headerIndexMap_(headers);
  const idx = UX_colIndexOf_(map, UX_CARRIER_CANDIDATES);
  return idx >= 0 ? idx : (typeof fallbackZeroBased === 'number' ? fallbackZeroBased : -1);
}

// ---- Quick diagnostic ----
function debugActiveSheetConfig_UX(){
  const sh = SpreadsheetApp.getActiveSheet();
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const trk = UX_pickTrackingIdx(hdr, Number(TRK_props_ && TRK_props_('TRACKING_DEFAULT_COL') || 1)-1);
  const car = UX_pickCarrierIdx(hdr, Number(TRK_props_ && TRK_props_('CARRIER_DEFAULT_COL')  || 25)-1);
  SpreadsheetApp.getUi().alert(
    'Tracking: ' + (trk>=0 ? hdr[trk]+' (col '+(trk+1)+')' : 'ei löytynyt') + '\n' +
    'Carrier: '  + (car>=0 ? hdr[car]+' (col '+(car+1)+')' : 'ei löytynyt')
  );
}
