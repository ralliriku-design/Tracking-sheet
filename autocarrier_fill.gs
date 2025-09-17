// autocarrier_fill.gs — Auto-detect and fill carriers + statuses for Vaatii_toimenpiteitä
// Uses existing TRK_track* functions. Safe to add alongside your current project.

/** Try carriers in order until one returns meaningful data. */
function TRK_trackAuto_(code){
  const orderProp = String(TRK_props_('AUTO_SCAN_ORDER') || '').trim();
  const order = (orderProp ? orderProp.split(',') : ['posti','gls','matkahuolto','dhl','bring'])
    .map(s => String(s||'').trim().toLowerCase()).filter(Boolean);
  for (const c of order){
    let res = null;
    try {
      if (c === 'gls') res = TRK_trackGLS(code);
      else if (c === 'posti') res = TRK_trackPosti(code);
      else if (c === 'dhl') res = TRK_trackDHL(code);
      else if (c === 'bring') res = TRK_trackBring(code);
      else if (c === 'matkahuolto') res = TRK_trackMH(code);
    } catch(e){ res = null; }
    if (res && (res.found || (res.status && String(res.status).toLowerCase() !== 'no_data'))){
      if (!res.carrier) res.carrier = c.charAt(0).toUpperCase()+c.slice(1);
      return res;
    }
  }
  return { carrier:'', status:'NO_MATCH_IN_AUTO_SCAN' };
}

/** Fix common numeric tracking formatting issues like "903289763248.0" */
function sanitizeTrackingCode_(val){
  if (val === null || typeof val === 'undefined') return '';
  let s = String(val).trim();
  if (/^[0-9]+(?:\.[0-9]+)?e\+[0-9]+$/i.test(s)){
    const m = s.toLowerCase().match(/^([0-9]+)(?:\.([0-9]+))?e\+([0-9]+)$/i);
    if (m){
      let int = m[1] || ''; let frac = m[2] || ''; let exp = Number(m[3]||'0');
      if (exp <= frac.length) s = int + frac.slice(0, exp) + (frac.slice(exp) ? frac.slice(exp) : '');
      else s = int + frac + '0'.repeat(exp - frac.length);
    }
  }
  if (/^[0-9]+\.0$/.test(s)) s = s.replace(/\.0$/, '');
  return s;
}

/** Use this if you want to run only autodetect for rows with empty RefreshCarrier. */
function autodetectAndRefresh_Vaatii(limit){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(typeof ACTION_SHEET!=='undefined'?ACTION_SHEET:'Vaatii_toimenpiteitä');
  if (!sh){ SpreadsheetApp.getUi().alert('Taulua Vaatii_toimenpiteitä ei löytynyt.'); return; }
  const data = sh.getDataRange().getValues();
  if (!data || data.length < 2){ SpreadsheetApp.getUi().alert('Ei dataa.'); return; }
  const hdr = data[0];
  const col = headerIndexMap_(hdr);
  const trackI = (typeof pickTrackingIndex_==='function') ? pickTrackingIndex_(hdr, data) : 0; // default A
  const rCarrierI = col['RefreshCarrier'] ?? -1;
  const rStatusI  = col['RefreshStatus']  ?? -1;
  const rTimeI    = col['RefreshTime']    ?? -1;
  const rLocI     = col['RefreshLocation']?? -1;
  const rRawI     = col['RefreshRaw']     ?? -1;

  let tried = 0, ok = 0, errors = 0, filledCarrier = 0;
  const maxRows = Math.max(2, sh.getLastRow());
  const cap = Number(limit||0) || 200; // default cap

  for (let r=1; r<maxRows && tried < cap; r++){
    const row = data[r];
    if (!row || !row.join('')) continue;
    const codeRaw = (trackI>=0 ? row[trackI] : row[0]);
    let code = sanitizeTrackingCode_(firstCode_(codeRaw||''));
    if (!code) continue;

    let carrierCell = (rCarrierI>=0 ? row[rCarrierI] : '');
    if (carrierCell) continue; // only autodetect missing carriers

    tried++;
    let res = TRK_trackAuto_(code);
    if (res && (res.found || res.status)){
      ok++;
      const timeVal = res.time ? (parseDateFlexible_(res.time) || new Date()) : new Date();
      const values = [
        res.carrier || '',
        res.status || '',
        timeVal,
        res.location || '',
        (res.raw || '').slice(0,2000)
      ];
      const firstCol = rCarrierI>=0 ? (rCarrierI+1) : (hdr.length+1);
      sh.getRange(r+1, firstCol, 1, values.length).setValues([values]);
      filledCarrier += res.carrier ? 1 : 0;
    } else {
      errors++;
    }
  }
  SpreadsheetApp.getUi().alert(`Autodetect done. Tried ${tried}, OK ${ok}, Errors ${errors}, Carriers filled ${filledCarrier}.`);
}
