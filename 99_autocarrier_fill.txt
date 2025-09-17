
// 99_autocarrier_fill.gs — optional autodetect scan if carrier missing
function TRK_trackAuto_(code){
  const orderProp = String(TRK_props_('AUTO_SCAN_ORDER') || '').trim();
  const order = (orderProp ? orderProp.split(',') : ['posti','gls','matkahuolto','dhl','bring'])
    .map(s => String(s||'').trim().toLowerCase()).filter(Boolean);
  for (const c of order){
    let res=null;
    try {
      if (c === 'gls') res = TRK_trackGLS(code);
      else if (c === 'posti') res = TRK_trackPosti(code);
      else if (c === 'dhl') res = TRK_trackDHL(code);
      else if (c === 'bring') res = TRK_trackBring(code);
      else if (c === 'matkahuolto') res = TRK_trackMH(code);
    } catch(e){ res = null; }
    if (res && (res.found || (res.status && String(res.status).toLowerCase()!=='no_data'))){
      if (!res.carrier) res.carrier = c.charAt(0).toUpperCase()+c.slice(1);
      return res;
    }
  }
  return { carrier:'', status:'NO_MATCH_IN_AUTO_SCAN' };
}

function autodetectAndRefresh_Vaatii(limit){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ACTION_SHEET);
  if (!sh){ SpreadsheetApp.getUi().alert('Taulua '+ACTION_SHEET+' ei löytynyt.'); return; }
  const data = sh.getDataRange().getValues(); if (!data || data.length<2){ SpreadsheetApp.getUi().alert('Ei dataa.'); return; }
  const hdr = data[0], col = headerIndexMap_(hdr);
  const trackI = pickTrackingIndex_(hdr, data);
  const rCarrierI = col['RefreshCarrier'] ?? -1;
  const rStatusI  = col['RefreshStatus']  ?? -1;
  const rTimeI    = col['RefreshTime']    ?? -1;
  const rLocI     = col['RefreshLocation']?? -1;
  const rRawI     = col['RefreshRaw']     ?? -1;

  let tried=0, ok=0, errors=0, filled=0;
  const cap = Number(limit||0) || 200;
  const maxRows = sh.getLastRow();

  for (let r=1; r<maxRows && tried < cap; r++){
    const row = data[r];
    if (!row || !row.join('')) continue;
    const code = firstCode_(row[trackI]||''); if (!code) continue;
    const hasCarrier = (rCarrierI>=0 ? row[rCarrierI] : '') || '';
    if (hasCarrier) continue;
    tried++;
    const res = TRK_trackAuto_(code);
    if (res && (res.found || res.status)){
      ok++;
      const values = [
        res.carrier || '',
        res.status || '',
        parseDateFlexible_(res.time) || new Date(),
        res.location || '',
        (res.raw || '').slice(0,2000)
      ];
      const firstCol = rCarrierI>=0 ? (rCarrierI+1) : (hdr.length+1);
      sh.getRange(r+1, firstCol, 1, values.length).setValues([values]);
      if (res.carrier) filled++;
    } else {
      errors++;
    }
  }
  SpreadsheetApp.getUi().alert('Autodetect done. Tried '+tried+', OK '+ok+', Errors '+errors+', Carriers filled '+filled+'.');
}

// tiny helper to parse flexible dates
function parseDateFlexible_(v) {
  if (!v) return null;
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5));
  const s = String(v).trim().replace(/[.;]/g,'-').replace(/T/,' ');
  const iso = s.match(/^\d{4}-\d{2}-\d{2}/) ? s : s.replace(/^(\d{1,2})[-.](\d{1,2})[-.](\d{2,4})/, '$3-$2-$1');
  const dt = new Date(iso);
  return isNaN(dt) ? null : dt;
}
