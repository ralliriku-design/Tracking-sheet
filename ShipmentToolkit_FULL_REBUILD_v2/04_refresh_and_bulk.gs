
// 04_refresh_and_bulk.gs — refresh with min-age, forced mode
function getMinAgeMinutes_(){
  const sp = PropertiesService.getScriptProperties();
  const v = sp.getProperty('REFRESH_MIN_AGE_MINUTES');
  return v ? Number(v) : 360; // default 6h
}

function refreshStatuses_Vaatii(){
  return refreshStatuses_Sheet(ACTION_SHEET, true);
}

function refreshStatuses_Sheet(sheetName, removeDelivered){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  if (!data.length) return;
  let hdr = data[0];
  // ensure output columns exist
  const outHdr = ['RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw'];
  outHdr.forEach(col => { if (!hdr.includes(col)) hdr.push(col); });
  sh.getRange(1,1,1,hdr.length).setValues([hdr]);
  const col = headerIndexMap_(hdr);
  const rFirst = hdr.indexOf('RefreshCarrier') + 1;

  const carrierI = pickCarrierIndex_(hdr);
  const trackI = pickTrackingIndex_(hdr, data);
  const timeI = col['RefreshTime'];
  const minAgeMs = getMinAgeMinutes_()*60*1000;
  const now = Date.now();

  const out = [];
  for (let r=1;r<data.length;r++){
    const row = data[r];
    if (!row || !row.join('')){ out.push(['','','','','']); continue; }
    const code = firstCode_(row[trackI]||'');
    let carrier = row[carrierI] || row[col['RefreshCarrier']] || row[col['Carrier']] || '';
    const last = row[timeI] || '';
    const lastTs = last ? new Date(last).getTime() : 0;
    const expired = !lastTs || (now - lastTs > minAgeMs);
    if (!code || !carrier || !expired){
      out.push([
        row[col['RefreshCarrier']] || carrier || '',
        row[col['RefreshStatus']] || '',
        row[col['RefreshTime']] || '',
        row[col['RefreshLocation']] || '',
        row[col['RefreshRaw']] || ''
      ]);
      continue;
    }
    let res = TRK_trackByCarrier_(carrier, code);
    out.push([
      res.carrier || carrier || '',
      res.status || '',
      res.time || fmtDateTime_(new Date()),
      res.location || '',
      (res.raw || '').slice(0,2000)
    ]);
  }
  sh.getRange(2, rFirst, out.length, 5).setValues(out);
}

function withMinAge_(minutes, fn){
  const sp = PropertiesService.getScriptProperties();
  const prev = sp.getProperty('REFRESH_MIN_AGE_MINUTES');
  try{ sp.setProperty('REFRESH_MIN_AGE_MINUTES', String(minutes)); return fn(); }
  finally{
    if (prev === null || typeof prev === 'undefined') sp.deleteProperty('REFRESH_MIN_AGE_MINUTES');
    else sp.setProperty('REFRESH_MIN_AGE_MINUTES', prev);
  }
}
function refreshStatuses_Vaatii_FORCE(){
  return withMinAge_(0, function(){ return refreshStatuses_Sheet(ACTION_SHEET, true); });
}
