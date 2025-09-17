
// 05_metrics_and_diagnostics.gs
function diagnoseRefresh_Vaatii(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ACTION_SHEET);
  if (!sh){ SpreadsheetApp.getUi().alert('Taulua '+ACTION_SHEET+' ei löytynyt.'); return; }
  const data = sh.getDataRange().getValues(); if (!data || data.length<2){ SpreadsheetApp.getUi().alert('Ei dataa.'); return; }
  const hdr = data[0], col = headerIndexMap_(hdr);
  const carrierI = pickCarrierIndex_(hdr);
  const trackI = pickTrackingIndex_(hdr, data);
  const timeI = col['RefreshTime'] ?? -1;
  const minAge = (Number(PropertiesService.getScriptProperties().getProperty('REFRESH_MIN_AGE_MINUTES')||'360')*60*1000);
  let nonEmpty=0, missCarrier=0, missTrack=0, notExpired=0;
  const eligible = [];
  const now = Date.now();
  for (let r=1;r<data.length;r++){
    const row = data[r]; if (!row || !row.join('')) continue;
    nonEmpty++;
    const code = firstCode_(row[trackI]||''); if (!code){ missTrack++; continue; }
    const car = row[carrierI] || row[col['RefreshCarrier']] || row[col['Carrier']] || ''; if (!car){ missCarrier++; continue; }
    const last = timeI>=0 ? row[timeI] : '';
    const ts = last ? new Date(last).getTime() : 0;
    const expired = !ts || (now - ts > minAge);
    if (!expired){ notExpired++; continue; }
    eligible.push(r);
  }
  const prev = ss.getSheetByName('Preview_Eligible') || ss.insertSheet('Preview_Eligible');
  prev.clear();
  prev.getRange(1,1,1,5).setValues([['Tracking','Carrier','Status','RefreshTime','Location']]);
  const rows = eligible.slice(0,50).map(r => [
    data[r][trackI], data[r][carrierI] || data[r][col['RefreshCarrier']] || '',
    data[r][col['RefreshStatus']] || '', data[r][timeI] || '', data[r][col['RefreshLocation']] || ''
  ]);
  if (rows.length) prev.getRange(2,1,rows.length,5).setValues(rows);
  SpreadsheetApp.getUi().alert(
    'Diagnostiikka valmis\n' +
    'Data rows (non-empty): '+nonEmpty+'\n' +
    'Missing carrier: '+missCarrier+'\n' +
    'Missing tracking code: '+missTrack+'\n' +
    'Not expired (<min-age): '+notExpired+'\n' +
    'Eligible now (would be Tried): '+eligible.length+'\n' +
    'Esikatselu: Preview_Eligible (max 50)'
  );
}

function refreshStatuses_Vaatii_FORCE_METRICS(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ACTION_SHEET);
  if (!sh){ SpreadsheetApp.getUi().alert('Taulua '+ACTION_SHEET+' ei löytynyt.'); return; }
  const data = sh.getDataRange().getValues(); if (!data || data.length<2){ SpreadsheetApp.getUi().alert('Ei dataa.'); return; }
  const hdr = data[0], col = headerIndexMap_(hdr);
  const carrierI = pickCarrierIndex_(hdr);
  const trackI = pickTrackingIndex_(hdr, data);
  const timeI = col['RefreshTime'] ?? -1;
  const targetRows = [];
  for (let r=1;r<data.length;r++){
    const row = data[r]; if (!row || !row.join('')) continue;
    const code = firstCode_(row[trackI]||''); if (!code) continue;
    const car = row[carrierI] || row[col['RefreshCarrier']] || row[col['Carrier']] || ''; if (!car) continue;
    targetRows.push(r);
  }
  // snapshot BEFORE times
  const before = {};
  for (const r of targetRows){
    const t = timeI>=0 ? data[r][timeI] : '';
    before[r] = t ? new Date(t).getTime() : 0;
  }
  // force refresh
  refreshStatuses_Vaatii_FORCE();
  // re-read
  const data2 = sh.getDataRange().getValues();
  const statusI = col['RefreshStatus'] ?? -1;
  let tried=0, success=0, errors=0, rate429=0;
  for (const r of targetRows){
    const afterCell = timeI>=0 ? data2[r][timeI] : '';
    const after = afterCell ? new Date(afterCell).getTime() : 0;
    if (after > before[r]){
      tried++;
      const st = String(statusI>=0 ? data2[r][statusI] : '').toLowerCase();
      if (!st || /^no_data|missing_|unknown_|unsupported_|http_|rate_limit/.test(st)){
        errors++;
        if (/rate_limit/.test(st)) rate429++;
      } else {
        success++;
      }
    }
  }
  const met = ss.getSheetByName('StatusMetrics') || ss.insertSheet('StatusMetrics');
  if (met.getLastRow() === 0){
    met.appendRow(['Timestamp','Mode','Sheet','Tried','Success','Errors','Rate429','Success%']);
  }
  const pct = tried ? (success/tried) : 0;
  met.appendRow([new Date(),'force-refresh',ACTION_SHEET,tried,success,errors,rate429,pct]);
  SpreadsheetApp.getUi().alert('Pakotus valmis.\nTried: '+tried+'\nSuccess: '+success+'\nErrors: '+errors+'\nRate429: '+rate429+'\nSuccess%: '+(tried?Math.round(100*success/tried):0)+'%');
}
