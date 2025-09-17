// actions.gs — pending builder, status refresh, quick actions, stats

function buildPendingFromPackagesAndArchive() {
  const ss = SpreadsheetApp.getActive();
  const shPack = ss.getSheetByName(TARGET_SHEET);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shPack || !shArch) { SpreadsheetApp.getUi().alert('Packages tai Packages_Archive -taulua ei löydy.'); return; }
  const dataPack = shPack.getDataRange().getValues();
  const dataArch = shArch.getDataRange().getValues();
  if (dataPack.length < 2) { SpreadsheetApp.getUi().alert('Packages-taulu on tyhjä.'); return; }
  const hdr = dataPack[0];
  const keyIdx = chooseKeyIndex_(hdr);
  const archMap = {};
  const archHdr = dataArch.length ? dataArch[0] : [];
  const archKeyIdx = archHdr.length ? chooseKeyIndex_(archHdr) : -1;
  for (let i=1; i<dataArch.length; i++) { const key = dataArch[i][archKeyIdx]; if (key) archMap[String(key)] = true; }
  const pending = [hdr];
  for (let j=1; j<dataPack.length; j++) {
    const key = dataPack[j][keyIdx];
    if (key && !archMap[String(key)]) pending.push(dataPack[j]);
  }
  let shPending = ss.getSheetByName(ACTION_SHEET) || ss.insertSheet(ACTION_SHEET);
  shPending.clear();
  shPending.getRange(1,1,pending.length,pending[0].length).setValues(pending);
}

function refreshStatuses_Vaatii() { refreshStatuses_Sheet(ACTION_SHEET, true); }

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
  let carrierI = pickCarrierIndex_(hdr);
  let trackI   = pickTrackingIndex_(hdr, data);
  // fallback header includes
  if (trackI < 0) trackI = findHeaderIndexIncludes_(hdr, ['track','barcode','waybill','awb','parcel']);
  if (carrierI < 0) carrierI = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper']);
  const rCarrierI = col['RefreshCarrier'];
  const rStatusI  = col['RefreshStatus'];
  const rTimeI    = col['RefreshTime'];
  const rLocI     = col['RefreshLocation'];
  const rRawI     = col['RefreshRaw'];

  /* HEURISTICS FALLBACK FOR TRACKING */
  if (trackI < 0) {
    const hIdx = chooseKeyIndexHeuristic_(hdr, data);
    if (hIdx >= 0) trackI = hIdx;
  }


  const now = Date.now();
  const cutoffMs = (getCfgInt_('REFRESH_MIN_AGE_MINUTES', 360)) * 60 * 1000;
  const writeRows = [];
  const toArchive = [];

  // stats
  let tried = 0, success = 0, errors = 0, rate429 = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row.join('')) continue;
    const carrier = String(row[carrierI] || row[rCarrierI] || '').toLowerCase();
    const code    = firstCode_(row[trackI] || '');
    const last    = row[rTimeI] ? new Date(row[rTimeI]).getTime() : 0;
    const expired = !last || (now - last > cutoffMs);
    let res = null;

    if (code && carrier && expired) { res = TRK_trackByCarrier_(carrier, code); tried++; }

    
    // Build write values with robust time: use API time if parseable, else NOW
    const timeVal = (res && res.time) ? (parseDateFlexible_(res.time) || new Date()) : new Date();
    const timeVal = (res && res.time) ? (parseDateFlexible_(res.time) || new Date()) : new Date();
    const values = [
      (res && res.carrier) || carrier || '',
      (res && res.status)  || row[rStatusI] || '',
      timeVal,
      (res && res.location)|| row[rLocI]   || '',
      (res && res.raw)     || row[rRawI]   || ''
    ];

    if (res){
      if (isSuccessResult_(res)) success++;
      else if (String(res.status||'')==='RATE_LIMIT_429') rate429++;
      else errors++;
    }
    const delivered = !!(res && String(res.status).toLowerCase().match(/delivered|toimitettu|luovutettu/));
    writeRows.push({ idx: i, values, delivered, rowSnapshot: row.slice() });
    if (removeDelivered && delivered) { toArchive.push({ idx: i, row: row.slice() }); }
  }

  const firstCol = hdr.indexOf('RefreshCarrier') + 1;
  writeRows.forEach(w => { try { sh.getRange(w.idx + 1, firstCol, 1, 5).setValues([w.values]); } catch(e){ logError_('refreshStatuses_Sheet setValues', e); } });

  // Ensure date format for RefreshTime column
  try {
    const timeCol = hdr.indexOf('RefreshTime') + 1;
    if (timeCol > 0 && sh.getLastRow() > 1) {
      sh.getRange(2, timeCol, sh.getLastRow()-1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    }
  } catch(e){ logError_('refreshStatuses_Sheet setNumberFormat', e); }

  if (removeDelivered && toArchive.length) {
    const shArch = ss.getSheetByName(ARCHIVE_SHEET) || ss.insertSheet(ARCHIVE_SHEET);
    toArchive.sort((a,b) => b.idx - a.idx).forEach(item => {
      try { shArch.appendRow(item.row); sh.deleteRow(item.idx + 1); } catch(e){ logError_('refreshStatuses_Sheet archive', e); }
    });
  }

  // stats out
  statsAppend_('manual-refresh', sheetName, tried, success, errors, rate429);
  SpreadsheetApp.getUi().alert(`Statuspäivitys valmis.\nKokeiltu: ${tried}\nOnnistui: ${success}\nVirheet: ${errors}\n429: ${rate429}`);
}

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
  if (statusCol === -1 || dateCol === -1) { SpreadsheetApp.getUi().alert('Ei löytynyt "Status" tai "Delivered at" -saraketta.'); return; }

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
    try { shArch.appendRow(item.row); shVaatii.deleteRow(item.idx+1); } catch(e){ logError_('confirmAndArchiveDelivered', e); }
  });
  SpreadsheetApp.getUi().alert(`Arkistoitu ${toMove.length} toimitettua riviä.`);
}


function diagnoseRefresh_(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName || ACTION_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('Diagnostiikka: taulua ei löydy.'); return; }
  const data = sh.getDataRange().getValues();
  if (data.length < 2){ SpreadsheetApp.getUi().alert('Diagnostiikka: taulu on tyhjä.'); return; }
  const hdr = data[0];
  const col = headerIndexMap_(hdr);
  let carrierI = pickCarrierIndex_(hdr);
  let trackI   = pickTrackingIndex_(hdr, data);
  const rTimeI = col['RefreshTime'];
  if (trackI < 0) trackI = findHeaderIndexIncludes_(hdr, ['track','barcode','waybill','awb','parcel']);
  if (carrierI < 0) carrierI = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper']);
  if (trackI < 0){ const hIdx = (typeof chooseKeyIndexHeuristic_==='function') ? chooseKeyIndexHeuristic_(hdr, data) : -1; if (hIdx>=0) trackI = hIdx; }

  let rows = 0, missingCarrier=0, missingCode=0, notExpired=0, eligible=0;
  for (let i=1;i<data.length;i++){
    const row = data[i];
    if (!row.join('')) continue;
    rows++;
    const carrier = String(trackI>=0 ? (row[carrierI] || row[col['RefreshCarrier']]) : (row[col['RefreshCarrier']] || '')).trim();
    const code = trackI>=0 ? String(row[trackI]||'').trim() : '';
    if (!carrier) missingCarrier++;
    if (!code) missingCode++;
    const last = rTimeI!==undefined ? row[rTimeI] : '';
    const cutoffMs = (getCfgInt_('REFRESH_MIN_AGE_MINUTES', 360)) * 60 * 1000;
    const expired = !last || (Date.now() - new Date(last).getTime() > cutoffMs);
    if (!expired) notExpired++;
    if (carrier && code && expired) eligible++;
  }

  const out = [
    ['Metric','Count'],
    ['Data rows (non-empty)', rows],
    ['Missing carrier', missingCarrier],
    ['Missing tracking code', missingCode],
    ['Not expired (<6h since last refresh)', notExpired],
    ['Eligible now (would be Tried)', eligible]
  ];
  const diag = ss.getSheetByName('RefreshDiag') || ss.insertSheet('RefreshDiag');
  diag.clear();
  diag.getRange(1,1,out.length,2).setValues(out);
  ss.setActiveSheet(diag);
  SpreadsheetApp.getUi().alert('Diagnostiikka valmis – katso RefreshDiag-välilehti.');
}


function highlightDetectedColumns_(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName || ACTION_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('Taulua ei löydy.'); return; }
  const data = sh.getDataRange().getValues();
  if (data.length < 1) { SpreadsheetApp.getUi().alert('Taulu on tyhjä.'); return; }
  const hdr = data[0];
  const col = headerIndexMap_(hdr);
  let carrierI = pickCarrierIndex_(hdr);  let trackI = pickTrackingIndex_(hdr, data);
  if (carrierI < 0) carrierI = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper']);

  // clear header formats
  sh.getRange(1,1,1,sh.getLastColumn()).setBackground(null).setFontWeight('normal');
  const notes = [];
  if (carrierI >= 0) {
    sh.getRange(1, carrierI+1).setBackground('#FFF4B8').setFontWeight('bold').setNote('Tunnistettu CARRIER-kentäksi');
    notes.push('Carrier: ' + hdr[carrierI]);
  }
  if (trackI >= 0) {
    sh.getRange(1, trackI+1).setBackground('#C7F0FF').setFontWeight('bold').setNote('Tunnistettu TRACKING-kentäksi');
    notes.push('Tracking: ' + hdr[trackI]);
  }
  SpreadsheetApp.getUi().alert(notes.length ? ('Korostettu: ' + notes.join(' | ')) : 'Ei tunnistettuja sarakkeita.');
}

function previewEligible_(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName || ACTION_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('Taulua ei löydy.'); return; }
  const data = sh.getDataRange().getValues();
  if (data.length < 2) { SpreadsheetApp.getUi().alert('Taulu on tyhjä.'); return; }

  const hdr = data[0];
  const col = headerIndexMap_(hdr);
  let carrierI = pickCarrierIndex_(hdr);  let trackI = pickTrackingIndex_(hdr, data);
  const rCarrierI = col['RefreshCarrier'];
  const rTimeI    = col['RefreshTime'];
  if (carrierI < 0) carrierI = findHeaderIndexIncludes_(hdr, ['carrier','courier','provider','forwarder','kuljetus','toimitustapa','shipper']);

  const out = [['Row','Carrier','Code','Last Refresh','Expired?']];
  let count = 0;
  for (let i=1;i<data.length;i++){
    const row = data[i];
    if (!row.join('')) continue;
    const carrier = String(row[carrierI] || row[rCarrierI] || '').trim();
    const code = String(row[trackI] || '').trim();
    const last = rTimeI!==undefined ? row[rTimeI] : '';
    const cutoffMs = (getCfgInt_('REFRESH_MIN_AGE_MINUTES', 360)) * 60 * 1000;
    const expired = !last || (Date.now() - new Date(last).getTime() > cutoffMs);
    if (carrier && code && expired){
      out.push([i+1, carrier, code, last, expired ? 'YES' : 'NO']);
      count++;
      if (count >= 50) break;
    }
  }
  const pv = ss.getSheetByName('RefreshPreview') || ss.insertSheet('RefreshPreview');
  pv.clear();
  if (count === 0){
    /* PREVIEW ZERO DIAGNOSTICS */
    const diag = [['Info','Value'],
      ['Detected carrier index', String(carrierI)],
      ['Detected tracking index', String(trackI)],
      ['RefreshTime index', String(rTimeI)]
    ];
    pv.getRange(1,1,diag.length,2).setValues(diag);
    const sample = [['Row','Carrier(sample)','Tracking(sample)','RefreshTime(sample)']];
    for (let i=1;i<Math.min(10, data.length); i++){
      sample.push([i+1, String(carrierI>=0?data[i][carrierI]:''), String(trackI>=0?data[i][trackI]:''), String(rTimeI!==undefined?data[i][rTimeI]:'')]);
    }
    pv.getRange(diag.length+2,1,sample.length,sample[0].length).setValues(sample);
  } else {
    pv.getRange(1,1,out.length,out[0].length).setValues(out);
  }
  ss.setActiveSheet(pv);
  SpreadsheetApp.getUi().alert(`Esikatselu valmis – ${count} riviä (max 50).`);
}


/* ===== Format fixer: enforce column formats so zeros & dates behave ===== */
function fixFormats_(sheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh){ SpreadsheetApp.getUi().alert('Taulua '+sheetName+' ei löytynyt.'); return; }
  const data = sh.getDataRange().getValues();
  if (!data.length){ SpreadsheetApp.getUi().alert('Taulu '+sheetName+' on tyhjä.'); return; }
  const hdr = data[0];
  const col = headerIndexMap_(hdr);
  const lastRow = Math.max(2, sh.getLastRow());
  const lastCol = sh.getLastColumn();

  // Detect indices
  const trackI   = pickTrackingIndex_(hdr, data);
  const carrierI = pickCarrierIndex_(hdr);
  const rTimeI   = col['RefreshTime'];
  const rCarI    = col['RefreshCarrier'];
  const rStatusI = col['RefreshStatus'];
  const rLocI    = col['RefreshLocation'];
  const rRawI    = col['RefreshRaw'];

  // Helper to format full column (rows 2..last)
  function fmtCol(idx, fmt){
    if (idx === undefined || idx < 0) return;
    const colN = idx + 1;
    if (lastRow >= 2) sh.getRange(2, colN, lastRow-1, 1).setNumberFormat(fmt);
  }
  function wrapCol(idx){
    if (idx === undefined || idx < 0) return;
    const colN = idx + 1;
    if (lastRow >= 2) sh.getRange(2, colN, lastRow-1, 1).setWrap(true);
  }

  // Enforce: tracking & carrier as text to preserve leading zeros etc.
  fmtCol(trackI, '@');
  fmtCol(carrierI, '@');
  // Our "Refresh*" columns
  if (rTimeI !== undefined) fmtCol(rTimeI, 'yyyy-mm-dd hh:mm');
  fmtCol(rCarI, '@'); fmtCol(rStatusI, '@'); fmtCol(rLocI, '@'); fmtCol(rRawI, '@');
  // Wrap long text fields
  wrapCol(rStatusI); wrapCol(rLocI); wrapCol(rRawI);

  SpreadsheetApp.getUi().alert('Formaattien korjaus tehty: '+sheetName);
}

function fixFormats_All(){
  try { fixFormats_(ACTION_SHEET); } catch(e){}
  try { fixFormats_(TARGET_SHEET); } catch(e){}
  try { fixFormats_(ARCHIVE_SHEET); } catch(e){}
  SpreadsheetApp.getUi().alert('Formaattien korjaus ajettu kaikille tauluille.');
}
