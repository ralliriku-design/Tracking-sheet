/******************************************************
 * Vaatii_toimenpiteitä (Pending) builder and refresh
 ******************************************************/

function collectTables_(){
  const ss = SpreadsheetApp.getActive();
  const tables = [];
  const add = (name) => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1){
      const d = sh.getDataRange().getDisplayValues();
      tables.push({ hdr: d[0].map(String), rows: d.slice(1) });
    }
  };
  add(TARGET_SHEET); add(ARCHIVE_SHEET);
  return tables;
}

function buildPendingFromPackagesAndArchive(){
  const ss = SpreadsheetApp.getActive();
  const tables = collectTables_();
  if (!tables.length) throw new Error('Ei rivejä lähdetauluissa.');
  let unionHdr = [];
  tables.forEach(t => unionHdr = mergeHeaders_(unionHdr, t.hdr));

  const deliveredI     = pickDeliveredIndex_(unionHdr);
  const deliveredFlagI = pickAnyIndex_(unionHdr, ['Status','Delivery status','Current status','State','Tila','Vaihe','Stage']);
  const keyI           = chooseKeyIndex_(unionHdr);

  const dst = new Map();
  const dstMap = headerIndexMap_(unionHdr);
  for (const T of tables){
    const srcMap = headerIndexMap_(T.hdr);
    for (const r of T.rows){
      if (!r.some(x => String(x||'').trim() !== '')) continue;
      const u = new Array(unionHdr.length).fill('');
      for (const [name, si] of Object.entries(srcMap)){
        const di = dstMap[name];
        if (typeof di === 'number') u[di] = r[si];
      }
      const key = String(u[keyI]||'').trim();
      if (!key) continue;
      const delivered = isDelivered_(u, deliveredFlagI, deliveredI);
      if (delivered) continue;
      dst.set(key, u);
    }
  }

  const out = Array.from(dst.values());
  const sh = getOrCreateSheet_(ACTION_SHEET);
  sh.clear();
  const headers = unionHdr.concat([
    'RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw',
    'RefreshAt','RefreshAttempts','RefreshNextAt','Delivered date (Confirmed)','Delivered_Source'
  ]);
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (out.length) sh.getRange(2,1,out.length,unionHdr.length).setValues(out);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(headers.length, 20));
  ss.toast(`Vaatii_toimenpiteitä päivitetty – ${out.length} riviä`);
}

function ensureRefreshCols_(sh, hdr){
  const need = [
    'RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw','RefreshAt',
    'RefreshAttempts','RefreshNextAt','Delivered date (Confirmed)','Delivered_Source'
  ];
  const have = hdr.slice();
  need.forEach(n => { if (!have.includes(n)) have.push(n); });

  if (have.length !== hdr.length){
    const cur = sh.getDataRange().getValues();
    const rows = cur.length > 1 ? cur.slice(1).map(r => { const rr = r.slice(); while (rr.length < have.length) rr.push(''); return rr; }) : [];
    sh.clear();
    sh.getRange(1,1,1,have.length).setValues([have]);
    if (rows.length) sh.getRange(2,1,rows.length,have.length).setValues(rows);
  }
  const map = headerIndexMap_(have);
  return {
    carrier: map['RefreshCarrier'],
    status: map['RefreshStatus'],
    time: map['RefreshTime'],
    location: map['RefreshLocation'],
    raw: map['RefreshRaw'],
    at: map['RefreshAt'],
    attempts: map['RefreshAttempts'],
    nextAt: map['RefreshNextAt'],
    delivConfirmed: map['Delivered date (Confirmed)'],
    delivSource: map['Delivered_Source']
  };
}

function refreshStatuses_Vaatii(){
  refreshStatuses_Sheet(ACTION_SHEET, true);
}

function refreshStatuses_Sheet(sheetName, removeDelivered = false){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2){ ss.toast(`${sheetName}: ei rivejä`); return; }

  const data = sh.getDataRange().getDisplayValues();
  const origHdr = data[0].map(v => String(v||'').trim());
  const rows = data.slice(1);

  const idxRef = ensureRefreshCols_(sh, origHdr);
  const finalHdr = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
  const carrierI   = pickAnyIndex_(finalHdr, CARRIER_CANDIDATES);
  const codeI      = pickAnyIndex_(finalHdr, TRACKING_CODE_CANDIDATES);
  const deliveredI = pickDeliveredIndex_(finalHdr);
  if (carrierI < 0 || codeI < 0){ ss.toast(`${sheetName}: ei löydy Carrier/Tracking -sarakkeita`); return; }

  const out = [];
  const keepRows = [];
  rows.forEach(r => {
    const row = (() => { const rr = r.slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr; })();
    const carrier = String(row[carrierI]||'').trim();
    const code    = firstCode_(row[codeI]);

    if (!carrier || !code){
      row[idxRef.status] = 'SKIP_NO_CODE';
      row[idxRef.at]     = fmtDateTime_(new Date());
      out.push(row); keepRows.push(true);
      return;
    }

    const res = TRK_trackByCarrier_(carrier, code);
    row[idxRef.carrier]  = res.carrier || carrier;
    row[idxRef.status]   = res.status  || '';
    row[idxRef.time]     = res.time    || '';
    row[idxRef.location] = res.location|| '';
    row[idxRef.raw]      = res.raw     || '';
    row[idxRef.at]       = fmtDateTime_(new Date());

    if (res.status === 'RATE_LIMIT_429' && typeof res.retryAfter === 'number'){
      const nextAt = trkComputeNextAtFromRetryAfter_(res.retryAfter);
      if (idxRef.nextAt >= 0) row[idxRef.nextAt] = fmtDateTime_(nextAt);
    }
    const isDeliveredNow = isDeliveredByText_(res.status || '');
    if (isDeliveredNow && res.time){
      if (typeof idxRef.delivConfirmed === 'number' && !row[idxRef.delivConfirmed]) row[idxRef.delivConfirmed] = res.time;
      if (typeof idxRef.delivSource === 'number') row[idxRef.delivSource] = 'tracking';
    }

    const deliveredAlready = deliveredI >= 0 && String(row[deliveredI]||'').trim();
    out.push(row);
    keepRows.push(!(removeDelivered && (deliveredAlready || isDeliveredNow)));
  });

  sh.clearContents();
  sh.getRange(1,1,1,finalHdr.length).setValues([finalHdr]);
  const filtered = out.filter((_, i) => keepRows[i]).map(r => {
    const rr = r.slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr;
  });
  if (filtered.length) sh.getRange(2,1,filtered.length, finalHdr.length).setValues(filtered);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(finalHdr.length, 20));
  const removed = out.length - filtered.length;
  ss.toast(`${sheetName}: status päivitetty (${removed} poistettu koska delivered)`);
}

function refreshStatuses_Filtered(sheetName, carriers, removeDelivered = false){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2){ ss.toast(`${sheetName}: ei rivejä`); return; }

  const data = sh.getDataRange().getDisplayValues();
  const origHdr = data[0].map(v => String(v||'').trim());
  const rows = data.slice(1);

  const idxRef = ensureRefreshCols_(sh, origHdr);
  const finalHdr = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
  const carrierI   = pickAnyIndex_(finalHdr, CARRIER_CANDIDATES);
  const codeI      = pickAnyIndex_(finalHdr, TRACKING_CODE_CANDIDATES);
  const deliveredI = pickDeliveredIndex_(finalHdr);
  if (carrierI < 0 || codeI < 0){ ss.toast(`${sheetName}: ei löydy Carrier/Tracking -sarakkeita`); return; }

  const want = (carriers || []).map(s => String(s||'').toLowerCase());
  const out = [], keep = [];
  rows.forEach(r => {
    const row = (() => { const rr = r.slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr; })();
    const carr = String(row[carrierI]||'').toLowerCase();
    const code = firstCode_(row[codeI]);
    if (want.some(w => carr.includes(w))){
      const res = TRK_trackByCarrier_(carr, code);
      row[idxRef.carrier]  = res.carrier || row[idxRef.carrier] || '';
      row[idxRef.status]   = res.status  || '';
      row[idxRef.time]     = res.time    || '';
      row[idxRef.location] = res.location|| '';
      row[idxRef.raw]      = res.raw     || '';
      row[idxRef.at]       = fmtDateTime_(new Date());
      const deliveredAlready = deliveredI >= 0 && String(row[deliveredI]||'').trim();
      const isDeliveredNow = isDeliveredByText_(res.status || '');
      out.push(row);
      keep.push(!(removeDelivered && (deliveredAlready || isDeliveredNow)));
    } else { out.push(row); keep.push(true); }
  });

  sh.clearContents();
  sh.getRange(1,1,1,finalHdr.length).setValues([finalHdr]);
  const filtered = out.filter((_, i) => keep[i]).map(r => { const rr = r.slice(0, finalHdr.length); while (rr.length < finalHdr.length) rr.push(''); return rr; });
  if (filtered.length) sh.getRange(2,1,filtered.length, finalHdr.length).setValues(filtered);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(finalHdr.length, 20));
}