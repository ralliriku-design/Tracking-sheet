/******************************************************
 * Rebuild Packages from latest import and archive removed
 ******************************************************/

function rebuildWithArchive_(values){
  if (!values || !values.length) throw new Error('Raportti on tyhjä.');
  const srcHdr = values[0].map(h => String(h||'').trim());
  const keyIdx = chooseKeyIndex_(srcHdr);
  if (keyIdx < 0) throw new Error('Yksilöivää avainsaraketta ei löytynyt.');

  const ss  = SpreadsheetApp.getActive();
  const sh  = getOrCreateSheet_(TARGET_SHEET);
  const shA = getOrCreateSheet_(ARCHIVE_SHEET);

  const oldHdr = (sh.getLastRow() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : []).map(v => String(v||''));
  const unionHdr = mergeHeaders_(oldHdr, srcHdr);
  if (oldHdr.join('|') !== unionHdr.join('|')) {
    sh.clear();
    sh.getRange(1,1,1,unionHdr.length).setValues([unionHdr]);
  }

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  let dstData = lastRow > 1 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
  const keyName = srcHdr[keyIdx];
  const dstKeyI = unionHdr.indexOf(keyName);
  const dstIdx  = new Map();
  for (let i = 0; i < dstData.length; i++){
    const k = String(dstData[i][dstKeyI] || '').trim();
    if (k) dstIdx.set(k, i);
  }
  const srcIdxMap = headerIndexMap_(srcHdr);
  const dstIdxMap = headerIndexMap_(unionHdr);

  const srcRows   = values.slice(1).filter(r => r.some(x => String(x||'').trim() !== ''));
  const srcMap    = new Map();
  for (const r of srcRows){
    const key = String(r[keyIdx]||'').trim();
    if (!key) continue;
    const row = Array(unionHdr.length).fill('');
    for (const [name, si] of Object.entries(srcIdxMap)){
      const di = dstIdxMap[name];
      if (typeof di === 'number') row[di] = r[si];
    }
    srcMap.set(key, row);
  }

  const now = new Date(), batchId = 'SRPT_'+now.getTime();
  const archHdr = ensureArchiveHeader_(shA, unionHdr);
  const toArchive = [];

  for (const [key, i] of dstIdx.entries()){
    if (!srcMap.has(key)) {
      toArchive.push(dstData[i]);
      dstData[i] = null;
    } else {
      const upd = srcMap.get(key);
      const cur = dstData[i] || Array(unionHdr.length).fill('');
      for (const [name, si] of Object.entries(srcIdxMap)){
        const di = dstIdxMap[name];
        cur[di] = upd[di];
      }
      dstData[i] = cur;
      srcMap.delete(key);
    }
  }

  if (toArchive.length) {
    const payload = toArchive.filter(r => r).map(r => r.concat([now, batchId, 'not in latest file']));
    shA.insertRowsAfter(1, payload.length);
    shA.getRange(2,1,payload.length, archHdr.length).setValues(payload);
  }

  dstData = dstData.filter(r => r !== null);
  const newRows = Array.from(srcMap.values());
  if (newRows.length) {
    newRows.reverse();
    sh.insertRowsAfter(1, newRows.length);
    sh.getRange(2,1,newRows.length, unionHdr.length).setValues(newRows);
  }

  if (dstData.length){
    const start = 2 + newRows.length;
    if (sh.getMaxRows() < start + dstData.length - 1){
      sh.insertRowsAfter(sh.getMaxRows(), start + dstData.length - 1 - sh.getMaxRows());
    }
    sh.getRange(start,1,dstData.length, unionHdr.length).setValues(dstData);
    const should = 1 + newRows.length + dstData.length;
    const extra = sh.getLastRow() - should;
    if (extra > 0) sh.deleteRows(should+1, extra);
  }

  applyFormats_(sh, unionHdr);
}

function ensureArchiveHeader_(shA, unionHdr){
  const want = unionHdr.concat(['ArchivedOn','BatchId','Reason']);
  const have = (shA.getLastRow() ? shA.getRange(1,1,1,shA.getLastColumn()).getValues()[0] : []).map(v => String(v||''));
  if (have.join('|') !== want.join('|')) {
    shA.clear();
    shA.getRange(1,1,1,want.length).setValues([want]);
  }
  return want;
}

function applyFormats_(sh, headers){
  const nrm = headers.map(normalize_);
  const idCols = headers.map((h,i) => KEY_CANDIDATES.flat().map(normalize_).includes(nrm[i]) ? i+1 : 0).filter(c => c>0);
  const dateCols = headers.map((h,i) => DATE_FIELDS.map(normalize_).includes(nrm[i]) ? i+1 : 0).filter(c => c>0);
  const lastRow  = sh.getLastRow();
  if (lastRow < 2) return;
  idCols.forEach(c   => sh.getRange(2,c,lastRow-1,1).setNumberFormat('@STRING@'));
  dateCols.forEach(c => sh.getRange(2,c,lastRow-1,1).setNumberFormat('yyyy-mm-dd hh:mm:ss'));
}

function checkArchiveDuplicates(){
  const ss  = SpreadsheetApp.getActive();
  const shA = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shA || shA.getLastRow() < 2) return;
  const data = shA.getDataRange().getDisplayValues();
  const hdr  = data[0].map(String);
  const keyI = chooseKeyIndex_(hdr);
  if (keyI < 0) return;

  const seen = new Map();
  const dups = [];
  for (let r = 1; r < data.length; r++){
    const k = String(data[r][keyI]||'').trim();
    if (!k) continue;
    if (seen.has(k)) dups.push([k, r+1, seen.get(k)]);
    else seen.set(k, r+1);
  }

  const log = getOrCreateSheet_(RUN_LOG_SHEET);
  writeHeaderOnce_(log, ['Step','Status','Message','Rows/Info','Duration (s)']);
  log.insertRowsAfter(1, 1);
  log.getRange(2,1,1,5).setValues([[
    'Archive duplicate check', 'OK', dups.length ? 'DUPLICATES FOUND' : 'None', `dups:${dups.length}`, ''
  ]]);

  if (dups.length){
    const name = 'Archive_Duplicates';
    const sh = getOrCreateSheet_(name);
    sh.clear();
    sh.getRange(1,1,1,3).setValues([['Key','Row','FirstRow']]);
    sh.getRange(2,1,dups.length,3).setValues(dups);
  }
}