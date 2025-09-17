
// 06_format_and_dedupe.gs
function fixFormatsAll_(){
  const ss = SpreadsheetApp.getActive();
  const sheets = ['Vaatii_toimenpiteitä','Packages','Packages_Archive'];
  const datePattern = 'yyyy-mm-dd hh:mm';
  sheets.forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const rg = sh.getDataRange(); const values = rg.getValues(); if (!values || !values.length) return;
    const hdr = values[0]; const map = headerIndexMap_(hdr);
    let tIdx = (map['Package Number'] ?? map['Tracking Number'] ?? 0);
    let cIdx = (map['RefreshCarrier'] ?? map['Carrier'] ?? -1);
    if (cIdx < 0){
      const def = Number(PropertiesService.getScriptProperties().getProperty('CARRIER_DEFAULT_COL')||'25');
      if (def >= 1 && def <= hdr.length) cIdx = def - 1;
    }
    const timeI = map['RefreshTime'] ?? -1;
    if (tIdx >= 0) sh.getRange(2, tIdx+1, Math.max(0, sh.getLastRow()-1), 1).setNumberFormat('@');
    if (cIdx >= 0) sh.getRange(2, cIdx+1, Math.max(0, sh.getLastRow()-1), 1).setNumberFormat('@');
    if (timeI >= 0) sh.getRange(2, timeI+1, Math.max(0, sh.getLastRow()-1), 1).setNumberFormat(datePattern);
    ['RefreshStatus','RefreshLocation','RefreshRaw'].forEach(k => {
      const i = map[k]; if (i>=0){ const rgc = sh.getRange(2, i+1, Math.max(0, sh.getLastRow()-1), 1); rgc.setWrap(true).setNumberFormat('@'); }
    });
  });
  SpreadsheetApp.getUi().alert('Formaattikorjaukset tehty.');
}

function checkArchiveDuplicates(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(ARCHIVE_SHEET);
  if (!sh){ SpreadsheetApp.getUi().alert('Packages_Archive ei löydy.'); return; }
  const rg = sh.getDataRange(); const v = rg.getValues(); if (v.length<2){ SpreadsheetApp.getUi().alert('Ei dataa.'); return; }
  const hdr = v[0], map = headerIndexMap_(hdr);
  const codeI = map['Package Number'] ?? map['Tracking Number'] ?? 0;
  const carI = map['RefreshCarrier'] ?? map['Carrier'] ?? -1;
  const joinWithCarrier = String(TRK_props_('DEDUP_JOIN_CARRIER')||'false').toLowerCase() === 'true';
  const seen = {}; const dups = [];
  for (let r=1;r<v.length;r++){
    const row = v[r];
    const key = (joinWithCarrier ? (row[codeI]+'|'+(row[carI]||'')) : row[codeI]);
    if (!key) continue;
    if (seen[key]) dups.push(r+1); else seen[key] = true;
  }
  const hl = sh.getRangeList(dups.map(r=> 'A'+r+':'+sh.getLastColumn()+r));
  if (hl) hl.activate().setBackground('#fff3cd');
  SpreadsheetApp.getUi().alert('Mahdollisia duplikaattirivejä: ' + dups.length);
}
