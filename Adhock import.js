/******************************************************
 * Adhoc: import Drive file (XLSX/CSV) to Adhoc_Tracking
 ******************************************************/

function adhocImportFromDriveFile(){
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Adhoc tuonti', 'Liitä Drive-URL tai -ID (.xlsx/.csv).', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const id = extractId(res.getResponseText());
  const file = DriveApp.getFileById(id);
  const raw  = readAttachmentToValues_(file.getBlob(), file.getName());
  buildAdhocFromValues_(raw.values, file.getName());
}

function buildAdhocFromValues_(values, label){
  if (!values || !values.length) throw new Error('Tuotu tiedosto on tyhjä.');
  const hdr = values[0].map(v => String(v||'').trim());
  const carrierI = pickAnyIndex_(hdr, CARRIER_CANDIDATES);
  const codeI    = pickAnyIndex_(hdr, TRACKING_CODE_CANDIDATES);
  if (carrierI < 0 || codeI < 0) throw new Error('Adhoc: ei löytynyt Carrier/Tracking -sarakkeita.');
  const rows = values.slice(1).filter(r => r.some(x => String(x||'').trim() !== ''));
  const out  = rows.map(r => [ r[carrierI], r[codeI] ]);
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateSheet_(ADHOC_SHEET);
  sh.clear();
  const baseHdr = ['Carrier','Tracking number'];
  const hdrAll = baseHdr.concat(['RefreshCarrier','RefreshStatus','RefreshTime','RefreshLocation','RefreshRaw','RefreshAt']);
  sh.getRange(1,1,1,hdrAll.length).setValues([hdrAll]);
  if (out.length) sh.getRange(2,1,out.length, 2).setValues(out);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(hdrAll.length, 20));
  ss.toast(`Adhoc_Tracking: tuotu ${out.length} riviä (${label})`);
}

function adhocRefresh(){
  refreshStatuses_Sheet(ADHOC_SHEET, false);
}

function extractId(urlOrId){
  const s = String(urlOrId||'').trim();
  const m = s.match(/[-\w]{25,}/);
  return m ? m[0] : s;
}