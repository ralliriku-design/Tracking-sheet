/******************************************************
 * Gmail → Read latest attachment (XLSX/CSV)
 * XLSX is converted via Drive Advanced Service.
 ******************************************************/

function findLatestAttachment_(){
  const threads = GmailApp.search(GMAIL_QUERY, 0, 50);
  let best = null;
  for (const th of threads){
    for (const msg of th.getMessages().reverse()){
      const atts = msg.getAttachments({includeInlineImages:false, includeAttachments:true}) || [];
      for (const a of atts){
        const n = (a.getName() || '').toLowerCase();
        if (!(n.endsWith('.xlsx') || n.endsWith('.csv'))) continue;
        if (!best || msg.getDate() > best.date) {
          best = { blob: a.copyBlob(), filename: a.getName(), date: msg.getDate() };
        }
      }
    }
  }
  return best;
}

function readAttachmentToValues_(blob, filename){
  const name = String(filename||'').toLowerCase();
  if (name.endsWith('.csv')){
    const csv = Utilities.parseCsv(blob.getDataAsString());
    return {values: sanitizeMatrix_(csv)};
  }
  if (name.endsWith('.xlsx')){
    // Convert to Google Sheet and read first sheet
    let tempId = null;
    try {
      const file = Drive.Files.insert(
        { title: filename, mimeType: 'application/vnd.google-apps.spreadsheet' },
        blob,
        { convert: true }
      );
      tempId = file.id;
      const tmpSs = SpreadsheetApp.openById(tempId);
      const sh = tmpSs.getSheets()[0];
      const values = sh.getDataRange().getValues();
      return {values: sanitizeMatrix_(values)};
    } finally {
      if (tempId) try { Drive.Files.remove(tempId); } catch(e){}
    }
  }
  throw new Error('Unsupported file type: ' + filename);
}

function fetchAndRebuild(){
  const att = findLatestAttachment_();
  if (!att) throw new Error('Ei löytynyt liitteitä labelilla "Shipment Report".');
  const raw = readAttachmentToValues_(att.blob, att.filename);
  const values = sanitizeMatrix_(raw.values);
  rebuildWithArchive_(values);
  SpreadsheetApp.getActive().toast('Raportti haettu ja yhdistetty ✅');
}