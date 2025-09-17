// gmail.gs — Gmail search, attachments, imports

function sanitizeMatrix_(matrix) {
  if (!matrix || !matrix.length) return matrix;
  let lastNonEmpty = matrix.length;
  for (let r = matrix.length - 1; r >= 0; r--) {
    if (!matrix[r].join('')) lastNonEmpty = r; else break;
  }
  return matrix.slice(0, lastNonEmpty);
}
function readAttachmentToValues_(blob, filename) {
  if (!blob) return null;
  filename = filename || blob.getName();
  try {
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function') {
      const file = Drive.Files.insert({ title: filename.replace(/\.(xlsx|csv)$/i,''), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true });
      const ss = SpreadsheetApp.openById(file.id);
      const values = ss.getSheets()[0].getDataRange().getDisplayValues();
      DriveApp.getFileById(file.id).setTrashed(true);
      return values;
    }
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.copy === 'function') {
      const up = DriveApp.createFile(blob);
      const copied = Drive.Files.copy({ title: filename.replace(/\.(xlsx|csv)$/i,''), mimeType: MimeType.GOOGLE_SHEETS }, up.getId());
      const ss = SpreadsheetApp.openById(copied.id);
      const values = ss.getSheets()[0].getDataRange().getDisplayValues();
      DriveApp.getFileById(up.getId()).setTrashed(true);
      DriveApp.getFileById(copied.id).setTrashed(true);
      return values;
    }
    return Utilities.parseCsv(blob.getDataAsString('UTF-8'));
  } catch (e) {
    SpreadsheetApp.getUi().alert('Virhe liitteen lukemisessa: ' + e);
    return null;
  }
}

const ALLOW_ATTACH_RX = new RegExp(PropertiesService.getScriptProperties().getProperty('ATTACH_ALLOW_REGEX') || '(?:^|\\b)(Packages[ _-]?Report)(?:\\b|$)', 'i');

function gmailQuery_() { return TRK_props_('GMAIL_QUERY') || 'subject:(Outbound) OR filename:(Outbound)'; }
function findLatestAttachment_() {
  const threads = GmailApp.search(gmailQuery_());
  if (!threads.length) return null;
  const msgs = GmailApp.getMessagesForThreads([threads[0]])[0];
  for (let i = msgs.length - 1; i >= 0; i--) {
    const atts = msgs[i].getAttachments();
    for (const att of atts) {
        const t0 = Date.now();
        const fname = att.getName();
        if (!ALLOW_ATTACH_RX.test(fname)) { shLog.appendRow([new Date(), 'History', 'SKIP', fname, 'Tiedoston nimi ei täsmää ALLOW_REGEXiin', ((Date.now()-t0)/1000).toFixed(2)]); continue; }
      if (ALLOW_ATTACH_RX.test(att.getName())) return { blob: att, name: att.getName() };
    }
  }
  return null;
}
function writeMerged_(sheet, existing, incoming){
  const old = existing.slice();
  const newM = incoming.slice();
  const hdrOld = old.shift() || [];
  const hdrNew = newM.shift() || [];
  const mergedHdr = mergeHeaders_(hdrOld, hdrNew);
  const mapOld = headerIndexMap_(hdrOld);
  const mapNew = headerIndexMap_(hdrNew);
  const out = [mergedHdr];
  old.forEach(row => { out.push(mergedHdr.map(col => (col in mapOld ? row[mapOld[col]] : ''))); });
  newM.forEach(row => { out.push(mergedHdr.map(col => (col in mapNew ? row[mapNew[col]] : ''))); });
  sheet.clear();
  sheet.getRange(1,1,out.length,out[0].length).setValues(out);
}
function fetchAndRebuild() {
  const attachment = findLatestAttachment_();
  if (!attachment) return;
  const values = readAttachmentToValues_(attachment.blob, attachment.name);
  if (!values) return;
  const matrix = sanitizeMatrix_(values);
  if (!matrix || matrix.length < 2) return;
  const ss = SpreadsheetApp.getActive();
  const shLog = ss.getSheetByName('ImportLog') || ss.insertSheet('ImportLog');
  if (shLog.getLastRow()===0) shLog.appendRow(['Timestamp','Step','Status','File','Reason','Duration (s)']);
  const startTs = Date.now();
  const sheetName = attachment.name.toLowerCase().includes('archive') ? ARCHIVE_SHEET : TARGET_SHEET;
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  if (sh.getLastRow() < 2) {
    sh.clear();
    sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
  } else {
    const existing = sh.getDataRange().getValues();
    writeMerged_(sh, existing, matrix); shLog.appendRow([new Date(), 'History', 'OK', fname, 'Merged', ((Date.now()-t0)/1000).toFixed(2)]);
  }
}
function fetchHistoryFromGmailOldToNew() {
  const threads = GmailApp.search(gmailQuery_());
  if (!threads.length) { SpreadsheetApp.getUi().alert('Ei löytynyt viestejä annetulla Gmail-haulla.'); return; }
  const ss = SpreadsheetApp.getActive();
  const shLog = ss.getSheetByName('ImportLog') || ss.insertSheet('ImportLog');
  if (shLog.getLastRow()===0) shLog.appendRow(['Timestamp','Step','Status','File','Reason','Duration (s)']);
  const startTs = Date.now();
  let sheetPack = ss.getSheetByName(TARGET_SHEET) || ss.insertSheet(TARGET_SHEET);
  let sheetArch = ss.getSheetByName(ARCHIVE_SHEET) || ss.insertSheet(ARCHIVE_SHEET);
  sheetPack.clear(); sheetArch.clear();
  let firstPack = true, firstArch = true;
  for (let ti = 0; ti < threads.length; ti++) {
    const msgs = GmailApp.getMessagesForThread(threads[ti]);
    for (let mi = 0; mi < msgs.length; mi++) {
      const atts = msgs[mi].getAttachments();
      for (const att of atts) {
        const t0 = Date.now();
        const fname = att.getName();
        if (!ALLOW_ATTACH_RX.test(fname)) { shLog.appendRow([new Date(), 'History', 'SKIP', fname, 'Tiedoston nimi ei täsmää ALLOW_REGEXiin', ((Date.now()-t0)/1000).toFixed(2)]); continue; }
        const name = att.getName();
        if (/Outbound.*\.(xlsx|csv)/i.test(name)) {
          const values = readAttachmentToValues_(att, name);
          if (!values) { shLog.appendRow([new Date(), 'History', 'FAIL', fname, 'Liite ei konvertoitunut (Drive/CSV)', ((Date.now()-t0)/1000).toFixed(2)]); continue; }
          const matrix = sanitizeMatrix_(values);
          if (!matrix || matrix.length < 2) { shLog.appendRow([new Date(), 'History', 'FAIL', fname, 'Tyhjä tai kelvoton taulukko', ((Date.now()-t0)/1000).toFixed(2)]); continue; }
          const isArchive = name.toLowerCase().includes('archive');
          const sh = isArchive ? sheetArch : sheetPack;
          if ((isArchive && firstArch) || (!isArchive && firstPack)) { shLog.appendRow([new Date(), 'History', 'OK', fname, (isArchive?'Init ARCHIVE':'Init PACKAGES'), ((Date.now()-t0)/1000).toFixed(2)]);
            sh.getRange(1,1,matrix.length,matrix[0].length).setValues(matrix);
            if (isArchive) firstArch = false; else firstPack = false;
          } else {
            const existing = sh.getDataRange().getValues();
            writeMerged_(sh, existing, matrix); shLog.appendRow([new Date(), 'History', 'OK', fname, 'Merged', ((Date.now()-t0)/1000).toFixed(2)]);
          }
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('Gmail-historia tuotu ja yhdistetty (Packages & Packages_Archive).');
}
