// powerbi.gs — Power BI helpers

function powerBiArchiveNew() {
  const ss = SpreadsheetApp.getActive();
  const shNew = ss.getSheetByName(PBI_NEW_SHEET);
  const shArch = ss.getSheetByName(ARCHIVE_SHEET);
  if (!shNew || !shArch) { SpreadsheetApp.getUi().alert('PBI_New tai Packages_Archive -taulua ei löydy.'); return; }
  const data = shNew.getDataRange().getValues();
  if (data.length > 1) { for (let r=1; r<data.length; r++) shArch.appendRow(data[r]); }
  shNew.clearContents();
  SpreadsheetApp.getUi().alert('PBI-uudet toimitukset arkistoitu.');
}
function triggerPowerBI() {
  const url = PropertiesService.getScriptProperties().getProperty('PBI_WEBHOOK_URL');
  if (!url) { SpreadsheetApp.getUi().alert('Power BI -webhook URL puuttuu asetuksista (PBI_WEBHOOK_URL)'); return; }
  try { UrlFetchApp.fetch(url, { muteHttpExceptions:true }); SpreadsheetApp.getUi().alert('Power BI -päivitys kutsuttu.'); }
  catch(err) { SpreadsheetApp.getUi().alert('Virhe Power BI -webhook-kutsussa: ' + err); }
}
