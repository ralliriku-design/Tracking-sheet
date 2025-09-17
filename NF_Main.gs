/** NF_Main.gs — NewFlow: Menus, triggers, orchestrations (additive, no conflicts) */

function NF_onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    const root = ui.createMenu('NewFlow');
    root
      .addItem('Gmail: tuo Packages (nShift) → Rebuild', 'NF_fetchGmailPackagesAndRebuild')
      .addItem('PBI: tuo Outbound (Drive-kansiosta/staging)', 'NF_importPowerBIOutbound')
      .addItem('ERP: tuo Stock Picking', 'NF_importERPStockPicking')
      .addSeparator()
      .addItem('Rakenna Maa-kohtainen toimitusaika (Keikka tehty → Toimitettu)', 'NF_buildCountryLeadtime')
      .addItem('Rakenna SOK & Kärkkäinen -viikkoraportit', 'NF_buildSokKarkkainenWeekly')
      .addSeparator()
      .addItem('Ajasta päivit. (arkipäivisin 12:00)', 'NF_setupDaily1200')
      .addItem('Ajasta viikko (ma 02:00)', 'NF_setupWeeklyMon0200')
      .addItem('Poista kaikki NewFlow-ajastukset', 'NF_clearAllNFTriggers')
      .addSeparator()
      .addItem('Aja: Päivittäinen NewFlow', 'NF_runDaily')
      .addItem('Aja: Viikkoraportit', 'NF_runWeekly')
      .addToUi();
  } catch (e) {}
}

function NF_setupDaily1200() {
  NF_clearTriggersBy_(['NF_runDaily']);
  ScriptApp.newTrigger('NF_runDaily').timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  SpreadsheetApp.getActive().toast('NewFlow: päivittäinen ajo ajastettu klo 12:00');
}

function NF_setupWeeklyMon0200() {
  NF_clearTriggersBy_(['NF_runWeekly']);
  ScriptApp.newTrigger('NF_runWeekly').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).nearMinute(0).create();
  SpreadsheetApp.getActive().toast('NewFlow: viikkoajo ajastettu ma 02:00');
}

function NF_clearAllNFTriggers() {
  const names = new Set(['NF_runDaily','NF_runWeekly']);
  ScriptApp.getProjectTriggers().forEach(tr => { if (names.has(tr.getHandlerFunction())) ScriptApp.deleteTrigger(tr); });
  SpreadsheetApp.getActive().toast('NewFlow: ajastukset poistettu');
}

function NF_clearTriggersBy_(names) {
  const set = new Set(names);
  ScriptApp.getProjectTriggers().forEach(tr => { if (set.has(tr.getHandlerFunction())) ScriptApp.deleteTrigger(tr); });
}

/** Daily: import sources then compute metrics */
function NF_runDaily() {
  try { NF_fetchGmailPackagesAndRebuild(); } catch (e) { Logger.log('NF Gmail import failed: ' + e); }
  try { NF_importPowerBIOutbound(); } catch (e) { Logger.log('NF PBI import failed: ' + e); }
  try { NF_importERPStockPicking(); } catch (e) { Logger.log('NF ERP import failed: ' + e); }
  try { NF_buildCountryLeadtime(); } catch (e) { Logger.log('NF leadtime build failed: ' + e); }
}

/** Weekly: SOK & Kärkkäinen */
function NF_runWeekly() {
  try { NF_buildSokKarkkainenWeekly(); } catch (e) { Logger.log('NF weekly failed: ' + e); }
}

/** Wrappers against existing repo functions to avoid duplication */
function NF_fetchGmailPackagesAndRebuild() {
  // Use existing fetchAndRebuild() if available
  if (typeof fetchAndRebuild === 'function') return fetchAndRebuild();
  // Fallback: try more specific names from variants
  if (typeof gmailImportLatestPackagesReport === 'function') return gmailImportLatestPackagesReport();
  SpreadsheetApp.getUi().alert('Gmail-tuonti ei ole käytettävissä tässä ympäristössä.');
}

function NF_importPowerBIOutbound() {
  if (typeof pbiImportOutbounds_OldestFirst === 'function') return pbiImportOutbounds_OldestFirst();
  SpreadsheetApp.getUi().alert('Power BI Outbound -tuonti ei ole käytettävissä tässä ympäristössä.');
}

function NF_importERPStockPicking() {
  // Prefer existing ERP updater if present
  if (typeof runERPUpdate === 'function') return runERPUpdate();
  // Otherwise noop with info
  SpreadsheetApp.getUi().alert('ERP Stock Picking -tuonti ei ole konfiguroitu. Aseta ERP-tuonti olemassa olevilla funktioilla tai Drive-kansioilla.');
}