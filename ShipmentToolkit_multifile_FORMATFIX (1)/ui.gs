// ui.gs — menu, creds UI, seeds, progress, triggers

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipment')
    .addItem('Credentials Hub', 'showCredentialsHub')
    .addItem('Hae Gmail → Päivitä taulut nyt', 'runDailyFlowOnce')
    .addItem('Rakenna Vaatii_toimenpiteitä (ei-toimitetut)', 'buildPendingFromPackagesAndArchive')
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Päivitä STATUS (valittu yhtiö)')
        .addItem('Matkahuolto', 'menuRefreshCarrier_MH')
        .addItem('Posti', 'menuRefreshCarrier_POSTI')
        .addItem('Bring', 'menuRefreshCarrier_BRING')
        .addItem('GLS', 'menuRefreshCarrier_GLS')
        .addItem('DHL', 'menuRefreshCarrier_DHL')
        .addItem('Kaikki (ei suodatusta)', 'menuRefreshCarrier_ALL')
    )
    .addSubMenu(
      ui.createMenu('Iso statusajo (turvallinen)')
        .addItem('Aloita (aktiivinen välilehti)', 'bulkStartForActiveSheet')
        .addItem('Aloita: Vaatii_toimenpiteitä', 'bulkStart_Vaatii')
        .addItem('Aloita: Packages', 'bulkStart_Packages')
        .addItem('Aloita: Packages_Archive', 'bulkStart_Archive')
        .addItem('Pysäytä bulk-ajo', 'bulkStop')
        .addItem('Näytä edistyminen', 'showProgressSidebar')
    )
    .addSeparator()
    .addItem('Gmail: tuo KAIKKI (vanhin → uusin)', 'fetchHistoryFromGmailOldToNew')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Power BI')
        .addItem('Arkistoi PBI-uudet → Packages_Archive', 'powerBiArchiveNew')
        .addItem('Käynnistä Power BI -päivitys (webhook)', 'triggerPowerBI')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Tarkistimet')
        .addItem('Diagnosoi statuspäivitys (Vaatii_toimenpiteitä)', 'diagnoseRefresh_Vaatii')
        .addItem('Korjaa formaatit (Kaikki taulut)', 'fixFormats_All')
        .addItem('Korosta tunnistetut sarakkeet (Vaatii_toimenpiteitä)', 'highlightDetectedColumns_Vaatii')
        .addItem('Esikatsele päivitettävät rivit (Vaatii_toimenpiteitä)', 'previewEligible_Vaatii')
        .addItem('Puuttuvat sarakeotsikot (aktiivinen välilehti)', 'showMissingColumns_Active')
        .addItem('Integraatioavaimet (Script Properties)', 'showMissingProperties')
        .addItem('Avaa ErrorLog', 'openErrorLog')
        .addItem('Tyhjennä ErrorLog', 'clearErrorLog')
        .addItem('Avaa RunStats', 'openRunStats')
    )
    .addSeparator()
    .addItem('Ajasta: arkipäivät klo 12:00', 'setupWeekday1200_DailyFlow')
    .addItem('Ajasta: viikkoraportti maanantai 02:00', 'setupWeeklyMon0200')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Pikatoiminnot / Vaatii')
        .addItem('Vahvista toimitukset (Vaatii)', 'confirmMissingDeliveredTimes')
        .addItem('Vahvista + Arkistoi (Vaatii)', 'confirmAndArchiveDelivered')
    )
    .addSeparator()
    .addItem('Tyhjennä seurannan välimuisti', 'trkClearCache')
    .addItem('Poista kaikki ajastukset', 'clearAllTriggers')
    .addToUi();
}
function onInstall() { onOpen(); }

/* ===== Creds Panel ===== */
function credGetProps() {
  const sp = PropertiesService.getScriptProperties();
  const keys = [
'MH_TRACK_URL','MH_BASIC',
    'POSTI_TOKEN_URL','POSTI_BASIC','POSTI_TRACK_URL','POSTI_TRK_URL','POSTI_TRK_BASIC','POSTI_TRK_USER','POSTI_TRK_PASS',
    'GLS_FI_TRACK_URL','GLS_FI_API_KEY','GLS_FI_SENDER_ID','GLS_FI_SENDER_IDS','GLS_FI_METHOD',
    'GLS_TOKEN_URL','GLS_BASIC','GLS_TRACK_URL',
    'DHL_TRACK_URL','DHL_API_KEY',
    'BRING_TRACK_URL','BRING_UID','BRING_KEY','BRING_CLIENT_URL',
    'BULK_BACKOFF_MINUTES_BASE',
    'PBI_WEBHOOK_URL',
    'GMAIL_QUERY',
    'TRACKING_HEADER_HINT','TRACKING_DEFAULT_COL','ATTACH_ALLOW_REGEX','REFRESH_MIN_AGE_MINUTES','CARRIER_HEADER_HINT'
];
  const props = {};
  keys.forEach(k => props[k] = sp.getProperty(k) || '');
  return props;
}
function credSaveProps(data) {
  const sp = PropertiesService.getScriptProperties();
  try {
    Object.keys(data).forEach(k => {
      if (data[k] !== '') sp.setProperty(k, String(data[k]));
      else sp.deleteProperty(k);
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: String(e) };
  }
}
function showCredentialsHub() {
  try { openCredsPanel(); } catch (e) { openControlPanel(); }
}
function openCredsPanel() {
  const t = HtmlService.createTemplateFromFile('Creds');
  t.props = credGetProps();
  const html = t.evaluate().setTitle('Integraatioavaimet').setWidth(760).setHeight(760);
  SpreadsheetApp.getUi().showSidebar(html);
}
function openControlPanel() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html><meta charset="utf-8"><title>Shipment – Asetuspaneeli</title>
<style>body{font:14px system-ui;margin:16px;color:#111}
h1{font-size:18px;margin:0 0 10px}.row{display:grid;grid-template-columns:190px 1fr;gap:8px;align-items:center;margin:6px 0}
input{width:100%}button{padding:6px 12px;margin:8px 8px 0 0}.info{font-size:12px;color:#555}</style>
<body>
  <h1>Asetuspaneeli (fallback)</h1>
  <div class="row"><span>GMAIL_QUERY</span><input id="gq" type="text" placeholder="label:ShipmentReports OR subject:(Outbound)"></div>
  <p class="info">Jätä kenttä tyhjäksi poistaaksesi asetuksen.</p>
  <button onclick="seed()">Täytä oletukset</button>
  <button onclick="flush()">Tyhjennä välimuisti</button>
  <script>
    google.script.run.withSuccessHandler(props => { document.getElementById('gq').value = props.GMAIL_QUERY || ''; }).credGetProps();
    function seed(){google.script.run.seedKnownAccountsAndKeys();}
    function flush(){google.script.run.trkClearCache();}
    window.onbeforeunload=()=>google.script.run.credSaveProps({GMAIL_QUERY:document.getElementById('gq').value});
  </script>
</body>`).setWidth(640).setHeight(320).setTitle('Asetuspaneeli');
  SpreadsheetApp.getUi().showModalDialog(html, 'Asetuspaneeli');
}
function seedKnownAccountsAndKeys() {
  const sp = PropertiesService.getScriptProperties();
  const setIfEmpty = (k,v) => { if (!sp.getProperty(k)) sp.setProperty(k, v); };
  setIfEmpty('GMAIL_QUERY','label:ShipmentReports OR subject:(Outbound) OR filename:(Outbound)');
  setIfEmpty('GLS_FI_API_KEY','<GLS_FI_API_KEY>');
  setIfEmpty('GLS_FI_TRACK_URL','https://<gls-fi-host>/customerapi/v2/get_tracking_events');
  setIfEmpty('GLS_FI_SENDER_ID','<GLS_SENDER_ID>');
  setIfEmpty('GLS_FI_SENDER_IDS','');
  setIfEmpty('GLS_FI_METHOD','POST');
  setIfEmpty('GLS_TOKEN_URL','https://api.gls-group.net/oauth2/v1/token');
  setIfEmpty('GLS_BASIC','<base64_clientId:secret>');
  setIfEmpty('GLS_TRACK_URL','https://api.gls-group.net/track-and-trace-v1/tracking/simple/references/{{code}}');
  setIfEmpty('POSTI_TOKEN_URL','https://oauth2.posti.com/oauth/token');
  setIfEmpty('POSTI_BASIC','<base64_clientId:secret>');
  setIfEmpty('POSTI_TRACK_URL','https://api.posti.fi/tracking/7/shipments/trackingnumbers/{{code}}');
  setIfEmpty('POSTI_TRK_URL','https://atlas.posti.fi/track-shipment-json?ShipmentId={{code}}');
  setIfEmpty('POSTI_TRK_BASIC','<base64_user:pass>');
  setIfEmpty('POSTI_TRK_USER','<ma_contract_user>');
  setIfEmpty('POSTI_TRK_PASS','<ma_contract_pass>');
  setIfEmpty('MH_TRACK_URL','https://extservices.matkahuolto.fi/mpaketti/public/tracking?ids={{code}}');
  setIfEmpty('MH_BASIC','<base64_user:pass>');
  setIfEmpty('DHL_TRACK_URL','https://api-eu.dhl.com/track/shipments?trackingNumber={{code}}');
  setIfEmpty('DHL_API_KEY','<DHL_API_KEY>');
  setIfEmpty('BRING_TRACK_URL','https://api.bring.com/tracking/api/v2/tracking.json?q={{code}}');
  setIfEmpty('BRING_UID','<BRING_API_UID>');
  setIfEmpty('BRING_KEY','<BRING_API_KEY>');
  setIfEmpty('BRING_CLIENT_URL','https://<yourapp>.domain.com');
  setIfEmpty('BULK_BACKOFF_MINUTES_BASE','30');
  setIfEmpty('PBI_WEBHOOK_URL','https://<oma-webhook-url>');
  SpreadsheetApp.getUi().alert('Oletusavaimet täytetty Script Properties -varastoon.');
}
function showProgressSidebar(){
  const sp = PropertiesService.getScriptProperties();
  const props = sp.getProperties();
  const jobs = Object.keys(props).filter(k => k.indexOf('BULKJOB|') === 0).map(k => JSON.parse(props[k]));
  const html = HtmlService.createHtmlOutput(
    '<h3 style="font-family:system-ui;margin:10px 0">Bulk-tila</h3>' +
    (jobs.length ? jobs.map(j => {
      const done = Math.max(0, Math.min((j.row||2)-2, j.total||0));
      const pct = j.total ? Math.round((done / j.total) * 100) : 0;
      const s = j.stats || {};
      return `<div style="font:13px system-ui;margin:6px 0"><b>${j.sheet}</b> — ${done}/${j.total} (${pct}%)<br>Success: ${s.success||0} / Tried: ${s.tried||0}</div>`;
    }).join('') : '<div style="font:13px system-ui">Ei käynnissä olevia ajoja.</div>')
  ).setTitle('Bulk-tila').setWidth(340).setHeight(220);
  SpreadsheetApp.getUi().showSidebar(html);
}
function openRunStats(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('RunStats');
  if (sh) ss.setActiveSheet(sh);
  else SpreadsheetApp.getUi().alert('RunStats-arkkia ei ole vielä luotu. Aja statuspäivitys ensin.');
}

/* ===== Triggers & orchestration ===== */
function runDailyFlowOnce(){
  const day = (new Date()).getDay(); // 0=Sun,6=Sat
  if (day === 0 || day === 6) return;
  try { fetchAndRebuild(); } catch(e){ logError_('runDailyFlowOnce.fetchAndRebuild', e); }
  try { buildPendingFromPackagesAndArchive(); } catch(e){ logError_('runDailyFlowOnce.buildPending', e); }
  try { refreshStatuses_Vaatii(); } catch(e){ logError_('runDailyFlowOnce.refreshVaatii', e); }
}
function setupWeekday1200_DailyFlow() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'runDailyFlowOnce') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('runDailyFlowOnce').timeBased().everyDays(1).atHour(12).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: päivittäinen ajo klo 12:00 (skripti ohittaa viikonloput).');
}
function setupWeeklyMon0200() {
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'makeWeeklyReportsSunSun') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('makeWeeklyReportsSunSun').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(2).nearMinute(0).create();
  SpreadsheetApp.getUi().alert('Ajastettu: viikkoraportti maanantaisin 02:00 (stub).');
}
function clearAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(tr => ScriptApp.deleteTrigger(tr));
  SpreadsheetApp.getUi().alert('Kaikki ajastukset poistettu.');
}

function diagnoseRefresh_Vaatii(){ diagnoseRefresh_(ACTION_SHEET); }

function highlightDetectedColumns_Vaatii(){ highlightDetectedColumns_(ACTION_SHEET); }
function previewEligible_Vaatii(){ previewEligible_(ACTION_SHEET); }
