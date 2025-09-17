/******************************************************
 * Shipment Tracking Toolkit (Clean Build)
 ******************************************************/

const TARGET_SHEET        = 'Packages';
const ARCHIVE_SHEET       = 'Packages_Archive';
const ACTION_SHEET        = 'Vaatii_toimenpiteitä';
const RUN_LOG_SHEET       = 'Run_All_Log';
const ADHOC_SHEET         = 'Adhoc_Tracking';

const GMAIL_QUERY     = 'label:"Shipment Report" newer_than:60d has:attachment (filename:xlsx OR filename:csv)';
const GMAIL_QUERY_ALL = 'label:"Shipment Report" has:attachment (filename:xlsx OR filename:csv)';

const PRIORITY_CARRIERS = ['posti','gls'];
const EXTRA_PRIORITY_CALLS = 150;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipment')
    .addSubMenu(
      ui.createMenu('Credentials')
        .addItem('Open Credentials Sidebar', 'showCredentialsSidebar')
        .addItem('Fill defaults', 'seedKnownAccountsAndKeys')
        .addItem('Show missing properties', 'showMissingProperties')
    )
    .addItem('Hae Gmail → Päivitä taulut nyt', 'runDailyFlowOnce')
    .addItem('Rakenna Vaatii_toimenpiteitä', 'buildPendingFromPackagesAndArchive')
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Päivitä STATUS (valittu)')
        .addItem('Posti', 'menuRefreshCarrier_POSTI')
        .addItem('GLS', 'menuRefreshCarrier_GLS')
        .addItem('DHL', 'menuRefreshCarrier_DHL')
        .addItem('Bring', 'menuRefreshCarrier_BRING')
        .addItem('Matkahuolto', 'menuRefreshCarrier_MH')
        .addItem('Kaikki (aktiivinen)', 'menuRefreshCarrier_ALL')
    )
    .addSubMenu(
      ui.createMenu('Iso statusajo (turvallinen)')
        .addItem('Aloita (aktiivinen välilehti)', 'bulkStartForActiveSheet')
        .addItem('Aloita: Vaatii_toimenpiteitä', 'bulkStart_Vaatii')
        .addItem('Aloita: Packages', 'bulkStart_Packages')
        .addItem('Aloita: Packages_Archive', 'bulkStart_Archive')
        .addItem('Pysäytä', 'bulkStop')
    )
    .addSeparator()
    .addItem('Adhoc: tuo Drive-URL/ID → Adhoc_Tracking', 'adhocImportFromDriveFile')
    .addItem('Adhoc: Päivitä statukset', 'adhocRefresh')
    .addToUi();
}

/** Full daily flow: Import latest → rebuild → Vaatii → refresh */
function runDailyFlowOnce() {
  const ss = SpreadsheetApp.getActive();
  const log = getOrCreateSheet_(RUN_LOG_SHEET);
  writeHeaderOnce_(log, ['Step','Status','Message','Rows/Info','Duration (s)']);
  log.clearContents();
  log.getRange(1,1,1,5).setValues([['Step','Status','Message','Rows/Info','Duration (s)']]);

  const steps = [
    { name: 'Gmail → Packages/Archive', fn: fetchAndRebuild, info: () => {
      const p = ss.getSheetByName(TARGET_SHEET), a = ss.getSheetByName(ARCHIVE_SHEET);
      return `Packages:${p ? Math.max(0,p.getLastRow()-1) : 0} Archive:${a ? Math.max(0,a.getLastRow()-1) : 0}`;
    }},
    { name: 'Rakenna Vaatii_toimenpiteitä', fn: buildPendingFromPackagesAndArchive, info: () => {
      const s = ss.getSheetByName(ACTION_SHEET);
      return `Rows:${s ? Math.max(0,s.getLastRow()-1) : 0}`;
    }},
    { name: 'Päivitä status: Vaatii_toimenpiteitä', fn: refreshStatuses_Vaatii, info: () => 'OK' },
    { name: 'Arkiston duplikaatit (nopea)', fn: checkArchiveDuplicates, info: () => 'OK' }
  ];

  const out = [];
  for (const step of steps) {
    const t0 = Date.now();
    let st = 'OK', msg = '';
    try { step.fn(); } catch (e) { st = 'FAIL'; msg = String(e && e.message || e); }
    const dur = ((Date.now() - t0)/1000).toFixed(2);
    out.push([step.name, st, msg, step.info ? step.info() : '', dur]);
  }
  log.getRange(2,1,out.length,5).setValues(out);
  ss.setActiveSheet(log);
}

/** Credentials UI */
function showCredentialsSidebar() {
  const t = HtmlService.createTemplateFromFile('UI_Credentials');
  t.props = credGetProps();
  const html = t.evaluate().setTitle('Credentials').setWidth(760);
  SpreadsheetApp.getUi().showSidebar(html);
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getActiveSheetName_() {
  return SpreadsheetApp.getActive().getActiveSheet().getName();
}

/** Menu helpers */
function menuRefreshCarrier_POSTI(){ refreshStatuses_Filtered(getActiveSheetName_(), ['posti'], false); }
function menuRefreshCarrier_GLS(){ refreshStatuses_Filtered(getActiveSheetName_(), ['gls'], false); }
function menuRefreshCarrier_DHL(){ refreshStatuses_Filtered(getActiveSheetName_(), ['dhl'], false); }
function menuRefreshCarrier_BRING(){ refreshStatuses_Filtered(getActiveSheetName_(), ['bring'], false); }
function menuRefreshCarrier_MH(){ refreshStatuses_Filtered(getActiveSheetName_(), ['matkahuolto'], false); }
function menuRefreshCarrier_ALL(){ refreshStatuses_Sheet(getActiveSheetName_(), false); }