
// 01_bootstrap_menu.gs
function onInstall(){ onOpen(); }
function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shipment')
    .addItem('Credentials Hub', 'showCredentialsHub')
    .addSeparator()
    .addItem('Esikatsele päivitettävät (Vaatii_)', 'diagnoseRefresh_Vaatii')
    .addItem('Korjaa formaatit (Kaikki taulut)', 'fixFormatsAll_')
    .addSeparator()
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
    .addItem('Pakota päivitys (Vaatii_toimenpiteitä)', 'refreshStatuses_Vaatii_FORCE')
    .addItem('Pakota päivitys + mittarit', 'refreshStatuses_Vaatii_FORCE_METRICS')
    .addSeparator()
    .addItem('Autodetect carrier + status (Vaatii_)', 'autodetectAndRefresh_Vaatii')
    .addSeparator()
    .addItem('Avaa ErrorLog', 'openErrorLog')
    .addToUi();
}
