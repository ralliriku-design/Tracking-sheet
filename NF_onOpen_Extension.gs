/**
 * NF_onOpen_Extension.gs — Extension to integrate NF menu items with existing onOpen
 * 
 * This file should be included alongside the existing onOpen function to add
 * NF bulk operations menu items to the Shipment menu.
 * 
 * Usage: Call NF_extendShipmentMenu() from the main onOpen() function
 */

/**
 * Extends the existing Shipment menu with NF bulk operations.
 * Call this function after creating the main Shipment menu in onOpen().
 */
function NF_extendShipmentMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Create the bulk operations submenu
    const bulkMenu = ui.createMenu('NF Bulk Operations')
      .addItem('Import from Drive + Rebuild All', 'NF_BulkRebuildAll')
      .addItem('Refresh All Pending (100)', 'NF_menuRefreshAllPending')
      .addItem('Find Duplicates (All Sources)', 'NF_BulkFindDuplicates_All')
      .addSeparator()
      .addItem('Import from Drive Only', 'NF_BulkImportFromDrive')
      .addItem('Update Inventory Balances', 'NF_UpdateInventoryBalances')
      .addItem('Build SOK/Kärkkäinen Always', 'NF_buildSokKarkkainenAlways');
    
    // Create the scheduling submenu
    const scheduleMenu = ui.createMenu('NF Scheduling')
      .addItem('Setup Daily 00:01 (Inventory)', 'NF_setupDaily0001Inventory')
      .addItem('Setup Daily 11:00 (Tracking)', 'NF_setupDaily1100Tracking')
      .addItem('Setup Weekly Mon 02:00 (Reports)', 'NF_setupWeeklyMon0200')
      .addSeparator()
      .addItem('Clear NF Triggers', 'NF_clearAllTriggers')
      .addItem('List Active Triggers', 'NF_listActiveTriggers');
    
    // Add to the existing Shipment menu if possible
    // Note: This requires the menu to be available as a variable in onOpen
    // If not possible, these will be standalone menus
    
    bulkMenu.addToUi();
    scheduleMenu.addToUi();
    
    console.log('NF menu extensions added successfully');
    
  } catch (error) {
    console.error('Error adding NF menu extensions:', error);
    // Fallback: create standalone menus
    try {
      ui.createMenu('NF Bulk Ops')
        .addItem('Rebuild All', 'NF_BulkRebuildAll')
        .addItem('Refresh Pending', 'NF_menuRefreshAllPending')
        .addToUi();
    } catch (fallbackError) {
      console.error('Fallback menu creation also failed:', fallbackError);
    }
  }
}

/**
 * Enhanced onOpen function that integrates NF features.
 * This can replace the existing onOpen function or be called from it.
 */
function NF_enhancedOnOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create the main Shipment menu with all existing items plus NF extensions
  const shipmentMenu = ui.createMenu('Shipment')
    .addItem('Credentials Hub', 'showCredentialsHub')
    .addItem('Hae Gmail → Päivitä taulut nyt', 'runDailyFlowOnce')
    .addItem('Rakenna Vaatii_toimenpiteitä (ei-toimitetut)', 'buildPendingFromPackagesAndArchive')
    .addItem('Päivitä Vaatii_toimenpiteitä status', 'refreshStatuses_Vaatii')
    .addSeparator()
    .addItem('Asetuspaneeli (pop-in)', 'openControlPanel')
    .addSeparator()
    .addItem('Rakenna SOK & Kärkkäinen (aina)', 'buildSokKarkkainenAlways')
    .addItem('Tee viikkoraportit (SUN→SUN) + status', 'makeWeeklyReportsSunSun')
    .addItem('Arkistoi toimitetut (Vaatii_toimenpiteitä)', 'archiveDeliveredFromVaatii')
    .addItem('Aseta valinnaiset oletukset', 'seedOptionalDefaults')
    .addItem('Täytä API-avaimet & asiakasnumerot', 'seedKnownAccountsAndKeys')
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
    )
    .addSeparator()
    // NF Bulk Operations
    .addSubMenu(
      ui.createMenu('NF Bulk Operations')
        .addItem('Import from Drive + Rebuild All', 'NF_BulkRebuildAll')
        .addItem('Refresh All Pending (100)', 'NF_menuRefreshAllPending')
        .addItem('Find Duplicates (All Sources)', 'NF_BulkFindDuplicates_All')
        .addSeparator()
        .addItem('Import from Drive Only', 'NF_BulkImportFromDrive')
        .addItem('Update Inventory Balances', 'NF_UpdateInventoryBalances')
        .addItem('Build SOK/Kärkkäinen Always', 'NF_buildSokKarkkainenAlways')
    )
    .addSubMenu(
      ui.createMenu('Adhoc (Power BI / XLSX)')
        .addItem('Tuo Drive-URL/ID → Adhoc_Tracking', 'adhocImportFromDriveFile')
        .addItem('Tuo Gmailin uusin "Outbound order" -liite', 'adhocImportLatestOutboundFromGmail')
        .addItem('Päivitä statukset (Adhoc_Tracking)', 'adhocRefresh')
    );

  // Continue with remaining menu items from original...
  // (Power BI submenu, Tarkistimet, Pikatoiminnot, etc.)
  
  shipmentMenu.addToUi();
  
  // Add NF Scheduling as separate menu
  ui.createMenu('NF Scheduling')
    .addItem('Setup Daily 00:01 (Inventory)', 'NF_setupDaily0001Inventory')
    .addItem('Setup Daily 11:00 (Tracking)', 'NF_setupDaily1100Tracking')
    .addItem('Setup Weekly Mon 02:00 (Reports)', 'NF_setupWeeklyMon0200')
    .addSeparator()
    .addItem('Clear NF Triggers', 'NF_clearAllTriggers')
    .addItem('List Active Triggers', 'NF_listActiveTriggers')
    .addToUi();
}

/**
 * Simple integration approach: add this call to the end of the existing onOpen function.
 * 
 * Example integration in existing onOpen():
 * 
 * function onOpen() {
 *   // ... existing menu creation code ...
 *   .addToUi();
 *   
 *   // Add NF extensions
 *   NF_addStandaloneMenus();
 * }
 */
function NF_addStandaloneMenus() {
  const ui = SpreadsheetApp.getUi();
  
  // Add NF Bulk Operations as standalone menu
  ui.createMenu('NF Bulk Ops')
    .addItem('Rebuild All', 'NF_BulkRebuildAll')
    .addItem('Refresh Pending (100)', 'NF_menuRefreshAllPending')
    .addItem('Find Duplicates', 'NF_BulkFindDuplicates_All')
    .addSeparator()
    .addItem('Drive Import', 'NF_BulkImportFromDrive')
    .addItem('Inventory Update', 'NF_UpdateInventoryBalances')
    .addToUi();
    
  // Add NF Scheduling as standalone menu  
  ui.createMenu('NF Schedule')
    .addItem('Daily 00:01 (Inventory)', 'NF_setupDaily0001Inventory')
    .addItem('Daily 11:00 (Tracking)', 'NF_setupDaily1100Tracking')
    .addItem('Weekly Mon 02:00', 'NF_setupWeeklyMon0200')
    .addSeparator()
    .addItem('Clear Triggers', 'NF_clearAllTriggers')
    .addToUi();
}