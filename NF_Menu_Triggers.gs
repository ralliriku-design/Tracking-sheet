/**
 * NF_Menu_Triggers.gs — Menu additions and trigger scheduling for bulk operations
 * 
 * Extends the existing menu system with bulk import and refresh capabilities.
 * Updates trigger schedules for automated daily and weekly operations.
 * 
 * New Menu Items:
 *   - Bulk: Import from Drive + Rebuild All
 *   - Bulk: Refresh All Pending (100)
 *   - Bulk: Find Duplicates (All Sources)
 * 
 * Updated Schedules:
 *   - Daily 00:01: Inventory/balances update via Drive imports
 *   - Daily 11:00: Gmail import and tracking refresh + SOK/KRK reports
 *   - Weekly Mon 02:00: Build weekly SOK/KRK reports
 */

/**
 * Extends the existing onOpen menu with bulk operation items.
 * This function should be called from the main onOpen() function.
 */
function NF_addBulkMenuItems() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Add bulk operations submenu
    const bulkMenu = ui.createMenu('Bulk Operations')
      .addItem('Import from Drive + Rebuild All', 'NF_BulkRebuildAll')
      .addItem('Refresh All Pending (100)', 'NF_menuRefreshAllPending')
      .addItem('Find Duplicates (All Sources)', 'NF_BulkFindDuplicates_All')
      .addSeparator()
      .addItem('Import from Drive Only', 'NF_BulkImportFromDrive')
      .addItem('Update Inventory Balances', 'NF_UpdateInventoryBalances')
      .addItem('Build SOK/Kärkkäinen Always', 'NF_buildSokKarkkainenAlways');
    
    // Add to existing Shipment menu (if it exists)
    const shipmentMenu = ui.getMenu('Shipment');
    if (shipmentMenu) {
      shipmentMenu.addSeparator().addSubMenu(bulkMenu);
    } else {
      // Create standalone menu if Shipment menu doesn't exist
      bulkMenu.addToUi();
    }
    
    // Add scheduling menu
    ui.createMenu('NF Scheduling')
      .addItem('Setup Daily 00:01 (Inventory)', 'NF_setupDaily0001Inventory')
      .addItem('Setup Daily 11:00 (Tracking)', 'NF_setupDaily1100Tracking')
      .addItem('Setup Weekly Mon 02:00 (Reports)', 'NF_setupWeeklyMon0200')
      .addSeparator()
      .addItem('Clear NF Triggers', 'NF_clearAllTriggers')
      .addItem('List Active Triggers', 'NF_listActiveTriggers')
      .addToUi();
      
    console.log('NF bulk menu items added successfully');
    
  } catch (error) {
    console.error('Error adding NF menu items:', error);
  }
}

/**
 * Menu wrapper for NF_RefreshAllPending with default limit.
 */
function NF_menuRefreshAllPending() {
  try {
    const result = NF_RefreshAllPending(100);
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Bulk Refresh Complete',
      `Updated ${result.totalUpdated} rows with ${result.totalCalls} API calls`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Sets up daily trigger at 00:01 for inventory and balance updates.
 * Runs NF_BulkImportFromDrive() then NF_UpdateInventoryBalances().
 */
function NF_setupDaily0001Inventory() {
  // Clear existing triggers for this function
  NF_clearTriggersForFunction_('NF_dailyInventoryUpdate');
  
  // Create new trigger
  ScriptApp.newTrigger('NF_dailyInventoryUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(1)
    .create();
    
  SpreadsheetApp.getUi().alert(
    'Trigger Scheduled',
    'Daily inventory update scheduled for 00:01 (Drive import + inventory balances)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  console.log('Daily 00:01 inventory trigger scheduled');
}

/**
 * Sets up daily trigger at 11:00 for tracking refresh and report updates.
 * Runs NF_RunDailyFlow(), NF_buildSokKarkkainenAlways(), and NF_ReconcileWeeklyFromImport().
 */
function NF_setupDaily1100Tracking() {
  // Clear existing triggers for this function
  NF_clearTriggersForFunction_('NF_dailyTrackingUpdate');
  
  // Create new trigger
  ScriptApp.newTrigger('NF_dailyTrackingUpdate')
    .timeBased()
    .everyDays(1)
    .atHour(11)
    .nearMinute(0)
    .create();
    
  SpreadsheetApp.getUi().alert(
    'Trigger Scheduled',
    'Daily tracking update scheduled for 11:00 (Gmail import + tracking refresh + SOK/KRK update)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  console.log('Daily 11:00 tracking trigger scheduled');
}

/**
 * Sets up weekly trigger on Monday at 02:00 for building weekly reports.
 * Maintains existing NF_BuildWeeklyReports() function or equivalent.
 */
function NF_setupWeeklyMon0200() {
  // Clear existing triggers for this function
  NF_clearTriggersForFunction_('NF_weeklyReportBuild');
  
  // Create new trigger
  ScriptApp.newTrigger('NF_weeklyReportBuild')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(2)
    .nearMinute(0)
    .create();
    
  SpreadsheetApp.getUi().alert(
    'Trigger Scheduled',
    'Weekly report build scheduled for Monday 02:00',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  console.log('Weekly Monday 02:00 report trigger scheduled');
}

/**
 * Clears all NF-related triggers.
 */
function NF_clearAllTriggers() {
  const nfFunctions = [
    'NF_dailyInventoryUpdate',
    'NF_dailyTrackingUpdate', 
    'NF_weeklyReportBuild'
  ];
  
  let cleared = 0;
  
  for (const functionName of nfFunctions) {
    cleared += NF_clearTriggersForFunction_(functionName);
  }
  
  SpreadsheetApp.getUi().alert(
    'Triggers Cleared',
    `Cleared ${cleared} NF-related triggers`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  console.log(`Cleared ${cleared} NF triggers`);
}

/**
 * Lists all currently active triggers in a dialog.
 */
function NF_listActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert('No active triggers found');
    return;
  }
  
  let triggerList = 'Active Triggers:\n\n';
  
  triggers.forEach((trigger, index) => {
    const func = trigger.getHandlerFunction();
    const triggerSource = trigger.getTriggerSource();
    
    if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
      const eventType = trigger.getEventType();
      triggerList += `${index + 1}. ${func} - ${eventType}\n`;
    } else {
      triggerList += `${index + 1}. ${func} - ${triggerSource}\n`;
    }
  });
  
  SpreadsheetApp.getUi().alert('Active Triggers', triggerList, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Helper function to clear triggers for a specific function.
 * @param {string} functionName - Name of the function to clear triggers for
 * @return {number} Number of triggers cleared
 */
function NF_clearTriggersForFunction_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  let cleared = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      cleared++;
    }
  });
  
  return cleared;
}

// === TRIGGER HANDLER FUNCTIONS ===

/**
 * Daily inventory update handler (00:01).
 * Imports from Drive and updates inventory balances.
 */
function NF_dailyInventoryUpdate() {
  console.log('NF_dailyInventoryUpdate: Starting daily inventory update at 00:01');
  
  try {
    // Step 1: Import from Drive
    console.log('Importing from Drive folder...');
    NF_BulkImportFromDrive();
    
    // Step 2: Update inventory balances
    console.log('Updating inventory balances...');
    NF_UpdateInventoryBalances();
    
    console.log('Daily inventory update completed successfully');
    
  } catch (error) {
    console.error('Daily inventory update failed:', error);
    
    // Optionally send error notification
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = ss.getSheetByName('Error_Log') || ss.insertSheet('Error_Log');
      const now = new Date();
      
      if (errorSheet.getLastRow() === 0) {
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Details']]);
      }
      
      errorSheet.appendRow([now, 'NF_dailyInventoryUpdate', error.message, error.stack || '']);
      
    } catch (logError) {
      console.error('Failed to log error:', logError);
    }
  }
}

/**
 * Daily tracking update handler (11:00).
 * Runs daily flow, builds SOK/KRK reports, and reconciles weekly data.
 */
function NF_dailyTrackingUpdate() {
  console.log('NF_dailyTrackingUpdate: Starting daily tracking update at 11:00');
  
  try {
    // Step 1: Run daily flow (Gmail import + tracking refresh)
    console.log('Running daily flow...');
    if (typeof NF_RunDailyFlow === 'function') {
      NF_RunDailyFlow();
    } else if (typeof runDailyFlowOnce === 'function') {
      runDailyFlowOnce();
    } else {
      console.warn('No daily flow function available');
    }
    
    // Step 2: Build SOK/Kärkkäinen reports
    console.log('Building SOK/Kärkkäinen reports...');
    NF_buildSokKarkkainenAlways();
    
    // Step 3: Reconcile weekly data from imports
    console.log('Reconciling weekly data...');
    if (typeof NF_ReconcileWeeklyFromImport === 'function') {
      NF_ReconcileWeeklyFromImport();
    }
    
    console.log('Daily tracking update completed successfully');
    
  } catch (error) {
    console.error('Daily tracking update failed:', error);
    
    // Log error
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = ss.getSheetByName('Error_Log') || ss.insertSheet('Error_Log');
      const now = new Date();
      
      if (errorSheet.getLastRow() === 0) {
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Details']]);
      }
      
      errorSheet.appendRow([now, 'NF_dailyTrackingUpdate', error.message, error.stack || '']);
      
    } catch (logError) {
      console.error('Failed to log error:', logError);
    }
  }
}

/**
 * Weekly report build handler (Monday 02:00).
 * Builds comprehensive weekly reports.
 */
function NF_weeklyReportBuild() {
  console.log('NF_weeklyReportBuild: Starting weekly report build on Monday 02:00');
  
  try {
    // Call existing weekly report function if available
    if (typeof NF_BuildWeeklyReports === 'function') {
      NF_BuildWeeklyReports();
    } else if (typeof makeWeeklyReportsSunSun === 'function') {
      makeWeeklyReportsSunSun();
    } else {
      console.log('Building SOK/Kärkkäinen reports as weekly fallback...');
      NF_buildSokKarkkainenAlways();
    }
    
    console.log('Weekly report build completed successfully');
    
  } catch (error) {
    console.error('Weekly report build failed:', error);
    
    // Log error
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const errorSheet = ss.getSheetByName('Error_Log') || ss.insertSheet('Error_Log');
      const now = new Date();
      
      if (errorSheet.getLastRow() === 0) {
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Details']]);
      }
      
      errorSheet.appendRow([now, 'NF_weeklyReportBuild', error.message, error.stack || '']);
      
    } catch (logError) {
      console.error('Failed to log error:', logError);
    }
  }
}