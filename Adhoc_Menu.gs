// Adhoc_Menu.gs — Safe menu installation for ad hoc tracker

/**
 * Safe menu installer that doesn't conflict with existing onOpen functions
 * Creates an installable trigger for menu management
 */

/**
 * Install menu trigger for ad hoc tracker
 * Call this once to set up the menu system
 */
function Adhoc_installMenuTrigger() {
  const ss = SpreadsheetApp.getActive();
  
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existingTrigger = triggers.find(trigger => 
    trigger.getHandlerFunction() === 'Adhoc_onOpen' &&
    trigger.getTriggerSource() === ScriptApp.TriggerSource.SPREADSHEETS &&
    trigger.getEventType() === ScriptApp.EventType.ON_OPEN
  );
  
  if (existingTrigger) {
    ss.toast('Ad hoc tracker menu trigger already installed');
    return;
  }
  
  // Create new trigger
  try {
    ScriptApp.newTrigger('Adhoc_onOpen')
      .onOpen()
      .create();
    
    ss.toast('Ad hoc tracker menu trigger installed successfully');
    
    // Immediately add the menu for this session
    Adhoc_onOpen();
    
  } catch (error) {
    ss.toast('Error installing menu trigger: ' + error.message);
    throw error;
  }
}

/**
 * Remove menu trigger for ad hoc tracker
 */
function Adhoc_removeMenuTrigger() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  
  const adhocTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'Adhoc_onOpen'
  );
  
  if (adhocTriggers.length === 0) {
    ss.toast('No ad hoc tracker menu triggers found');
    return;
  }
  
  adhocTriggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  ss.toast(`Removed ${adhocTriggers.length} ad hoc tracker menu trigger(s)`);
}

/**
 * OnOpen handler - adds ad hoc tracker menu items
 * This function is called by the installable trigger
 */
function Adhoc_onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Create main tracking menu if it doesn't exist
    let trackingMenu;
    
    try {
      // Try to get existing Tracking menu
      trackingMenu = ui.createMenu('Tracking');
    } catch (error) {
      // If menu creation fails, create a standalone menu
      trackingMenu = ui.createMenu('Tracking');
    }
    
    // Add ad hoc tracker submenu
    const adhocSubmenu = ui.createMenu('Ad hoc -tracker')
      .addItem('Aja tuonti & päivitys', 'ADHOC_RunFromUrl')
      .addItem('Päivitä olemassa olevat', 'ADHOC_RefreshResults')
      .addSeparator()
      .addItem('Asenna menu trigger', 'Adhoc_installMenuTrigger')
      .addItem('Poista menu trigger', 'Adhoc_removeMenuTrigger')
      .addSeparator()
      .addItem('Näytä tulokset', 'ADHOC_ShowResults')
      .addItem('Näytä KPI:t', 'ADHOC_ShowKPI');
    
    // Add submenu to tracking menu
    trackingMenu.addSubMenu(adhocSubmenu)
      .addToUi();
      
  } catch (error) {
    // Fallback: create standalone menu if main menu fails
    try {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('Ad hoc Tracker')
        .addItem('Aja tuonti & päivitys', 'ADHOC_RunFromUrl')
        .addItem('Päivitä olemassa olevat', 'ADHOC_RefreshResults')
        .addSeparator()
        .addItem('Asenna menu trigger', 'Adhoc_installMenuTrigger')
        .addItem('Poista menu trigger', 'Adhoc_removeMenuTrigger')
        .addSeparator()
        .addItem('Näytä tulokset', 'ADHOC_ShowResults')
        .addItem('Näytä KPI:t', 'ADHOC_ShowKPI')
        .addToUi();
    } catch (fallbackError) {
      Logger.log('Failed to create ad hoc tracker menu: ' + fallbackError.message);
    }
  }
}

/**
 * Manual menu installation (alternative to trigger)
 * Can be called directly if triggers don't work
 */
function Adhoc_installMenuManual() {
  Adhoc_onOpen();
  SpreadsheetApp.getActive().toast('Ad hoc tracker menu added manually');
}

/**
 * Show Adhoc_Results sheet
 */
function ADHOC_ShowResults() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Adhoc_Results');
  
  if (!sheet) {
    ss.toast('Adhoc_Results sheet not found. Run import first.');
    return;
  }
  
  ss.setActiveSheet(sheet);
}

/**
 * Show Adhoc_KPI sheet
 */
function ADHOC_ShowKPI() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Adhoc_KPI');
  
  if (!sheet) {
    ss.toast('Adhoc_KPI sheet not found. Run import first.');
    return;
  }
  
  ss.setActiveSheet(sheet);
}

/**
 * Check trigger installation status
 */
function Adhoc_checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const adhocTriggers = triggers.filter(trigger => 
    trigger.getHandlerFunction() === 'Adhoc_onOpen'
  );
  
  const ss = SpreadsheetApp.getActive();
  
  if (adhocTriggers.length === 0) {
    ss.toast('No ad hoc tracker menu triggers installed');
  } else {
    ss.toast(`${adhocTriggers.length} ad hoc tracker menu trigger(s) installed`);
  }
  
  return adhocTriggers.length;
}

/**
 * Diagnostic function to test menu system
 */
function Adhoc_testMenu() {
  const ss = SpreadsheetApp.getActive();
  
  try {
    Adhoc_onOpen();
    ss.toast('Menu test successful');
  } catch (error) {
    ss.toast('Menu test failed: ' + error.message);
    throw error;
  }
}

/**
 * Initialize ad hoc tracker system
 * Sets up triggers and menus for first-time use
 */
function Adhoc_initialize() {
  const ss = SpreadsheetApp.getActive();
  
  try {
    // Install trigger
    Adhoc_installMenuTrigger();
    
    // Create example sheets if they don't exist
    if (!ss.getSheetByName('Adhoc_Results')) {
      const sheet = ss.insertSheet('Adhoc_Results');
      const headers = [
        'Carrier', 'Tracking', 'Country', 'Status', 'CreatedISO', 'DeliveredISO',
        'DaysToDeliver', 'WeekISO', 'RefreshAt', 'Location', 'Raw'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    
    if (!ss.getSheetByName('Adhoc_KPI')) {
      const sheet = ss.insertSheet('Adhoc_KPI');
      const headers = [
        'Country', 'ISO Week', 'Deliveries', 'Avg Days', 'Median Days', 'Min Days', 'Max Days'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    
    ss.toast('Ad hoc tracker initialized successfully');
    
  } catch (error) {
    ss.toast('Initialization failed: ' + error.message);
    throw error;
  }
}