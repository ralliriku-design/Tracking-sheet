/**
 * NF_Menu_Triggers.gs - Menu and Trigger Management
 * 
 * Provides non-conflicting menu integration and trigger management for the New Flow system.
 * Ensures triggers are installed/removed safely without affecting existing functionality.
 */

/********************* MENU INTEGRATION ***************************/

/**
 * Non-conflicting onOpen trigger for New Flow menu
 * This should be installed once via NF_installMenuTrigger()
 */
function NF_onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  // Create New Flow submenu
  const newFlowMenu = ui.createMenu('New Flow')
    .addItem('Run daily flow now', 'NF_RunDailyFlow')
    .addItem('Build SOK & Kärkkäinen (last week)', 'NF_BuildWeeklyReports')
    .addSeparator()
    .addItem('Delivery Times list', 'NF_BuildDeliveryTimes')
    .addItem('Country Week Leadtime', 'NF_MakeCountryWeekLeadtime')
    .addSeparator()
    .addItem('Install weekday 12:00', 'NF_setupWeekday1200')
    .addItem('Install weekly Mon 02:00', 'NF_setupWeeklyMon0200')
    .addItem('Remove NF triggers', 'NF_clearNFTriggers')
    .addSeparator()
    .addItem('Show NF status', 'NF_showStatus');
  
  // Check if main Tracking menu exists, if so add as submenu
  try {
    const existingMenu = ui.createMenu('Tracking');
    existingMenu.addSubMenu(newFlowMenu);
    existingMenu.addToUi();
  } catch (e) {
    // If Tracking menu doesn't exist or fails, add as standalone
    newFlowMenu.addToUi();
  }
}

/**
 * Install the New Flow menu trigger (call this once during setup)
 */
function NF_installMenuTrigger() {
  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.find(t => 
    t.getEventType() === ScriptApp.EventType.ON_OPEN && 
    t.getHandlerFunction() === 'NF_onOpen'
  );
  
  if (existing) {
    SpreadsheetApp.getUi().alert('New Flow menu trigger already installed.');
    return;
  }
  
  // Install trigger
  try {
    ScriptApp.newTrigger('NF_onOpen')
      .onOpen()
      .create();
    
    SpreadsheetApp.getUi().alert('New Flow menu trigger installed successfully.\nRefresh the page to see the menu.');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Failed to install menu trigger: ${error.message}`);
  }
}

/********************* TRIGGER MANAGEMENT ***************************/

/**
 * Setup weekday 12:00 trigger for daily flow
 */
function NF_setupWeekday1200() {
  const functionName = 'NF_RunDailyFlow';
  
  // Remove any existing daily triggers for this function
  NF_removeTriggersByFunction_(functionName);
  
  try {
    // Create weekday triggers (Monday through Friday at 12:00)
    for (let day = ScriptApp.WeekDay.MONDAY; day <= ScriptApp.WeekDay.FRIDAY; day++) {
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyWeeks(1)
        .onWeekDay(day)
        .atHour(12)
        .create();
    }
    
    SpreadsheetApp.getUi().alert('Daily flow triggers installed for weekdays at 12:00.');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Failed to setup daily triggers: ${error.message}`);
  }
}

/**
 * Setup weekly Monday 02:00 trigger for weekly reports
 */
function NF_setupWeeklyMon0200() {
  const functionName = 'NF_BuildWeeklyReports';
  
  // Remove any existing weekly triggers for this function
  NF_removeTriggersByFunction_(functionName);
  
  try {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(2)
      .create();
    
    SpreadsheetApp.getUi().alert('Weekly reports trigger installed for Mondays at 02:00.');
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Failed to setup weekly trigger: ${error.message}`);
  }
}

/**
 * Clear all New Flow triggers (NF_* functions only)
 */
function NF_clearNFTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const nfTriggers = triggers.filter(t => t.getHandlerFunction().startsWith('NF_'));
  
  if (nfTriggers.length === 0) {
    SpreadsheetApp.getUi().alert('No New Flow triggers found.');
    return;
  }
  
  let removed = 0;
  for (const trigger of nfTriggers) {
    try {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    } catch (e) {
      console.log(`Warning: Could not remove trigger ${trigger.getHandlerFunction()}:`, e.message);
    }
  }
  
  SpreadsheetApp.getUi().alert(`Removed ${removed} New Flow triggers.`);
}

/**
 * Show New Flow status and configuration
 */
function NF_showStatus() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  // Check triggers
  const triggers = ScriptApp.getProjectTriggers();
  const nfTriggers = triggers.filter(t => t.getHandlerFunction().startsWith('NF_'));
  
  // Check Script Properties
  const props = PropertiesService.getScriptProperties();
  const requiredProps = [
    'SOK_FREIGHT_ACCOUNT',
    'KARKKAINEN_NUMBERS',
    'TARGET_SHEET',
    'ARCHIVE_SHEET',
    'GMAIL_QUERY'
  ];
  
  const missingProps = requiredProps.filter(prop => !props.getProperty(prop));
  
  // Check sheets
  const requiredSheets = ['Packages', 'Packages_Archive'];
  const missingSheets = requiredSheets.filter(name => !ss.getSheetByName(name));
  
  // Build status message
  let status = '=== NEW FLOW STATUS ===\n\n';
  
  status += `TRIGGERS (${nfTriggers.length}):\n`;
  if (nfTriggers.length === 0) {
    status += '  None installed\n';
  } else {
    nfTriggers.forEach(t => {
      const type = t.getEventType() === ScriptApp.EventType.ON_OPEN ? 'onOpen' : 'timeBased';
      status += `  ${t.getHandlerFunction()} (${type})\n`;
    });
  }
  
  status += `\nSCRIPT PROPERTIES:\n`;
  requiredProps.forEach(prop => {
    const value = props.getProperty(prop);
    status += `  ${prop}: ${value ? '✓' : '✗'}\n`;
  });
  
  if (missingProps.length > 0) {
    status += `\nMISSING PROPERTIES: ${missingProps.join(', ')}\n`;
  }
  
  status += `\nSHEETS:\n`;
  requiredSheets.forEach(name => {
    const exists = ss.getSheetByName(name) ? '✓' : '✗';
    status += `  ${name}: ${exists}\n`;
  });
  
  if (missingSheets.length > 0) {
    status += `\nMISSING SHEETS: ${missingSheets.join(', ')}\n`;
  }
  
  // Check for tracking engine
  const hasEnhanced = typeof TRK_trackByCarrierEnhanced === 'function';
  const hasFallback = typeof TRK_trackByCarrier === 'function';
  status += `\nTRACKING ENGINE:\n`;
  status += `  TRK_trackByCarrierEnhanced: ${hasEnhanced ? '✓' : '✗'}\n`;
  status += `  TRK_trackByCarrier (fallback): ${hasFallback ? '✓' : '✗'}\n`;
  
  ui.alert('New Flow Status', status, ui.ButtonSet.OK);
}

/********************* HELPER FUNCTIONS ***************************/

/**
 * Remove triggers by function name (safely, only NF_ functions)
 */
function NF_removeTriggersByFunction_(functionName) {
  if (!functionName.startsWith('NF_')) {
    throw new Error('Can only remove NF_ triggers for safety');
  }
  
  const triggers = ScriptApp.getProjectTriggers();
  const toRemove = triggers.filter(t => t.getHandlerFunction() === functionName);
  
  for (const trigger of toRemove) {
    try {
      ScriptApp.deleteTrigger(trigger);
    } catch (e) {
      console.log(`Warning: Could not remove trigger ${functionName}:`, e.message);
    }
  }
  
  return toRemove.length;
}

/**
 * List all project triggers (for debugging)
 */
function NF_listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  console.log('=== ALL PROJECT TRIGGERS ===');
  triggers.forEach((trigger, index) => {
    const eventType = trigger.getEventType();
    const handlerFunction = trigger.getHandlerFunction();
    
    let details = '';
    if (eventType === ScriptApp.EventType.CLOCK) {
      // Time-based trigger
      details = `every ${trigger.getTimeBased()}`;
    } else if (eventType === ScriptApp.EventType.ON_OPEN) {
      details = 'onOpen';
    } else if (eventType === ScriptApp.EventType.ON_EDIT) {
      details = 'onEdit';
    } else {
      details = eventType.toString();
    }
    
    console.log(`${index + 1}. ${handlerFunction} (${details})`);
  });
  
  return triggers.length;
}

/**
 * Safe trigger installation helper
 */
function NF_installTriggerSafely_(functionName, triggerBuilder) {
  // Check if function exists
  try {
    const func = eval(functionName);
    if (typeof func !== 'function') {
      throw new Error(`Function ${functionName} is not available`);
    }
  } catch (e) {
    throw new Error(`Function ${functionName} is not available: ${e.message}`);
  }
  
  // Remove existing triggers for this function
  const removed = NF_removeTriggersByFunction_(functionName);
  
  // Install new trigger
  const trigger = triggerBuilder(functionName);
  
  return { trigger, removed };
}