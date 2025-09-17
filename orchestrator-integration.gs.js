/**
 * Orchestrator Integration Helper
 * 
 * This file provides integration functions to connect the enhanced tracking orchestrator
 * with the existing Google Apps Script tracking system.
 */

/**
 * Helper functions needed for integration
 */

/**
 * Normalize text for header matching
 * @param {string} s - Text to normalize
 * @returns {string} Normalized text
 */
function normalize_(s) {
  if (typeof normalize_ === 'function' && normalize_ !== arguments.callee) {
    // Use existing function if available
    return normalize_(s);
  }
  return String(s || '').toLowerCase().replace(/\s+/g,' ').trim().replace(/[^\p{L}\p{N}]+/gu,' ');
}

/**
 * Create header index map
 * @param {Array} hdr - Header array
 * @returns {Object} Header index map
 */
function headerIndexMap_(hdr) {
  if (typeof headerIndexMap_ === 'function' && headerIndexMap_ !== arguments.callee) {
    // Use existing function if available
    return headerIndexMap_(hdr);
  }
  
  const m = {};
  (hdr || []).forEach((h, i) => {
    try {
      const n = normalize_(h || '');
      m[h] = i;        // original
      m[n] = i;        // normalized
    } catch (e) {
      m[h] = i;
      m[String(h).toLowerCase()] = i;
    }
  });
  return m;
}

/**
 * Find column index from candidates
 * @param {Object} hdrMap - Header index map
 * @param {Array} candidates - Candidate header names
 * @returns {number} Column index or -1
 */
function colIndexOf_(hdrMap, candidates) {
  if (typeof colIndexOf_ === 'function' && colIndexOf_ !== arguments.callee) {
    // Use existing function if available
    return colIndexOf_(hdrMap, candidates);
  }
  
  for (const label of candidates) {
    if (label in hdrMap) return hdrMap[label];
    const n = normalize_(label);
    if (n in hdrMap) return hdrMap[n];
  }
  return -1;
}

/**
 * Canonicalize carrier name
 * @param {string} s - Carrier name
 * @returns {string} Canonical carrier name
 */
function canonicalCarrier_(s) {
  if (typeof canonicalCarrier_ === 'function' && canonicalCarrier_ !== arguments.callee) {
    // Use existing function if available
    return canonicalCarrier_(s);
  }
  
  const carrier = String(s || '').toLowerCase().trim();
  if (carrier.includes('posti')) return 'Posti';
  if (carrier.includes('gls')) return 'GLS';
  if (carrier.includes('dhl')) return 'DHL';
  if (carrier.includes('bring')) return 'Bring';
  if (carrier.includes('matka')) return 'Matkahuolto';
  return s || '';
}

/**
 * Enhanced credential validation functions for the UI
 */

/**
 * Enhanced version of credSaveProps with validation
 * @param {Object} data - Credential data to save
 * @returns {Object} Save result with validation
 */
function enhancedCredSaveProps(data) {
  if (!data || typeof data !== 'object') {
    return { ok: false, msg: 'No data provided' };
  }

  const credManager = new SafeCredentialManager();
  const results = {};
  let hasErrors = false;

  try {
    Object.keys(data).forEach(key => {
      const value = data[key];
      
      // Determine validation options based on key
      const options = getCredentialValidationOptions_(key);
      
      if (value === '' || value === null || value === undefined) {
        // Delete empty credentials
        const success = credManager.safeSet(key, null);
        results[key] = { success, action: 'deleted' };
      } else {
        // Validate and set credential
        const success = credManager.safeSet(key, value, options);
        results[key] = { 
          success, 
          action: 'set',
          validated: true,
          options: options
        };
        
        if (!success) hasErrors = true;
      }
    });

    return {
      ok: !hasErrors,
      msg: hasErrors ? 'Some credentials failed validation' : 'All credentials saved successfully',
      details: results
    };

  } catch (error) {
    return {
      ok: false,
      msg: `Error saving credentials: ${error.message}`,
      details: results
    };
  }
}

/**
 * Enhanced version of credGetProps with validation status
 * @returns {Object} Credentials with validation status
 */
function enhancedCredGetProps() {
  const credManager = new SafeCredentialManager();
  const sp = PropertiesService.getScriptProperties();
  
  const keys = [
    'MH_TRACK_URL', 'MH_BASIC',
    'POSTI_TOKEN_URL', 'POSTI_BASIC', 'POSTI_TRACK_URL', 'POSTI_TRK_URL', 'POSTI_TRK_BASIC', 'POSTI_TRK_USER', 'POSTI_TRK_PASS',
    'GLS_FI_TRACK_URL', 'GLS_FI_API_KEY', 'GLS_FI_SENDER_ID', 'GLS_FI_SENDER_IDS', 'GLS_FI_METHOD',
    'GLS_TOKEN_URL', 'GLS_BASIC', 'GLS_TRACK_URL',
    'DHL_TRACK_URL', 'DHL_API_KEY',
    'BRING_TRACK_URL', 'BRING_UID', 'BRING_KEY', 'BRING_CLIENT_URL',
    'BULK_BACKOFF_MINUTES_BASE',
    'PBI_WEBHOOK_URL',
    'GMAIL_QUERY',
    'TRACKING_HEADER_HINT', 'TRACKING_DEFAULT_COL', 'ATTACH_ALLOW_REGEX', 'REFRESH_MIN_AGE_MINUTES'
  ];

  const props = {};
  const validation = {};

  keys.forEach(key => {
    const rawValue = sp.getProperty(key) || '';
    props[key] = rawValue;
    
    // Add validation status
    const options = getCredentialValidationOptions_(key);
    const validatedValue = credManager.safeGet(key, options);
    validation[key] = {
      present: !!rawValue,
      valid: !!validatedValue,
      required: isRequiredCredential_(key),
      options: options
    };
  });

  return {
    credentials: props,
    validation: validation,
    summary: {
      total: keys.length,
      present: Object.values(validation).filter(v => v.present).length,
      valid: Object.values(validation).filter(v => v.valid).length,
      required: Object.values(validation).filter(v => v.required).length
    }
  };
}

/**
 * Test all carrier credentials
 * @returns {Object} Test results for all carriers
 */
async function testAllCarrierCredentials() {
  const carriers = ['POSTI', 'GLS', 'DHL', 'BRING', 'MATKAHUOLTO'];
  const credManager = new SafeCredentialManager();
  const results = {};

  for (const carrier of carriers) {
    try {
      const validation = credManager.validateCarrierCredentials(carrier);
      const testResult = await credManager.testCredential(carrier, 'TEST123');
      
      results[carrier] = {
        validation: validation,
        test: testResult,
        overall: validation.valid && testResult.success
      };
    } catch (error) {
      results[carrier] = {
        validation: { valid: false, error: error.message },
        test: { success: false, error: error.message },
        overall: false
      };
    }
  }

  return results;
}

/**
 * Enhanced bulk start functions using the orchestrator
 */

/**
 * Enhanced bulk start for Vaatii_toimenpiteitä sheet
 */
async function enhancedBulkStart_Vaatii() {
  await enhancedBulkStart('Vaatii_toimenpiteitä', null);
}

/**
 * Enhanced bulk start for Packages sheet
 */
async function enhancedBulkStart_Packages() {
  await enhancedBulkStart('Packages', null);
}

/**
 * Enhanced bulk start for Packages_Archive sheet
 */
async function enhancedBulkStart_Archive() {
  await enhancedBulkStart('Packages_Archive', null);
}

/**
 * Enhanced bulk start with carrier filter
 */
async function enhancedBulkStartCarrier(sheetName, carrier) {
  await enhancedBulkStart(sheetName, carrier);
}

/**
 * Menu functions for enhanced tracking
 */
async function menuEnhancedRefreshCarrier_MH() {
  await enhancedBulkStartCarrier('Vaatii_toimenpiteitä', 'matkahuolto');
}

async function menuEnhancedRefreshCarrier_POSTI() {
  await enhancedBulkStartCarrier('Vaatii_toimenpiteitä', 'posti');
}

async function menuEnhancedRefreshCarrier_BRING() {
  await enhancedBulkStartCarrier('Vaatii_toimenpiteitä', 'bring');
}

async function menuEnhancedRefreshCarrier_GLS() {
  await enhancedBulkStartCarrier('Vaatii_toimenpiteitä', 'gls');
}

async function menuEnhancedRefreshCarrier_DHL() {
  await enhancedBulkStartCarrier('Vaatii_toimenpiteitä', 'dhl');
}

async function menuEnhancedRefreshCarrier_ALL() {
  await enhancedBulkStart('Vaatii_toimenpiteitä', null);
}

/**
 * Orchestrator management functions
 */

/**
 * Initialize the orchestrator for use
 */
async function setupTrackingOrchestrator() {
  try {
    const orchestrator = await initializeTrackingOrchestrator();
    SpreadsheetApp.getUi().alert('Enhanced Tracking Orchestrator initialized successfully!');
    return orchestrator;
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Failed to initialize orchestrator: ${error.message}`);
    throw error;
  }
}

/**
 * Show detailed credential validation report
 */
function showCredentialValidationReport() {
  const credManager = new SafeCredentialManager();
  const carriers = ['POSTI', 'GLS', 'DHL', 'BRING', 'MATKAHUOLTO'];
  let report = 'Credential Validation Report:\n\n';

  carriers.forEach(carrier => {
    const validation = credManager.validateCarrierCredentials(carrier);
    report += `${carrier}:\n`;
    report += `  Overall Valid: ${validation.valid ? '✓' : '✗'}\n`;
    
    Object.entries(validation.credentials).forEach(([key, status]) => {
      report += `  ${key}: ${status.valid ? '✓' : '✗'} (${status.message})\n`;
    });
    report += '\n';
  });

  SpreadsheetApp.getUi().alert(report);
}

/**
 * Emergency stop for orchestrator
 */
async function emergencyStopOrchestrator() {
  const orchestrator = getTrackingOrchestrator();
  await orchestrator.shutdown();
  SpreadsheetApp.getUi().alert('Orchestrator has been stopped.');
}

/**
 * Helper functions
 */

/**
 * Get validation options for a credential key
 * @private
 */
function getCredentialValidationOptions_(key) {
  const urlPattern = /^https?:\/\//;
  const codeUrlPattern = /^https?:\/\/.*\{\{code\}\}/;
  
  const validationMap = {
    // URLs that should contain {{code}} placeholder
    'MH_TRACK_URL': { notEmpty: true, pattern: codeUrlPattern },
    'POSTI_TRACK_URL': { notEmpty: true, pattern: codeUrlPattern },
    'POSTI_TRK_URL': { notEmpty: true, pattern: codeUrlPattern },
    'GLS_TRACK_URL': { notEmpty: true, pattern: codeUrlPattern },
    'DHL_TRACK_URL': { notEmpty: true, pattern: codeUrlPattern },
    'BRING_TRACK_URL': { notEmpty: true, pattern: codeUrlPattern },
    
    // URLs without code placeholder
    'POSTI_TOKEN_URL': { notEmpty: true, pattern: urlPattern },
    'GLS_FI_TRACK_URL': { notEmpty: true, pattern: urlPattern },
    'GLS_TOKEN_URL': { notEmpty: true, pattern: urlPattern },
    'PBI_WEBHOOK_URL': { notEmpty: true, pattern: urlPattern },
    'BRING_CLIENT_URL': { notEmpty: true, pattern: urlPattern },
    
    // API Keys (should be reasonably long)
    'MH_BASIC': { notEmpty: true, minLength: 10 },
    'POSTI_BASIC': { notEmpty: true, minLength: 10 },
    'POSTI_TRK_BASIC': { notEmpty: true, minLength: 10 },
    'GLS_FI_API_KEY': { notEmpty: true, minLength: 10 },
    'GLS_BASIC': { notEmpty: true, minLength: 10 },
    'DHL_API_KEY': { notEmpty: true, minLength: 10 },
    'BRING_KEY': { notEmpty: true, minLength: 5 },
    
    // User credentials
    'POSTI_TRK_USER': { notEmpty: true, minLength: 3 },
    'POSTI_TRK_PASS': { notEmpty: true, minLength: 3 },
    'BRING_UID': { notEmpty: true, minLength: 3 },
    
    // Other fields
    'GLS_FI_SENDER_ID': { notEmpty: true },
    'GLS_FI_SENDER_IDS': { notEmpty: true },
    'GMAIL_QUERY': { notEmpty: true }
  };

  return validationMap[key] || { notEmpty: false };
}

/**
 * Check if a credential is required for basic functionality
 * @private
 */
function isRequiredCredential_(key) {
  const requiredCredentials = [
    'POSTI_TOKEN_URL', 'POSTI_BASIC', 'POSTI_TRACK_URL',
    'GLS_FI_TRACK_URL', 'GLS_FI_API_KEY',
    'DHL_TRACK_URL', 'DHL_API_KEY'
  ];
  
  return requiredCredentials.includes(key);
}

/**
 * Compatibility wrapper for existing bulkTick function
 * This allows gradual migration to the new orchestrator
 */
function enhancedBulkTick() {
  try {
    // Try to use orchestrator if available
    const orchestrator = getTrackingOrchestrator();
    if (orchestrator && orchestrator.isRunning) {
      orchestrator.processQueue();
      return;
    }
  } catch (error) {
    // Fall back to original bulkTick if orchestrator not available
    if (typeof logError_ === 'function') {
      logError_('enhancedBulkTick', `Orchestrator not available, falling back: ${error.message}`);
    }
  }
  
  // Call original bulkTick as fallback
  if (typeof bulkTick === 'function') {
    bulkTick();
  }
}

/**
 * Create enhanced menu items
 */
function addEnhancedMenuItems() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('🚀 Enhanced Tracking')
      .addSubMenu(ui.createMenu('Initialize')
        .addItem('Setup Orchestrator', 'setupTrackingOrchestrator')
        .addItem('Show Status', 'showOrchestratorStatus')
        .addItem('Reset Circuit Breakers', 'resetAllCircuitBreakers')
        .addItem('Emergency Stop', 'emergencyStopOrchestrator'))
      .addSubMenu(ui.createMenu('Enhanced Bulk Refresh')
        .addItem('All Carriers', 'menuEnhancedRefreshCarrier_ALL')
        .addItem('Posti Only', 'menuEnhancedRefreshCarrier_POSTI')
        .addItem('GLS Only', 'menuEnhancedRefreshCarrier_GLS')
        .addItem('DHL Only', 'menuEnhancedRefreshCarrier_DHL')
        .addItem('Bring Only', 'menuEnhancedRefreshCarrier_BRING')
        .addItem('Matkahuolto Only', 'menuEnhancedRefreshCarrier_MH'))
      .addSubMenu(ui.createMenu('Enhanced Sheets')
        .addItem('Vaatii_toimenpiteitä', 'enhancedBulkStart_Vaatii')
        .addItem('Packages', 'enhancedBulkStart_Packages')
        .addItem('Packages_Archive', 'enhancedBulkStart_Archive'))
      .addSubMenu(ui.createMenu('Credentials')
        .addItem('Validation Report', 'showCredentialValidationReport')
        .addItem('Test All Carriers', 'testAllCarrierCredentials'))
      .addToUi();
      
  } catch (error) {
    console.error('Failed to add enhanced menu items:', error);
  }
}

/**
 * Enhanced onOpen function that adds the new menu
 */
function enhancedOnOpen() {
  // Call original onOpen if it exists
  if (typeof onOpen === 'function') {
    onOpen();
  }
  
  // Add enhanced menu
  addEnhancedMenuItems();
}