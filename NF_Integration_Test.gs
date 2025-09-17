/**
 * NF_Integration_Test.gs - Integration tests with existing functions
 * Tests the actual integration with existing repo functions
 */

function NF_testExistingFunctionAvailability() {
  const results = [];
  
  // Test for fetchAndRebuild
  if (typeof fetchAndRebuild === 'function') {
    results.push('✓ fetchAndRebuild function is available');
  } else {
    results.push('✗ fetchAndRebuild function NOT found');
  }
  
  // Test for gmailImportLatestPackagesReport
  if (typeof gmailImportLatestPackagesReport === 'function') {
    results.push('✓ gmailImportLatestPackagesReport function is available');
  } else {
    results.push('✗ gmailImportLatestPackagesReport function NOT found');
  }
  
  // Test for pbiImportOutbounds_OldestFirst
  if (typeof pbiImportOutbounds_OldestFirst === 'function') {
    results.push('✓ pbiImportOutbounds_OldestFirst function is available');
  } else {
    results.push('✗ pbiImportOutbounds_OldestFirst function NOT found');
  }
  
  // Test for runERPUpdate
  if (typeof runERPUpdate === 'function') {
    results.push('✓ runERPUpdate function is available');
  } else {
    results.push('✗ runERPUpdate function NOT found');
  }
  
  // Test for getLastFinishedWeekSunWindow_
  if (typeof getLastFinishedWeekSunWindow_ === 'function') {
    results.push('✓ getLastFinishedWeekSunWindow_ function is available');
  } else {
    results.push('✗ getLastFinishedWeekSunWindow_ function NOT found');
  }
  
  // Test for writeWeeklySheet_
  if (typeof writeWeeklySheet_ === 'function') {
    results.push('✓ writeWeeklySheet_ function is available');
  } else {
    results.push('✗ writeWeeklySheet_ function NOT found');
  }
  
  Logger.log('Existing Function Availability:\n' + results.join('\n'));
  return results;
}

function NF_testSheetAccess() {
  const results = [];
  const ss = SpreadsheetApp.getActive();
  
  // Test access to main sheets
  const requiredSheets = ['Packages', 'Packages_Archive'];
  const optionalSheets = ['PBI_Outbound_Staging', 'PowerBI_Import', 'PowerBI_New'];
  
  for (const sheetName of requiredSheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      results.push(`✓ Required sheet "${sheetName}" exists`);
    } else {
      results.push(`✗ Required sheet "${sheetName}" MISSING`);
    }
  }
  
  for (const sheetName of optionalSheets) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      results.push(`✓ Optional sheet "${sheetName}" exists`);
    } else {
      results.push(`ⓘ Optional sheet "${sheetName}" not found (okay)`);
    }
  }
  
  Logger.log('Sheet Access Test:\n' + results.join('\n'));
  return results;
}

function NF_testConstantsAvailability() {
  const results = [];
  
  // Test for SOK and Kärkkäinen constants
  if (typeof SOK_FREIGHT_ACCOUNT !== 'undefined') {
    results.push(`✓ SOK_FREIGHT_ACCOUNT constant available: ${SOK_FREIGHT_ACCOUNT}`);
  } else {
    results.push('✗ SOK_FREIGHT_ACCOUNT constant NOT found');
  }
  
  if (typeof KARKKAINEN_NUMBERS !== 'undefined') {
    results.push(`✓ KARKKAINEN_NUMBERS constant available: ${JSON.stringify(KARKKAINEN_NUMBERS)}`);
  } else {
    results.push('✗ KARKKAINEN_NUMBERS constant NOT found');
  }
  
  // Test for sheet name constants
  if (typeof TARGET_SHEET !== 'undefined') {
    results.push(`✓ TARGET_SHEET constant available: ${TARGET_SHEET}`);
  } else {
    results.push('✗ TARGET_SHEET constant NOT found');
  }
  
  if (typeof ARCHIVE_SHEET !== 'undefined') {
    results.push(`✓ ARCHIVE_SHEET constant available: ${ARCHIVE_SHEET}`);
  } else {
    results.push('✗ ARCHIVE_SHEET constant NOT found');
  }
  
  Logger.log('Constants Availability:\n' + results.join('\n'));
  return results;
}

function NF_testHelperFunctionCompatibility() {
  const results = [];
  
  // Test for normalize_ function
  if (typeof normalize_ === 'function') {
    try {
      const test = normalize_('Test String  123');
      results.push(`✓ normalize_ function works: "${test}"`);
    } catch (e) {
      results.push('✗ normalize_ function exists but failed: ' + e);
    }
  } else {
    results.push('✗ normalize_ function NOT found');
  }
  
  // Test for headerIndexMap_ function
  if (typeof headerIndexMap_ === 'function') {
    try {
      const test = headerIndexMap_(['Column1', 'Column2']);
      results.push('✓ headerIndexMap_ function works');
    } catch (e) {
      results.push('✗ headerIndexMap_ function exists but failed: ' + e);
    }
  } else {
    results.push('✗ headerIndexMap_ function NOT found');
  }
  
  // Test for colIndexOf_ function
  if (typeof colIndexOf_ === 'function') {
    try {
      const map = { 'test': 0 };
      const test = colIndexOf_(map, ['test']);
      results.push(`✓ colIndexOf_ function works: ${test}`);
    } catch (e) {
      results.push('✗ colIndexOf_ function exists but failed: ' + e);
    }
  } else {
    results.push('✗ colIndexOf_ function NOT found');
  }
  
  Logger.log('Helper Function Compatibility:\n' + results.join('\n'));
  return results;
}

function NF_testTriggerManagement() {
  const results = [];
  
  try {
    // Test creating and removing triggers
    const initialTriggers = ScriptApp.getProjectTriggers().length;
    
    // Test NF trigger setup (should not interfere with existing)
    NF_setupDaily1200();
    const afterDaily = ScriptApp.getProjectTriggers().length;
    
    NF_setupWeeklyMon0200();
    const afterWeekly = ScriptApp.getProjectTriggers().length;
    
    // Test removal
    NF_clearAllNFTriggers();
    const afterClear = ScriptApp.getProjectTriggers().length;
    
    if (afterDaily > initialTriggers && afterWeekly > afterDaily && afterClear === initialTriggers) {
      results.push('✓ Trigger management works correctly');
      results.push(`  Initial: ${initialTriggers}, After daily: ${afterDaily}, After weekly: ${afterWeekly}, After clear: ${afterClear}`);
    } else {
      results.push('✗ Trigger management test failed');
      results.push(`  Counts: ${initialTriggers} → ${afterDaily} → ${afterWeekly} → ${afterClear}`);
    }
  } catch (e) {
    results.push('✗ Trigger management threw error: ' + e);
  }
  
  Logger.log('Trigger Management Test:\n' + results.join('\n'));
  return results;
}

function NF_runIntegrationTests() {
  Logger.log('=== NewFlow Integration Test Suite ===\n');
  
  const results = {
    existingFunctions: NF_testExistingFunctionAvailability(),
    sheetAccess: NF_testSheetAccess(),
    constants: NF_testConstantsAvailability(),
    helperFunctions: NF_testHelperFunctionCompatibility(),
    triggerManagement: NF_testTriggerManagement()
  };
  
  // Count passed tests
  let passed = 0;
  let total = 0;
  
  Object.values(results).forEach(testResults => {
    testResults.forEach(result => {
      total++;
      if (result.startsWith('✓')) passed++;
    });
  });
  
  Logger.log(`\n=== Integration Test Summary: ${passed}/${total} passed ===`);
  
  if (passed === total) {
    SpreadsheetApp.getActive().toast('NewFlow integration tests passed! ✓', 'Integration Tests', 5);
  } else {
    SpreadsheetApp.getActive().toast(`Integration tests: ${passed}/${total} passed. See logs for details.`, 'Integration Tests', 10);
  }
  
  return { passed, total, results };
}

/**
 * Test complete NewFlow workflow with sample data
 */
function NF_testCompleteWorkflow() {
  Logger.log('=== NewFlow Complete Workflow Test ===');
  
  try {
    // Step 1: Test data source gathering (should handle missing sheets gracefully)
    Logger.log('Testing data source gathering...');
    NF_buildCountryLeadtime();
    Logger.log('✓ Country leadtime build completed (check sheets for results)');
    
    // Step 2: Test weekly reporting
    Logger.log('Testing weekly reporting...');
    NF_buildSokKarkkainenWeekly();
    Logger.log('✓ SOK & Kärkkäinen weekly reports completed');
    
    // Step 3: Test daily automation function
    Logger.log('Testing daily automation...');
    NF_runDaily();
    Logger.log('✓ Daily automation completed');
    
    // Step 4: Test weekly automation function
    Logger.log('Testing weekly automation...');
    NF_runWeekly();
    Logger.log('✓ Weekly automation completed');
    
    Logger.log('\n✓ Complete workflow test passed!');
    SpreadsheetApp.getActive().toast('NewFlow complete workflow test passed! ✓', 'Workflow Test', 5);
    
    return true;
  } catch (e) {
    Logger.log('✗ Complete workflow test failed: ' + e);
    SpreadsheetApp.getActive().toast('Workflow test failed. See logs for details.', 'Workflow Test', 10);
    return false;
  }
}