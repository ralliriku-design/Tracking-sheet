/**
 * NF_Test_Integration.gs — Integration tests for NF extension functionality
 * 
 * Simple tests to verify that the NF functions work correctly before full deployment.
 * These tests check basic functionality without making external API calls.
 */

/**
 * Test Drive import folder ID configuration.
 */
function NF_test_DriveImportConfig() {
  console.log('Testing Drive import configuration...');
  
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_IMPORT_FOLDER_ID');
  
  if (!folderId) {
    console.error('❌ DRIVE_IMPORT_FOLDER_ID not configured');
    return false;
  }
  
  if (folderId !== '1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh') {
    console.warn('⚠️ DRIVE_IMPORT_FOLDER_ID differs from expected value');
  }
  
  console.log('✅ Drive import folder ID configured:', folderId);
  return true;
}

/**
 * Test file pattern matching logic.
 */
function NF_test_FilePatternMatching() {
  console.log('Testing file pattern matching...');
  
  const testCases = [
    { filename: 'ERP_Quants_2025-01-15.xlsx', expectedType: 'quants' },
    { filename: 'Warehouse Balance Report.csv', expectedType: 'warehouse balance' },
    { filename: 'warehouse_balance_daily.xlsx', expectedType: 'warehouse balance' },
    { filename: 'Stock Picking Export.csv', expectedType: 'stock picking' },
    { filename: 'PBI_Deliveries_Weekly.xlsx', expectedType: 'deliveries' },
    { filename: 'pbi_shipment_data.csv', expectedType: 'deliveries' },
    { filename: 'PBI Balance Report.xlsx', expectedType: 'pbi balance' },
    { filename: 'Random_File.xlsx', expectedType: null }
  ];
  
  const importTypes = [
    { patterns: ['quants'], type: 'quants' },
    { patterns: ['warehouse balance', 'warehouse_balance', '3pl balance'], type: 'warehouse balance' },
    { patterns: ['stock picking', 'erp picking'], type: 'stock picking' },
    { patterns: ['deliveries', 'pbi deliveries', 'pbi_shipment', 'pbi_outbound'], type: 'deliveries' },
    { patterns: ['pbi balance', 'pbi stock', 'pbi inventory'], type: 'pbi balance' }
  ];
  
  let passed = 0;
  let failed = 0;
  
  for (const testCase of testCases) {
    const fileName = testCase.filename.toLowerCase();
    let matchedType = null;
    
    for (const importType of importTypes) {
      const matches = importType.patterns.some(pattern => 
        fileName.includes(pattern.toLowerCase())
      );
      if (matches) {
        matchedType = importType.type;
        break;
      }
    }
    
    if (matchedType === testCase.expectedType) {
      console.log(`✅ ${testCase.filename} -> ${matchedType || 'no match'}`);
      passed++;
    } else {
      console.error(`❌ ${testCase.filename} -> expected: ${testCase.expectedType}, got: ${matchedType}`);
      failed++;
    }
  }
  
  console.log(`Pattern matching test: ${passed} passed, ${failed} failed`);
  return failed === 0;
}

/**
 * Test header normalization.
 */
function NF_test_HeaderNormalization() {
  console.log('Testing header normalization...');
  
  const testMatrix = [
    [' Article ', '  Location  ', 'Quantity '],
    ['Product A', 'Warehouse 1', '100'],
    ['Product B', 'Warehouse 2', '200']
  ];
  
  const normalized = NF_Drive_NormalizeHeaders_(testMatrix);
  
  const expectedHeaders = ['Article', 'Location', 'Quantity'];
  const actualHeaders = normalized[0];
  
  let success = true;
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (actualHeaders[i] !== expectedHeaders[i]) {
      console.error(`❌ Header ${i}: expected '${expectedHeaders[i]}', got '${actualHeaders[i]}'`);
      success = false;
    }
  }
  
  // Check that data rows are unchanged
  if (normalized[1][0] !== 'Product A' || normalized[2][0] !== 'Product B') {
    console.error('❌ Data rows were modified during header normalization');
    success = false;
  }
  
  if (success) {
    console.log('✅ Header normalization works correctly');
  }
  
  return success;
}

/**
 * Test SOK/Kärkkäinen payer digit normalization.
 */
function NF_test_PayerNormalization() {
  console.log('Testing payer digit normalization...');
  
  const testCases = [
    { input: '5010', expected: '5010' },
    { input: 'Account: 5010', expected: '5010' },
    { input: '1234-ABC', expected: '1234' },
    { input: 'No digits here!', expected: '' }
  ];
  
  let passed = 0;
  
  for (const testCase of testCases) {
    const result = NF_normalizeDigits_(testCase.input);
    if (result === testCase.expected) {
      console.log(`✅ '${testCase.input}' -> '${result}'`);
      passed++;
    } else {
      console.error(`❌ '${testCase.input}' -> expected: '${testCase.expected}', got: '${result}'`);
    }
  }
  
  console.log(`Payer normalization test: ${passed}/${testCases.length} passed`);
  return passed === testCases.length;
}

/**
 * Test date window calculation.
 */
function NF_test_DateWindow() {
  console.log('Testing date window calculation...');
  
  try {
    const { start, end } = NF_getLastFinishedWeekSunWindow_();
    
    if (!start || !end) {
      console.error('❌ Date window calculation returned null values');
      return false;
    }
    
    if (start >= end) {
      console.error('❌ Start date is not before end date');
      return false;
    }
    
    const daysBetween = (end - start) / (24 * 60 * 60 * 1000);
    if (daysBetween !== 7) {
      console.error(`❌ Expected 7 days between start and end, got ${daysBetween}`);
      return false;
    }
    
    // Check that start is a Sunday (day 0)
    if (start.getDay() !== 0) {
      console.error(`❌ Start date is not a Sunday, got day ${start.getDay()}`);
      return false;
    }
    
    console.log(`✅ Date window: ${start.toISOString().split('T')[0]} to ${end.toISOString().split('T')[0]}`);
    return true;
    
  } catch (error) {
    console.error('❌ Date window calculation failed:', error);
    return false;
  }
}

/**
 * Test menu function availability.
 */
function NF_test_MenuFunctions() {
  console.log('Testing menu function availability...');
  
  const requiredFunctions = [
    'NF_BulkRebuildAll',
    'NF_BulkImportFromDrive', 
    'NF_RefreshAllPending',
    'NF_BulkFindDuplicates_All',
    'NF_UpdateInventoryBalances',
    'NF_buildSokKarkkainenAlways',
    'NF_setupDaily0001Inventory',
    'NF_setupDaily1100Tracking',
    'NF_setupWeeklyMon0200'
  ];
  
  let available = 0;
  
  for (const funcName of requiredFunctions) {
    if (typeof eval(funcName) === 'function') {
      console.log(`✅ ${funcName} available`);
      available++;
    } else {
      console.error(`❌ ${funcName} not available`);
    }
  }
  
  console.log(`Function availability test: ${available}/${requiredFunctions.length} available`);
  return available === requiredFunctions.length;
}

/**
 * Run all integration tests.
 */
function NF_runAllTests() {
  console.log('🧪 Starting NF Integration Tests...\n');
  
  const tests = [
    { name: 'Drive Import Config', func: NF_test_DriveImportConfig },
    { name: 'File Pattern Matching', func: NF_test_FilePatternMatching },
    { name: 'Header Normalization', func: NF_test_HeaderNormalization },
    { name: 'Payer Normalization', func: NF_test_PayerNormalization },
    { name: 'Date Window Calculation', func: NF_test_DateWindow },
    { name: 'Menu Functions', func: NF_test_MenuFunctions }
  ];
  
  let passed = 0;
  let failed = 0;
  
  for (const test of tests) {
    console.log(`\n--- ${test.name} ---`);
    try {
      if (test.func()) {
        passed++;
      } else {
        failed++;
      }
    } catch (error) {
      console.error(`❌ ${test.name} threw error:`, error);
      failed++;
    }
  }
  
  console.log(`\n📊 Test Summary: ${passed} passed, ${failed} failed`);
  
  if (failed === 0) {
    console.log('🎉 All tests passed! NF extension is ready for use.');
  } else {
    console.log('⚠️ Some tests failed. Review issues before deployment.');
  }
  
  return { passed, failed };
}

/**
 * Helper function for payer digit normalization (duplicated from NF_SOK_KRK_Weekly.gs for testing).
 */
function NF_normalizeDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}