/**
 * NF_Test.gs - Simple tests for New Flow functionality
 * 
 * Basic validation tests for key functions to ensure they work correctly.
 */

function NF_testBasicFunctionality() {
  const testResults = [];
  
  try {
    // Test 1: Header index mapping
    const headers = ['Package Number', 'Carrier', 'Country', 'Pick up date', 'Submitted date'];
    const map = NF_headerIndexMap_(headers);
    testResults.push({
      test: 'headerIndexMap_',
      result: map['Carrier'] === 1 ? 'PASS' : 'FAIL',
      details: `Carrier index: ${map['Carrier']}`
    });
    
    // Test 2: Start time parsing logic
    const mockRow = ['PKG123', 'Posti', 'FI', '2025-01-10', '2025-01-08', '2025-01-05'];
    const startInfo = NF_parseStartTime_(mockRow, map);
    testResults.push({
      test: 'parseStartTime_ (pickup priority)',
      result: startInfo.source === 'pickup' && startInfo.time === '2025-01-10' ? 'PASS' : 'FAIL',
      details: `Source: ${startInfo.source}, Time: ${startInfo.time}`
    });
    
    // Test 3: Week window calculation
    const testDate = new Date('2025-01-15'); // Wednesday
    const fakeNow = new Date(testDate);
    const window = calculateTestWeekWindow_(fakeNow);
    testResults.push({
      test: 'week window calculation',
      result: window.start.getDay() === 0 ? 'PASS' : 'FAIL', // Should be Sunday
      details: `Start: ${window.start.toISOString().split('T')[0]}, End: ${window.end.toISOString().split('T')[0]}`
    });
    
    // Test 4: SOK/Kärkkäinen number normalization
    const sokTest = NF_normalizeDigits_('990-719-901');
    const krkTest = NF_normalizeDigits_('615471');
    testResults.push({
      test: 'digit normalization',
      result: sokTest === '990719901' && krkTest === '615471' ? 'PASS' : 'FAIL',
      details: `SOK: ${sokTest}, KRK: ${krkTest}`
    });
    
    // Test 5: ISO week formatting
    const testISOWeek = NF_yearIsoWeek_(new Date('2025-01-15'));
    testResults.push({
      test: 'ISO week formatting',
      result: testISOWeek.includes('2025') && testISOWeek.includes('W') ? 'PASS' : 'FAIL',
      details: `ISO Week: ${testISOWeek}`
    });
    
  } catch (error) {
    testResults.push({
      test: 'exception handling',
      result: 'FAIL',
      details: error.message
    });
  }
  
  // Display results
  let summary = '=== NEW FLOW TEST RESULTS ===\n\n';
  let passCount = 0;
  
  testResults.forEach((result, index) => {
    summary += `${index + 1}. ${result.test}: ${result.result}\n`;
    summary += `   ${result.details}\n\n`;
    if (result.result === 'PASS') passCount++;
  });
  
  summary += `SUMMARY: ${passCount}/${testResults.length} tests passed`;
  
  SpreadsheetApp.getUi().alert('Test Results', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  
  return testResults;
}

/**
 * Test the menu trigger installation
 */
function NF_testMenuInstallation() {
  try {
    // Check if onOpen function exists and is callable
    if (typeof NF_onOpen !== 'function') {
      throw new Error('NF_onOpen function not found');
    }
    
    // Check if installation function exists
    if (typeof NF_installMenuTrigger !== 'function') {
      throw new Error('NF_installMenuTrigger function not found');
    }
    
    // Check current triggers
    const triggers = ScriptApp.getProjectTriggers();
    const nfTriggers = triggers.filter(t => t.getHandlerFunction().startsWith('NF_'));
    
    SpreadsheetApp.getUi().alert(
      'Menu Test Results',
      `✓ NF_onOpen function: Available\n✓ NF_installMenuTrigger function: Available\n✓ Current NF triggers: ${nfTriggers.length}\n\nReady for installation!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    SpreadsheetApp.getUi().alert('Menu Test Failed', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * Test Script Properties configuration
 */
function NF_testScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  const requiredProps = {
    'SOK_FREIGHT_ACCOUNT': '990719901',
    'KARKKAINEN_NUMBERS': '615471,802669,7030057',
    'TARGET_SHEET': 'Packages',
    'ARCHIVE_SHEET': 'Packages_Archive'
  };
  
  let report = '=== SCRIPT PROPERTIES TEST ===\n\n';
  let allGood = true;
  
  Object.keys(requiredProps).forEach(key => {
    const value = props.getProperty(key);
    const expected = requiredProps[key];
    const status = value ? '✓' : '✗';
    
    report += `${key}: ${status}\n`;
    report += `  Current: ${value || 'NOT SET'}\n`;
    report += `  Expected: ${expected}\n\n`;
    
    if (!value) allGood = false;
  });
  
  if (allGood) {
    report += 'All required properties are configured!';
  } else {
    report += 'Some properties need to be set. Use File → Project properties → Script properties in Apps Script.';
  }
  
  SpreadsheetApp.getUi().alert('Properties Test', report, SpreadsheetApp.getUi().ButtonSet.OK);
  return allGood;
}

/**
 * Helper function for testing week calculation
 */
function calculateTestWeekWindow_(testDate) {
  const now = new Date(testDate);
  now.setHours(0, 0, 0, 0);
  const dayOfWeek = now.getDay(); // 0 = Sunday
  
  // Calculate this Sunday
  const thisSunday = new Date(now);
  thisSunday.setDate(now.getDate() - dayOfWeek);
  
  // Last finished week is the previous week
  const end = thisSunday;
  const start = new Date(end);
  start.setDate(end.getDate() - 7);
  
  return { start, end };
}

/**
 * Run all tests at once
 */
function NF_runAllTests() {
  console.log('Running all New Flow tests...');
  
  const basicResults = NF_testBasicFunctionality();
  const menuResults = NF_testMenuInstallation();
  const propsResults = NF_testScriptProperties();
  
  console.log('Test results:', { basic: basicResults, menu: menuResults, props: propsResults });
}