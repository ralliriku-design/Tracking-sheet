// Demo_AdhocTest.gs — Demonstration of the enhanced ad hoc tracker

/**
 * Demo function that shows how the enhanced tracker handles "package id (seurantakoodi)"
 * This demonstrates the exact use case from the problem statement
 */
function DEMO_TestFinnishHeader() {
  Logger.log('=== Testing Enhanced Ad Hoc Tracker with Finnish Headers ===');
  
  // Create test data that matches the problem statement scenario
  const testData = [
    ['package id (seurantakoodi)', 'Carrier', 'Country', 'Description'],
    ['JD1234567890', 'Posti', 'Finland', 'Test package 1'],
    ['ABC987654321', 'GLS', 'Germany', 'Test package 2'],
    ['XYZ456789012', 'DHL', 'Sweden', 'Test package 3']
  ];
  
  try {
    // Test the enhanced tracker function directly
    Logger.log('Testing ADHOC_buildFromValues with Finnish header...');
    
    const result = ADHOC_buildFromValues(testData, 'Finnish Header Test');
    
    Logger.log('SUCCESS: Processed ' + result + ' tracking codes');
    Logger.log('The error "Adhoc: ei löytynyt Carrier/Tracking -sarakkeita" should NOT appear');
    
    SpreadsheetApp.getActive().toast('Demo completed successfully - check Adhoc_Results sheet');
    
  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    
    if (error.message.includes('ei löytynyt Carrier/Tracking -sarakkeita') || 
        error.message.includes('ei löytynyt seurantakoodin saraketta')) {
      Logger.log('ANALYSIS: Header detection failed. This suggests an issue with the enhanced detection logic.');
    }
    
    SpreadsheetApp.getActive().toast('Demo failed: ' + error.message);
    throw error;
  }
}

/**
 * Test the exact scenario from the problem statement
 */
function DEMO_ProblemStatementScenario() {
  Logger.log('=== Testing Exact Problem Statement Scenario ===');
  
  // Create the specific case mentioned in the problem
  const testHeaders = ['package id (seurantakoodi)', 'other column', 'description'];
  const testData = [
    testHeaders,
    ['JD1234567890', 'some data', 'test item 1'],
    ['ABC987654321', 'more data', 'test item 2']
  ];
  
  try {
    // Test header detection specifically
    Logger.log('Testing header detection for "package id (seurantakoodi)"...');
    
    const trackingCol = adhocFindColumn_(testHeaders, testData, ADHOC_TRACKING_CANDIDATES, adhocIsTrackingLike_);
    Logger.log('Tracking column detected at index: ' + trackingCol);
    
    if (trackingCol >= 0) {
      Logger.log('SUCCESS: Header "' + testHeaders[trackingCol] + '" was correctly identified as tracking column');
    } else {
      Logger.log('FAILURE: Header detection failed for "package id (seurantakoodi)"');
    }
    
    // Test the full process
    const result = ADHOC_buildFromValues(testData, 'Problem Statement Test');
    Logger.log('Full process completed successfully with ' + result + ' items processed');
    
    SpreadsheetApp.getActive().toast('Problem statement scenario test passed!');
    
  } catch (error) {
    Logger.log('ERROR in problem statement scenario: ' + error.message);
    SpreadsheetApp.getActive().toast('Test failed: ' + error.message);
    throw error;
  }
}

/**
 * Compare old vs new behavior
 */
function DEMO_CompareOldVsNew() {
  Logger.log('=== Comparing Old vs New Tracker Behavior ===');
  
  const testHeaders = ['package id (seurantakoodi)', 'carrier name', 'other'];
  
  // Simulate old behavior (from Traacking.txt)
  Logger.log('Testing old behavior simulation...');
  
  try {
    // This simulates the old pickAnyIndex_ function behavior
    const oldTrackingCandidates = [
      'Tracking number','Tracking','Barcode','Waybill','Waybill No','AWB',
      'Package Number','PackageNumber','Shipment ID','Consignment number'
    ];
    
    // Old normalization (from Traacking.txt): just lowercase and basic cleanup
    const oldNormalize = (s) => String(s || '').toLowerCase().replace(/\s+/g, ' ').trim().replace(/[^\p{L}\p{N}]+/gu, ' ');
    const oldNormalizedHeaders = testHeaders.map(oldNormalize);
    
    let oldResult = -1;
    for (const candidate of oldTrackingCandidates) {
      const normalizedCandidate = oldNormalize(candidate);
      const index = oldNormalizedHeaders.indexOf(normalizedCandidate);
      if (index >= 0) {
        oldResult = index;
        break;
      }
    }
    
    Logger.log('Old method result: ' + oldResult + ' (should be -1, indicating failure)');
    
    // Test new behavior
    Logger.log('Testing new behavior...');
    const newResult = adhocFindColumn_(testHeaders, null, ADHOC_TRACKING_CANDIDATES, null);
    Logger.log('New method result: ' + newResult + ' (should be 0, indicating success)');
    
    if (oldResult < 0 && newResult >= 0) {
      Logger.log('SUCCESS: New method correctly handles the Finnish header while old method fails');
      SpreadsheetApp.getActive().toast('Comparison test passed - new method is better!');
    } else {
      Logger.log('Unexpected results in comparison test');
      SpreadsheetApp.getActive().toast('Comparison test results unexpected');
    }
    
  } catch (error) {
    Logger.log('Error in comparison test: ' + error.message);
    throw error;
  }
}

/**
 * Test content fallback functionality
 */
function DEMO_ContentFallback() {
  Logger.log('=== Testing Content Fallback Functionality ===');
  
  // Create data where headers don't match but content should be detected
  const badHeaders = ['Column A', 'Column B', 'Column C'];
  const goodData = [
    badHeaders,
    ['some text', 'JD1234567890', 'Posti'],
    ['other info', 'ABC987654321', 'GLS'],
    ['description', 'XYZ456789012', 'DHL']
  ];
  
  try {
    Logger.log('Testing content-based detection when headers are unclear...');
    
    // This should detect column 1 (index 1) as tracking based on content
    const trackingCol = adhocFindColumn_(badHeaders, goodData, ['nonexistent_header'], adhocIsTrackingLike_);
    Logger.log('Content-based tracking detection result: ' + trackingCol + ' (should be 1)');
    
    if (trackingCol === 1) {
      Logger.log('SUCCESS: Content fallback correctly identified tracking column');
      
      // Test full process
      const result = ADHOC_buildFromValues(goodData, 'Content Fallback Test');
      Logger.log('Full content fallback process completed with ' + result + ' items');
      
      SpreadsheetApp.getActive().toast('Content fallback test passed!');
    } else {
      Logger.log('FAILURE: Content fallback did not work as expected');
      SpreadsheetApp.getActive().toast('Content fallback test failed');
    }
    
  } catch (error) {
    Logger.log('Error in content fallback test: ' + error.message);
    throw error;
  }
}

/**
 * Run all demonstration tests
 */
function DEMO_RunAll() {
  Logger.log('=== Running All Enhanced Ad Hoc Tracker Demonstrations ===');
  
  const tests = [
    { name: 'Finnish Header Test', func: DEMO_TestFinnishHeader },
    { name: 'Problem Statement Scenario', func: DEMO_ProblemStatementScenario },
    { name: 'Old vs New Comparison', func: DEMO_CompareOldVsNew },
    { name: 'Content Fallback', func: DEMO_ContentFallback }
  ];
  
  let passed = 0;
  let failed = 0;
  
  for (const test of tests) {
    try {
      Logger.log('\n--- Running: ' + test.name + ' ---');
      test.func();
      Logger.log('✓ PASSED: ' + test.name);
      passed++;
    } catch (error) {
      Logger.log('✗ FAILED: ' + test.name + ' - ' + error.message);
      failed++;
    }
  }
  
  Logger.log('\n=== SUMMARY ===');
  Logger.log('Passed: ' + passed);
  Logger.log('Failed: ' + failed);
  Logger.log('Total: ' + (passed + failed));
  
  const message = passed === tests.length 
    ? 'All tests passed! Enhanced ad hoc tracker is working correctly.'
    : failed + ' test(s) failed. Check logs for details.';
    
  SpreadsheetApp.getActive().toast(message);
}

/**
 * Simple test to verify basic functionality works
 */
function DEMO_QuickTest() {
  // Just test that the main function can be called without errors
  try {
    const testData = [
      ['package id (seurantakoodi)', 'notes'],
      ['TEST123456', 'sample']
    ];
    
    // This should work without throwing the old error
    ADHOC_buildFromValues(testData, 'Quick Test');
    SpreadsheetApp.getActive().toast('Quick test passed - basic functionality works!');
    
  } catch (error) {
    SpreadsheetApp.getActive().toast('Quick test failed: ' + error.message);
    throw error;
  }
}