// Test_AdhocTracker.gs — Test functions for the enhanced ad hoc tracker

/**
 * Test the header detection functionality
 */
function TEST_AdhocHeaderDetection() {
  // Test data with various header formats
  const testHeaders = [
    'Package ID (seurantakoodi)', 'Carrier Name', 'Country',
    'Other Column', 'Description', 'Weight'
  ];
  
  const testData = [
    testHeaders,
    ['ABC123456', 'Posti', 'Finland', 'Test', 'Sample package', '1.5'],
    ['DEF789012', 'GLS', 'Germany', 'Test2', 'Another package', '2.1'],
    ['GHI345678', 'DHL', 'Sweden', 'Test3', 'Third package', '0.8']
  ];
  
  // Test tracking column detection
  Logger.log('Testing header detection...');
  
  // Should find column 0 (Package ID (seurantakoodi))
  const trackingCol = adhocFindColumn_(testHeaders, testData, ADHOC_TRACKING_CANDIDATES, adhocIsTrackingLike_);
  Logger.log('Tracking column index: ' + trackingCol + ' (should be 0)');
  
  // Should find column 1 (Carrier Name)
  const carrierCol = adhocFindColumn_(testHeaders, testData, ADHOC_CARRIER_CANDIDATES, adhocIsCarrierLike_);
  Logger.log('Carrier column index: ' + carrierCol + ' (should be 1)');
  
  // Should find column 2 (Country)
  const countryCol = adhocFindColumn_(testHeaders, testData, ADHOC_COUNTRY_CANDIDATES, null);
  Logger.log('Country column index: ' + countryCol + ' (should be 2)');
  
  // Test normalize function
  Logger.log('Testing normalize function...');
  const testNormalization = [
    'Package ID (seurantakoodi)',
    'TRACKING NUMBER',
    'Carrier-Name',
    'Country_Code'
  ];
  
  testNormalization.forEach(header => {
    Logger.log(header + ' -> ' + adhocNormalize_(header));
  });
  
  SpreadsheetApp.getActive().toast('Header detection test completed - check logs');
}

/**
 * Test the tracking code validation
 */
function TEST_AdhocTrackingValidation() {
  const testCodes = [
    'ABC123456789',     // Should be valid
    'JD0001234567',     // Should be valid
    '1234567890123',    // Should be valid
    'FI123456789FI',    // Should be valid (postal format)
    'ABC',              // Should be invalid (too short)
    '   ',              // Should be invalid (empty)
    '!!!@@@',           // Should be invalid (no alphanumeric)
    '12345678'          // Should be valid (numeric)
  ];
  
  Logger.log('Testing tracking code validation...');
  
  testCodes.forEach(code => {
    const isValid = adhocIsTrackingLike_(code);
    Logger.log(code + ' -> ' + (isValid ? 'VALID' : 'INVALID'));
  });
  
  SpreadsheetApp.getActive().toast('Tracking validation test completed - check logs');
}

/**
 * Test the carrier validation
 */
function TEST_AdhocCarrierValidation() {
  const testCarriers = [
    'Posti',           // Should be valid
    'GLS Finland',     // Should be valid
    'DHL Express',     // Should be valid
    'Unknown Carrier', // Should be valid
    'ABC',             // Should be valid
    '123',             // Should be invalid
    '',                // Should be invalid
    '   '              // Should be invalid
  ];
  
  Logger.log('Testing carrier validation...');
  
  testCarriers.forEach(carrier => {
    const isValid = adhocIsCarrierLike_(carrier);
    Logger.log(carrier + ' -> ' + (isValid ? 'VALID' : 'INVALID'));
  });
  
  SpreadsheetApp.getActive().toast('Carrier validation test completed - check logs');
}

/**
 * Test content-based fallback detection
 */
function TEST_AdhocContentFallback() {
  // Test data where headers don't match but content should be detected
  const testHeaders = ['Col1', 'Col2', 'Col3', 'Col4'];
  const testData = [
    testHeaders,
    ['Some text', 'ABC123456789', 'Posti', 'Finland'],
    ['Other data', 'DEF987654321', 'GLS', 'Germany'],
    ['More info', 'GHI567890123', 'DHL', 'Sweden'],
    ['Description', 'JKL234567890', 'Bring', 'Norway']
  ];
  
  Logger.log('Testing content-based fallback detection...');
  
  // Should detect column 1 as tracking (based on content)
  const trackingCol = adhocFindColumn_(testHeaders, testData, ['nonexistent'], adhocIsTrackingLike_);
  Logger.log('Tracking column (content-based): ' + trackingCol + ' (should be 1)');
  
  // Should detect column 2 as carrier (based on content)
  const carrierCol = adhocFindColumn_(testHeaders, testData, ['nonexistent'], adhocIsCarrierLike_);
  Logger.log('Carrier column (content-based): ' + carrierCol + ' (should be 2)');
  
  SpreadsheetApp.getActive().toast('Content fallback test completed - check logs');
}

/**
 * Test ISO week calculation
 */
function TEST_AdhocISOWeek() {
  const testDates = [
    '2024-01-01',   // Should be 2024-W01
    '2024-12-31',   // Should be 2025-W01 (week belongs to next year)
    '2024-06-15',   // Should be 2024-W24
    '2023-12-31'    // Should be 2023-W52
  ];
  
  Logger.log('Testing ISO week calculation...');
  
  testDates.forEach(dateStr => {
    const date = new Date(dateStr);
    const week = getISOWeek_(date);
    Logger.log(dateStr + ' -> ' + week);
  });
  
  SpreadsheetApp.getActive().toast('ISO week test completed - check logs');
}

/**
 * Run all tests
 */
function TEST_AdhocAll() {
  Logger.log('=== Running all ad hoc tracker tests ===');
  
  try {
    TEST_AdhocHeaderDetection();
    TEST_AdhocTrackingValidation(); 
    TEST_AdhocCarrierValidation();
    TEST_AdhocContentFallback();
    TEST_AdhocISOWeek();
    
    SpreadsheetApp.getActive().toast('All ad hoc tracker tests completed successfully');
  } catch (error) {
    Logger.log('Test failed: ' + error.message);
    SpreadsheetApp.getActive().toast('Test failed: ' + error.message);
  }
}

/**
 * Create test data sheet for manual testing
 */
function TEST_CreateAdhocTestData() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('TEST_AdhocData') || ss.insertSheet('TEST_AdhocData');
  
  const testData = [
    ['Package ID (seurantakoodi)', 'Carrier Name', 'Destination Country', 'Description', 'Weight'],
    ['ABC123456789', 'Posti', 'Finland', 'Test package 1', '1.5'],
    ['DEF987654321', 'GLS', 'Germany', 'Test package 2', '2.1'],
    ['GHI567890123', 'DHL', 'Sweden', 'Test package 3', '0.8'],
    ['JKL234567890', 'Bring', 'Norway', 'Test package 4', '1.2'],
    ['MNO345678901', 'Matkahuolto', 'Finland', 'Test package 5', '3.0']
  ];
  
  sheet.clear();
  sheet.getRange(1, 1, testData.length, testData[0].length).setValues(testData);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, testData[0].length);
  
  ss.setActiveSheet(sheet);
  ss.toast('Test data created in TEST_AdhocData sheet');
}