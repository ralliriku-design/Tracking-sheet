/**
 * Tests for the Enhanced Tracking Orchestrator
 * 
 * This file contains test functions to verify the functionality of the enhanced
 * tracking orchestrator and its components.
 */

/**
 * Test the SafeCredentialManager
 */
function testSafeCredentialManager() {
  console.log('Testing SafeCredentialManager...');
  
  const credManager = new SafeCredentialManager();
  let passed = 0;
  let failed = 0;
  
  // Test 1: Valid credential setting and getting
  try {
    const success = credManager.safeSet('TEST_KEY', 'test-value-123', { minLength: 5 });
    if (success) {
      const value = credManager.safeGet('TEST_KEY', { minLength: 5 });
      if (value === 'test-value-123') {
        console.log('✓ Test 1 passed: Valid credential set/get');
        passed++;
      } else {
        console.log('✗ Test 1 failed: Retrieved value mismatch');
        failed++;
      }
    } else {
      console.log('✗ Test 1 failed: Failed to set credential');
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 1 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 2: Invalid credential (too short)
  try {
    const success = credManager.safeSet('TEST_KEY_SHORT', 'x', { minLength: 5 });
    if (!success) {
      console.log('✓ Test 2 passed: Correctly rejected short credential');
      passed++;
    } else {
      console.log('✗ Test 2 failed: Should have rejected short credential');
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 2 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 3: URL pattern validation
  try {
    const validUrl = credManager.safeSet('TEST_URL', 'https://api.example.com/track/{{code}}', 
      { pattern: /^https?:\/\/.*\{\{code\}\}/ });
    const invalidUrl = credManager.safeSet('TEST_URL_INVALID', 'invalid-url', 
      { pattern: /^https?:\/\/.*\{\{code\}\}/ });
    
    if (validUrl && !invalidUrl) {
      console.log('✓ Test 3 passed: URL pattern validation works');
      passed++;
    } else {
      console.log('✗ Test 3 failed: URL pattern validation failed');
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 3 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 4: Carrier credential validation
  try {
    // Set up some test credentials for DHL
    credManager.safeSet('DHL_TRACK_URL', 'https://api.dhl.com/track/{{code}}');
    credManager.safeSet('DHL_API_KEY', 'test-api-key-12345');
    
    const validation = credManager.validateCarrierCredentials('DHL');
    if (validation.valid) {
      console.log('✓ Test 4 passed: Carrier credential validation works');
      passed++;
    } else {
      console.log('✗ Test 4 failed: Carrier credential validation failed');
      console.log('  Validation result:', validation);
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 4 failed: Exception -', error.message);
    failed++;
  }
  
  // Cleanup test credentials
  try {
    credManager.safeSet('TEST_KEY', null);
    credManager.safeSet('TEST_KEY_SHORT', null);
    credManager.safeSet('TEST_URL', null);
    credManager.safeSet('TEST_URL_INVALID', null);
    credManager.safeSet('DHL_TRACK_URL', null);
    credManager.safeSet('DHL_API_KEY', null);
  } catch (error) {
    console.log('Warning: Failed to cleanup test credentials:', error.message);
  }
  
  console.log(`SafeCredentialManager tests completed: ${passed} passed, ${failed} failed`);
  return { passed, failed };
}

/**
 * Test the CircuitBreaker
 */
function testCircuitBreaker() {
  console.log('Testing CircuitBreaker...');
  
  let passed = 0;
  let failed = 0;
  
  // Test 1: Basic circuit breaker functionality
  try {
    const breaker = new CircuitBreaker('TEST_CARRIER', { 
      failureThreshold: 2, 
      recoveryTimeout: 100 
    });
    
    // Initially should be CLOSED
    if (breaker.getStatus().state === 'CLOSED') {
      console.log('✓ Test 1a passed: Initial state is CLOSED');
      passed++;
    } else {
      console.log('✗ Test 1a failed: Initial state should be CLOSED');
      failed++;
    }
    
    // First failure
    try {
      breaker.execute(() => { throw new Error('Test failure'); });
    } catch (e) {
      // Expected to fail
    }
    
    // Second failure should open the circuit
    try {
      breaker.execute(() => { throw new Error('Test failure'); });
    } catch (e) {
      // Expected to fail
    }
    
    // Circuit should now be OPEN
    if (breaker.getStatus().state === 'OPEN') {
      console.log('✓ Test 1b passed: Circuit opened after failures');
      passed++;
    } else {
      console.log('✗ Test 1b failed: Circuit should be OPEN after failures');
      failed++;
    }
    
    // Reset for cleanup
    breaker.reset();
    
  } catch (error) {
    console.log('✗ Test 1 failed: Exception -', error.message);
    failed++;
  }
  
  console.log(`CircuitBreaker tests completed: ${passed} passed, ${failed} failed`);
  return { passed, failed };
}

/**
 * Test the RetryManager
 */
function testRetryManager() {
  console.log('Testing RetryManager...');
  
  let passed = 0;
  let failed = 0;
  
  // Test 1: Successful execution without retries
  try {
    const retryManager = new RetryManager({ maxRetries: 2, baseDelay: 10 });
    let callCount = 0;
    
    const result = retryManager.executeWithRetry(() => {
      callCount++;
      return 'success';
    });
    
    if (result === 'success' && callCount === 1) {
      console.log('✓ Test 1 passed: Successful execution without retries');
      passed++;
    } else {
      console.log('✗ Test 1 failed: Unexpected behavior');
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 1 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 2: Retry on retryable error
  try {
    const retryManager = new RetryManager({ maxRetries: 2, baseDelay: 10 });
    let callCount = 0;
    
    try {
      retryManager.executeWithRetry(() => {
        callCount++;
        if (callCount < 3) {
          const error = new Error('HTTP_429');
          error.status = 'RATE_LIMIT_429';
          throw error;
        }
        return 'success';
      });
      
      console.log('✓ Test 2 passed: Retry mechanism works');
      passed++;
    } catch (error) {
      if (callCount === 3) {
        console.log('✓ Test 2 passed: Retried correct number of times');
        passed++;
      } else {
        console.log('✗ Test 2 failed: Incorrect retry count');
        failed++;
      }
    }
  } catch (error) {
    console.log('✗ Test 2 failed: Exception -', error.message);
    failed++;
  }
  
  console.log(`RetryManager tests completed: ${passed} passed, ${failed} failed`);
  return { passed, failed };
}

/**
 * Test the TrackingOrchestrator
 */
function testTrackingOrchestrator() {
  console.log('Testing TrackingOrchestrator...');
  
  let passed = 0;
  let failed = 0;
  
  // Test 1: Orchestrator initialization
  try {
    const orchestrator = new TrackingOrchestrator();
    orchestrator.initialize();
    
    if (orchestrator.isRunning) {
      console.log('✓ Test 1 passed: Orchestrator initializes correctly');
      passed++;
    } else {
      console.log('✗ Test 1 failed: Orchestrator failed to initialize');
      failed++;
    }
    
    // Test 2: Job submission
    try {
      const jobId = orchestrator.submitJob({
        carrier: 'TEST_CARRIER',
        trackingCode: 'TEST123'
      });
      
      if (jobId && typeof jobId === 'string') {
        console.log('✓ Test 2 passed: Job submission works');
        passed++;
      } else {
        console.log('✗ Test 2 failed: Job submission failed');
        failed++;
      }
    } catch (error) {
      console.log('✗ Test 2 failed: Exception -', error.message);
      failed++;
    }
    
    // Test 3: Status retrieval
    try {
      const status = orchestrator.getStatus();
      
      if (status && typeof status === 'object' && 'isRunning' in status) {
        console.log('✓ Test 3 passed: Status retrieval works');
        passed++;
      } else {
        console.log('✗ Test 3 failed: Status retrieval failed');
        failed++;
      }
    } catch (error) {
      console.log('✗ Test 3 failed: Exception -', error.message);
      failed++;
    }
    
    // Cleanup
    orchestrator.shutdown();
    
  } catch (error) {
    console.log('✗ Test 1 failed: Exception -', error.message);
    failed++;
  }
  
  console.log(`TrackingOrchestrator tests completed: ${passed} passed, ${failed} failed`);
  return { passed, failed };
}

/**
 * Test helper functions
 */
function testHelperFunctions() {
  console.log('Testing helper functions...');
  
  let passed = 0;
  let failed = 0;
  
  // Test 1: normalize_ function
  try {
    const normalized = normalize_('Tracking Number');
    if (normalized === 'tracking number') {
      console.log('✓ Test 1 passed: normalize_ function works');
      passed++;
    } else {
      console.log('✗ Test 1 failed: normalize_ function result:', normalized);
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 1 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 2: headerIndexMap_ function
  try {
    const headers = ['Package Number', 'Tracking Code', 'Carrier'];
    const map = headerIndexMap_(headers);
    
    if (map['Package Number'] === 0 && map['Tracking Code'] === 1 && map['Carrier'] === 2) {
      console.log('✓ Test 2 passed: headerIndexMap_ function works');
      passed++;
    } else {
      console.log('✗ Test 2 failed: headerIndexMap_ function result:', map);
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 2 failed: Exception -', error.message);
    failed++;
  }
  
  // Test 3: colIndexOf_ function
  try {
    const headers = ['Package Number', 'Tracking Code', 'Carrier'];
    const map = headerIndexMap_(headers);
    const trackingIndex = colIndexOf_(map, TRACKING_CANDIDATES);
    
    if (trackingIndex === 1) { // Should find 'Tracking Code'
      console.log('✓ Test 3 passed: colIndexOf_ function works');
      passed++;
    } else {
      console.log('✗ Test 3 failed: colIndexOf_ function result:', trackingIndex);
      failed++;
    }
  } catch (error) {
    console.log('✗ Test 3 failed: Exception -', error.message);
    failed++;
  }
  
  console.log(`Helper function tests completed: ${passed} passed, ${failed} failed`);
  return { passed, failed };
}

/**
 * Run all tests
 */
function runAllOrchestratorTests() {
  console.log('=== Running Enhanced Tracking Orchestrator Tests ===\n');
  
  const results = {
    credentialManager: testSafeCredentialManager(),
    circuitBreaker: testCircuitBreaker(),
    retryManager: testRetryManager(),
    orchestrator: testTrackingOrchestrator(),
    helpers: testHelperFunctions()
  };
  
  const totalPassed = Object.values(results).reduce((sum, r) => sum + r.passed, 0);
  const totalFailed = Object.values(results).reduce((sum, r) => sum + r.failed, 0);
  
  console.log('\n=== Test Summary ===');
  Object.entries(results).forEach(([component, result]) => {
    console.log(`${component}: ${result.passed} passed, ${result.failed} failed`);
  });
  
  console.log(`\nOverall: ${totalPassed} passed, ${totalFailed} failed`);
  
  if (totalFailed === 0) {
    console.log('🎉 All tests passed!');
    SpreadsheetApp.getUi().alert('All orchestrator tests passed! ✓');
  } else {
    console.log('❌ Some tests failed. Check the console for details.');
    SpreadsheetApp.getUi().alert(`Tests completed: ${totalPassed} passed, ${totalFailed} failed. Check console for details.`);
  }
  
  return results;
}

/**
 * Quick smoke test for basic functionality
 */
function quickOrchestratorTest() {
  try {
    console.log('Running quick orchestrator smoke test...');
    
    // Test credential manager
    const credManager = new SafeCredentialManager();
    credManager.safeSet('TEST_SMOKE', 'smoke-test-value');
    const value = credManager.safeGet('TEST_SMOKE');
    
    if (value !== 'smoke-test-value') {
      throw new Error('Credential manager smoke test failed');
    }
    
    // Test orchestrator creation
    const orchestrator = new TrackingOrchestrator();
    orchestrator.initialize();
    
    if (!orchestrator.isRunning) {
      throw new Error('Orchestrator initialization smoke test failed');
    }
    
    // Cleanup
    credManager.safeSet('TEST_SMOKE', null);
    orchestrator.shutdown();
    
    console.log('✓ Quick smoke test passed');
    SpreadsheetApp.getUi().alert('Quick smoke test passed! ✓');
    
  } catch (error) {
    console.error('✗ Quick smoke test failed:', error.message);
    SpreadsheetApp.getUi().alert(`Quick smoke test failed: ${error.message}`);
  }
}