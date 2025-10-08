/**
 * Enhanced Tracking Orchestrator Usage Examples
 * 
 * This file demonstrates how to use the enhanced tracking orchestrator
 * in your Google Apps Script projects.
 */

/**
 * Example 1: Basic setup and credential validation
 */
function example1_BasicSetup() {
  console.log('=== Example 1: Basic Setup ===');
  
  try {
    // Initialize the orchestrator
    const orchestrator = getTrackingOrchestrator();
    orchestrator.initialize();
    
    console.log('Orchestrator initialized successfully');
    
    // Validate credentials for a specific carrier
    const credManager = new SafeCredentialManager();
    const validation = credManager.validateCarrierCredentials('DHL');
    
    console.log('DHL Credential Validation:', validation);
    
    if (!validation.valid) {
      console.log('Setting up example DHL credentials...');
      credManager.safeSet('DHL_TRACK_URL', 'https://api.dhl.com/track/{{code}}');
      credManager.safeSet('DHL_API_KEY', 'your-dhl-api-key-here');
    }
    
    // Show orchestrator status
    const status = orchestrator.getStatus();
    console.log('Orchestrator Status:', status);
    
  } catch (error) {
    console.error('Example 1 failed:', error.message);
  }
}

/**
 * Example 2: Submit tracking jobs programmatically
 */
function example2_SubmitJobs() {
  console.log('=== Example 2: Submit Tracking Jobs ===');
  
  try {
    const orchestrator = getTrackingOrchestrator();
    
    // Example tracking jobs
    const jobs = [
      { carrier: 'DHL', trackingCode: 'DHL123456789' },
      { carrier: 'Posti', trackingCode: 'POSTI987654321' },
      { carrier: 'GLS', trackingCode: 'GLS555666777' }
    ];
    
    // Submit jobs
    const jobIds = [];
    jobs.forEach(job => {
      try {
        const jobId = orchestrator.submitJob(job);
        jobIds.push(jobId);
        console.log(`Submitted job ${jobId} for ${job.carrier}: ${job.trackingCode}`);
      } catch (error) {
        console.error(`Failed to submit job for ${job.carrier}:`, error.message);
      }
    });
    
    // Process the queue
    orchestrator.processQueue();
    
    console.log(`Submitted ${jobIds.length} jobs total`);
    
  } catch (error) {
    console.error('Example 2 failed:', error.message);
  }
}

/**
 * Example 3: Enhanced bulk tracking with error handling
 */
async function example3_EnhancedBulkTracking() {
  console.log('=== Example 3: Enhanced Bulk Tracking ===');
  
  try {
    // This would typically be called for a real sheet
    // For demo, we'll show the process
    
    const sheetName = 'Packages'; // Example sheet
    const carrierFilter = null; // No filter, process all carriers
    
    console.log(`Starting enhanced bulk tracking for sheet: ${sheetName}`);
    
    // Initialize orchestrator
    await setupTrackingOrchestrator();
    
    // In real usage, this would process the actual sheet
    await enhancedBulkStart(sheetName, carrierFilter);
    
    console.log('Enhanced bulk tracking initiated');
    
    // Monitor progress
    setTimeout(() => {
      const status = getOrchestratorStatus();
      console.log('Current orchestrator status:', status);
    }, 5000);
    
  } catch (error) {
    console.error('Example 3 failed:', error.message);
  }
}

/**
 * Example 4: Circuit breaker demonstration
 */
function example4_CircuitBreakerDemo() {
  console.log('=== Example 4: Circuit Breaker Demo ===');
  
  try {
    // Create a circuit breaker for testing
    const breaker = new CircuitBreaker('DEMO_CARRIER', {
      failureThreshold: 3,
      recoveryTimeout: 5000
    });
    
    console.log('Initial status:', breaker.getStatus());
    
    // Simulate failures
    for (let i = 0; i < 5; i++) {
      try {
        breaker.execute(() => {
          throw new Error(`Simulated failure ${i + 1}`);
        });
      } catch (error) {
        console.log(`Attempt ${i + 1}: ${error.message}`);
        console.log('Circuit breaker status:', breaker.getStatus().state);
      }
    }
    
    // Reset the circuit breaker
    breaker.reset();
    console.log('Reset status:', breaker.getStatus());
    
  } catch (error) {
    console.error('Example 4 failed:', error.message);
  }
}

/**
 * Example 5: Retry mechanism demonstration
 */
async function example5_RetryDemo() {
  console.log('=== Example 5: Retry Mechanism Demo ===');
  
  try {
    const retryManager = new RetryManager({
      maxRetries: 3,
      baseDelay: 1000,
      multiplier: 2
    });
    
    let attemptCount = 0;
    
    // Function that fails first 2 times, then succeeds
    const flakyFunction = () => {
      attemptCount++;
      console.log(`Attempt ${attemptCount}`);
      
      if (attemptCount < 3) {
        const error = new Error('Temporary failure (rate limit)');
        error.status = 'RATE_LIMIT_429';
        throw error;
      }
      
      return 'Success!';
    };
    
    // Execute with retry
    const result = await retryManager.executeWithRetry(flakyFunction, {
      carrier: 'DEMO',
      trackingCode: 'DEMO123'
    });
    
    console.log('Final result:', result);
    console.log(`Succeeded after ${attemptCount} attempts`);
    
  } catch (error) {
    console.error('Example 5 failed:', error.message);
  }
}

/**
 * Example 6: Credential testing and validation
 */
async function example6_CredentialTesting() {
  console.log('=== Example 6: Credential Testing ===');
  
  try {
    // Test all carrier credentials
    const results = await testAllCarrierCredentials();
    
    console.log('Credential test results:');
    Object.entries(results).forEach(([carrier, result]) => {
      console.log(`${carrier}:`);
      console.log(`  Validation: ${result.validation.valid ? '✓' : '✗'}`);
      console.log(`  Test call: ${result.test.success ? '✓' : '✗'}`);
      console.log(`  Overall: ${result.overall ? '✓' : '✗'}`);
    });
    
    // Show detailed validation report
    showCredentialValidationReport();
    
  } catch (error) {
    console.error('Example 6 failed:', error.message);
  }
}

/**
 * Example 7: Performance monitoring
 */
function example7_PerformanceMonitoring() {
  console.log('=== Example 7: Performance Monitoring ===');
  
  try {
    const orchestrator = getTrackingOrchestrator();
    
    // Get detailed status
    const status = orchestrator.getStatus();
    
    console.log('=== Orchestrator Performance Metrics ===');
    console.log(`Running: ${status.isRunning}`);
    console.log(`Active Jobs: ${status.activeJobs}`);
    console.log(`Queued Jobs: ${status.queuedJobs}`);
    console.log(`Total Jobs Processed: ${status.metrics.totalJobs}`);
    console.log(`Successful Jobs: ${status.metrics.successfulJobs}`);
    console.log(`Failed Jobs: ${status.metrics.failedJobs}`);
    console.log(`Jobs with Retries: ${status.metrics.retriedJobs}`);
    console.log(`Uptime: ${Math.round(status.uptime / 1000)} seconds`);
    
    console.log('\n=== Circuit Breaker Status ===');
    Object.entries(status.circuitBreakers).forEach(([carrier, cb]) => {
      console.log(`${carrier}: ${cb.state} (failures: ${cb.failureCount})`);
    });
    
    // Calculate success rate
    const successRate = status.metrics.totalJobs > 0 ? 
      (status.metrics.successfulJobs / status.metrics.totalJobs * 100).toFixed(1) : 0;
    console.log(`\nOverall Success Rate: ${successRate}%`);
    
  } catch (error) {
    console.error('Example 7 failed:', error.message);
  }
}

/**
 * Run all examples
 */
async function runAllExamples() {
  console.log('🚀 Running Enhanced Tracking Orchestrator Examples\n');
  
  try {
    example1_BasicSetup();
    console.log('\n---\n');
    
    example2_SubmitJobs();
    console.log('\n---\n');
    
    await example3_EnhancedBulkTracking();
    console.log('\n---\n');
    
    example4_CircuitBreakerDemo();
    console.log('\n---\n');
    
    await example5_RetryDemo();
    console.log('\n---\n');
    
    await example6_CredentialTesting();
    console.log('\n---\n');
    
    example7_PerformanceMonitoring();
    
    console.log('\n✅ All examples completed!');
    
  } catch (error) {
    console.error('❌ Example execution failed:', error.message);
  }
}

/**
 * Quick demo for UI
 */
function quickDemo() {
  try {
    console.log('Running quick demo of enhanced tracking orchestrator...');
    
    // Initialize
    const orchestrator = getTrackingOrchestrator();
    orchestrator.initialize();
    
    // Submit a test job
    const jobId = orchestrator.submitJob({
      carrier: 'TEST',
      trackingCode: 'DEMO123'
    });
    
    console.log(`Demo job submitted with ID: ${jobId}`);
    
    // Show status
    const status = orchestrator.getStatus();
    console.log('Demo status:', {
      running: status.isRunning,
      activeJobs: status.activeJobs,
      queuedJobs: status.queuedJobs
    });
    
    SpreadsheetApp.getUi().alert(
      `Enhanced Tracking Orchestrator Demo Complete!\n\n` +
      `Job ID: ${jobId}\n` +
      `Running: ${status.isRunning}\n` +
      `Active Jobs: ${status.activeJobs}\n` +
      `Queued Jobs: ${status.queuedJobs}`
    );
    
  } catch (error) {
    console.error('Quick demo failed:', error.message);
    SpreadsheetApp.getUi().alert(`Demo failed: ${error.message}`);
  }
}

/**
 * Menu functions for easy access
 */
function menuRunExample1() { example1_BasicSetup(); }
function menuRunExample2() { example2_SubmitJobs(); }
function menuRunExample3() { example3_EnhancedBulkTracking(); }
function menuRunExample4() { example4_CircuitBreakerDemo(); }
function menuRunExample5() { example5_RetryDemo(); }
function menuRunExample6() { example6_CredentialTesting(); }
function menuRunExample7() { example7_PerformanceMonitoring(); }

/**
 * Add examples menu
 */
function addExamplesMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('📚 Orchestrator Examples')
      .addItem('Quick Demo', 'quickDemo')
      .addSeparator()
      .addItem('1. Basic Setup', 'menuRunExample1')
      .addItem('2. Submit Jobs', 'menuRunExample2')
      .addItem('3. Enhanced Bulk Tracking', 'menuRunExample3')
      .addItem('4. Circuit Breaker Demo', 'menuRunExample4')
      .addItem('5. Retry Mechanism', 'menuRunExample5')
      .addItem('6. Credential Testing', 'menuRunExample6')
      .addItem('7. Performance Monitoring', 'menuRunExample7')
      .addSeparator()
      .addItem('Run All Examples', 'runAllExamples')
      .addToUi();
      
  } catch (error) {
    console.error('Failed to add examples menu:', error);
  }
}