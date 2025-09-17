# Enhanced Tracking Orchestrator

## Overview

The Enhanced Tracking Orchestrator is a modular, enterprise-grade system for managing shipment tracking operations in Google Apps Script. It provides robust error handling, credential management, retry mechanisms, and performance monitoring.

## Architecture

The orchestrator consists of several key components:

### 1. SafeCredentialManager
- **Purpose**: Secure credential storage and validation
- **Features**: 
  - Pattern-based validation (URLs, API keys)
  - Caching for performance
  - Carrier-specific credential requirements
  - Test functionality

### 2. CircuitBreaker
- **Purpose**: Fault tolerance and failure isolation
- **Features**:
  - CLOSED/OPEN/HALF_OPEN states
  - Configurable failure thresholds
  - Automatic recovery
  - Persistent state storage

### 3. RetryManager
- **Purpose**: Intelligent retry logic with backoff
- **Features**:
  - Exponential backoff with jitter
  - Smart error classification
  - Rate limit awareness
  - Configurable retry policies

### 4. TrackingOrchestrator
- **Purpose**: Main coordination and job management
- **Features**:
  - Job queue management
  - Concurrent processing
  - Status monitoring
  - Integration with existing systems

## Installation

1. **Copy the Files**: Add these files to your Google Apps Script project:
   - `tracking-orchestrator.gs` - Main orchestrator classes
   - `orchestrator-integration.gs` - Integration helpers
   - `orchestrator-tests.gs` - Test functions
   - `orchestrator-examples.gs` - Usage examples

2. **Enable Required Services**:
   - Google Drive API (for XLSX conversion if needed)
   - Any external APIs for carriers

3. **Set Up Credentials**: Use the enhanced credential management system to store API keys securely.

## Quick Start

### Basic Setup

```javascript
// Initialize the orchestrator
const orchestrator = await initializeTrackingOrchestrator();

// Set up credentials safely
const credManager = new SafeCredentialManager();
credManager.safeSet('DHL_TRACK_URL', 'https://api.dhl.com/track/{{code}}', {
  pattern: /^https?:\/\/.*\{\{code\}\}/
});
credManager.safeSet('DHL_API_KEY', 'your-api-key', {
  minLength: 10
});
```

### Submit Tracking Jobs

```javascript
// Submit individual jobs
const jobId = orchestrator.submitJob({
  carrier: 'DHL',
  trackingCode: 'DHL123456789'
});

// Process the queue
await orchestrator.processQueue();
```

### Enhanced Bulk Tracking

```javascript
// Use enhanced bulk tracking for existing sheets
await enhancedBulkStart('Packages', null); // All carriers
await enhancedBulkStart('Packages', 'DHL'); // DHL only
```

## Configuration

### Orchestrator Configuration

The orchestrator uses these default settings (can be customized):

```javascript
const ORCHESTRATOR_CONFIG = {
  MAX_CONCURRENT_JOBS: 5,
  MAX_QUEUE_SIZE: 1000,
  JOB_TIMEOUT_MS: 30000,
  MAX_RETRIES: 3,
  BASE_RETRY_DELAY_MS: 1000,
  MAX_RETRY_DELAY_MS: 30000,
  FAILURE_THRESHOLD: 5,
  RECOVERY_TIMEOUT_MS: 60000
};
```

### Credential Validation

Each carrier has specific credential requirements:

- **URLs**: Must be valid HTTP/HTTPS with `{{code}}` placeholder for tracking URLs
- **API Keys**: Minimum length requirements
- **Pattern Validation**: Configurable regex patterns

## Usage Examples

### 1. Credential Management

```javascript
// Validate all carrier credentials
const credManager = new SafeCredentialManager();
const validation = credManager.validateCarrierCredentials('DHL');

if (!validation.valid) {
  console.log('Missing credentials:', validation.credentials);
}

// Test credentials with actual API call
const testResult = await credManager.testCredential('DHL', 'TEST123');
console.log('Test result:', testResult);
```

### 2. Circuit Breaker Usage

```javascript
// Create circuit breaker
const breaker = new CircuitBreaker('MY_CARRIER', {
  failureThreshold: 3,
  recoveryTimeout: 60000
});

// Execute with protection
try {
  const result = await breaker.execute(async () => {
    return await someApiCall();
  });
} catch (error) {
  console.log('Circuit breaker prevented call or call failed');
}
```

### 3. Retry Logic

```javascript
// Execute with retry
const retryManager = new RetryManager({
  maxRetries: 3,
  baseDelay: 1000
});

const result = await retryManager.executeWithRetry(async () => {
  return await unreliableApiCall();
}, { carrier: 'DHL', trackingCode: 'ABC123' });
```

## Integration with Existing System

The orchestrator integrates seamlessly with existing Google Apps Script tracking systems:

### Menu Integration

```javascript
// Enhanced menus are automatically added
function onOpen() {
  enhancedOnOpen(); // Adds orchestrator menus
}
```

### Backward Compatibility

```javascript
// Use enhanced functions as drop-in replacements
enhancedBulkStart_Vaatii(); // Instead of bulkStart_Vaatii()
enhancedCredSaveProps(data); // Instead of credSaveProps(data)
```

### Gradual Migration

The system supports gradual migration:

1. Keep existing `bulkTick()` as fallback
2. Use `enhancedBulkTick()` for new features
3. Migrate sheet by sheet using enhanced functions

## Monitoring and Status

### Status Monitoring

```javascript
// Get comprehensive status
const status = getOrchestratorStatus();
console.log('Active jobs:', status.activeJobs);
console.log('Success rate:', status.metrics.successfulJobs / status.metrics.totalJobs);
```

### Health Checks

```javascript
// Check circuit breaker health
Object.entries(status.circuitBreakers).forEach(([carrier, cb]) => {
  if (cb.state !== 'CLOSED') {
    console.log(`Warning: ${carrier} circuit breaker is ${cb.state}`);
  }
});
```

## Error Handling

The orchestrator provides comprehensive error handling:

- **Credential Errors**: Safe validation prevents invalid credentials
- **Network Errors**: Circuit breakers isolate failing services
- **Rate Limits**: Intelligent backoff and retry
- **Logging**: Comprehensive error logging via `logError_()` function

## Testing

### Run Tests

```javascript
// Run all tests
runAllOrchestratorTests();

// Quick smoke test
quickOrchestratorTest();

// Test specific components
testSafeCredentialManager();
testCircuitBreaker();
testRetryManager();
```

### Test Coverage

Tests cover:
- Credential validation and storage
- Circuit breaker state transitions
- Retry logic and backoff calculations
- Job submission and processing
- Integration with existing functions

## Performance Considerations

### Optimization Features

- **Credential Caching**: Reduces property access overhead
- **Circuit Breakers**: Prevent wasted calls to failing services
- **Intelligent Retry**: Avoids unnecessary retries for non-recoverable errors
- **Concurrent Processing**: Configurable job concurrency

### Monitoring Metrics

- Total jobs processed
- Success/failure rates
- Retry statistics
- Circuit breaker states
- Processing times

## Troubleshooting

### Common Issues

1. **Credentials Not Working**
   ```javascript
   // Use credential validation
   showCredentialValidationReport();
   ```

2. **Circuit Breaker Stuck Open**
   ```javascript
   // Reset circuit breakers
   resetAllCircuitBreakers();
   ```

3. **Jobs Not Processing**
   ```javascript
   // Check orchestrator status
   showOrchestratorStatus();
   
   // Emergency stop and restart
   await emergencyStopOrchestrator();
   await setupTrackingOrchestrator();
   ```

### Debug Mode

Enable detailed logging by checking console output in Google Apps Script editor.

## API Reference

### SafeCredentialManager

- `safeGet(key, options)` - Get credential with validation
- `safeSet(key, value, options)` - Set credential with validation
- `validateCarrierCredentials(carrier)` - Validate carrier setup
- `testCredential(carrier, testCode)` - Test API call

### CircuitBreaker

- `execute(fn)` - Execute function with circuit breaker protection
- `reset()` - Reset circuit breaker state
- `getStatus()` - Get current status

### RetryManager

- `executeWithRetry(fn, context)` - Execute with retry logic

### TrackingOrchestrator

- `initialize()` - Initialize orchestrator
- `submitJob(job)` - Submit tracking job
- `processQueue()` - Process job queue
- `getStatus()` - Get status and metrics
- `shutdown()` - Graceful shutdown

## Migration Guide

### From Existing System

1. **Install Files**: Add orchestrator files to your project
2. **Test Setup**: Run `quickOrchestratorTest()`
3. **Update Menus**: Call `enhancedOnOpen()` in your `onOpen()` function
4. **Migrate Gradually**: Use enhanced functions alongside existing ones
5. **Monitor**: Use status functions to monitor performance

### Breaking Changes

None - the orchestrator is designed for backward compatibility.

## Support

For issues or questions:
1. Check the test functions for examples
2. Review console logs for detailed error information
3. Use the status monitoring functions to diagnose issues
4. Refer to the examples file for common usage patterns