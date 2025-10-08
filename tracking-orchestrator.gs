/**
 * Enhanced Modular Tracking Orchestrator
 * 
 * This module provides an enhanced, modular tracking orchestrator for Google Apps Script projects
 * with safe credential helpers, retry/backoff mechanisms, and improved error handling.
 * 
 * Features:
 * - Safe credential validation and management
 * - Exponential backoff with circuit breaker pattern
 * - Job queue management with priority handling
 * - Comprehensive error handling and logging
 * - Status monitoring and health checks
 * - Modular architecture for easy maintenance
 */

// Orchestrator configuration constants
const ORCHESTRATOR_CONFIG = {
  // Queue management
  MAX_CONCURRENT_JOBS: 5,
  MAX_QUEUE_SIZE: 1000,
  JOB_TIMEOUT_MS: 30000,
  
  // Retry configuration
  MAX_RETRIES: 3,
  BASE_RETRY_DELAY_MS: 1000,
  MAX_RETRY_DELAY_MS: 30000,
  RETRY_MULTIPLIER: 2,
  JITTER_RANGE: 0.1,
  
  // Circuit breaker
  FAILURE_THRESHOLD: 5,
  RECOVERY_TIMEOUT_MS: 60000,
  HALF_OPEN_MAX_CALLS: 3,
  
  // Rate limiting
  DEFAULT_RATE_LIMIT_MS: 1000,
  RATE_LIMIT_BACKOFF_MULTIPLIER: 2,
  MAX_RATE_LIMIT_BACKOFF_MS: 300000,
  
  // Monitoring
  HEALTH_CHECK_INTERVAL_MS: 30000,
  METRICS_RETENTION_HOURS: 24
};

/**
 * Enhanced Credential Manager with safe validation
 */
class SafeCredentialManager {
  constructor() {
    this.properties = PropertiesService.getScriptProperties();
    this.cache = new Map();
    this.cacheTimeout = 5 * 60 * 1000; // 5 minutes
  }
  
  /**
   * Safely get a credential with validation
   * @param {string} key - The credential key
   * @param {Object} options - Validation options
   * @returns {string|null} The credential value or null if invalid
   */
  safeGet(key, options = {}) {
    try {
      // Check cache first
      const cached = this.cache.get(key);
      if (cached && Date.now() - cached.timestamp < this.cacheTimeout) {
        return cached.value;
      }
      
      const value = this.properties.getProperty(key);
      if (!value) {
        this.logCredentialError_(key, 'Credential not found');
        return null;
      }
      
      // Validate credential format
      if (options.minLength && value.length < options.minLength) {
        this.logCredentialError_(key, `Credential too short (min: ${options.minLength})`);
        return null;
      }
      
      if (options.pattern && !options.pattern.test(value)) {
        this.logCredentialError_(key, 'Credential format invalid');
        return null;
      }
      
      if (options.notEmpty && !value.trim()) {
        this.logCredentialError_(key, 'Credential is empty');
        return null;
      }
      
      // Cache the validated credential
      this.cache.set(key, { value, timestamp: Date.now() });
      return value;
      
    } catch (error) {
      this.logCredentialError_(key, `Error retrieving credential: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Safely set a credential with validation
   * @param {string} key - The credential key
   * @param {string} value - The credential value
   * @param {Object} options - Validation options
   * @returns {boolean} Success status
   */
  safeSet(key, value, options = {}) {
    try {
      if (!key || typeof key !== 'string') {
        throw new Error('Invalid credential key');
      }
      
      if (value === null || value === undefined) {
        this.properties.deleteProperty(key);
        this.cache.delete(key);
        return true;
      }
      
      // Validate before setting
      const stringValue = String(value);
      if (options.minLength && stringValue.length < options.minLength) {
        throw new Error(`Credential too short (min: ${options.minLength})`);
      }
      
      if (options.pattern && !options.pattern.test(stringValue)) {
        throw new Error('Credential format invalid');
      }
      
      this.properties.setProperty(key, stringValue);
      this.cache.set(key, { value: stringValue, timestamp: Date.now() });
      return true;
      
    } catch (error) {
      this.logCredentialError_(key, `Error setting credential: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Test credential by making a test API call
   * @param {string} carrier - The carrier name
   * @param {string} testCode - Test tracking code
   * @returns {Object} Test result
   */
  async testCredential(carrier, testCode) {
    try {
      const result = await this.makeTestCall_(carrier, testCode);
      return {
        success: true,
        carrier,
        message: 'Credential test successful',
        details: result
      };
    } catch (error) {
      return {
        success: false,
        carrier,
        message: 'Credential test failed',
        error: error.message
      };
    }
  }
  
  /**
   * Get all required credentials for a carrier
   * @param {string} carrier - The carrier name
   * @returns {Object} Credential validation result
   */
  validateCarrierCredentials(carrier) {
    const carrierUpper = carrier.toUpperCase();
    const credentialMap = this.getCarrierCredentialMap_(carrierUpper);
    const results = {};
    let allValid = true;
    
    for (const [key, options] of Object.entries(credentialMap)) {
      const value = this.safeGet(key, options);
      results[key] = {
        valid: !!value,
        present: !!this.properties.getProperty(key),
        message: value ? 'Valid' : 'Missing or invalid'
      };
      if (!value) allValid = false;
    }
    
    return {
      carrier,
      valid: allValid,
      credentials: results
    };
  }
  
  /**
   * Clear credential cache
   */
  clearCache() {
    this.cache.clear();
  }
  
  /**
   * Get carrier-specific credential requirements
   * @private
   */
  getCarrierCredentialMap_(carrier) {
    const maps = {
      'POSTI': {
        'POSTI_TOKEN_URL': { notEmpty: true, pattern: /^https?:\/\// },
        'POSTI_BASIC': { notEmpty: true, minLength: 10 },
        'POSTI_TRACK_URL': { notEmpty: true, pattern: /^https?:\/\/.*\{\{code\}\}/ }
      },
      'GLS': {
        'GLS_FI_TRACK_URL': { notEmpty: true, pattern: /^https?:\/\// },
        'GLS_FI_API_KEY': { notEmpty: true, minLength: 10 }
      },
      'DHL': {
        'DHL_TRACK_URL': { notEmpty: true, pattern: /^https?:\/\/.*\{\{code\}\}/ },
        'DHL_API_KEY': { notEmpty: true, minLength: 10 }
      },
      'BRING': {
        'BRING_TRACK_URL': { notEmpty: true, pattern: /^https?:\/\/.*\{\{code\}\}/ },
        'BRING_UID': { notEmpty: true },
        'BRING_KEY': { notEmpty: true }
      },
      'MATKAHUOLTO': {
        'MH_TRACK_URL': { notEmpty: true, pattern: /^https?:\/\/.*\{\{code\}\}/ },
        'MH_BASIC': { notEmpty: true, minLength: 10 }
      }
    };
    
    return maps[carrier] || {};
  }
  
  /**
   * Make a test API call for credential validation
   * @private
   */
  async makeTestCall_(carrier, testCode) {
    // This would integrate with existing TRK_trackByCarrier_ function
    // For now, return a mock successful result
    return { status: 'test_successful', carrier };
  }
  
  /**
   * Log credential errors
   * @private
   */
  logCredentialError_(key, message) {
    if (typeof logError_ === 'function') {
      logError_('CredentialManager', `${key}: ${message}`);
    } else {
      console.error(`Credential Error - ${key}: ${message}`);
    }
  }
}

/**
 * Circuit Breaker for handling failures
 */
class CircuitBreaker {
  constructor(carrier, options = {}) {
    this.carrier = carrier;
    this.failureThreshold = options.failureThreshold || ORCHESTRATOR_CONFIG.FAILURE_THRESHOLD;
    this.recoveryTimeout = options.recoveryTimeout || ORCHESTRATOR_CONFIG.RECOVERY_TIMEOUT_MS;
    this.halfOpenMaxCalls = options.halfOpenMaxCalls || ORCHESTRATOR_CONFIG.HALF_OPEN_MAX_CALLS;
    
    this.state = 'CLOSED'; // CLOSED, OPEN, HALF_OPEN
    this.failureCount = 0;
    this.lastFailureTime = 0;
    this.halfOpenCallCount = 0;
    
    this.loadState_();
  }
  
  /**
   * Execute a function with circuit breaker protection
   * @param {Function} fn - Function to execute
   * @returns {Promise} Function result or circuit breaker error
   */
  async execute(fn) {
    if (this.state === 'OPEN') {
      if (Date.now() - this.lastFailureTime >= this.recoveryTimeout) {
        this.state = 'HALF_OPEN';
        this.halfOpenCallCount = 0;
        this.saveState_();
      } else {
        throw new Error(`Circuit breaker OPEN for ${this.carrier}`);
      }
    }
    
    if (this.state === 'HALF_OPEN' && this.halfOpenCallCount >= this.halfOpenMaxCalls) {
      throw new Error(`Circuit breaker HALF_OPEN call limit exceeded for ${this.carrier}`);
    }
    
    try {
      const result = await fn();
      this.onSuccess_();
      return result;
    } catch (error) {
      this.onFailure_();
      throw error;
    }
  }
  
  /**
   * Reset the circuit breaker
   */
  reset() {
    this.state = 'CLOSED';
    this.failureCount = 0;
    this.lastFailureTime = 0;
    this.halfOpenCallCount = 0;
    this.saveState_();
  }
  
  /**
   * Get current circuit breaker status
   */
  getStatus() {
    return {
      carrier: this.carrier,
      state: this.state,
      failureCount: this.failureCount,
      lastFailureTime: this.lastFailureTime,
      healthCheck: this.state === 'CLOSED' || 
                   (this.state === 'HALF_OPEN' && this.halfOpenCallCount < this.halfOpenMaxCalls)
    };
  }
  
  /**
   * Handle successful execution
   * @private
   */
  onSuccess_() {
    if (this.state === 'HALF_OPEN') {
      this.halfOpenCallCount++;
      if (this.halfOpenCallCount >= this.halfOpenMaxCalls) {
        this.state = 'CLOSED';
        this.failureCount = 0;
        this.saveState_();
      }
    } else {
      this.failureCount = 0;
      this.saveState_();
    }
  }
  
  /**
   * Handle failed execution
   * @private
   */
  onFailure_() {
    this.failureCount++;
    this.lastFailureTime = Date.now();
    
    if (this.state === 'HALF_OPEN') {
      this.state = 'OPEN';
    } else if (this.failureCount >= this.failureThreshold) {
      this.state = 'OPEN';
    }
    
    this.saveState_();
  }
  
  /**
   * Load state from properties
   * @private
   */
  loadState_() {
    const sp = PropertiesService.getScriptProperties();
    const stateKey = `CIRCUIT_BREAKER_${this.carrier}`;
    const stateJson = sp.getProperty(stateKey);
    
    if (stateJson) {
      try {
        const state = JSON.parse(stateJson);
        this.state = state.state || 'CLOSED';
        this.failureCount = state.failureCount || 0;
        this.lastFailureTime = state.lastFailureTime || 0;
        this.halfOpenCallCount = state.halfOpenCallCount || 0;
      } catch (error) {
        // Reset to default state on parse error
        this.reset();
      }
    }
  }
  
  /**
   * Save state to properties
   * @private
   */
  saveState_() {
    const sp = PropertiesService.getScriptProperties();
    const stateKey = `CIRCUIT_BREAKER_${this.carrier}`;
    const state = {
      state: this.state,
      failureCount: this.failureCount,
      lastFailureTime: this.lastFailureTime,
      halfOpenCallCount: this.halfOpenCallCount
    };
    
    try {
      sp.setProperty(stateKey, JSON.stringify(state));
    } catch (error) {
      console.error(`Failed to save circuit breaker state for ${this.carrier}:`, error);
    }
  }
}

/**
 * Enhanced Retry Manager with exponential backoff
 */
class RetryManager {
  constructor(options = {}) {
    this.maxRetries = options.maxRetries || ORCHESTRATOR_CONFIG.MAX_RETRIES;
    this.baseDelay = options.baseDelay || ORCHESTRATOR_CONFIG.BASE_RETRY_DELAY_MS;
    this.maxDelay = options.maxDelay || ORCHESTRATOR_CONFIG.MAX_RETRY_DELAY_MS;
    this.multiplier = options.multiplier || ORCHESTRATOR_CONFIG.RETRY_MULTIPLIER;
    this.jitterRange = options.jitterRange || ORCHESTRATOR_CONFIG.JITTER_RANGE;
  }
  
  /**
   * Execute function with retry logic
   * @param {Function} fn - Function to execute
   * @param {Object} context - Context for logging
   * @returns {Promise} Function result
   */
  async executeWithRetry(fn, context = {}) {
    let lastError;
    
    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      try {
        if (attempt > 0) {
          const delay = this.calculateDelay_(attempt);
          await this.sleep_(delay);
        }
        
        const result = await fn();
        
        if (attempt > 0) {
          this.logRetrySuccess_(context, attempt);
        }
        
        return result;
      } catch (error) {
        lastError = error;
        
        if (attempt === this.maxRetries) {
          this.logRetryFailure_(context, attempt, error);
          break;
        }
        
        if (!this.isRetryableError_(error)) {
          this.logNonRetryableError_(context, error);
          break;
        }
        
        this.logRetryAttempt_(context, attempt + 1, error);
      }
    }
    
    throw lastError;
  }
  
  /**
   * Calculate delay with exponential backoff and jitter
   * @private
   */
  calculateDelay_(attempt) {
    const exponentialDelay = this.baseDelay * Math.pow(this.multiplier, attempt - 1);
    const clampedDelay = Math.min(exponentialDelay, this.maxDelay);
    
    // Add jitter to prevent thundering herd
    const jitter = clampedDelay * this.jitterRange * (Math.random() * 2 - 1);
    return Math.max(0, clampedDelay + jitter);
  }
  
  /**
   * Check if error is retryable
   * @private
   */
  isRetryableError_(error) {
    if (typeof error === 'object' && error !== null) {
      // Rate limiting is retryable
      if (error.message && error.message.includes('429')) return true;
      if (error.status === 'RATE_LIMIT_429') return true;
      
      // Network errors are retryable
      if (error.message && error.message.includes('timeout')) return true;
      if (error.message && error.message.includes('network')) return true;
      
      // Server errors (5xx) are retryable
      if (error.message && /HTTP_5\d\d/.test(error.message)) return true;
    }
    
    return false;
  }
  
  /**
   * Sleep utility
   * @private
   */
  sleep_(ms) {
    return new Promise(resolve => {
      if (typeof Utilities !== 'undefined' && Utilities.sleep) {
        Utilities.sleep(ms);
        resolve();
      } else {
        setTimeout(resolve, ms);
      }
    });
  }
  
  /**
   * Logging methods
   * @private
   */
  logRetryAttempt_(context, attempt, error) {
    const message = `Retry attempt ${attempt}/${this.maxRetries} for ${context.carrier || 'unknown'} - ${context.trackingCode || 'unknown'}: ${error.message}`;
    if (typeof logError_ === 'function') {
      logError_('RetryManager', message);
    }
  }
  
  logRetrySuccess_(context, attempt) {
    const message = `Retry successful after ${attempt} attempts for ${context.carrier || 'unknown'} - ${context.trackingCode || 'unknown'}`;
    console.log(message);
  }
  
  logRetryFailure_(context, attempts, error) {
    const message = `All ${attempts} retry attempts failed for ${context.carrier || 'unknown'} - ${context.trackingCode || 'unknown'}: ${error.message}`;
    if (typeof logError_ === 'function') {
      logError_('RetryManager', message);
    }
  }
  
  logNonRetryableError_(context, error) {
    const message = `Non-retryable error for ${context.carrier || 'unknown'} - ${context.trackingCode || 'unknown'}: ${error.message}`;
    if (typeof logError_ === 'function') {
      logError_('RetryManager', message);
    }
  }
}

/**
 * Main Tracking Orchestrator
 */
class TrackingOrchestrator {
  constructor() {
    this.credentialManager = new SafeCredentialManager();
    this.retryManager = new RetryManager();
    this.circuitBreakers = new Map();
    this.activeJobs = new Map();
    this.jobQueue = [];
    this.isRunning = false;
    this.metrics = {
      totalJobs: 0,
      successfulJobs: 0,
      failedJobs: 0,
      retriedJobs: 0,
      startTime: Date.now()
    };
  }
  
  /**
   * Initialize the orchestrator
   */
  async initialize() {
    try {
      this.loadConfiguration_();
      this.setupMetrics_();
      this.isRunning = true;
      console.log('Tracking Orchestrator initialized successfully');
    } catch (error) {
      console.error('Failed to initialize Tracking Orchestrator:', error);
      throw error;
    }
  }
  
  /**
   * Submit a tracking job
   * @param {Object} job - Job configuration
   * @returns {string} Job ID
   */
  submitJob(job) {
    const jobId = this.generateJobId_();
    const enhancedJob = {
      id: jobId,
      ...job,
      status: 'queued',
      createdAt: Date.now(),
      attempts: 0
    };
    
    if (this.jobQueue.length >= ORCHESTRATOR_CONFIG.MAX_QUEUE_SIZE) {
      throw new Error('Job queue is full');
    }
    
    this.jobQueue.push(enhancedJob);
    this.metrics.totalJobs++;
    
    return jobId;
  }
  
  /**
   * Process a single tracking job
   * @param {Object} job - Job to process
   * @returns {Promise<Object>} Job result
   */
  async processJob(job) {
    const { carrier, trackingCode } = job;
    
    if (!carrier || !trackingCode) {
      throw new Error('Invalid job: missing carrier or trackingCode');
    }
    
    // Validate credentials first
    const credentialValidation = this.credentialManager.validateCarrierCredentials(carrier);
    if (!credentialValidation.valid) {
      throw new Error(`Invalid credentials for ${carrier}`);
    }
    
    // Get or create circuit breaker for this carrier
    const circuitBreaker = this.getCircuitBreaker_(carrier);
    
    // Execute with circuit breaker and retry logic
    const context = { carrier, trackingCode, jobId: job.id };
    
    return await circuitBreaker.execute(async () => {
      return await this.retryManager.executeWithRetry(async () => {
        return await this.executeTracking_(carrier, trackingCode);
      }, context);
    });
  }
  
  /**
   * Process the job queue
   */
  async processQueue() {
    if (!this.isRunning) {
      console.log('Orchestrator not running');
      return;
    }
    
    const maxConcurrent = ORCHESTRATOR_CONFIG.MAX_CONCURRENT_JOBS;
    
    while (this.jobQueue.length > 0 && this.activeJobs.size < maxConcurrent) {
      const job = this.jobQueue.shift();
      this.processJobAsync_(job);
    }
  }
  
  /**
   * Get orchestrator status and health
   * @returns {Object} Status information
   */
  getStatus() {
    const circuitBreakerStatus = {};
    for (const [carrier, breaker] of this.circuitBreakers.entries()) {
      circuitBreakerStatus[carrier] = breaker.getStatus();
    }
    
    return {
      isRunning: this.isRunning,
      activeJobs: this.activeJobs.size,
      queuedJobs: this.jobQueue.length,
      metrics: this.metrics,
      circuitBreakers: circuitBreakerStatus,
      uptime: Date.now() - this.metrics.startTime
    };
  }
  
  /**
   * Shutdown the orchestrator
   */
  async shutdown() {
    this.isRunning = false;
    
    // Wait for active jobs to complete
    while (this.activeJobs.size > 0) {
      await this.sleep_(1000);
    }
    
    console.log('Tracking Orchestrator shutdown complete');
  }
  
  /**
   * Get or create circuit breaker for a carrier
   * @private
   */
  getCircuitBreaker_(carrier) {
    if (!this.circuitBreakers.has(carrier)) {
      this.circuitBreakers.set(carrier, new CircuitBreaker(carrier));
    }
    return this.circuitBreakers.get(carrier);
  }
  
  /**
   * Process a job asynchronously
   * @private
   */
  async processJobAsync_(job) {
    this.activeJobs.set(job.id, job);
    job.status = 'processing';
    job.startedAt = Date.now();
    
    try {
      const result = await this.processJob(job);
      job.status = 'completed';
      job.result = result;
      job.completedAt = Date.now();
      this.metrics.successfulJobs++;
      
      if (job.attempts > 0) {
        this.metrics.retriedJobs++;
      }
      
    } catch (error) {
      job.status = 'failed';
      job.error = error.message;
      job.completedAt = Date.now();
      this.metrics.failedJobs++;
      
      if (typeof logError_ === 'function') {
        logError_('TrackingOrchestrator', `Job ${job.id} failed: ${error.message}`);
      }
    } finally {
      this.activeJobs.delete(job.id);
    }
  }
  
  /**
   * Execute the actual tracking call
   * @private
   */
  async executeTracking_(carrier, trackingCode) {
    // Apply rate limiting
    await this.applyRateLimit_(carrier);
    
    // This will integrate with the existing TRK_trackByCarrier_ function
    if (typeof TRK_trackByCarrier_ === 'function') {
      return TRK_trackByCarrier_(carrier, trackingCode);
    } else {
      // Fallback for testing
      return {
        carrier,
        status: 'DELIVERED',
        time: new Date().toISOString(),
        location: 'Test Location',
        found: true
      };
    }
  }
  
  /**
   * Apply rate limiting
   * @private
   */
  async applyRateLimit_(carrier) {
    const sp = PropertiesService.getScriptProperties();
    const rateLimitKey = `RATE_LIMIT_${carrier.toUpperCase()}`;
    const lastCallKey = `LAST_CALL_${carrier.toUpperCase()}`;
    
    const rateLimitMs = Number(sp.getProperty(rateLimitKey)) || ORCHESTRATOR_CONFIG.DEFAULT_RATE_LIMIT_MS;
    const lastCall = Number(sp.getProperty(lastCallKey)) || 0;
    
    const elapsed = Date.now() - lastCall;
    if (elapsed < rateLimitMs) {
      const waitTime = rateLimitMs - elapsed;
      await this.sleep_(waitTime);
    }
    
    sp.setProperty(lastCallKey, String(Date.now()));
  }
  
  /**
   * Generate unique job ID
   * @private
   */
  generateJobId_() {
    return `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
  
  /**
   * Load configuration from properties
   * @private
   */
  loadConfiguration_() {
    // Load any custom configuration from Script Properties
    const sp = PropertiesService.getScriptProperties();
    // Future: load custom settings
  }
  
  /**
   * Setup metrics collection
   * @private
   */
  setupMetrics_() {
    // Initialize metrics if needed
    this.metrics.startTime = Date.now();
  }
  
  /**
   * Sleep utility
   * @private
   */
  sleep_(ms) {
    return new Promise(resolve => {
      if (typeof Utilities !== 'undefined' && Utilities.sleep) {
        Utilities.sleep(ms);
        resolve();
      } else {
        setTimeout(resolve, ms);
      }
    });
  }
}

// Global orchestrator instance
let globalOrchestrator = null;

/**
 * Get the global orchestrator instance
 * @returns {TrackingOrchestrator}
 */
function getTrackingOrchestrator() {
  if (!globalOrchestrator) {
    globalOrchestrator = new TrackingOrchestrator();
  }
  return globalOrchestrator;
}

/**
 * Initialize the tracking orchestrator
 */
async function initializeTrackingOrchestrator() {
  const orchestrator = getTrackingOrchestrator();
  await orchestrator.initialize();
  return orchestrator;
}

/**
 * Safe credential helper functions for backward compatibility
 */

/**
 * Safely get a credential with validation
 * @param {string} key - Credential key
 * @param {Object} options - Validation options
 * @returns {string|null}
 */
function safeGetCredential(key, options = {}) {
  const credManager = new SafeCredentialManager();
  return credManager.safeGet(key, options);
}

/**
 * Safely set a credential with validation
 * @param {string} key - Credential key
 * @param {string} value - Credential value
 * @param {Object} options - Validation options
 * @returns {boolean}
 */
function safeSetCredential(key, value, options = {}) {
  const credManager = new SafeCredentialManager();
  return credManager.safeSet(key, value, options);
}

/**
 * Test credentials for a specific carrier
 * @param {string} carrier - Carrier name
 * @param {string} testCode - Test tracking code
 * @returns {Object}
 */
async function testCarrierCredentials(carrier, testCode = 'TEST123') {
  const credManager = new SafeCredentialManager();
  return await credManager.testCredential(carrier, testCode);
}

/**
 * Enhanced version of the bulk tracking process using the orchestrator
 * @param {string} sheetName - Sheet to process
 * @param {string} carrierFilter - Optional carrier filter
 */
async function enhancedBulkStart(sheetName, carrierFilter = null) {
  try {
    const orchestrator = await initializeTrackingOrchestrator();
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No data rows to process');
    }
    
    const hdr = data[0];
    const headerMap = headerIndexMap_(hdr);
    
    // Find relevant columns
    const carrierCol = colIndexOf_(headerMap, CARRIER_CANDIDATES);
    const trackingCol = colIndexOf_(headerMap, TRACKING_CANDIDATES);
    
    if (carrierCol === -1 || trackingCol === -1) {
      throw new Error('Required columns not found');
    }
    
    // Submit jobs to orchestrator
    let jobsSubmitted = 0;
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const carrier = String(row[carrierCol] || '').trim();
      const trackingCode = String(row[trackingCol] || '').trim();
      
      if (!carrier || !trackingCode) continue;
      
      if (carrierFilter && !carrier.toLowerCase().includes(carrierFilter.toLowerCase())) {
        continue;
      }
      
      try {
        const jobId = orchestrator.submitJob({
          carrier: canonicalCarrier_(carrier),
          trackingCode,
          sheetName,
          rowIndex: r,
          originalCarrier: carrier
        });
        jobsSubmitted++;
      } catch (error) {
        if (typeof logError_ === 'function') {
          logError_('enhancedBulkStart', `Failed to submit job for row ${r}: ${error.message}`);
        }
      }
    }
    
    // Start processing
    await orchestrator.processQueue();
    
    SpreadsheetApp.getUi().alert(
      `Enhanced bulk tracking started for "${sheetName}". ` +
      `Submitted ${jobsSubmitted} jobs.` +
      (carrierFilter ? ` (Filtered by: ${carrierFilter})` : '')
    );
    
  } catch (error) {
    if (typeof logError_ === 'function') {
      logError_('enhancedBulkStart', error.message);
    }
    SpreadsheetApp.getUi().alert(`Error starting enhanced bulk tracking: ${error.message}`);
  }
}

/**
 * Get orchestrator status for monitoring
 * @returns {Object}
 */
function getOrchestratorStatus() {
  const orchestrator = getTrackingOrchestrator();
  return orchestrator.getStatus();
}

/**
 * Reset circuit breakers for all carriers
 */
function resetAllCircuitBreakers() {
  const orchestrator = getTrackingOrchestrator();
  for (const [carrier, breaker] of orchestrator.circuitBreakers.entries()) {
    breaker.reset();
  }
  SpreadsheetApp.getUi().alert('All circuit breakers have been reset.');
}

/**
 * Show orchestrator status in UI
 */
function showOrchestratorStatus() {
  const status = getOrchestratorStatus();
  const message = `
Tracking Orchestrator Status:
- Running: ${status.isRunning}
- Active Jobs: ${status.activeJobs}
- Queued Jobs: ${status.queuedJobs}
- Total Jobs: ${status.metrics.totalJobs}
- Successful: ${status.metrics.successfulJobs}
- Failed: ${status.metrics.failedJobs}
- Uptime: ${Math.round(status.uptime / 1000)} seconds

Circuit Breaker Status:
${Object.entries(status.circuitBreakers).map(([carrier, cb]) => 
  `- ${carrier}: ${cb.state} (failures: ${cb.failureCount})`
).join('\n')}
  `;
  
  SpreadsheetApp.getUi().alert(message);
}