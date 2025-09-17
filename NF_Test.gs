/**
 * NF_Test.gs - Simple manual testing functions for NewFlow
 * These functions help verify NewFlow components work correctly
 */

function NF_testMenuSetup() {
  try {
    NF_onOpen();
    Logger.log('✓ NF_onOpen() executed without errors');
    return true;
  } catch (e) {
    Logger.log('✗ NF_onOpen() failed: ' + e.toString());
    return false;
  }
}

function NF_testHelperFunctions() {
  const results = [];
  
  // Test property helper
  try {
    const testProp = NF_Prop_('NON_EXISTENT_KEY', 'default_value');
    if (testProp === 'default_value') {
      results.push('✓ NF_Prop_ works with default');
    } else {
      results.push('✗ NF_Prop_ failed default test');
    }
  } catch (e) {
    results.push('✗ NF_Prop_ threw error: ' + e);
  }
  
  // Test date parsing
  try {
    const date1 = NF_parseDate_('2024-01-15 10:30:00');
    const date2 = NF_parseDate_('invalid date');
    const date3 = NF_parseDate_('');
    
    if (date1 instanceof Date && !date2 && !date3) {
      results.push('✓ NF_parseDate_ handles valid/invalid dates correctly');
    } else {
      results.push('✗ NF_parseDate_ failed validation');
    }
  } catch (e) {
    results.push('✗ NF_parseDate_ threw error: ' + e);
  }
  
  // Test digits normalization
  try {
    const digits1 = NF_digits_('ABC-123-XYZ');
    const digits2 = NF_digits_('990719901');
    
    if (digits1 === '123' && digits2 === '990719901') {
      results.push('✓ NF_digits_ extracts digits correctly');
    } else {
      results.push(`✗ NF_digits_ failed: got "${digits1}" and "${digits2}"`);
    }
  } catch (e) {
    results.push('✗ NF_digits_ threw error: ' + e);
  }
  
  Logger.log('NF Helper Function Tests:\n' + results.join('\n'));
  return results;
}

function NF_testWeekWindow() {
  try {
    const { start, end } = NF_lastFinishedWeek_();
    const now = new Date();
    
    Logger.log(`Current time: ${now}`);
    Logger.log(`Week start: ${start}`);
    Logger.log(`Week end: ${end}`);
    
    // Verify it's a proper week window
    const diffDays = (end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000);
    
    if (Math.abs(diffDays - 7) < 0.1 && end <= now) {
      Logger.log('✓ NF_lastFinishedWeek_ returns valid 7-day window in the past');
      return true;
    } else {
      Logger.log(`✗ Week window validation failed: ${diffDays} days, end > now: ${end > now}`);
      return false;
    }
  } catch (e) {
    Logger.log('✗ NF_lastFinishedWeek_ threw error: ' + e);
    return false;
  }
}

function NF_testWrapperFunctions() {
  const results = [];
  
  // Test Gmail wrapper (should handle missing function gracefully)
  try {
    // This should not throw even if fetchAndRebuild doesn't exist
    NF_fetchGmailPackagesAndRebuild();
    results.push('✓ NF_fetchGmailPackagesAndRebuild handled gracefully');
  } catch (e) {
    results.push('✗ NF_fetchGmailPackagesAndRebuild threw error: ' + e);
  }
  
  // Test PBI wrapper
  try {
    NF_importPowerBIOutbound();
    results.push('✓ NF_importPowerBIOutbound handled gracefully');
  } catch (e) {
    results.push('✗ NF_importPowerBIOutbound threw error: ' + e);
  }
  
  // Test ERP wrapper
  try {
    NF_importERPStockPicking();
    results.push('✓ NF_importERPStockPicking handled gracefully');
  } catch (e) {
    results.push('✗ NF_importERPStockPicking threw error: ' + e);
  }
  
  Logger.log('NF Wrapper Function Tests:\n' + results.join('\n'));
  return results;
}

function NF_testLeadtimeLogic() {
  // Create test data
  const mockPackages = [
    {
      source: 'Packages',
      trackingCode: 'TEST123',
      country: 'Finland',
      carrier: 'Posti',
      jobDoneTs: new Date('2024-01-10 10:00:00'),
      deliveredTs: new Date('2024-01-12 14:30:00'),
      rawRow: [],
      headers: []
    }
  ];
  
  const mockPowerBI = [
    {
      source: 'PowerBI',
      trackingCode: 'TEST123',
      country: 'Finland',
      carrier: 'Posti',
      jobDoneTs: new Date('2024-01-09 08:00:00'), // Earlier, should be overridden
      deliveredTs: null,
      rawRow: [],
      headers: []
    }
  ];
  
  const mockERP = [
    {
      source: 'ERP',
      trackingCode: 'TEST123',
      country: 'Finland',
      carrier: 'Posti',
      jobDoneTs: new Date('2024-01-10 12:00:00'), // ERP should have priority
      deliveredTs: null,
      rawRow: [],
      headers: []
    }
  ];
  
  try {
    const merged = NF_mergeSourcesByTrackingCode_(mockPackages, mockPowerBI, mockERP);
    const analyzed = NF_analyzeLeadtimes_(merged);
    
    if (analyzed.length === 1) {
      const result = analyzed[0];
      
      // Should use ERP job done time (priority), Packages delivered time
      if (result.jobDoneSource === 'ERP' && 
          result.deliveredTs && 
          result.leadTimeDays && 
          parseFloat(result.leadTimeDays) > 0) {
        Logger.log('✓ Leadtime logic works correctly');
        Logger.log(`  Job done source: ${result.jobDoneSource}`);
        Logger.log(`  Lead time: ${result.leadTimeDays} days`);
        return true;
      } else {
        Logger.log('✗ Leadtime logic failed validation');
        Logger.log(`  Job done source: ${result.jobDoneSource}`);
        Logger.log(`  Lead time: ${result.leadTimeDays}`);
        return false;
      }
    } else {
      Logger.log(`✗ Expected 1 merged result, got ${analyzed.length}`);
      return false;
    }
  } catch (e) {
    Logger.log('✗ Leadtime logic threw error: ' + e);
    return false;
  }
}

function NF_runAllTests() {
  Logger.log('=== NewFlow Test Suite ===\n');
  
  const results = {
    menuSetup: NF_testMenuSetup(),
    helperFunctions: NF_testHelperFunctions(),
    weekWindow: NF_testWeekWindow(),
    wrapperFunctions: NF_testWrapperFunctions(),
    leadtimeLogic: NF_testLeadtimeLogic()
  };
  
  const passed = Object.values(results).filter(r => r === true || (Array.isArray(r) && r.every(s => s.startsWith('✓')))).length;
  const total = Object.keys(results).length;
  
  Logger.log(`\n=== Test Summary: ${passed}/${total} passed ===`);
  
  if (passed === total) {
    SpreadsheetApp.getActive().toast('NewFlow tests passed! ✓', 'Test Results', 3);
  } else {
    SpreadsheetApp.getActive().toast(`NewFlow tests: ${passed}/${total} passed`, 'Test Results', 5);
  }
  
  return results;
}