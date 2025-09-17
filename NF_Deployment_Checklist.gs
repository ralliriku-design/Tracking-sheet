/**
 * NF_Deployment_Checklist.gs — Deployment validation and checklist
 * 
 * Comprehensive validation script to ensure all NF components are properly
 * configured and ready for production use.
 */

/**
 * Complete deployment validation checklist.
 * Run this before using the NF extension in production.
 */
function NF_validateDeployment() {
  console.log('🔍 NF Extension Deployment Validation');
  console.log('=====================================\n');
  
  const results = [];
  
  // 1. Check Script Properties
  console.log('1. Script Properties Configuration');
  console.log('----------------------------------');
  
  const driveImportId = PropertiesService.getScriptProperties().getProperty('DRIVE_IMPORT_FOLDER_ID');
  if (driveImportId) {
    console.log('✅ DRIVE_IMPORT_FOLDER_ID configured:', driveImportId);
    results.push({ check: 'Drive Import Folder ID', status: 'PASS' });
  } else {
    console.log('❌ DRIVE_IMPORT_FOLDER_ID not configured');
    results.push({ check: 'Drive Import Folder ID', status: 'FAIL', fix: 'Set DRIVE_IMPORT_FOLDER_ID = 1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh' });
  }
  
  // 2. Check Advanced Services
  console.log('\n2. Advanced Services');
  console.log('-------------------');
  
  try {
    // Try to access Drive API
    const testCall = Drive.Files;
    console.log('✅ Drive API advanced service enabled');
    results.push({ check: 'Drive API Service', status: 'PASS' });
  } catch (error) {
    console.log('❌ Drive API advanced service not enabled');
    results.push({ check: 'Drive API Service', status: 'FAIL', fix: 'Enable Drive API in Services' });
  }
  
  // 3. Check NF Functions
  console.log('\n3. NF Function Availability');
  console.log('--------------------------');
  
  const nfFunctions = [
    'NF_Drive_ImportLatestAll',
    'NF_BulkRebuildAll',
    'NF_RefreshAllPending',
    'NF_UpdateInventoryBalances',
    'NF_buildSokKarkkainenAlways'
  ];
  
  let functionsOk = 0;
  for (const funcName of nfFunctions) {
    try {
      const func = eval(funcName);
      if (typeof func === 'function') {
        console.log(`✅ ${funcName}`);
        functionsOk++;
      } else {
        console.log(`❌ ${funcName} not a function`);
      }
    } catch (error) {
      console.log(`❌ ${funcName} not available`);
    }
  }
  
  results.push({ 
    check: 'NF Functions', 
    status: functionsOk === nfFunctions.length ? 'PASS' : 'FAIL',
    details: `${functionsOk}/${nfFunctions.length} available`
  });
  
  // 4. Check Helper Functions
  console.log('\n4. Helper Function Dependencies');
  console.log('------------------------------');
  
  const helperStatus = NF_checkHelperFunctions();
  
  if (helperStatus.critical.length === 0) {
    console.log('✅ All critical helper functions available');
    results.push({ check: 'Critical Helpers', status: 'PASS' });
  } else {
    console.log('❌ Missing critical helpers:', helperStatus.critical.join(', '));
    results.push({ 
      check: 'Critical Helpers', 
      status: 'FAIL', 
      fix: 'Ensure Traacking.txt or equivalent helper functions are included'
    });
  }
  
  if (helperStatus.missing.length > 0) {
    console.log('⚠️ Missing optional helpers:', helperStatus.missing.join(', '));
    console.log('  (Fallback implementations will be used)');
  }
  
  // 5. Check File Patterns
  console.log('\n5. File Pattern Recognition Test');
  console.log('------------------------------');
  
  const patternTest = NF_test_FilePatternMatching();
  results.push({ 
    check: 'File Patterns', 
    status: patternTest ? 'PASS' : 'FAIL'
  });
  
  // 6. Check Menu Integration
  console.log('\n6. Menu Integration');
  console.log('-----------------');
  
  try {
    // Test menu creation (won't actually show menu)
    const ui = SpreadsheetApp.getUi();
    console.log('✅ UI access available for menu creation');
    results.push({ check: 'Menu Integration', status: 'PASS' });
  } catch (error) {
    console.log('❌ UI access failed:', error.message);
    results.push({ check: 'Menu Integration', status: 'FAIL' });
  }
  
  // 7. Generate Summary
  console.log('\n📊 Validation Summary');
  console.log('====================');
  
  const passed = results.filter(r => r.status === 'PASS').length;
  const failed = results.filter(r => r.status === 'FAIL').length;
  const total = results.length;
  
  console.log(`Total Checks: ${total}`);
  console.log(`Passed: ${passed}`);
  console.log(`Failed: ${failed}`);
  
  if (failed === 0) {
    console.log('\n🎉 All validations passed! NF extension is ready for deployment.');
  } else {
    console.log('\n⚠️ Some validations failed. Review the following:');
    results.filter(r => r.status === 'FAIL').forEach(result => {
      console.log(`  • ${result.check}: ${result.fix || 'See documentation'}`);
    });
  }
  
  return { passed, failed, results };
}

/**
 * Show installation instructions.
 */
function NF_showInstallationInstructions() {
  const instructions = `
NF Extension Installation Instructions
=====================================

1. SCRIPT PROPERTIES
   Add to Project Settings → Script Properties:
   DRIVE_IMPORT_FOLDER_ID = 1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh

2. ADVANCED SERVICES
   Enable in Apps Script Editor → Services:
   • Google Drive API

3. MENU INTEGRATION
   Add to existing onOpen() function:
   
   function onOpen() {
     // ... existing menu code ...
     .addToUi();
     
     // Add NF extensions
     NF_addStandaloneMenus();
   }

4. VERIFY INSTALLATION
   Run: NF_validateDeployment()

5. TEST FUNCTIONALITY
   Run: NF_runAllTests()

6. SCHEDULE AUTOMATION
   Use NF Scheduling menu to set up:
   • Daily 00:01: Inventory updates
   • Daily 11:00: Tracking refresh
   • Weekly Mon 02:00: Reports

For detailed documentation, see NF_README.md
`;

  console.log(instructions);
  
  // Also try to show in UI if available
  try {
    SpreadsheetApp.getUi().alert('NF Installation Instructions', instructions, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    // UI not available, just log
  }
  
  return instructions;
}

/**
 * Quick setup function that attempts to configure basic settings.
 */
function NF_quickSetup() {
  console.log('🚀 NF Extension Quick Setup');
  console.log('===========================\n');
  
  try {
    // Set Drive Import Folder ID if not already set
    const sp = PropertiesService.getScriptProperties();
    const currentId = sp.getProperty('DRIVE_IMPORT_FOLDER_ID');
    
    if (!currentId) {
      sp.setProperty('DRIVE_IMPORT_FOLDER_ID', '1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh');
      console.log('✅ Set DRIVE_IMPORT_FOLDER_ID');
    } else {
      console.log('✅ DRIVE_IMPORT_FOLDER_ID already configured');
    }
    
    // Test Drive API access
    try {
      Drive.Files.list({ maxResults: 1 });
      console.log('✅ Drive API access confirmed');
    } catch (error) {
      console.log('❌ Drive API not accessible. Enable in Services → Drive API');
    }
    
    // Add menus if possible
    try {
      NF_addStandaloneMenus();
      console.log('✅ NF menus added');
    } catch (error) {
      console.log('⚠️ Could not add menus automatically. Add NF_addStandaloneMenus() to onOpen()');
    }
    
    console.log('\n🎯 Quick setup completed! Run NF_validateDeployment() for full validation.');
    
  } catch (error) {
    console.error('❌ Quick setup failed:', error);
    console.log('Please follow manual installation instructions.');
  }
}

/**
 * Show available NF functions and their descriptions.
 */
function NF_showAvailableFunctions() {
  const functions = [
    {
      name: 'NF_BulkRebuildAll()',
      description: 'Complete daily workflow: Drive import + inventory + tracking + reports',
      category: 'Main'
    },
    {
      name: 'NF_BulkImportFromDrive()',
      description: 'Import all files from configured Drive folder',
      category: 'Import'
    },
    {
      name: 'NF_RefreshAllPending(limitPerRun)',
      description: 'Batch refresh tracking statuses for pending shipments',
      category: 'Tracking'
    },
    {
      name: 'NF_UpdateInventoryBalances()',
      description: 'Build inventory aggregates and reconciliation reports',
      category: 'Inventory'
    },
    {
      name: 'NF_buildSokKarkkainenAlways()',
      description: 'Merge all historical data into SOK/Kärkkäinen reports',
      category: 'Reports'
    },
    {
      name: 'NF_BulkFindDuplicates_All()',
      description: 'Find duplicates across all data sources',
      category: 'Quality'
    },
    {
      name: 'NF_validateDeployment()',
      description: 'Validate installation and configuration',
      category: 'Setup'
    },
    {
      name: 'NF_runAllTests()',
      description: 'Run integration tests',
      category: 'Testing'
    }
  ];
  
  console.log('📚 Available NF Functions');
  console.log('=========================\n');
  
  const categories = [...new Set(functions.map(f => f.category))];
  
  for (const category of categories) {
    console.log(`${category}:`);
    console.log('-'.repeat(category.length + 1));
    
    functions.filter(f => f.category === category).forEach(func => {
      console.log(`  ${func.name}`);
      console.log(`    ${func.description}\n`);
    });
  }
  
  return functions;
}