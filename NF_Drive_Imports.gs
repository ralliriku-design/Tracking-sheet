/**
 * NF_Drive_Imports.gs — Drive-based bulk imports for inventory reconciliation
 * 
 * Scans Google Drive folder for latest files by type and imports them into designated sheets.
 * Supports CSV and XLSX files with case-insensitive filename pattern matching.
 * 
 * Script Property Required:
 *   DRIVE_IMPORT_FOLDER_ID = 1yAkYYR6hetV3XATEJqg7qvy5NAJrFgKh
 * 
 * File Types & Target Sheets:
 *   - Quants ("quants") → Import_Quants (ERP)
 *   - Warehouse balance ("warehouse balance"|"warehouse_balance"|"3pl balance") → Import_Warehouse_Balance
 *   - ERP stock picking ("stock picking"|"erp picking") → Import_ERP_StockPicking
 *   - PBI deliveries ("deliveries"|"pbi deliveries"|"pbi_shipment"|"pbi_outbound") → Import_Weekly
 *   - PBI balances ("pbi balance"|"pbi stock"|"pbi inventory") → Import_PBI_Balance
 */

/**
 * Main function: scans the Drive folder and imports the latest matching files by type into import sheets.
 */
function NF_Drive_ImportLatestAll() {
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_IMPORT_FOLDER_ID');
  if (!folderId) {
    throw new Error('Script property DRIVE_IMPORT_FOLDER_ID not configured');
  }
  
  console.log(`NF_Drive_ImportLatestAll: Scanning folder ${folderId}`);
  
  // Define file type patterns and their target sheets
  const importTypes = [
    {
      patterns: ['quants'],
      sheetName: 'Import_Quants',
      description: 'ERP Quants'
    },
    {
      patterns: ['warehouse balance', 'warehouse_balance', '3pl balance'],
      sheetName: 'Import_Warehouse_Balance', 
      description: 'Warehouse Balance'
    },
    {
      patterns: ['stock picking', 'erp picking'],
      sheetName: 'Import_ERP_StockPicking',
      description: 'ERP Stock Picking'
    },
    {
      patterns: ['deliveries', 'pbi deliveries', 'pbi_shipment', 'pbi_outbound'],
      sheetName: 'Import_Weekly',
      description: 'PBI Deliveries'
    },
    {
      patterns: ['pbi balance', 'pbi stock', 'pbi inventory'],
      sheetName: 'Import_PBI_Balance',
      description: 'PBI Balance'
    }
  ];
  
  const results = [];
  
  for (const importType of importTypes) {
    try {
      const file = NF_Drive_PickLatestFileByPattern_(folderId, importType.patterns);
      if (file) {
        console.log(`Found ${importType.description}: ${file.getName()} (${file.getDateCreated()})`);
        NF_Drive_ReadCsvOrXlsxToSheet_(file, importType.sheetName);
        results.push({
          type: importType.description,
          file: file.getName(),
          sheet: importType.sheetName,
          status: 'success'
        });
      } else {
        console.log(`No file found for ${importType.description}`);
        results.push({
          type: importType.description,
          file: null,
          sheet: importType.sheetName,
          status: 'no_file'
        });
      }
    } catch (error) {
      console.error(`Error importing ${importType.description}:`, error);
      results.push({
        type: importType.description,
        file: null,
        sheet: importType.sheetName,
        status: 'error',
        error: error.message
      });
    }
  }
  
  // Log summary
  const successful = results.filter(r => r.status === 'success').length;
  const total = results.length;
  console.log(`NF_Drive_ImportLatestAll completed: ${successful}/${total} successful`);
  
  return results;
}

/**
 * Returns file metadata for the most recent file matching any of the patterns.
 * @param {string} folderId - Google Drive folder ID
 * @param {string[]} patterns - Array of filename patterns to match (case-insensitive)
 * @return {DriveApp.File|null} Most recent matching file or null
 */
function NF_Drive_PickLatestFileByPattern_(folderId, patterns) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    let latestFile = null;
    let latestDate = null;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName().toLowerCase();
      
      // Check if file is CSV or XLSX
      if (!fileName.endsWith('.csv') && !fileName.endsWith('.xlsx')) {
        continue;
      }
      
      // Check if filename contains any of the patterns
      const matches = patterns.some(pattern => 
        fileName.includes(pattern.toLowerCase())
      );
      
      if (matches) {
        const fileDate = file.getDateCreated();
        if (!latestFile || fileDate > latestDate) {
          latestFile = file;
          latestDate = fileDate;
        }
      }
    }
    
    return latestFile;
  } catch (error) {
    console.error('Error in NF_Drive_PickLatestFileByPattern_:', error);
    return null;
  }
}

/**
 * Reads CSV or XLSX file and writes raw data to the target sheet.
 * Clears the sheet first and sets frozen header row.
 * @param {DriveApp.File} file - Drive file to read
 * @param {string} sheetName - Target sheet name
 */
function NF_Drive_ReadCsvOrXlsxToSheet_(file, sheetName) {
  if (!file) {
    throw new Error('File is null or undefined');
  }
  
  console.log(`Reading file ${file.getName()} to sheet ${sheetName}`);
  
  // Read file content based on type
  let matrix;
  const fileName = file.getName().toLowerCase();
  
  if (fileName.endsWith('.csv')) {
    // Read CSV file
    const csvContent = file.getBlob().getDataAsString('UTF-8');
    matrix = Utilities.parseCsv(csvContent);
  } else if (fileName.endsWith('.xlsx')) {
    // Convert XLSX to Google Sheets and read data
    try {
      // Use Advanced Drive Service to convert XLSX to Google Sheets
      const tempFile = Drive.Files.insert({
        title: 'TempImport_' + Date.now(),
        mimeType: MimeType.GOOGLE_SHEETS
      }, file.getBlob(), {
        convert: true
      });
      
      const tempSs = SpreadsheetApp.openById(tempFile.id);
      const tempSheet = tempSs.getSheets()[0];
      matrix = tempSheet.getDataRange().getValues();
      
      // Clean up temporary file
      DriveApp.getFileById(tempFile.id).setTrashed(true);
    } catch (error) {
      console.error('Error converting XLSX file:', error);
      throw new Error(`Failed to convert XLSX file: ${error.message}`);
    }
  } else {
    throw new Error(`Unsupported file type: ${fileName}`);
  }
  
  if (!matrix || matrix.length === 0) {
    throw new Error('File contains no data');
  }
  
  // Normalize headers
  matrix = NF_Drive_NormalizeHeaders_(matrix);
  
  // Get or create target sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // Clear existing data
  sheet.clear();
  
  // Write data to sheet
  if (matrix.length > 0 && matrix[0].length > 0) {
    sheet.getRange(1, 1, matrix.length, matrix[0].length).setValues(matrix);
    
    // Set frozen header row
    sheet.setFrozenRows(1);
    
    // Auto-resize columns (limit to reasonable number)
    const columnsToResize = Math.min(matrix[0].length, 20);
    sheet.autoResizeColumns(1, columnsToResize);
  }
  
  console.log(`Successfully imported ${matrix.length} rows to ${sheetName}`);
}

/**
 * Normalizes header row by trimming extra spaces without altering data rows.
 * @param {any[][]} matrix - 2D array of data
 * @return {any[][]} Matrix with normalized headers
 */
function NF_Drive_NormalizeHeaders_(matrix) {
  if (!matrix || matrix.length === 0) {
    return matrix;
  }
  
  // Clone the matrix to avoid modifying the original
  const normalized = matrix.map(row => row.slice());
  
  // Normalize only the header row (first row)
  if (normalized.length > 0) {
    normalized[0] = normalized[0].map(header => 
      String(header || '').trim()
    );
  }
  
  return normalized;
}