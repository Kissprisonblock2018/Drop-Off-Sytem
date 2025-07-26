/**
 * COMPLETE TRANSPORT MANAGEMENT SYSTEM
 * 
 * This system manages transport across Dashboard, Transport, Completed Transports, and Completed Pick Up / Ship sheets
 * 
 * Process Flow:
 * 1. Analyze Dashboard Start/End locations to determine transport needs
 * 2. Populate Transport sheet with items needing transport (if not already in completed sheets)
 * 3. Process completed transports and move them to Completed Transports
 * 4. Update Dashboard Transported status
 */

/**
 * MAIN EXECUTION FUNCTION - Run this to execute the complete transport process
 */
function runCompleteTransportSystem() {
  try {
    Logger.log('🚀 Starting Complete Transport Management System...');
    
    const result = processCompleteTransportSystem();
    
    Logger.log('✅ Transport Management System completed successfully!');
    Logger.log(`📊 Results: ${JSON.stringify(result, null, 2)}`);
    
    // Show UI alert with results
    const message = `Transport System Complete!\n\n` +
      `Dashboard processed: ${result.dashboardProcessed} rows\n` +
      `Needs Transport (YES): ${result.needsTransportYes}\n` +
      `No Transport (NO): ${result.needsTransportNo}\n` +
      `Added to Transport: ${result.addedToTransport}\n` +
      `Marked as Transported: ${result.markedTransported}\n` +
      `Completed transports moved: ${result.completedTransportsMoved}`;
    
    SpreadsheetApp.getUi().alert('Transport System Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return result;
    
  } catch (error) {
    Logger.log(`❌ Error in runCompleteTransportSystem: ${error.message}`);
    Logger.log(`❌ Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert('Transport System Error', `Error: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * CORE TRANSPORT PROCESSING FUNCTION
 */
function processCompleteTransportSystem() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all required sheets
  const sheets = {
    dashboard: getSheetByPartialName(spreadsheet, 'Dashboard'),
    transport: getSheetByPartialName(spreadsheet, 'Transport'),
    completedTransports: getSheetByPartialName(spreadsheet, 'Completed Transports'),
    completedPickupShip: getSheetByPartialName(spreadsheet, 'Completed Pick Up')
  };
  
  // Validate required sheets
  if (!sheets.dashboard) {
    throw new Error('Dashboard sheet not found');
  }
  if (!sheets.transport) {
    throw new Error('Transport sheet not found');
  }
  if (!sheets.completedTransports) {
    throw new Error('Completed Transports sheet not found');
  }
  
  Logger.log('📋 Found sheets:');
  Object.entries(sheets).forEach(([key, sheet]) => {
    Logger.log(`  ${key}: ${sheet ? sheet.getName() : 'NOT FOUND'}`);
  });
  
  // STEP 1: Process Dashboard to determine transport needs
  Logger.log('\n🔍 STEP 1: Processing Dashboard for transport needs...');
  const dashboardResult = processDashboardTransportNeeds(sheets.dashboard, sheets.transport, sheets.completedTransports, sheets.completedPickupShip);
  
  // STEP 2: Process completed transports from Transport sheet
  Logger.log('\n📦 STEP 2: Processing completed transports...');
  const completedResult = processCompletedTransports(sheets.transport, sheets.completedTransports, sheets.dashboard);
  
  // STEP 3: Setup checkbox formatting
  Logger.log('\n✅ STEP 3: Setting up checkbox formatting...');
  setupCheckboxFormatting(sheets);
  
  const totalResult = {
    dashboardProcessed: dashboardResult.processed,
    needsTransportYes: dashboardResult.needsTransportYes,
    needsTransportNo: dashboardResult.needsTransportNo,
    addedToTransport: dashboardResult.addedToTransport,
    markedTransported: dashboardResult.markedTransported,
    completedTransportsMoved: completedResult.moved,
    skipped: dashboardResult.skipped
  };
  
  Logger.log('\n📊 FINAL RESULTS:');
  Logger.log(JSON.stringify(totalResult, null, 2));
  
  return totalResult;
}

/**
 * STEP 1: Process Dashboard to determine transport needs
 */
function processDashboardTransportNeeds(dashboardSheet, transportSheet, completedTransportsSheet, completedPickupShipSheet) {
  const data = dashboardSheet.getDataRange().getValues();
  
  // Find header row (should be around row 12 based on CSV analysis)
  let headerRowIndex = -1;
  let headers = null;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Look for row containing 'ID' as first column with data
    if (row[0] && row[0].toString().trim() === 'ID') {
      headerRowIndex = i;
      headers = row;
      break;
    }
  }
  
  if (headerRowIndex === -1) {
    throw new Error('Could not find header row with ID column in Dashboard');
  }
  
  Logger.log(`📍 Found headers at row ${headerRowIndex + 1}`);
  
  // Find column indices
  const columns = {
    id: findColumnIndex(headers, 'ID'),
    seller: findColumnIndex(headers, 'Seller'),
    customer: findColumnIndex(headers, 'Customer'),
    startLocation: findColumnIndex(headers, 'Start Location'),
    endLocation: findColumnIndex(headers, 'End Location'),
    needsTransport: findColumnIndex(headers, 'Needs Transport'),
    transported: findColumnIndex(headers, 'Transported')
  };
  
  // Validate required columns
  const missingColumns = [];
  Object.entries(columns).forEach(([key, index]) => {
    if (index === -1) {
      missingColumns.push(key);
    }
  });
  
  if (missingColumns.length > 0) {
    throw new Error(`Missing columns in Dashboard: ${missingColumns.join(', ')}`);
  }
  
  Logger.log('📋 Column indices:', columns);
  
  // Get existing records from completed sheets to avoid duplicates
  const existingRecords = getExistingTransportRecords(transportSheet, completedTransportsSheet, completedPickupShipSheet);
  Logger.log(`🔍 Found ${existingRecords.size} existing transport records`);
  
  let processed = 0;
  let needsTransportYes = 0;
  let needsTransportNo = 0;
  let addedToTransport = 0;
  let markedTransported = 0;
  let skipped = 0;
  
  // Process each data row (start after header row)
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    const id = row[columns.id];
    const seller = row[columns.seller] || '';
    const customer = row[columns.customer] || '';
    const startLocation = (row[columns.startLocation] || '').toString().trim();
    const endLocation = (row[columns.endLocation] || '').toString().trim();
    
    // Skip empty rows
    if (!id) {
      skipped++;
      continue;
    }
    
    Logger.log(`\n📝 Processing Row ${i + 1} (ID: ${id}):`);
    Logger.log(`  Start: "${startLocation}" | End: "${endLocation}"`);
    
    // Skip rows with no Start Location
    if (startLocation === '') {
      Logger.log(`  → SKIPPING (no Start Location)`);
      skipped++;
      continue;
    }
    
    // Skip rows with no End Location  
    if (endLocation === '') {
      Logger.log(`  → SKIPPING (no End Location)`);
      skipped++;
      continue;
    }
    
    // Determine transport needs (only process if both locations exist)
    let needsTransportValue;
    let shouldBeTransported = false;
    
    if (startLocation !== endLocation) {
      // Different locations = transport needed
      needsTransportValue = 'YES';
      shouldBeTransported = false;
      needsTransportYes++;
      Logger.log(`  → TRANSPORT NEEDED (${startLocation} → ${endLocation})`);
    } else {
      // Same locations = no transport needed
      needsTransportValue = 'NO';
      shouldBeTransported = true;
      needsTransportNo++;
      Logger.log(`  → NO TRANSPORT NEEDED (same location: ${startLocation})`);
    }
    
    // Update Needs Transport column
    dashboardSheet.getRange(i + 1, columns.needsTransport + 1).setValue(needsTransportValue);
    
    // Handle based on transport needs
    if (needsTransportValue === 'YES') {
      // Add to Transport sheet if not already in completed sheets
      const recordKey = `${id}|${seller.trim()}|${customer.trim()}`;
      
      if (!existingRecords.has(recordKey)) {
        addToTransportSheet(transportSheet, {
          id: id,
          seller: seller.trim(),
          customer: customer.trim(),
          transported: false
        });
        addedToTransport++;
        Logger.log(`  → ADDED TO TRANSPORT SHEET`);
      } else {
        Logger.log(`  → ALREADY EXISTS in transport/completed sheets`);
      }
    } else if (needsTransportValue === 'NO' && shouldBeTransported) {
      // Mark as transported since no transport is needed
      const transportedCell = dashboardSheet.getRange(i + 1, columns.transported + 1);
      transportedCell.setValue(true);
      markedTransported++;
      Logger.log(`  → MARKED AS TRANSPORTED (no transport needed)`);
    }
    
    processed++;
  }
  
  Logger.log(`\n📊 Dashboard Processing Results:`);
  Logger.log(`  Processed: ${processed}`);
  Logger.log(`  Needs Transport YES: ${needsTransportYes}`);
  Logger.log(`  Needs Transport NO: ${needsTransportNo}`);
  Logger.log(`  Added to Transport: ${addedToTransport}`);
  Logger.log(`  Marked Transported: ${markedTransported}`);
  Logger.log(`  Skipped: ${skipped}`);
  
  return {
    processed,
    needsTransportYes,
    needsTransportNo,
    addedToTransport,
    markedTransported,
    skipped
  };
}

/**
 * STEP 2: Process completed transports from Transport sheet
 */
function processCompletedTransports(transportSheet, completedTransportsSheet, dashboardSheet) {
  const data = transportSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('No data in Transport sheet to process');
    return { moved: 0 };
  }
  
  const headers = data[0];
  const transportedCol = findColumnIndex(headers, 'Transported');
  const idCol = findColumnIndex(headers, 'ID');
  const sellerCol = findColumnIndex(headers, 'Seller');
  const customerCol = findColumnIndex(headers, 'Customer');
  
  if (transportedCol === -1 || idCol === -1) {
    Logger.log('Required columns not found in Transport sheet');
    return { moved: 0 };
  }
  
  let moved = 0;
  
  // Process from bottom to top to avoid index issues when deleting rows
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const isTransported = row[transportedCol];
    const id = row[idCol];
    const seller = row[sellerCol] || '';
    const customer = row[customerCol] || '';
    
    // Check if transported is true
    if (isTransported === true || isTransported === 'TRUE' || isTransported === 'True') {
      Logger.log(`📦 Moving completed transport: ID ${id}`);
      
      // Add to Completed Transports with current date
      addToCompletedTransports(completedTransportsSheet, row, headers);
      
      // Update Dashboard Transported status
      updateDashboardTransported(dashboardSheet, id, seller, customer);
      
      // Remove from Transport sheet
      transportSheet.deleteRow(i + 1);
      
      moved++;
    }
  }
  
  Logger.log(`📦 Moved ${moved} completed transports`);
  return { moved };
}

/**
 * Get existing transport records from all relevant sheets
 */
function getExistingTransportRecords(transportSheet, completedTransportsSheet, completedPickupShipSheet) {
  const existingRecords = new Set();
  
  // Check Transport sheet
  if (transportSheet) {
    const records = getRecordsFromSheet(transportSheet);
    records.forEach(record => existingRecords.add(record));
    Logger.log(`📋 Transport sheet: ${records.size} records`);
  }
  
  // Check Completed Transports sheet
  if (completedTransportsSheet) {
    const records = getRecordsFromSheet(completedTransportsSheet);
    records.forEach(record => existingRecords.add(record));
    Logger.log(`📋 Completed Transports sheet: ${records.size} records`);
  }
  
  // Check Completed Pick Up / Ship sheet
  if (completedPickupShipSheet) {
    const records = getRecordsFromSheet(completedPickupShipSheet);
    records.forEach(record => existingRecords.add(record));
    Logger.log(`📋 Completed Pick Up/Ship sheet: ${records.size} records`);
  }
  
  return existingRecords;
}

/**
 * Get records from a sheet as Set of "ID|Seller|Customer" keys
 */
function getRecordsFromSheet(sheet) {
  const records = new Set();
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return records;
    
    const headers = data[0];
    const idCol = findColumnIndex(headers, 'ID');
    const sellerCol = findColumnIndex(headers, 'Seller');
    const customerCol = findColumnIndex(headers, 'Customer');
    
    if (idCol === -1) return records;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = row[idCol];
      const seller = (row[sellerCol] || '').toString().trim();
      const customer = (row[customerCol] || '').toString().trim();
      
      if (id) {
        const recordKey = `${id}|${seller}|${customer}`;
        records.add(recordKey);
      }
    }
  } catch (error) {
    Logger.log(`Warning: Error reading records from ${sheet.getName()}: ${error.message}`);
  }
  
  return records;
}

/**
 * Add a record to the Transport sheet
 */
function addToTransportSheet(transportSheet, record) {
  try {
    // Check if Transport sheet has headers
    let data = transportSheet.getDataRange().getValues();
    
    if (data.length === 0 || !data[0] || data[0].length === 0 || data[0][0] === '') {
      // Add headers if sheet is empty
      Logger.log('📝 Adding headers to Transport sheet...');
      transportSheet.getRange(1, 1, 1, 4).setValues([['ID', 'Seller', 'Customer', 'Transported']]);
      data = transportSheet.getDataRange().getValues();
    }
    
    // Check if record already exists
    const idCol = findColumnIndex(data[0], 'ID');
    const sellerCol = findColumnIndex(data[0], 'Seller');
    const customerCol = findColumnIndex(data[0], 'Customer');
    
    for (let i = 1; i < data.length; i++) {
      const existingId = data[i][idCol];
      const existingSeller = (data[i][sellerCol] || '').toString().trim();
      const existingCustomer = (data[i][customerCol] || '').toString().trim();
      
      if (existingId === record.id && 
          existingSeller === record.seller && 
          existingCustomer === record.customer) {
        Logger.log(`  Record already exists in Transport sheet`);
        return;
      }
    }
    
    // Add new record
    const nextRow = transportSheet.getLastRow() + 1;
    transportSheet.getRange(nextRow, 1).setValue(record.id);
    transportSheet.getRange(nextRow, 2).setValue(record.seller);
    transportSheet.getRange(nextRow, 3).setValue(record.customer);
    
    // Set transported column as checkbox
    const transportedCell = transportSheet.getRange(nextRow, 4);
    transportedCell.insertCheckboxes();
    transportedCell.setValue(record.transported);
    
  } catch (error) {
    Logger.log(`❌ Error adding to Transport sheet: ${error.message}`);
  }
}

/**
 * Add a record to Completed Transports with current date
 */
function addToCompletedTransports(completedTransportsSheet, transportRow, transportHeaders) {
  try {
    // Check if Completed Transports has headers
    let data = completedTransportsSheet.getDataRange().getValues();
    
    if (data.length === 0 || !data[0] || data[0].length === 0 || data[0][0] === '') {
      // Add headers
      const headers = [...transportHeaders, 'Date Transported'];
      completedTransportsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    
    // Add the record with current date
    const nextRow = completedTransportsSheet.getLastRow() + 1;
    const currentDate = new Date();
    
    // Copy transport data and add date
    for (let col = 0; col < transportRow.length; col++) {
      completedTransportsSheet.getRange(nextRow, col + 1).setValue(transportRow[col]);
    }
    
    // Add date transported
    completedTransportsSheet.getRange(nextRow, transportRow.length + 1).setValue(currentDate);
    
    // Set transported column as checkbox (should be column 4)
    const transportedCol = findColumnIndex(data[0] || transportHeaders, 'Transported');
    if (transportedCol !== -1) {
      const transportedCell = completedTransportsSheet.getRange(nextRow, transportedCol + 1);
      transportedCell.insertCheckboxes();
      transportedCell.setValue(true);
    }
    
  } catch (error) {
    Logger.log(`❌ Error adding to Completed Transports: ${error.message}`);
  }
}

/**
 * Update Dashboard Transported status for a specific record
 */
function updateDashboardTransported(dashboardSheet, targetId, targetSeller, targetCustomer) {
  try {
    const data = dashboardSheet.getDataRange().getValues();
    
    // Find header row
    let headerRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === 'ID') {
        headerRowIndex = i;
        break;
      }
    }
    
    if (headerRowIndex === -1) return;
    
    const headers = data[headerRowIndex];
    const idCol = findColumnIndex(headers, 'ID');
    const sellerCol = findColumnIndex(headers, 'Seller');
    const customerCol = findColumnIndex(headers, 'Customer');
    const transportedCol = findColumnIndex(headers, 'Transported');
    
    if (idCol === -1 || transportedCol === -1) return;
    
    // Find matching row
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const id = data[i][idCol];
      const seller = (data[i][sellerCol] || '').toString().trim();
      const customer = (data[i][customerCol] || '').toString().trim();
      
      if (id === targetId && seller === targetSeller && customer === targetCustomer) {
        const transportedCell = dashboardSheet.getRange(i + 1, transportedCol + 1);
        transportedCell.insertCheckboxes();
        transportedCell.setValue(true);
        Logger.log(`  ✅ Updated Dashboard Transported for ID ${targetId}`);
        break;
      }
    }
  } catch (error) {
    Logger.log(`❌ Error updating Dashboard: ${error.message}`);
  }
}

/**
 * Setup checkbox formatting for all boolean columns
 */
function setupCheckboxFormatting(sheets) {
  // Format Dashboard boolean columns
  if (sheets.dashboard) {
    formatBooleanColumns(sheets.dashboard);
  }
  
  // Format Transport boolean columns
  if (sheets.transport) {
    formatBooleanColumns(sheets.transport);
  }
  
  // Format Completed Transports boolean columns
  if (sheets.completedTransports) {
    formatBooleanColumns(sheets.completedTransports);
  }
}

/**
 * Format boolean columns in a sheet as checkboxes
 */
function formatBooleanColumns(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return;
    
    // Find header row
    let headerRowIndex = 0;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(cell => cell && cell.toString().includes('ID'))) {
        headerRowIndex = i;
        break;
      }
    }
    
    const headers = data[headerRowIndex];
    
    // Define boolean column names
    const booleanColumns = [
      'Created', 'Paid', 'Dropped Off', 'Transported', 
      'Ready For (Pickup or Ship)', 'Picked Up Or Shiped'
    ];
    
    booleanColumns.forEach(columnName => {
      const colIndex = findColumnIndex(headers, columnName);
      if (colIndex !== -1 && data.length > headerRowIndex + 1) {
        try {
          const range = sheet.getRange(
            headerRowIndex + 2, // Start after header row
            colIndex + 1,
            data.length - headerRowIndex - 1,
            1
          );
          range.insertCheckboxes();
        } catch (columnError) {
          // Column might not need formatting
        }
      }
    });
    
  } catch (error) {
    Logger.log(`Warning: Error formatting checkboxes in ${sheet.getName()}: ${error.message}`);
  }
}

/**
 * UTILITY FUNCTIONS
 */

/**
 * Find sheet by partial name (case-insensitive)
 */
function getSheetByPartialName(spreadsheet, partialName) {
  const sheets = spreadsheet.getSheets();
  const normalizedPartial = partialName.toLowerCase();
  
  // Try exact match first
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase() === normalizedPartial) {
      return sheet;
    }
  }
  
  // Try partial match
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().includes(normalizedPartial)) {
      return sheet;
    }
  }
  
  return null;
}

/**
 * Find column index by name (case-insensitive, handles extra spaces)
 */
function findColumnIndex(headers, columnName) {
  const searchName = columnName.toLowerCase().trim();
  
  for (let i = 0; i < headers.length; i++) {
    const headerName = (headers[i] || '').toString().toLowerCase().trim();
    if (headerName === searchName) {
      return i;
    }
  }
  
  return -1;
}

/**
 * DEBUG FUNCTIONS
 */

/**
 * Show current status of all sheets
 */
function debugTransportSystem() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log('=== TRANSPORT SYSTEM DEBUG ===');
    
    const sheets = spreadsheet.getSheets();
    Logger.log(`📋 Available sheets: ${sheets.map(s => s.getName()).join(', ')}`);
    
    // Check each relevant sheet
    const sheetNames = ['Dashboard', 'Transport', 'Completed Transports', 'Completed Pick Up'];
    
    sheetNames.forEach(partialName => {
      const sheet = getSheetByPartialName(spreadsheet, partialName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        Logger.log(`\n📊 ${sheet.getName()}:`);
        Logger.log(`  Rows: ${data.length}`);
        if (data.length > 0) {
          Logger.log(`  Headers: ${data[0].join(', ')}`);
        }
      } else {
        Logger.log(`\n❌ ${partialName}: NOT FOUND`);
      }
    });
    
  } catch (error) {
    Logger.log(`❌ Debug error: ${error.message}`);
  }
}

/**
 * Test function to check Dashboard transport logic without making changes
 */
function testTransportLogic() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = getSheetByPartialName(spreadsheet, 'Dashboard');
    
    if (!dashboardSheet) {
      throw new Error('Dashboard sheet not found');
    }
    
    const data = dashboardSheet.getDataRange().getValues();
    
    // Find header row
    let headerRowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === 'ID') {
        headerRowIndex = i;
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      throw new Error('Header row not found');
    }
    
    const headers = data[headerRowIndex];
    const columns = {
      id: findColumnIndex(headers, 'ID'),
      startLocation: findColumnIndex(headers, 'Start Location'),
      endLocation: findColumnIndex(headers, 'End Location'),
      needsTransport: findColumnIndex(headers, 'Needs Transport'),
      transported: findColumnIndex(headers, 'Transported')
    };
    
    Logger.log('=== TRANSPORT LOGIC TEST ===');
    Logger.log(`📍 Header row: ${headerRowIndex + 1}`);
    Logger.log(`📋 Columns: ${JSON.stringify(columns)}`);
    
    Logger.log('\n📊 First 10 rows analysis:');
    
    for (let i = headerRowIndex + 1; i < Math.min(data.length, headerRowIndex + 11); i++) {
      const row = data[i];
      const id = row[columns.id];
      const startLocation = (row[columns.startLocation] || '').toString().trim();
      const endLocation = (row[columns.endLocation] || '').toString().trim();
      const currentNeedsTransport = row[columns.needsTransport];
      const currentTransported = row[columns.transported];
      
      if (!id) continue;
      
      let suggestedNeedsTransport;
      let suggestedTransported;
      
      if (startLocation === '') {
        suggestedNeedsTransport = 'SKIP (no Start Location)';
        suggestedTransported = 'SKIP (no Start Location)';
      } else if (endLocation === '') {
        suggestedNeedsTransport = 'SKIP (no End Location)';
        suggestedTransported = 'SKIP (no End Location)';
      } else if (startLocation !== endLocation) {
        suggestedNeedsTransport = 'YES';
        suggestedTransported = 'FALSE (add to Transport)';
      } else {
        suggestedNeedsTransport = 'NO';
        suggestedTransported = 'TRUE';
      }
      
      Logger.log(`\nRow ${i + 1} (ID: ${id}):`);
      Logger.log(`  Start: "${startLocation}" | End: "${endLocation}"`);
      Logger.log(`  Current Needs Transport: "${currentNeedsTransport}"`);
      Logger.log(`  Current Transported: "${currentTransported}"`);
      Logger.log(`  Suggested Needs Transport: ${suggestedNeedsTransport}`);
      Logger.log(`  Suggested Action: ${suggestedTransported}`);
    }
    
  } catch (error) {
    Logger.log(`❌ Test error: ${error.message}`);
  }
}

/**
 * CLEAR FUNCTIONS (for testing)
 */

/**
 * Clear Transport sheet (for testing purposes)
 */
function clearTransportSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const transportSheet = getSheetByPartialName(spreadsheet, 'Transport');
    
    if (!transportSheet) {
      Logger.log('Transport sheet not found');
      return;
    }
    
    // Keep headers but clear all data
    const lastRow = transportSheet.getLastRow();
    if (lastRow > 1) {
      transportSheet.getRange(2, 1, lastRow - 1, transportSheet.getLastColumn()).clearContent();
      Logger.log(`Cleared ${lastRow - 1} rows from Transport sheet`);
    } else {
      Logger.log('Transport sheet is already empty');
    }
    
  } catch (error) {
    Logger.log(`Error clearing Transport sheet: ${error.message}`);
  }
}

/**
 * Clear Completed Transports sheet (for testing purposes)
 */
function clearCompletedTransportsSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const completedSheet = getSheetByPartialName(spreadsheet, 'Completed Transports');
    
    if (!completedSheet) {
      Logger.log('Completed Transports sheet not found');
      return;
    }
    
    // Keep headers but clear all data
    const lastRow = completedSheet.getLastRow();
    if (lastRow > 1) {
      completedSheet.getRange(2, 1, lastRow - 1, completedSheet.getLastColumn()).clearContent();
      Logger.log(`Cleared ${lastRow - 1} rows from Completed Transports sheet`);
    } else {
      Logger.log('Completed Transports sheet is already empty');
    }
    
  } catch (error) {
    Logger.log(`Error clearing Completed Transports sheet: ${error.message}`);
  }
}

/**
 * EXECUTION SHORTCUTS
 */

/**
 * Quick setup - run this once to initialize the Transport sheet structure
 */
function setupTransportSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let transportSheet = getSheetByPartialName(spreadsheet, 'Transport');
    
    if (!transportSheet) {
      // Create Transport sheet if it doesn't exist
      transportSheet = spreadsheet.insertSheet('Transport');
      Logger.log('Created new Transport sheet');
    }
    
    // Add headers if sheet is empty
    const data = transportSheet.getDataRange().getValues();
    if (data.length === 0 || !data[0] || data[0][0] === '') {
      transportSheet.getRange(1, 1, 1, 4).setValues([['ID', 'Seller', 'Customer', 'Transported']]);
      Logger.log('Added headers to Transport sheet');
    }
    
    // Setup Completed Transports sheet
    let completedSheet = getSheetByPartialName(spreadsheet, 'Completed Transports');
    if (!completedSheet) {
      completedSheet = spreadsheet.insertSheet('Completed Transports');
      Logger.log('Created new Completed Transports sheet');
    }
    
    const completedData = completedSheet.getDataRange().getValues();
    if (completedData.length === 0 || !completedData[0] || completedData[0][0] === '') {
      completedSheet.getRange(1, 1, 1, 5).setValues([['ID', 'Seller', 'Customer', 'Transported', 'Date Transported']]);
      Logger.log('Added headers to Completed Transports sheet');
    }
    
    Logger.log('✅ Transport system setup complete');
    
  } catch (error) {
    Logger.log(`❌ Setup error: ${error.message}`);
  }
}

/**
 * Manual test run (processes only first 10 rows for testing)
 */
function testRunTransportSystem() {
  try {
    Logger.log('🧪 Running transport system in TEST MODE (first 10 rows only)...');
    
    // This would need to be implemented with a row limit for testing
    // For now, just run the debug functions
    debugTransportSystem();
    testTransportLogic();
    
    Logger.log('🧪 Test run complete - check logs for analysis');
    
  } catch (error) {
    Logger.log(`❌ Test run error: ${error.message}`);
  }
}

/**
 * TRIGGER SETUP
 */

/**
 * Set up automatic triggers (optional - run once to automate)
 */
function setupAutoTriggers() {
  try {
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    
    // Create trigger to run every hour
    ScriptApp.newTrigger('runCompleteTransportSystem')
      .timeBased()
      .everyHours(1)
      .create();
    
    Logger.log('✅ Auto-trigger set up to run every hour');
    
  } catch (error) {
    Logger.log(`❌ Trigger setup error: ${error.message}`);
  }
}

/**
 * Remove all triggers
 */
function removeAutoTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    Logger.log('✅ All auto-triggers removed');
  } catch (error) {
    Logger.log(`❌ Error removing triggers: ${error.message}`);
  }
}

/**
 * SUMMARY FUNCTION
 */
function showTransportSystemSummary() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get sheet statistics
    const dashboardSheet = getSheetByPartialName(spreadsheet, 'Dashboard');
    const transportSheet = getSheetByPartialName(spreadsheet, 'Transport');
    const completedSheet = getSheetByPartialName(spreadsheet, 'Completed Transports');
    
    let summary = "=== TRANSPORT SYSTEM SUMMARY ===\n\n";
    
    if (dashboardSheet) {
      const dashData = dashboardSheet.getDataRange().getValues();
      summary += `📊 Dashboard: ${dashData.length - 1} total rows\n`;
    }
    
    if (transportSheet) {
      const transportData = transportSheet.getDataRange().getValues();
      const activeTransports = transportData.length > 1 ? transportData.length - 1 : 0;
      summary += `🚛 Transport (active): ${activeTransports} items\n`;
    }
    
    if (completedSheet) {
      const completedData = completedSheet.getDataRange().getValues();
      const completedTransports = completedData.length > 1 ? completedData.length - 1 : 0;
      summary += `✅ Completed Transports: ${completedTransports} items\n`;
    }
    
    summary += "\n🚀 To run the complete system: runCompleteTransportSystem()";
    summary += "\n🧪 To test first: testRunTransportSystem()";
    summary += "\n🔧 To setup sheets: setupTransportSheet()";
    
    Logger.log(summary);
    
    // Also show in UI
    SpreadsheetApp.getUi().alert('Transport System Summary', summary, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    Logger.log(`❌ Summary error: ${error.message}`);
  }
}
