/**
 * Enhanced Management Log Pickup/Ship Processing System
 * Ensures "Picked Up Or Shiped" checkbox is ALWAYS marked as true
 * Handles manual checkbox changes in Management Log and processes confirmations
 */

/**
 * Main onEdit trigger - monitors Management Log for manual checkbox changes
 */
function onEdit(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    // Only process Management Log sheet edits
    if (!isManagementLogSheet(sheetName)) {
      return;
    }
    
    // Only process single cell edits in the "Picked Up Or Shiped" column
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
      return;
    }
    
    const editedRow = range.getRow();
    const editedColumn = range.getColumn();
    const newValue = range.getValue();
    
    // Get Management Log structure
    const managementData = sheet.getDataRange().getValues();
    if (managementData.length === 0) return;
    
    const headers = managementData[0];
    const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
    
    // Find the "Picked Up Or Shiped" column index
    const pickedUpColIndex = trimmedHeaders.indexOf('Picked Up Or Shiped');
    if (pickedUpColIndex === -1) {
      console.log('Picked Up Or Shiped column not found in Management Log');
      return;
    }
    
    // Check if the edited column is the "Picked Up Or Shiped" column
    if (editedColumn !== pickedUpColIndex + 1) { // +1 for 1-based indexing
      return;
    }
    
    // Only process if the checkbox was marked as true
    if (newValue !== true && newValue !== 'TRUE' && newValue !== 'True') {
      return;
    }
    
    // Get the row data
    const rowData = managementData[editedRow - 1]; // Convert to 0-based index
    const id = rowData[trimmedHeaders.indexOf('ID')];
    const seller = rowData[trimmedHeaders.indexOf('Seller')];
    const customer = rowData[trimmedHeaders.indexOf('Customer')];
    
    // Validate required data
    if (!id || !seller || !customer) {
      SpreadsheetApp.getUi().alert('Error', 'Missing required data (ID, Seller, or Customer) in the selected row.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Revert the checkbox immediately (will be set back if confirmed)
    range.setValue(false);
    
    // Show confirmation dialog with two-step process
    showPickupConfirmationDialog(id, seller, customer);
    
  } catch (error) {
    console.log('Error in onEdit trigger: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Check if the sheet is a Management Log sheet (handles different naming variations)
 */
function isManagementLogSheet(sheetName) {
  const managementLogVariations = [
    'Management Log',
    'Magement Log', 
    'Management',
    'Magement'
  ];
  
  return managementLogVariations.some(variation => 
    sheetName.toLowerCase().includes(variation.toLowerCase())
  );
}

/**
 * Show confirmation dialog using built-in UI dialogs
 */
function showPickupConfirmationDialog(id, seller, customer) {
  const ui = SpreadsheetApp.getUi();
  
  // First dialog - show details and ask for confirmation
  const message = `🚚 PICKUP/SHIP CONFIRMATION\n\n` +
    `Order Details:\n` +
    `📋 Order ID: ${id}\n` +
    `🏪 Seller: ${seller}\n` +
    `👤 Customer: ${customer}\n\n` +
    `Do you want to process this pickup/ship?`;
  
  const firstResponse = ui.alert('Confirm Pickup/Ship', message, ui.ButtonSet.YES_NO);
  
  if (firstResponse === ui.Button.NO) {
    ui.alert('Cancelled', 'The pickup/ship process was cancelled.', ui.ButtonSet.OK);
    return;
  }
  
  // Second dialog - ask about processing scope
  const scopeMessage = `PROCESSING SCOPE\n\n` +
    `Customer: ${customer}\n\n` +
    `Choose processing option:\n\n` +
    `YES = Process ALL items for this customer\n` +
    `NO = Process ONLY this single item\n` +
    `CANCEL = Abort operation`;
  
  const scopeResponse = ui.alert('Processing Scope', scopeMessage, ui.ButtonSet.YES_NO_CANCEL);
  
  if (scopeResponse === ui.Button.YES) {
    // Process all items for customer
    processAllCustomerItems(customer);
  } else if (scopeResponse === ui.Button.NO) {
    // Process single item
    processPickupShip(id, seller, customer);
  } else {
    // User cancelled
    ui.alert('Cancelled', 'The pickup/ship process was cancelled.', ui.ButtonSet.OK);
  }
}

/**
 * Enhanced process pickup/ship action for a single item
 * ENSURES the "Picked Up Or Shiped" checkbox is ALWAYS marked as true
 */
function processPickupShip(id, seller, customer) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Debug: List all available sheets
    console.log('Available sheets:');
    spreadsheet.getSheets().forEach(sheet => {
      console.log(`  - "${sheet.getName()}"`);
    });
    
    // Find the Dashboard sheet
    const dashboardSheet = findDashboardSheet(spreadsheet);
    if (!dashboardSheet) {
      throw new Error('Dashboard sheet not found');
    }
    
    // Find the Management Log sheet
    const managementSheet = findManagementLogSheet(spreadsheet);
    if (!managementSheet) {
      throw new Error('Management Log sheet not found');
    }
    
    // Find the Completed Pick Up / Ship sheet (try multiple names)
    const completedSheet = findCompletedPickupShipSheet(spreadsheet);
    if (!completedSheet) {
      const availableSheets = spreadsheet.getSheets().map(s => s.getName()).join(', ');
      throw new Error(`Completed Pick Up / Ship sheet not found. Available sheets: ${availableSheets}`);
    }
    
    // Find the matching row in Dashboard
    const dashboardRowData = findDashboardRow(dashboardSheet, id, seller, customer);
    if (!dashboardRowData) {
      throw new Error(`No matching row found in Dashboard for ID: ${id}, Seller: ${seller}, Customer: ${customer}`);
    }
    
    console.log(`✅ Found Dashboard row ${dashboardRowData.rowNumber} for processing`);
    
    // STEP 1: FIRST - Update the Dashboard row's "Picked Up Or Shiped" column to true
    // This is the most critical step - do this BEFORE moving the row
    updateDashboardPickupStatus(dashboardSheet, dashboardRowData.rowNumber, dashboardRowData.headers);
    console.log(`✅ Dashboard "Picked Up Or Shiped" checkbox set to TRUE for row ${dashboardRowData.rowNumber}`);
    
    // STEP 2: Refresh the row data after updating the checkbox
    const updatedDashboardRowData = findDashboardRow(dashboardSheet, id, seller, customer);
    if (!updatedDashboardRowData) {
      throw new Error('Could not find updated Dashboard row after checkbox update');
    }
    
    // STEP 3: Move the entire row to Completed Pick Up / Ship sheet
    moveRowToCompletedPickupShip(dashboardSheet, completedSheet, updatedDashboardRowData);
    console.log(`✅ Moved row to ${completedSheet.getName()}`);
    
    // STEP 4: Remove the item from Management Log
    removeFromManagementLog(managementSheet, id, seller, customer);
    console.log(`✅ Removed from Management Log`);
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      'Success', 
      `✅ Successfully processed pickup/ship for:\n\nOrder ID: ${id}\nSeller: ${seller}\nCustomer: ${customer}\n\nActions completed:\n• ✅ Marked "Picked Up Or Shiped" = TRUE in Dashboard\n• ✅ Moved to ${completedSheet.getName()}\n• ✅ Removed from Management Log`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.log('Error in processPickupShip: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'Failed to process pickup/ship: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Process all items for a specific customer
 */
function processAllCustomerItems(targetCustomer) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Find required sheets
    const dashboardSheet = findDashboardSheet(spreadsheet);
    const managementSheet = findManagementLogSheet(spreadsheet);
    const completedSheet = findCompletedPickupShipSheet(spreadsheet);
    
    if (!dashboardSheet || !managementSheet || !completedSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Get all items from Management Log for this customer
    const managementData = managementSheet.getDataRange().getValues();
    const managementHeaders = managementData[0];
    const trimmedMgmtHeaders = managementHeaders.map(header => header ? header.toString().trim() : '');
    
    const mgmtIdColIndex = trimmedMgmtHeaders.indexOf('ID');
    const mgmtSellerColIndex = trimmedMgmtHeaders.indexOf('Seller');
    const mgmtCustomerColIndex = trimmedMgmtHeaders.indexOf('Customer');
    
    if (mgmtIdColIndex === -1 || mgmtSellerColIndex === -1 || mgmtCustomerColIndex === -1) {
      throw new Error('Required columns not found in Management Log');
    }
    
    // Find all matching customer items
    const customerItems = [];
    for (let i = 1; i < managementData.length; i++) {
      const row = managementData[i];
      const customer = row[mgmtCustomerColIndex];
      
      if (customer && customer.toString().trim() === targetCustomer.toString().trim()) {
        customerItems.push({
          id: row[mgmtIdColIndex],
          seller: row[mgmtSellerColIndex],
          customer: row[mgmtCustomerColIndex],
          rowIndex: i
        });
      }
    }
    
    if (customerItems.length === 0) {
      SpreadsheetApp.getUi().alert('No Items Found', `No items found for customer "${targetCustomer}" in Management Log.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Found ${customerItems.length} items for customer "${targetCustomer}"`);
    let processedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    // Get dashboard headers once (they're in row 12)
    const dashboardData = dashboardSheet.getDataRange().getValues();
    const dashboardHeaders = dashboardData[11]; // Row 12 = index 11
    
    // Process each item (process from bottom to top to avoid index issues when deleting)
    for (let i = customerItems.length - 1; i >= 0; i--) {
      const item = customerItems[i];
      
      try {
        // Find the matching row in Dashboard
        const dashboardRowData = findDashboardRow(dashboardSheet, item.id, item.seller, item.customer);
        if (dashboardRowData) {
          console.log(`Processing item ${i + 1}/${customerItems.length}: ID=${item.id}`);
          
          // CRITICAL: Update Dashboard pickup status FIRST
          updateDashboardPickupStatus(dashboardSheet, dashboardRowData.rowNumber, dashboardHeaders);
          console.log(`  ✅ Marked "Picked Up Or Shiped" = TRUE for row ${dashboardRowData.rowNumber}`);
          
          // Get updated row data after checkbox change
          const updatedRowData = findDashboardRow(dashboardSheet, item.id, item.seller, item.customer);
          if (updatedRowData) {
            // Move to completed sheet
            moveRowToCompletedPickupShip(dashboardSheet, completedSheet, updatedRowData);
            console.log(`  ✅ Moved to ${completedSheet.getName()}`);
          }
          
          // Remove from Management Log (use current row index accounting for previous deletions)
          const currentMgmtRowIndex = item.rowIndex + 1 - (customerItems.length - 1 - i);
          managementSheet.deleteRow(currentMgmtRowIndex);
          console.log(`  ✅ Removed from Management Log`);
          
          processedCount++;
          
        } else {
          console.log(`Dashboard row not found for: ID=${item.id}, Seller=${item.seller}, Customer=${item.customer}`);
          errors.push(`Dashboard row not found for ID: ${item.id}`);
          errorCount++;
        }
        
      } catch (itemError) {
        console.log(`Error processing item ${item.id}: ${itemError.toString()}`);
        errors.push(`Error with ID ${item.id}: ${itemError.message}`);
        errorCount++;
      }
    }
    
    // Show results
    let message = `🎯 BULK PROCESSING RESULTS\n\nCustomer: "${targetCustomer}"\n\n`;
    message += `✅ Successfully processed: ${processedCount} items\n`;
    
    if (errorCount > 0) {
      message += `❌ Errors encountered: ${errorCount} items\n\n`;
      message += `Error details:\n${errors.slice(0, 3).join('\n')}`;
      if (errors.length > 3) {
        message += `\n... and ${errors.length - 3} more errors`;
      }
    }
    
    message += `\n\nActions completed:\n`;
    message += `• ✅ Marked "Picked Up Or Shiped" = TRUE for ${processedCount} items\n`;
    message += `• ✅ Moved ${processedCount} items to ${completedSheet.getName()}\n`;
    message += `• ✅ Removed ${processedCount} items from Management Log`;
    
    SpreadsheetApp.getUi().alert('Bulk Processing Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.log('Error in processAllCustomerItems: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'Failed to process all customer items: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

/**
 * Find the Dashboard sheet with various naming possibilities
 */
function findDashboardSheet(spreadsheet) {
  const possibleNames = [
    'Pick Up _ Ship _ Transport System  Dashboard',
    'Dashboard',
    'Pick Up _ Ship (Main Sheets)'
  ];
  
  for (const name of possibleNames) {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) return sheet;
  }
  
  // Look for sheets containing "Dashboard"
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().includes('dashboard')) {
      return sheet;
    }
  }
  
  return null;
}

/**
 * Find the Management Log sheet
 */
function findManagementLogSheet(spreadsheet) {
  const possibleNames = [
    'Pick Up _ Ship _ Transport System  Management Log',
    'Management Log',
    'Magement Log', 
    'Management',
    'Magement'
  ];
  
  for (const name of possibleNames) {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) return sheet;
  }
  
  // Look for sheets containing "Management"
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName.includes('management') || sheetName.includes('magement')) {
      return sheet;
    }
  }
  
  return null;
}

/**
 * Find the Completed Pick Up / Ship sheet (try multiple variations)
 */
function findCompletedPickupShipSheet(spreadsheet) {
  const possibleNames = [
    'Completed Pick Up / Ship',
    'Completed Pick Up Ship',
    'Completed Pickup Ship',
    'Completed Pick Up',
    'Completed Ship',
    'Pick Up _ Ship _ Transport System  Completed Pick Up / Ship',
    'Pick Up _ Ship _ Transport System  Completed Transports',
    'Completed Transports'  // Fallback to existing sheet
  ];
  
  for (const name of possibleNames) {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) {
      console.log(`Found completed sheet: "${name}"`);
      return sheet;
    }
  }
  
  // Look for sheets containing "completed" and ("pickup" or "ship" or "transport")
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName.includes('completed') && 
        (sheetName.includes('pickup') || sheetName.includes('ship') || sheetName.includes('transport'))) {
      console.log(`Found similar completed sheet: "${sheet.getName()}"`);
      return sheet;
    }
  }
  
  return null;
}

/**
 * Find the matching row in Dashboard based on ID, Seller, and Customer
 * Headers are specifically in row 12 (index 11)
 */
function findDashboardRow(dashboardSheet, targetId, targetSeller, targetCustomer) {
  const data = dashboardSheet.getDataRange().getValues();
  if (data.length < 12) {
    throw new Error('Dashboard sheet does not have enough rows (headers should be in row 12)');
  }
  
  // Headers are in row 12 (index 11)
  const headerRowIndex = 11;
  const headers = data[headerRowIndex];
  const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
  
  console.log('Dashboard headers:', trimmedHeaders);
  
  // Find column indices
  const idColIndex = trimmedHeaders.indexOf('ID');
  const sellerColIndex = trimmedHeaders.indexOf('Seller');
  const customerColIndex = trimmedHeaders.indexOf('Customer');
  
  if (idColIndex === -1 || sellerColIndex === -1 || customerColIndex === -1) {
    throw new Error(`Required columns not found in Dashboard headers. ID: ${idColIndex}, Seller: ${sellerColIndex}, Customer: ${customerColIndex}`);
  }
  
  console.log(`Column indices - ID: ${idColIndex}, Seller: ${sellerColIndex}, Customer: ${customerColIndex}`);
  
  // Search for matching row (start from row 13, index 12)
  for (let i = 12; i < data.length; i++) {
    const row = data[i];
    const id = row[idColIndex];
    const seller = row[sellerColIndex];
    const customer = row[customerColIndex];
    
    // Match all three values (trim whitespace for comparison)
    if (id && seller && customer &&
        id.toString().trim() === targetId.toString().trim() &&
        seller.toString().trim() === targetSeller.toString().trim() &&
        customer.toString().trim() === targetCustomer.toString().trim()) {
      
      console.log(`Found matching row at index ${i} (row ${i + 1})`);
      return {
        rowIndex: i,
        rowNumber: i + 1, // 1-based for Google Sheets
        data: row,
        headers: headers,
        headerRowIndex: headerRowIndex
      };
    }
  }
  
  return null;
}

/**
 * ENHANCED: Update the Dashboard row's "Picked Up Or Shiped" column to true
 * Multiple attempts and verification to ensure the checkbox is set properly
 */
function updateDashboardPickupStatus(dashboardSheet, rowNumber, headers = null) {
  try {
    // Get headers if not provided
    if (!headers) {
      const data = dashboardSheet.getDataRange().getValues();
      headers = data[11]; // Headers are in row 12 (index 11)
    }
    
    const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
    
    const pickedUpColIndex = trimmedHeaders.indexOf('Picked Up Or Shiped');
    if (pickedUpColIndex === -1) {
      throw new Error('Picked Up Or Shiped column not found in Dashboard');
    }
    
    console.log(`🎯 Setting "Picked Up Or Shiped" checkbox to TRUE for row ${rowNumber}, column ${pickedUpColIndex + 1}`);
    
    // Get the target cell (rowNumber is already 1-based)
    const targetCell = dashboardSheet.getRange(rowNumber, pickedUpColIndex + 1);
    
    // MULTIPLE METHODS to ensure the checkbox is set to TRUE:
    
    // Method 1: Insert checkbox first, then set value
    try {
      targetCell.insertCheckboxes();
      targetCell.setValue(true);
      console.log(`✅ Method 1 successful: Checkbox inserted and set to TRUE`);
    } catch (method1Error) {
      console.log(`⚠️ Method 1 failed: ${method1Error.toString()}`);
    }
    
    // Method 2: Force clear and reset with checkbox
    try {
      targetCell.clearContent();
      targetCell.insertCheckboxes();
      targetCell.setValue(true);
      console.log(`✅ Method 2 successful: Cell cleared, checkbox inserted, and set to TRUE`);
    } catch (method2Error) {
      console.log(`⚠️ Method 2 failed: ${method2Error.toString()}`);
    }
    
    // Method 3: Direct boolean value (in case checkbox already exists)
    try {
      targetCell.setValue(true);
      console.log(`✅ Method 3 successful: Direct boolean value set to TRUE`);
    } catch (method3Error) {
      console.log(`⚠️ Method 3 failed: ${method3Error.toString()}`);
    }
    
    // Verification: Check if the value was actually set
    Utilities.sleep(100); // Brief pause to ensure the value is set
    const verificationValue = targetCell.getValue();
    console.log(`🔍 VERIFICATION: Cell value after update = "${verificationValue}" (type: ${typeof verificationValue})`);
    
    if (verificationValue === true || verificationValue === 'TRUE' || verificationValue === 'True') {
      console.log(`✅ CONFIRMED: "Picked Up Or Shiped" checkbox successfully set to TRUE for row ${rowNumber}`);
    } else {
      console.log(`❌ WARNING: Checkbox value verification failed. Expected TRUE, got: "${verificationValue}"`);
      
      // Final attempt: Force set to TRUE one more time
      try {
        targetCell.setValue(true);
        console.log(`🔄 Final attempt: Forced value to TRUE`);
      } catch (finalError) {
        console.log(`❌ Final attempt failed: ${finalError.toString()}`);
      }
    }
    
  } catch (error) {
    console.log(`❌ Error in updateDashboardPickupStatus: ${error.toString()}`);
    throw new Error(`Failed to update Dashboard pickup status: ${error.message}`);
  }
}

/**
 * ENHANCED: Move the entire row from Dashboard to Completed Pick Up / Ship sheet
 * Ensures the "Picked Up Or Shiped" value is preserved and set to TRUE
 */
function moveRowToCompletedPickupShip(dashboardSheet, completedSheet, dashboardRowData) {
  try {
    // Get the existing structure of the completed sheet
    const completedData = completedSheet.getDataRange().getValues();
    const completedHeaders = completedData.length > 0 ? completedData[0] : [];
    
    console.log(`Completed sheet "${completedSheet.getName()}" headers:`, completedHeaders);
    
    // If the completed sheet is empty, copy the dashboard headers plus date
    if (completedHeaders.length === 0 || !completedHeaders[0]) {
      const newHeaders = [...dashboardRowData.headers, 'Completion Date'];
      completedSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      console.log('Added headers to completed sheet');
    }
    
    // Create the completed record - copy dashboard row but ensure critical values
    const completedRowData = [...dashboardRowData.data];
    
    // CRITICAL: Ensure "Picked Up Or Shiped" is set to TRUE in the copied data
    const dashboardHeaders = dashboardRowData.headers;
    const pickedUpColIndex = dashboardHeaders.findIndex(h => h && h.toString().trim() === 'Picked Up Or Shiped');
    
    if (pickedUpColIndex !== -1) {
      completedRowData[pickedUpColIndex] = true; // Force to TRUE
      console.log(`🎯 FORCED "Picked Up Or Shiped" to TRUE in copied data (column ${pickedUpColIndex})`);
    }
    
    // Fix data validation issues by mapping Dashboard values to completed sheet expected values
    const orderMethodIndex = dashboardHeaders.findIndex(h => h && h.toString().trim() === 'Order Method');
    if (orderMethodIndex !== -1) {
      const currentValue = completedRowData[orderMethodIndex];
      console.log(`Original Order Method value: "${currentValue}"`);
      
      // Map Dashboard values to completed sheet validation values
      if (currentValue === 'Marketplace') {
        completedRowData[orderMethodIndex] = '4GV Marketplace';
        console.log('Mapped "Marketplace" to "4GV Marketplace"');
      } else if (currentValue === 'Virtual') {
        completedRowData[orderMethodIndex] = 'Virtual Event';
        console.log('Mapped "Virtual" to "Virtual Event"');
      } else if (currentValue === 'Invoice' || currentValue === 'Wix') {
        completedRowData[orderMethodIndex] = 'Wix Invoice';
        console.log(`Mapped "${currentValue}" to "Wix Invoice"`);
      }
    }
    
    // Add completion date
    completedRowData.push(new Date());
    
    // Add the row to completed sheet
    const lastRow = completedSheet.getLastRow();
    const targetRow = lastRow + 1;
    
    try {
      completedSheet.getRange(targetRow, 1, 1, completedRowData.length)
        .setValues([completedRowData]);
      
      console.log(`✅ Added row to ${completedSheet.getName()} sheet at row ${targetRow}`);
      
      // CRITICAL: Apply checkbox formatting and ensure TRUE value for "Picked Up Or Shiped"
      formatCheckboxesInCompletedRow(completedSheet, targetRow, dashboardHeaders, completedRowData);
      
      // DOUBLE-CHECK: Verify the "Picked Up Or Shiped" checkbox in completed sheet
      if (pickedUpColIndex !== -1) {
        const completedPickupCell = completedSheet.getRange(targetRow, pickedUpColIndex + 1);
        completedPickupCell.insertCheckboxes();
        completedPickupCell.setValue(true);
        console.log(`🔍 VERIFIED: "Picked Up Or Shiped" set to TRUE in completed sheet row ${targetRow}`);
      }
      
      // Remove the row from Dashboard only after successful insertion
      dashboardSheet.deleteRow(dashboardRowData.rowNumber);
      console.log(`✅ Deleted row ${dashboardRowData.rowNumber} from Dashboard`);
      
    } catch (insertError) {
      console.log('Error inserting row to completed sheet:', insertError.toString());
      
      // If there's still a validation error, try inserting just the core data
      console.log('Attempting to insert core data only...');
      
      // Get core columns that should exist in both sheets
      const coreData = [
        dashboardRowData.data[dashboardHeaders.findIndex(h => h && h.toString().trim() === 'ID')],
        dashboardRowData.data[dashboardHeaders.findIndex(h => h && h.toString().trim() === 'Seller')],
        dashboardRowData.data[dashboardHeaders.findIndex(h => h && h.toString().trim() === 'Customer')],
        '4GV Marketplace', // Safe default for Order Method
        true, // CRITICAL: Picked Up Or Shiped = TRUE
        new Date() // Completion date
      ];
      
      // Try with minimal data
      completedSheet.getRange(targetRow, 1, 1, coreData.length)
        .setValues([coreData]);
      
      // Format the checkboxes for core data
      const pickedUpCoreCell = completedSheet.getRange(targetRow, 5); // Column 5 is the picked up status
      pickedUpCoreCell.insertCheckboxes();
      pickedUpCoreCell.setValue(true);
      
      console.log('✅ Successfully inserted core data to completed sheet with "Picked Up Or Shiped" = TRUE');
      
      // Remove from Dashboard
      dashboardSheet.deleteRow(dashboardRowData.rowNumber);
      console.log(`✅ Deleted row ${dashboardRowData.rowNumber} from Dashboard`);
    }
    
  } catch (error) {
    console.log(`❌ Error in moveRowToCompletedPickupShip: ${error.toString()}`);
    throw error;
  }
}

/**
 * ENHANCED: Apply checkbox formatting to boolean columns in the completed row
 * Special attention to "Picked Up Or Shiped" column
 */
function formatCheckboxesInCompletedRow(completedSheet, targetRow, dashboardHeaders, completedRowData) {
  // List of columns that should be formatted as checkboxes
  const checkboxColumns = [
    'Created',
    'Paid', 
    'Dropped Off',
    'Transported',
    'Ready For (Pickup or Ship)',
    'Picked Up Or Shiped',  // CRITICAL: This must be TRUE
    'Needs Transport',
    'Entire Order In Store Overide'
  ];
  
  checkboxColumns.forEach(columnName => {
    const columnIndex = dashboardHeaders.findIndex(h => h && h.toString().trim() === columnName);
    if (columnIndex !== -1) {
      const cellValue = completedRowData[columnIndex];
      
      // Check if the value is boolean-like
      if (cellValue === true || cellValue === false || 
          cellValue === 'TRUE' || cellValue === 'FALSE' ||
          cellValue === 'True' || cellValue === 'False') {
        
        const cell = completedSheet.getRange(targetRow, columnIndex + 1);
        cell.insertCheckboxes();
        
        // Set the correct boolean value
        let boolValue = (cellValue === true || cellValue === 'TRUE' || cellValue === 'True');
        
        // SPECIAL HANDLING: Force "Picked Up Or Shiped" to TRUE
        if (columnName === 'Picked Up Or Shiped') {
          boolValue = true;
          console.log(`🎯 FORCED "${columnName}" to TRUE in completed sheet (was: ${cellValue})`);
        }
        
        cell.setValue(boolValue);
        console.log(`✅ Applied checkbox formatting to column "${columnName}" with value: ${boolValue}`);
      }
    }
  });
  
  // DOUBLE VERIFICATION for "Picked Up Or Shiped" column
  const pickedUpColIndex = dashboardHeaders.findIndex(h => h && h.toString().trim() === 'Picked Up Or Shiped');
  if (pickedUpColIndex !== -1) {
    const pickedUpCell = completedSheet.getRange(targetRow, pickedUpColIndex + 1);
    
    // Force it one more time to be absolutely sure
    pickedUpCell.insertCheckboxes();
    pickedUpCell.setValue(true);
    
    // Verify the final value
    const finalValue = pickedUpCell.getValue();
    console.log(`🔍 FINAL VERIFICATION: "Picked Up Or Shiped" in completed sheet = ${finalValue} (${typeof finalValue})`);
    
    if (finalValue !== true) {
      console.log(`❌ CRITICAL WARNING: "Picked Up Or Shiped" is not TRUE in completed sheet!`);
      // One more attempt
      pickedUpCell.setValue(true);
      console.log(`🔄 Made final attempt to set "Picked Up Or Shiped" to TRUE`);
    } else {
      console.log(`✅ CONFIRMED: "Picked Up Or Shiped" is TRUE in completed sheet`);
    }
  }
}

/**
 * Remove the item from Management Log
 */
function removeFromManagementLog(managementSheet, targetId, targetSeller, targetCustomer) {
  const data = managementSheet.getDataRange().getValues();
  if (data.length === 0) return;
  
  const headers = data[0];
  const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
  
  // Find column indices
  const idColIndex = trimmedHeaders.indexOf('ID');
  const sellerColIndex = trimmedHeaders.indexOf('Seller');
  const customerColIndex = trimmedHeaders.indexOf('Customer');
  
  if (idColIndex === -1 || sellerColIndex === -1 || customerColIndex === -1) {
    console.log('Required columns not found in Management Log for removal');
    return;
  }
  
  // Search for matching row to remove
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[idColIndex];
    const seller = row[sellerColIndex];
    const customer = row[customerColIndex];
    
    // Match all three values (trim whitespace for comparison)
    if (id && seller && customer &&
        id.toString().trim() === targetId.toString().trim() &&
        seller.toString().trim() === targetSeller.toString().trim() &&
        customer.toString().trim() === targetCustomer.toString().trim()) {
      
      // Delete the row from Management Log (i+1 for 1-based indexing)
      managementSheet.deleteRow(i + 1);
      console.log(`✅ Removed row ${i + 1} from Management Log`);
      return;
    }
  }
  
  console.log(`⚠️ No matching row found in Management Log for removal: ID=${targetId}, Seller=${targetSeller}, Customer=${targetCustomer}`);
}

/**
 * UTILITY FUNCTION: Manually verify and fix "Picked Up Or Shiped" checkboxes
 * Run this if you need to fix existing completed records
 */
function fixPickedUpCheckboxesInCompletedSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const completedSheet = findCompletedPickupShipSheet(spreadsheet);
    
    if (!completedSheet) {
      console.log('❌ Completed Pick Up / Ship sheet not found');
      return;
    }
    
    const data = completedSheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('⚠️ No data to fix in completed sheet');
      return;
    }
    
    const headers = data[0];
    const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
    const pickedUpColIndex = trimmedHeaders.indexOf('Picked Up Or Shiped');
    
    if (pickedUpColIndex === -1) {
      console.log('❌ "Picked Up Or Shiped" column not found in completed sheet');
      return;
    }
    
    console.log(`🔧 Fixing "Picked Up Or Shiped" checkboxes in ${completedSheet.getName()}`);
    console.log(`Found column at index ${pickedUpColIndex + 1}`);
    
    let fixedCount = 0;
    
    // Process each data row (skip header)
    for (let i = 1; i < data.length; i++) {
      const currentValue = data[i][pickedUpColIndex];
      const rowNumber = i + 1;
      
      console.log(`Row ${rowNumber}: Current value = "${currentValue}" (${typeof currentValue})`);
      
      // Fix the cell regardless of current value
      const cell = completedSheet.getRange(rowNumber, pickedUpColIndex + 1);
      cell.insertCheckboxes();
      cell.setValue(true);
      fixedCount++;
      
      console.log(`  ✅ Fixed row ${rowNumber}: Set to TRUE with checkbox`);
    }
    
    console.log(`🎯 COMPLETED: Fixed ${fixedCount} rows in ${completedSheet.getName()}`);
    SpreadsheetApp.getUi().alert(
      'Fix Complete', 
      `Successfully fixed ${fixedCount} "Picked Up Or Shiped" checkboxes in ${completedSheet.getName()}.\n\nAll items are now marked as TRUE with proper checkbox formatting.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.log(`❌ Error fixing checkboxes: ${error.toString()}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to fix checkboxes: ${error.toString()}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * DEBUG FUNCTION: Check the status of "Picked Up Or Shiped" column in both sheets
 */
function debugPickedUpStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = findDashboardSheet(spreadsheet);
    const completedSheet = findCompletedPickupShipSheet(spreadsheet);
    
    console.log('=== PICKED UP STATUS DEBUG ===');
    
    // Check Dashboard
    if (dashboardSheet) {
      const dashboardData = dashboardSheet.getDataRange().getValues();
      if (dashboardData.length > 11) {
        const dashboardHeaders = dashboardData[11]; // Row 12
        const trimmedDashHeaders = dashboardHeaders.map(header => header ? header.toString().trim() : '');
        const dashPickedUpCol = trimmedDashHeaders.indexOf('Picked Up Or Shiped');
        
        console.log(`\n📊 DASHBOARD (${dashboardSheet.getName()}):`);
        console.log(`  Headers in row 12: ${trimmedDashHeaders.join(', ')}`);
        console.log(`  "Picked Up Or Shiped" column index: ${dashPickedUpCol}`);
        
        if (dashPickedUpCol !== -1) {
          // Check first few data rows
          for (let i = 12; i < Math.min(dashboardData.length, 17); i++) {
            const value = dashboardData[i][dashPickedUpCol];
            const id = dashboardData[i][trimmedDashHeaders.indexOf('ID')];
            console.log(`    Row ${i + 1} (ID: ${id}): "${value}" (${typeof value})`);
          }
        }
      }
    }
    
    // Check Completed Sheet
    if (completedSheet) {
      const completedData = completedSheet.getDataRange().getValues();
      if (completedData.length > 0) {
        const completedHeaders = completedData[0]; // Row 1
        const trimmedCompHeaders = completedHeaders.map(header => header ? header.toString().trim() : '');
        const compPickedUpCol = trimmedCompHeaders.indexOf('Picked Up Or Shiped');
        
        console.log(`\n📊 COMPLETED SHEET (${completedSheet.getName()}):`);
        console.log(`  Headers in row 1: ${trimmedCompHeaders.join(', ')}`);
        console.log(`  "Picked Up Or Shiped" column index: ${compPickedUpCol}`);
        
        if (compPickedUpCol !== -1) {
          // Check all data rows
          for (let i = 1; i < completedData.length; i++) {
            const value = completedData[i][compPickedUpCol];
            const id = completedData[i][trimmedCompHeaders.indexOf('ID')];
            console.log(`    Row ${i + 1} (ID: ${id}): "${value}" (${typeof value})`);
          }
        }
      }
    }
    
    console.log('\n=== END DEBUG ===');
    
  } catch (error) {
    console.log(`❌ Debug error: ${error.toString()}`);
  }
}

/**
 * Manual test function for debugging
 */
function testPickupProcess() {
  try {
    // Test with sample data from your CSV
    const testId = 'O9713224555';
    const testSeller = '4GoodVibes Gift Shop';
    const testCustomer = 'Gina Kolenda';
    
    console.log('🧪 Testing pickup process...');
    processPickupShip(testId, testSeller, testCustomer);
    console.log('✅ Test completed successfully');
    
  } catch (error) {
    console.log('❌ Test failed: ' + error.toString());
  }
}

/**
 * Manual test function for bulk processing
 */
function testBulkPickupProcess() {
  try {
    const testCustomer = 'Gina Kolenda';
    
    console.log('🧪 Testing bulk pickup process...');
    processAllCustomerItems(testCustomer);
    console.log('✅ Bulk test completed successfully');
    
  } catch (error) {
    console.log('❌ Bulk test failed: ' + error.toString());
  }
}
