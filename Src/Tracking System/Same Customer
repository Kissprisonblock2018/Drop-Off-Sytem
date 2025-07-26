function mergeDataAndUpdateReadyStatus() {
  // Get the spreadsheet and sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = spreadsheet.getSheetByName('Dashboard');
  const sellerDropoffSheet = spreadsheet.getSheetByName('Seller Drop Off'); // Adjust sheet name as needed
  const managementLogSheet = spreadsheet.getSheetByName('Management Log') || 
                            spreadsheet.getSheetByName('Magement Log') ||
                            spreadsheet.getSheetByName('Management') ||
                            spreadsheet.getSheetByName('Magement');
  
  if (!dashboardSheet) {
    throw new Error('Dashboard sheet not found');
  }
  
  // STEP 1: Update Management Log with completed transported items from Dashboard
  if (managementLogSheet) {
    console.log('Updating Management Log with completed transported items...');
    updateManagementLogFromDashboard(dashboardSheet, managementLogSheet);
  } else {
    console.log('Management Log sheet not found - skipping auto-population');
  }
  
  // Get all data from Dashboard sheet starting from the correct position
  // Based on CSV analysis, headers are in row 12, starting from column A (not column L)
  const startColumn = 1; // Column A (1st column) 
  const startRow = 12; // Row 12 where headers are located
  const lastColumn = dashboardSheet.getLastColumn();
  const lastRow = dashboardSheet.getLastRow();
  
  // Get the range starting from column A, row 12
  const dataRange = dashboardSheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn);
  const dashboardData = dataRange.getValues();
  const dashboardHeaders = dashboardData[0];
  
  // Find column indices - trim headers to handle spaces
  const trimmedHeaders = dashboardHeaders.map(header => header ? header.toString().trim() : '');
  
  const idColIndex = trimmedHeaders.indexOf('ID');
  const customerColIndex = trimmedHeaders.indexOf('Customer');
  const transportedColIndex = trimmedHeaders.indexOf('Transported');
  const sameCustomerOrdersColIndex = trimmedHeaders.indexOf('Same Customer Orders');
  const readyForPickupShipColIndex = trimmedHeaders.indexOf('Ready For (Pickup or Ship)');
  const sellerColIndex = trimmedHeaders.indexOf('Seller');
  
  console.log('Column indices found:');
  console.log('ID:', idColIndex);
  console.log('Customer:', customerColIndex);
  console.log('Transported:', transportedColIndex);
  console.log('Same Customer Orders:', sameCustomerOrdersColIndex);
  console.log('Ready For (Pickup or Ship):', readyForPickupShipColIndex);
  console.log('Seller:', sellerColIndex);
  
  if (customerColIndex === -1 || transportedColIndex === -1 || 
      sameCustomerOrdersColIndex === -1 || readyForPickupShipColIndex === -1 || 
      idColIndex === -1) {
    throw new Error('Required columns not found in Dashboard sheet');
  }
  
  // Get Management Log overrides if the sheet exists
  let managementOverrides = new Map();
  console.log('Looking for Management Log sheet...');
  
  if (managementLogSheet) {
    console.log('Management Log sheet found, processing overrides...');
    managementOverrides = getManagementLogOverrides(managementLogSheet);
    console.log('Management overrides loaded:', managementOverrides.size, 'entries');
  } else {
    console.log('Management Log sheet NOT found. Available sheets:');
    const allSheets = spreadsheet.getSheets();
    allSheets.forEach(sheet => console.log('  - ' + sheet.getName()));
  }
  
  // Merge with Seller Drop Off data if the sheet exists
  if (sellerDropoffSheet) {
    mergeSellersData(dashboardSheet, sellerDropoffSheet, sellerColIndex);
  }
  
  // Process each row starting from row 2 (skip header, which is now row 13 in the sheet)
  for (let i = 1; i < dashboardData.length; i++) {
    const currentCustomer = dashboardData[i][customerColIndex];
    const currentId = dashboardData[i][idColIndex];
    const currentSeller = dashboardData[i][sellerColIndex];
    
    if (!currentCustomer) continue; // Skip empty customer names
    
    // Count same customer orders and check if all are transported
    const customerOrdersInfo = analyzeCustomerOrders(dashboardData, currentCustomer, customerColIndex, transportedColIndex);
    
    // Update Same Customer Orders column (adjust for actual sheet position)
    const actualRow = startRow + i;
    const actualCol = startColumn + sameCustomerOrdersColIndex;
    dashboardSheet.getRange(actualRow, actualCol).setValue(customerOrdersInfo.count);
    
    // Check for Management Log override
    const overrideKey = `${currentId}|${currentSeller}|${currentCustomer}`;
    let readyStatus = customerOrdersInfo.allTransported;
    
    console.log(`Checking override for: "${overrideKey}"`);
    console.log(`  - Management overrides available: ${managementOverrides.size}`);
    
    if (managementOverrides.has(overrideKey)) {
      readyStatus = true; // Override to true if found in Management Log
      console.log(`  -> OVERRIDE APPLIED for: ${overrideKey}`);
    } else {
      console.log(`  -> No override found for: ${overrideKey}`);
      if (managementOverrides.size > 0) {
        console.log(`  -> Available override keys:`, Array.from(managementOverrides.keys()));
      }
    }
    
    // Update Ready For (Pickup or Ship) column (adjust for actual sheet position)
    const readyActualCol = startColumn + readyForPickupShipColIndex;
    dashboardSheet.getRange(actualRow, readyActualCol).setValue(readyStatus);
  }
  
  Logger.log('Dashboard updated successfully');
}

function analyzeCustomerOrders(dashboardData, customerName, customerColIndex, transportedColIndex) {
  let count = 0;
  let allTransported = true;
  
  // Go through all rows to find matching customers
  for (let i = 1; i < dashboardData.length; i++) {
    const rowCustomer = dashboardData[i][customerColIndex];
    
    if (rowCustomer === customerName) {
      count++;
      const isTransported = dashboardData[i][transportedColIndex];
      
      // Check if transported value is not TRUE
      if (isTransported !== true && isTransported !== 'TRUE' && isTransported !== 'True') {
        allTransported = false;
      }
    }
  }
  
  return {
    count: count,
    allTransported: allTransported
  };
}

function updateManagementLogFromDashboard(dashboardSheet, managementLogSheet) {
  // Get Dashboard data
  const startColumn = 1;
  const startRow = 12;
  const lastColumn = dashboardSheet.getLastColumn();
  const lastRow = dashboardSheet.getLastRow();
  
  const dataRange = dashboardSheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn);
  const dashboardData = dataRange.getValues();
  const dashboardHeaders = dashboardData[0];
  
  // Find Dashboard column indices
  const trimmedHeaders = dashboardHeaders.map(header => header ? header.toString().trim() : '');
  const idColIndex = trimmedHeaders.indexOf('ID');
  const sellerColIndex = trimmedHeaders.indexOf('Seller');
  const customerColIndex = trimmedHeaders.indexOf('Customer');
  const transportedColIndex = trimmedHeaders.indexOf('Transported');
  
  console.log('Dashboard columns found:');
  console.log('ID index:', idColIndex);
  console.log('Seller index:', sellerColIndex);
  console.log('Customer index:', customerColIndex);
  console.log('Transported index:', transportedColIndex);
  
  if (idColIndex === -1 || sellerColIndex === -1 || customerColIndex === -1 || transportedColIndex === -1) {
    console.log('Required Dashboard columns not found for Management Log update');
    return;
  }
  
  // Get Management Log data
  const managementData = managementLogSheet.getDataRange().getValues();
  const managementHeaders = managementData.length > 0 ? managementData[0] : [];
  const trimmedMgmtHeaders = managementHeaders.map(header => header ? header.toString().trim() : '');
  
  // Find Management Log column indices - looking for SEPARATE columns
  const mgmtIdColIndex = trimmedMgmtHeaders.indexOf('ID');
  const mgmtSellerColIndex = trimmedMgmtHeaders.indexOf('Seller');
  const mgmtCustomerColIndex = trimmedMgmtHeaders.indexOf('Customer');
  const mgmtPickupColIndex = trimmedMgmtHeaders.indexOf('Picked Up Or Shiped');
  const mgmtOverrideColIndex = trimmedMgmtHeaders.indexOf('Entire Order In Store Overide');
  
  console.log('Management Log structure (separate columns):');
  console.log('ID column index:', mgmtIdColIndex);
  console.log('Seller column index:', mgmtSellerColIndex);
  console.log('Customer column index:', mgmtCustomerColIndex);
  console.log('Pickup column index:', mgmtPickupColIndex);
  console.log('Override column index:', mgmtOverrideColIndex);
  console.log('Management Log headers:', trimmedMgmtHeaders);
  
  if (mgmtIdColIndex === -1 || mgmtSellerColIndex === -1 || mgmtCustomerColIndex === -1) {
    console.log('ERROR: Required Management Log columns not found');
    console.log('Available headers:', trimmedMgmtHeaders);
    return;
  }
  
  // Create a set of existing Management Log entries for exact matching
  // Only consider rows that have actual ID, Seller, Customer data
  // Skip rows that are just formatting (empty ID/Seller/Customer)
  const existingEntries = new Set();
  const emptyRowIndices = []; // Track empty rows for filling
  
  for (let i = 1; i < managementData.length; i++) {
    const id = managementData[i][mgmtIdColIndex];
    const seller = managementData[i][mgmtSellerColIndex];
    const customer = managementData[i][mgmtCustomerColIndex];
    const pickedUp = managementData[i][mgmtPickupColIndex];
    const override = managementData[i][mgmtOverrideColIndex];
    
    // Check if this row has actual data (not just formatting)
    if (id && seller && customer) {
      const matchKey = `${id}|${seller}|${customer}`;
      existingEntries.add(matchKey);
      console.log(`Existing Management Log entry: ${matchKey}`);
    } 
    // Check if this is an empty row available for filling
    else if (!id && !seller && !customer) {
      emptyRowIndices.push(i + 1); // Store 1-based row number for Google Sheets
      console.log(`Empty row found at index ${i + 1} (available for filling)`);
    }
    // Log partially filled rows for debugging
    else {
      console.log(`Partially filled row ${i + 1}: ID="${id}", Seller="${seller}", Customer="${customer}"`);
    }
  }
  
  console.log('Total existing Management Log entries:', existingEntries.size);
  console.log('Empty rows available for filling:', emptyRowIndices.length);
  
  // Find all transported items from Dashboard and match them exactly
  const newEntries = [];
  let emptyRowIndex = 0; // Track which empty row to fill next
  
  for (let i = 1; i < dashboardData.length; i++) {
    const id = dashboardData[i][idColIndex];
    const seller = dashboardData[i][sellerColIndex];
    const customer = dashboardData[i][customerColIndex];
    const transported = dashboardData[i][transportedColIndex];
    
    // Skip if missing data
    if (!id || !seller || !customer) {
      console.log(`Skipping row ${i + 1}: missing data - ID: "${id}", Seller: "${seller}", Customer: "${customer}"`);
      continue;
    }
    
    // Only process if transported = TRUE
    if (transported === true || transported === 'TRUE' || transported === 'True') {
      // Create exact match key for comparison
      const matchKey = `${id}|${seller}|${customer}`;
      
      // Check if this exact combination already exists in Management Log
      if (!existingEntries.has(matchKey)) {
        console.log(`Adding transported item to Management Log:`);
        console.log(`  ID: "${id}"`);
        console.log(`  Seller: "${seller}"`);
        console.log(`  Customer: "${customer}"`);
        console.log(`  Match key: "${matchKey}"`);
        
        // Determine where to place this entry
        let targetRow;
        if (emptyRowIndex < emptyRowIndices.length) {
          // Use an existing empty row
          targetRow = emptyRowIndices[emptyRowIndex];
          console.log(`  -> Using existing empty row ${targetRow}`);
          emptyRowIndex++;
        } else {
          // Add to the end
          targetRow = managementLogSheet.getLastRow() + 1 + (newEntries.length - emptyRowIndex);
          console.log(`  -> Adding new row ${targetRow}`);
        }
        
        newEntries.push({
          row: targetRow,
          data: [id, seller, customer, false, false] // ID, Seller, Customer, Picked Up, Override
        });
        
        // Add to existing entries to prevent duplicates within this run
        existingEntries.add(matchKey);
      } else {
        console.log(`Entry already exists in Management Log for: ${matchKey}`);
      }
    } else {
      console.log(`Item not transported, skipping: ID="${id}", Seller="${seller}", Customer="${customer}", Transported="${transported}"`);
    }
  }
  
  // Add new entries to Management Log
  if (newEntries.length > 0) {
    console.log(`Adding ${newEntries.length} new entries to Management Log`);
    
    // If Management Log is empty, add headers first
    if (managementData.length === 0) {
      const headers = ['ID', 'Seller', 'Customer', 'Picked Up Or Shiped', 'Entire Order In Store Overide'];
      managementLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      console.log('Added headers to empty Management Log');
    }
    
    // Add entries to their specific rows WITH CHECKBOXES
    newEntries.forEach((entry, index) => {
      // Add the basic data (ID, Seller, Customer)
      managementLogSheet.getRange(entry.row, 1, 1, 3).setValues([[entry.data[0], entry.data[1], entry.data[2]]]);
      
      // Set checkbox fields for columns 4 and 5 (Picked Up Or Shiped, Entire Order In Store Overide)
      const pickupCell = managementLogSheet.getRange(entry.row, 4);
      const overrideCell = managementLogSheet.getRange(entry.row, 5);
      
      // Insert checkboxes and set values
      pickupCell.insertCheckboxes();
      pickupCell.setValue(entry.data[3]); // false for Picked Up
      
      overrideCell.insertCheckboxes();
      overrideCell.setValue(entry.data[4]); // false for Override
      
      console.log(`  Row ${entry.row}: ID="${entry.data[0]}" | Seller="${entry.data[1]}" | Customer="${entry.data[2]}" | Pickup: ☐ ${entry.data[3]} | Override: ☐ ${entry.data[4]}`);
    });
    
    console.log(`Successfully added ${newEntries.length} entries to Management Log with checkboxes`);
  } else {
    console.log('No new transported entries needed for Management Log');
  }
}

function getManagementLogOverrides(managementLogSheet) {
  // Get all data from Management Log sheet starting from row 1
  const managementData = managementLogSheet.getDataRange().getValues();
  
  console.log('Management Log total rows:', managementData.length);
  
  if (managementData.length === 0) {
    console.log('Management Log sheet is empty');
    return new Map();
  }
  
  const managementHeaders = managementData[0];
  console.log('Raw Management Log headers:', managementHeaders);
  
  const trimmedMgmtHeaders = managementHeaders.map(header => header ? header.toString().trim() : '');
  console.log('Trimmed Management Log headers:', trimmedMgmtHeaders);
  
  // Look for separate ID, Seller, Customer columns (the correct format)
  const idColIndex = trimmedMgmtHeaders.indexOf('ID');
  const sellerColIndex = trimmedMgmtHeaders.indexOf('Seller');
  const customerColIndex = trimmedMgmtHeaders.indexOf('Customer');
  let overrideColIndex = trimmedMgmtHeaders.indexOf('Entire Order In Store Overide');
  if (overrideColIndex === -1) {
    overrideColIndex = trimmedMgmtHeaders.indexOf('Entire Order In Store Override');
  }
  
  console.log('Management Log column indices (separate columns):');
  console.log('ID:', idColIndex);
  console.log('Seller:', sellerColIndex);
  console.log('Customer:', customerColIndex);
  console.log('Override column index:', overrideColIndex);
  
  if (idColIndex === -1 || sellerColIndex === -1 || customerColIndex === -1) {
    console.log('ERROR: Could not find required ID, Seller, or Customer columns');
    console.log('Available headers:', trimmedMgmtHeaders);
    return new Map();
  }
  
  if (overrideColIndex === -1) {
    console.log('ERROR: Could not find override column');
    console.log('Available headers:', trimmedMgmtHeaders);
    return new Map();
  }
  
  const overrideMap = new Map();
  
  // Process Management Log data (skip header row)
  console.log('Processing Management Log data rows...');
  for (let i = 1; i < managementData.length; i++) {
    const id = managementData[i][idColIndex];
    const seller = managementData[i][sellerColIndex];
    const customer = managementData[i][customerColIndex];
    const overrideValue = managementData[i][overrideColIndex];
    
    console.log(`Row ${i + 1}: ID="${id}", Seller="${seller}", Customer="${customer}", Override="${overrideValue}"`);
    
    // Skip rows with missing data
    if (!id || !seller || !customer) {
      console.log('  -> Skipping row with missing ID, Seller, or Customer data');
      continue;
    }
    
    // Check if override is marked as true
    if (overrideValue === true || overrideValue === 'TRUE' || overrideValue === 'True' || overrideValue === 'true') {
      console.log('  -> Override is TRUE, adding to map...');
      
      const key = `${id}|${seller}|${customer}`;
      overrideMap.set(key, true);
      console.log(`  -> Override added for: ${key}`);
    } else {
      console.log('  -> Override is not TRUE, skipping');
    }
  }
  
  console.log('Total overrides found:', overrideMap.size);
  console.log('Override keys:', Array.from(overrideMap.keys()));
  
  return overrideMap;
}

/**
 * NEW FUNCTION: Setup existing Management Log columns as checkboxes
 * Run this once to convert any existing TRUE/FALSE text to checkboxes
 */
function setupManagementLogCheckboxes() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const managementLogSheet = spreadsheet.getSheetByName('Management Log') || 
                              spreadsheet.getSheetByName('Magement Log') ||
                              spreadsheet.getSheetByName('Management') ||
                              spreadsheet.getSheetByName('Magement');
    
    if (!managementLogSheet) {
      console.log('Management Log sheet not found');
      return;
    }
    
    const data = managementLogSheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('Management Log sheet has no data rows');
      return;
    }
    
    const headers = data[0];
    const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
    
    const pickupColIndex = trimmedHeaders.indexOf('Picked Up Or Shiped');
    const overrideColIndex = trimmedHeaders.indexOf('Entire Order In Store Overide') !== -1 ? 
                            trimmedHeaders.indexOf('Entire Order In Store Overide') : 
                            trimmedHeaders.indexOf('Entire Order In Store Override');
    
    console.log('Setting up checkboxes for Management Log...');
    console.log('Pickup column index:', pickupColIndex);
    console.log('Override column index:', overrideColIndex);
    
    const lastRow = managementLogSheet.getLastRow();
    
    // Setup Picked Up Or Shiped column as checkboxes
    if (pickupColIndex !== -1 && lastRow > 1) {
      const pickupRange = managementLogSheet.getRange(2, pickupColIndex + 1, lastRow - 1, 1);
      
      // Get current values and convert to boolean
      const pickupValues = pickupRange.getValues().map(row => {
        const value = row[0];
        if (value === true || value === 'TRUE' || value === 'True' || value === 'true') {
          return [true];
        } else {
          return [false];
        }
      });
      
      // Insert checkboxes and set converted values
      pickupRange.insertCheckboxes();
      pickupRange.setValues(pickupValues);
      console.log('✅ Converted Picked Up Or Shiped column to checkboxes');
    }
    
    // Setup Override column as checkboxes
    if (overrideColIndex !== -1 && lastRow > 1) {
      const overrideRange = managementLogSheet.getRange(2, overrideColIndex + 1, lastRow - 1, 1);
      
      // Get current values and convert to boolean
      const overrideValues = overrideRange.getValues().map(row => {
        const value = row[0];
        if (value === true || value === 'TRUE' || value === 'True' || value === 'true') {
          return [true];
        } else {
          return [false];
        }
      });
      
      // Insert checkboxes and set converted values
      overrideRange.insertCheckboxes();
      overrideRange.setValues(overrideValues);
      console.log('✅ Converted Override column to checkboxes');
    }
    
    console.log('✅ Management Log checkbox setup completed');
    
  } catch (error) {
    console.log('❌ Error setting up Management Log checkboxes:', error.message);
  }
}

function mergeSellersData(dashboardSheet, sellerDropoffSheet, sellerColIndex) {
  // Get seller drop off data
  const sellerData = sellerDropoffSheet.getDataRange().getValues();
  const sellerHeaders = sellerData[0];
  
  // Find Seller ID column in both sheets
  const sellerIdColIndex = sellerHeaders.indexOf('Seller ID') || sellerHeaders.indexOf('ID');
  
  if (sellerIdColIndex === -1) {
    Logger.log('Seller ID column not found in Seller Drop Off sheet');
    return;
  }
  
  // Get dashboard data starting from row 12, column A
  const startColumn = 1;
  const startRow = 12;
  const lastColumn = dashboardSheet.getLastColumn();
  const lastRow = dashboardSheet.getLastRow();
  
  const dataRange = dashboardSheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn);
  const dashboardData = dataRange.getValues();
  
  // Create a map of seller data for quick lookup
  const sellerMap = new Map();
  for (let i = 1; i < sellerData.length; i++) {
    const sellerId = sellerData[i][sellerIdColIndex];
    if (sellerId) {
      sellerMap.set(sellerId, sellerData[i]);
    }
  }
  
  // Merge data based on Seller ID
  for (let i = 1; i < dashboardData.length; i++) {
    const sellerId = dashboardData[i][sellerColIndex];
    
    if (sellerId && sellerMap.has(sellerId)) {
      const sellerInfo = sellerMap.get(sellerId);
      
      // Here you can add logic to merge specific fields from seller drop off
      // For example, if there are additional fields to copy:
      // const actualRow = startRow + i;
      // const actualCol = startColumn + someColumnIndex;
      // dashboardSheet.getRange(actualRow, actualCol).setValue(sellerInfo[someIndex]);
      
      Logger.log(`Merged data for seller: ${sellerId}`);
    }
  }
}

// Function to run the merge and update process
function runDashboardUpdate() {
  try {
    mergeDataAndUpdateReadyStatus();
    SpreadsheetApp.getUi().alert('Dashboard updated successfully!');
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error updating dashboard: ' + error.toString());
  }
}



// Function to set up triggers or manual execution
function onEdit(e) {
  // Optional: Automatically run when certain cells are edited
  const range = e.range;
  const sheet = range.getSheet();
  
  if (sheet.getName() === 'Dashboard') {
    const column = range.getColumn();
    const row = range.getRow();
    
    // Only process if editing is in the data area (starting from row 12)
    if (row >= 12) {
      const startColumn = 1;
      const startRow = 12;
      const lastColumn = sheet.getLastColumn();
      
      const headerRange = sheet.getRange(startRow, startColumn, 1, lastColumn);
      const headers = headerRange.getValues()[0];
      const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
      
      // If Transported column was edited, update the ready status
      const editedColumnIndex = column - startColumn;
      if (trimmedHeaders[editedColumnIndex] === 'Transported') {
        // Run update for the specific customer
        updateSpecificCustomer(sheet, row);
      }
    }
  }
}

function updateSpecificCustomer(sheet, rowNumber) {
  const startColumn = 1;
  const startRow = 12;
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // Get data starting from the correct position
  const dataRange = sheet.getRange(startRow, startColumn, lastRow - startRow + 1, lastColumn);
  const data = dataRange.getValues();
  const headers = data[0];
  const trimmedHeaders = headers.map(header => header ? header.toString().trim() : '');
  
  const idColIndex = trimmedHeaders.indexOf('ID');
  const customerColIndex = trimmedHeaders.indexOf('Customer');
  const transportedColIndex = trimmedHeaders.indexOf('Transported');
  const sameCustomerOrdersColIndex = trimmedHeaders.indexOf('Same Customer Orders');
  const readyForPickupShipColIndex = trimmedHeaders.indexOf('Ready For (Pickup or Ship)');
  const sellerColIndex = trimmedHeaders.indexOf('Seller');
  
  // Get Management Log overrides
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const managementLogSheet = spreadsheet.getSheetByName('Management Log') || 
                            spreadsheet.getSheetByName('Magement Log') ||
                            spreadsheet.getSheetByName('Management') ||
                            spreadsheet.getSheetByName('Magement');
  let managementOverrides = new Map();
  if (managementLogSheet) {
    managementOverrides = getManagementLogOverrides(managementLogSheet);
  }
  
  // Get customer name from the edited row (adjust for array indexing)
  const dataRowIndex = rowNumber - startRow;
  const customerName = data[dataRowIndex][customerColIndex];
  
  if (customerName) {
    const customerOrdersInfo = analyzeCustomerOrders(data, customerName, customerColIndex, transportedColIndex);
    
    // Update all rows for this customer
    for (let i = 1; i < data.length; i++) {
      if (data[i][customerColIndex] === customerName) {
        const currentId = data[i][idColIndex];
        const currentSeller = data[i][sellerColIndex];
        const overrideKey = `${currentId}|${currentSeller}|${customerName}`;
        
        let readyStatus = customerOrdersInfo.allTransported;
        
        // Check for Management Log override
        if (managementOverrides.has(overrideKey)) {
          readyStatus = true;
        }
        
        const actualRow = startRow + i;
        sheet.getRange(actualRow, startColumn + sameCustomerOrdersColIndex).setValue(customerOrdersInfo.count);
        sheet.getRange(actualRow, startColumn + readyForPickupShipColIndex).setValue(readyStatus);
      }
    }
  }
}
