/**
 * Order Sync Script - Syncs marketplace orders from Dashboard to Pick Up Orders
 * Runs every 4 hours or can be triggered manually
 * FIXED: Updated pickup location mapping to use 'End Location' from Dashboard
 */

// Configuration
const CONFIG = {
  DASHBOARD_SHEET_URL: 'https://docs.google.com/spreadsheets/d/1MWCzWUDfjKr2mGlQcIgKF0uqFoRZsg8QUiGlGiCmFaM/',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  PICKUP_ORDERS_SHEET_NAME: 'Pick Up Orders',
  HEADER_ROW: 12, // Headers are on row 12, not row 1
  
  // Column mappings for Dashboard sheet (exact names from the sheet)
  DASHBOARD_COLUMNS: {
    ORDER_ID: 'ID',
    ORDER_METHOD: 'Order Method',
    DROPPED_OFF: 'Dropped Off',
    SELLER: 'Seller',
    BUYER: 'Customer',
    ORDERED_DATE: 'Date Created',
    PICKUP_LOCATION: 'End Location', // FIXED: Changed from 'Pickup Location ' to 'End Location'
    PICKED_UP_OR_SHIPPED: 'Picked Up Or Shiped', // Note: typo in original
    TRANSPORTED: 'Transported',
    STATUS: 'Status'
  },
  
  // Column mappings for Pick Up Orders sheet (matching existing sheet structure)
  PICKUP_COLUMNS: {
    ID: 'Order ID',  // Changed from 'ID' to 'Order ID'
    SELLER: 'Seller',
    CUSTOMER: 'Buyer',  // Changed from 'Customer' to 'Buyer'
    PICKUP_LOCATION: 'Pickup Location',
    LAST_SYNCED: 'Last Synced' // Optional: to track when record was added
  },
  
  // Filter criteria
  FILTER_ORDER_METHOD: 'Marketplace',
  FILTER_DROPPED_OFF: false // Changed from 'FALSE' to boolean false
};

/**
 * Main function to sync orders - can be called manually
 */
function syncMarketplaceOrders() {
  try {
    console.log('=== STARTING ORDER SYNC PROCESS ===');
    
    // Get dashboard data
    console.log('Step 1: Getting dashboard data...');
    const dashboardData = getDashboardData();
    console.log(`Dashboard data result: ${dashboardData ? dashboardData.length : 'null'} records`);
    
    if (!dashboardData || dashboardData.length === 0) {
      console.log('ERROR: No data found in Dashboard sheet');
      return;
    }
    
    // Show sample of dashboard data
    console.log('Sample dashboard data (first 2 records):');
    dashboardData.slice(0, 2).forEach((record, index) => {
      console.log(`Record ${index + 1}:`, JSON.stringify(record, null, 2));
    });
    
    // Filter orders based on criteria
    console.log('\nStep 2: Filtering orders...');
    const filteredOrders = filterOrders(dashboardData);
    console.log(`Filtering result: ${filteredOrders.length} orders match criteria`);
    
    if (filteredOrders.length === 0) {
      console.log('ERROR: No orders match filtering criteria');
      return;
    }
    
    console.log('Sample filtered orders (first 2):');
    filteredOrders.slice(0, 2).forEach((order, index) => {
      console.log(`Filtered order ${index + 1}:`, JSON.stringify(order, null, 2));
    });
    
    // Get existing pickup orders to avoid duplicates
    console.log('\nStep 3: Getting existing pickup orders...');
    const existingOrders = getExistingPickupOrders();
    console.log(`Existing orders: ${existingOrders.length}`);
    console.log('Existing order IDs:', existingOrders);
    
    // Find new orders to add
    console.log('\nStep 4: Finding new orders...');
    const newOrders = findNewOrders(filteredOrders, existingOrders);
    console.log(`New orders to add: ${newOrders.length}`);
    
    if (newOrders.length > 0) {
      console.log('Sample new orders (first 2):');
      newOrders.slice(0, 2).forEach((order, index) => {
        console.log(`New order ${index + 1}:`, JSON.stringify(order, null, 2));
      });
      
      // Add new orders to Pick Up Orders sheet
      console.log('\nStep 5: Adding orders to pickup sheet...');
      addOrdersToPickupSheet(newOrders);
      console.log(`SUCCESS: Added ${newOrders.length} orders to Pick Up Orders sheet`);
    } else {
      console.log('No new orders to add (all are duplicates)');
    }
    
    console.log('\n=== ORDER SYNC COMPLETED ===');
    
  } catch (error) {
    console.error('=== ERROR DURING ORDER SYNC ===');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    
    // Optional: Send email notification on error
    // sendErrorNotification(error);
  }
}

/**
 * Get data from the Dashboard sheet
 */
function getDashboardData() {
  try {
    console.log(`Attempting to open Dashboard sheet at: ${CONFIG.DASHBOARD_SHEET_URL}`);
    
    // Open the external spreadsheet
    const dashboardSpreadsheet = SpreadsheetApp.openByUrl(CONFIG.DASHBOARD_SHEET_URL);
    console.log('✓ Dashboard spreadsheet opened successfully');
    
    const dashboardSheet = dashboardSpreadsheet.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
    console.log(`✓ Found sheet: ${CONFIG.DASHBOARD_SHEET_NAME}`);
    
    if (!dashboardSheet) {
      throw new Error(`Sheet "${CONFIG.DASHBOARD_SHEET_NAME}" not found in Dashboard spreadsheet`);
    }
    
    // Get total data range
    const totalRows = dashboardSheet.getLastRow();
    const totalCols = dashboardSheet.getLastColumn();
    console.log(`Sheet dimensions: ${totalRows} rows x ${totalCols} columns`);
    
    if (totalRows < CONFIG.HEADER_ROW) {
      console.log(`ERROR: Not enough rows. Expected headers on row ${CONFIG.HEADER_ROW}, but sheet only has ${totalRows} rows`);
      return [];
    }
    
    // Get headers from the specified row
    console.log(`Reading headers from row ${CONFIG.HEADER_ROW}...`);
    const headers = dashboardSheet.getRange(CONFIG.HEADER_ROW, 1, 1, totalCols).getValues()[0];
    console.log('Headers found:', headers);
    
    // Verify required columns exist
    const requiredColumns = Object.values(CONFIG.DASHBOARD_COLUMNS);
    console.log('Required columns:', requiredColumns);
    
    const missingColumns = requiredColumns.filter(col => !headers.includes(col));
    if (missingColumns.length > 0) {
      console.error('MISSING COLUMNS:', missingColumns);
      throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
    }
    console.log('✓ All required columns found');
    
    // Get data starting from the row after headers
    if (totalRows <= CONFIG.HEADER_ROW) {
      console.log('No data rows found after headers');
      return [];
    }
    
    const dataRows = totalRows - CONFIG.HEADER_ROW;
    console.log(`Reading ${dataRows} data rows starting from row ${CONFIG.HEADER_ROW + 1}...`);
    
    const data = dashboardSheet.getRange(CONFIG.HEADER_ROW + 1, 1, dataRows, totalCols).getValues();
    console.log(`✓ Successfully read ${data.length} data rows`);
    
    // Convert to objects for easier handling
    const objects = data.map((row, index) => {
      const rowObj = {};
      headers.forEach((header, colIndex) => {
        rowObj[header] = row[colIndex];
      });
      
      // Log first few rows for debugging
      if (index < 3) {
        console.log(`Data row ${index + 1}:`, {
          'ID': rowObj[CONFIG.DASHBOARD_COLUMNS.ORDER_ID],
          'Order Method': rowObj[CONFIG.DASHBOARD_COLUMNS.ORDER_METHOD],
          'Dropped Off': rowObj[CONFIG.DASHBOARD_COLUMNS.DROPPED_OFF],
          'End Location': rowObj[CONFIG.DASHBOARD_COLUMNS.PICKUP_LOCATION] // Updated logging
        });
      }
      
      return rowObj;
    });
    
    console.log(`✓ Converted ${objects.length} rows to objects`);
    return objects;
    
  } catch (error) {
    console.error('ERROR in getDashboardData:', error.message);
    console.error('Error stack:', error.stack);
    throw error;
  }
}

/**
 * Filter orders based on criteria
 */
function filterOrders(orders) {
  console.log(`Filtering ${orders.length} orders from Dashboard...`);
  
  const filteredOrders = orders.filter(order => {
    const orderMethod = order[CONFIG.DASHBOARD_COLUMNS.ORDER_METHOD];
    const droppedOff = order[CONFIG.DASHBOARD_COLUMNS.DROPPED_OFF];
    const orderId = order[CONFIG.DASHBOARD_COLUMNS.ORDER_ID];
    
    // Skip orders without ID
    if (!orderId || orderId.toString().trim() === '') {
      console.log(`Skipping order with no ID`);
      return false;
    }
    
    // Check if Order Method is "Marketplace"
    const isMarketplace = orderMethod === CONFIG.FILTER_ORDER_METHOD;
    
    // Check if Dropped Off is false (boolean false, not string "FALSE")
    const isNotDroppedOff = droppedOff === CONFIG.FILTER_DROPPED_OFF;
    
    const shouldInclude = isMarketplace && isNotDroppedOff;
    
    if (shouldInclude) {
      console.log(`✓ Including Order ${orderId}: Method=${orderMethod}, DroppedOff=${droppedOff}`);
    } else {
      // Only log exclusions for first few to avoid spam
      const orderIndex = orders.indexOf(order);
      if (orderIndex < 5) {
        console.log(`✗ Excluding Order ${orderId}: Method=${orderMethod} (${typeof orderMethod}), DroppedOff=${droppedOff} (${typeof droppedOff})`);
        console.log(`  -> Marketplace check: ${isMarketplace}, NotDropped check: ${isNotDroppedOff}`);
      }
    }
    
    return shouldInclude;
  });
  
  console.log(`Filtered result: ${filteredOrders.length} orders match criteria`);
  return filteredOrders;
}

/**
 * Get existing orders from Pick Up Orders sheet
 */
function getExistingPickupOrders() {
  try {
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const pickupSheet = currentSpreadsheet.getSheetByName(CONFIG.PICKUP_ORDERS_SHEET_NAME);
    
    if (!pickupSheet) {
      console.log(`Creating new sheet: ${CONFIG.PICKUP_ORDERS_SHEET_NAME}`);
      return createPickupOrdersSheet();
    }
    
    const dataRange = pickupSheet.getDataRange();
    if (dataRange.getNumRows() === 0) {
      return [];
    }
    
    const data = dataRange.getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Find the ID column to check for existing orders
    const idColumnIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.ID);
    if (idColumnIndex === -1) {
      console.log('ID column not found, treating as empty sheet');
      return [];
    }
    
    // Extract existing Order IDs
    return rows.map(row => row[idColumnIndex]).filter(id => id && id.toString().trim() !== '');
    
  } catch (error) {
    console.error('Error getting existing pickup orders:', error);
    return [];
  }
}

/**
 * Find orders that don't already exist in pickup sheet
 */
function findNewOrders(filteredOrders, existingOrderIds) {
  return filteredOrders.filter(order => {
    const orderId = order[CONFIG.DASHBOARD_COLUMNS.ORDER_ID];
    return orderId && !existingOrderIds.includes(orderId.toString());
  });
}

/**
 * Add new orders to Pick Up Orders sheet
 */
function addOrdersToPickupSheet(newOrders) {
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let pickupSheet = currentSpreadsheet.getSheetByName(CONFIG.PICKUP_ORDERS_SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!pickupSheet) {
    pickupSheet = createPickupOrdersSheet();
  }
  
  // Get or create headers
  const headers = ensureHeaders(pickupSheet);
  
  // Prepare data rows
  const dataRows = newOrders.map(order => {
    const row = new Array(headers.length).fill('');
    
    // Map data to correct columns
    const sellerIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.SELLER);
    const idIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.ID);
    const pickupLocationIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.PICKUP_LOCATION);
    const customerIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.CUSTOMER);
    const lastSyncedIndex = headers.indexOf(CONFIG.PICKUP_COLUMNS.LAST_SYNCED);
    
    if (sellerIndex !== -1) row[sellerIndex] = order[CONFIG.DASHBOARD_COLUMNS.SELLER] || '';
    if (idIndex !== -1) row[idIndex] = order[CONFIG.DASHBOARD_COLUMNS.ORDER_ID] || '';
    if (pickupLocationIndex !== -1) row[pickupLocationIndex] = order[CONFIG.DASHBOARD_COLUMNS.PICKUP_LOCATION] || ''; // Now maps to End Location
    if (customerIndex !== -1) row[customerIndex] = order[CONFIG.DASHBOARD_COLUMNS.BUYER] || '';
    if (lastSyncedIndex !== -1) row[lastSyncedIndex] = new Date();
    
    return row;
  });
  
  // Add data to sheet
  if (dataRows.length > 0) {
    const lastRow = pickupSheet.getLastRow();
    const range = pickupSheet.getRange(lastRow + 1, 1, dataRows.length, headers.length);
    range.setValues(dataRows);
  }
}

/**
 * Create Pick Up Orders sheet with headers
 */
function createPickupOrdersSheet() {
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const pickupSheet = currentSpreadsheet.insertSheet(CONFIG.PICKUP_ORDERS_SHEET_NAME);
  
  // Set up headers - only the needed columns
  const headers = [
    CONFIG.PICKUP_COLUMNS.SELLER,         // 'Seller'
    CONFIG.PICKUP_COLUMNS.ID,             // 'Order ID' 
    CONFIG.PICKUP_COLUMNS.PICKUP_LOCATION, // 'Pickup Location'
    CONFIG.PICKUP_COLUMNS.CUSTOMER,       // 'Buyer'
    CONFIG.PICKUP_COLUMNS.LAST_SYNCED     // 'Last Synced'
  ];
  
  pickupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = pickupSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e1f5fe');
  
  return pickupSheet;
}

/**
 * Ensure headers exist in Pick Up Orders sheet
 */
function ensureHeaders(sheet) {
  const expectedHeaders = [
    CONFIG.PICKUP_COLUMNS.SELLER,         // 'Seller'
    CONFIG.PICKUP_COLUMNS.ID,             // 'Order ID'
    CONFIG.PICKUP_COLUMNS.PICKUP_LOCATION, // 'Pickup Location'
    CONFIG.PICKUP_COLUMNS.CUSTOMER,       // 'Buyer'
    CONFIG.PICKUP_COLUMNS.LAST_SYNCED     // 'Last Synced'
  ];
  
  if (sheet.getLastRow() === 0) {
    // No data, add headers
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e1f5fe');
  }
  
  // Return current headers (expand if needed to match expected length)
  const currentCols = Math.max(sheet.getLastColumn(), expectedHeaders.length);
  const currentHeaders = sheet.getRange(1, 1, 1, currentCols).getValues()[0];
  
  // Ensure we have all expected headers
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (!currentHeaders[i]) {
      currentHeaders[i] = expectedHeaders[i];
    }
  }
  
  return currentHeaders;
}

/**
 * Set up automatic trigger to run every 4 hours
 * Run this function once to set up the trigger
 */
function setupAutoTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncMarketplaceOrders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every 4 hours
  ScriptApp.newTrigger('syncMarketplaceOrders')
    .timeBased()
    .everyHours(4)
    .create();
    
  console.log('Auto-trigger set up successfully - will run every 4 hours');
}

/**
 * Remove the automatic trigger
 * Run this function to stop automatic syncing
 */
function removeAutoTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncMarketplaceOrders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  console.log('Auto-trigger removed successfully');
}

/**
 * Test function to check configuration and connectivity
 */
function testConfiguration() {
  console.log('Testing configuration...');
  
  try {
    // Test Dashboard access
    const dashboardSpreadsheet = SpreadsheetApp.openByUrl(CONFIG.DASHBOARD_SHEET_URL);
    const dashboardSheet = dashboardSpreadsheet.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
    console.log('✓ Dashboard sheet accessible');
    
    // Test headers location
    const totalRows = dashboardSheet.getLastRow();
    const totalCols = dashboardSheet.getLastColumn();
    console.log(`Sheet has ${totalRows} rows and ${totalCols} columns`);
    
    if (totalRows >= CONFIG.HEADER_ROW) {
      const headers = dashboardSheet.getRange(CONFIG.HEADER_ROW, 1, 1, totalCols).getValues()[0];
      console.log(`Headers from row ${CONFIG.HEADER_ROW}:`, headers);
      
      // Check required columns
      const requiredColumns = Object.values(CONFIG.DASHBOARD_COLUMNS);
      console.log('Required columns:', requiredColumns);
      
      requiredColumns.forEach(col => {
        const found = headers.includes(col);
        console.log(`Column "${col}": ${found ? '✓ Found' : '✗ Missing'}`);
      });
      
      // Test sample data
      if (totalRows > CONFIG.HEADER_ROW) {
        const sampleDataRange = dashboardSheet.getRange(CONFIG.HEADER_ROW + 1, 1, Math.min(5, totalRows - CONFIG.HEADER_ROW), totalCols);
        const sampleData = sampleDataRange.getValues();
        console.log('Sample data rows:', sampleData.length);
        
        // Convert sample to objects and show raw data
        const sampleObjects = sampleData.map((row, index) => {
          const rowObj = {};
          headers.forEach((header, colIndex) => {
            rowObj[header] = row[colIndex];
          });
          
          console.log(`\nRow ${index + 1} data:`);
          console.log(`  ID: "${rowObj[CONFIG.DASHBOARD_COLUMNS.ORDER_ID]}"`);
          console.log(`  Order Method: "${rowObj[CONFIG.DASHBOARD_COLUMNS.ORDER_METHOD]}"`);
          console.log(`  Dropped Off: "${rowObj[CONFIG.DASHBOARD_COLUMNS.DROPPED_OFF]}"`);
          console.log(`  Seller: "${rowObj[CONFIG.DASHBOARD_COLUMNS.SELLER]}"`);
          console.log(`  Customer: "${rowObj[CONFIG.DASHBOARD_COLUMNS.BUYER]}"`);
          console.log(`  End Location: "${rowObj[CONFIG.DASHBOARD_COLUMNS.PICKUP_LOCATION]}"`); // Updated logging
          
          return rowObj;
        });
        
        // Test filtering with detailed output
        console.log('\n=== FILTERING TEST ===');
        const filteredSample = filterOrders(sampleObjects);
        console.log(`Sample filtering result: ${sampleObjects.length} total, ${filteredSample.length} matching criteria`);
        
        if (filteredSample.length > 0) {
          console.log('Sample matching orders:', filteredSample.map(o => o[CONFIG.DASHBOARD_COLUMNS.ORDER_ID]));
        }
      }
    } else {
      console.error(`Headers expected on row ${CONFIG.HEADER_ROW}, but sheet only has ${totalRows} rows`);
    }
    
    // Test local sheet access
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('✓ Current spreadsheet accessible:', currentSpreadsheet.getName());
    
    console.log('Configuration test completed');
    
  } catch (error) {
    console.error('Configuration test failed:', error);
    console.error('Error details:', error.toString());
  }
}
