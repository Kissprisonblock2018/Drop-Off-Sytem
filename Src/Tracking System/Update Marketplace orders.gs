/**
 * Main function to merge seller drop-off data with dashboard data
 * This function can be called from other scripts or triggered manually
 */
function mergeSellerDropOffData() {
  try {
    // Get the spreadsheet (assuming all sheets are in the same spreadsheet)
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // First, let's see what sheets are available
    const allSheets = spreadsheet.getSheets();
    Logger.log('Available sheets:');
    allSheets.forEach(sheet => {
      Logger.log(`  - "${sheet.getName()}"`);
    });
    
    // Try to find sheets by partial name matching
    const dashboardSheet = findSheetByPartialName(spreadsheet, 'Dashboard');
    const medfordSheet = findSheetByPartialName(spreadsheet, 'Medford');
    const kingstonSheet = findSheetByPartialName(spreadsheet, 'Kingston');
    
    if (!dashboardSheet) {
      throw new Error('Dashboard sheet not found. Available sheets: ' + allSheets.map(s => s.getName()).join(', '));
    }
    if (!medfordSheet) {
      throw new Error('Medford drop-off sheet not found. Available sheets: ' + allSheets.map(s => s.getName()).join(', '));
    }
    if (!kingstonSheet) {
      throw new Error('Kingston drop-off sheet not found. Available sheets: ' + allSheets.map(s => s.getName()).join(', '));
    }
    
    Logger.log(`Using sheets: Dashboard="${dashboardSheet.getName()}", Medford="${medfordSheet.getName()}", Kingston="${kingstonSheet.getName()}"`);
    
    // Get dashboard data
    const dashboardData = getDashboardData(dashboardSheet);
    
    // Get drop-off data from both locations
    const medfordDropOffs = getDropOffData(medfordSheet, 'Medford');
    const kingstonDropOffs = getDropOffData(kingstonSheet, 'Kingston');
    
    // Combine all drop-offs
    const allDropOffs = [...medfordDropOffs, ...kingstonDropOffs];
    
    // Update dashboard with drop-off information
    updateDashboardWithDropOffs(dashboardSheet, dashboardData, allDropOffs);
    
    Logger.log(`Successfully processed ${allDropOffs.length} drop-off records`);
    Logger.log(`Updated ${dashboardData.dataRange.getNumRows() - dashboardData.headerRowIndex - 1} dashboard rows`);
    
  } catch (error) {
    Logger.log('Error in mergeSellerDropOffData: ' + error.toString());
    throw error;
  }
}

/**
 * Helper function to find a sheet by partial name matching (case-insensitive)
 */
function findSheetByPartialName(spreadsheet, partialName) {
  const sheets = spreadsheet.getSheets();
  const normalizedPartial = partialName.toLowerCase();
  
  // First try exact match
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase() === normalizedPartial) {
      return sheet;
    }
  }
  
  // Then try partial match
  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().includes(normalizedPartial)) {
      return sheet;
    }
  }
  
  return null;
}

/**
 * Function to list all available sheets and their basic info
 */
function listAvailableSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    Logger.log('=== AVAILABLE SHEETS ===');
    sheets.forEach((sheet, index) => {
      const dataRange = sheet.getDataRange();
      Logger.log(`${index + 1}. Sheet Name: "${sheet.getName()}"`);
      Logger.log(`   Rows: ${dataRange.getNumRows()}, Columns: ${dataRange.getNumColumns()}`);
      
      // Show first row (headers) if available
      if (dataRange.getNumRows() > 0) {
        const firstRow = sheet.getRange(1, 1, 1, Math.min(10, dataRange.getNumColumns())).getValues()[0];
        Logger.log(`   Headers (first 10): ${firstRow.join(', ')}`);
      }
      Logger.log('');
    });
    
  } catch (error) {
    Logger.log('Error listing sheets: ' + error.toString());
  }
}

/**
 * Get dashboard data with headers and create a lookup map
 */
function getDashboardData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length === 0) {
    throw new Error('Dashboard sheet is empty');
  }
  
  // Find the actual header row that contains "ID" in the first column with data
  let headerRowIndex = -1;
  let headers = null;
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    // Look for a row that has "ID" as one of the first few columns
    for (let j = 0; j < Math.min(5, row.length); j++) {
      if (row[j] && row[j].toString().trim() === 'ID') {
        headerRowIndex = i;
        headers = row;
        break;
      }
    }
    if (headerRowIndex !== -1) break;
  }
  
  if (headerRowIndex === -1) {
    throw new Error('Could not find header row with ID column in Dashboard sheet');
  }
  
  Logger.log(`Found header row at index ${headerRowIndex}`);
  Logger.log(`Headers: ${headers.map((h, i) => `${i}:"${h}"`).join(', ')}`);
  
  // Find column indices - note some headers have trailing spaces
  const columnIndices = {
    id: findColumnIndex(headers, 'ID'),
    seller: findColumnIndex(headers, 'Seller'),
    startLocation: findColumnIndex(headers, 'Start Location'),
    droppedOff: findColumnIndex(headers, 'Dropped Off')
  };
  
  // Log what columns we found for debugging
  Logger.log(`Dashboard column indices: ID=${columnIndices.id}, Seller=${columnIndices.seller}, Start Location=${columnIndices.startLocation}, Dropped Off=${columnIndices.droppedOff}`);
  
  // Validate required columns exist
  if (columnIndices.id === -1) {
    throw new Error('ID column not found in Dashboard');
  }
  if (columnIndices.seller === -1) {
    throw new Error('Seller column not found in Dashboard');
  }
  if (columnIndices.startLocation === -1) {
    throw new Error('Start Location column not found in Dashboard');
  }
  if (columnIndices.droppedOff === -1) {
    throw new Error('Dropped Off column not found in Dashboard');
  }
  
  // Create lookup map for faster searching (skip header row and any empty rows above it)
  // Key format: "OrderID|SellerName" for compound matching
  const idToRowMap = new Map();
  for (let i = headerRowIndex + 1; i < values.length; i++) {
    const orderId = values[i][columnIndices.id];
    const sellerName = values[i][columnIndices.seller];
    if (orderId && sellerName) {
      const key = `${orderId.toString().trim()}|${sellerName.toString().trim()}`;
      idToRowMap.set(key, i);
    }
  }
  
  return {
    dataRange: dataRange,
    values: values,
    headers: headers,
    headerRowIndex: headerRowIndex,
    columnIndices: columnIndices,
    idToRowMap: idToRowMap
  };
}

/**
 * Get drop-off data from a sheet
 */
function getDropOffData(sheet, location) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  if (values.length <= 1) {
    Logger.log(`No data found in ${location} drop-off sheet`);
    return [];
  }
  
  // Find the actual header row that contains "Order ID (if Market Place)"
  let headerRowIndex = -1;
  let headers = null;
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    for (let j = 0; j < row.length; j++) {
      if (row[j] && row[j].toString().includes('Order ID (if Market Place)')) {
        headerRowIndex = i;
        headers = row;
        break;
      }
    }
    if (headerRowIndex !== -1) break;
  }
  
  if (headerRowIndex === -1) {
    Logger.log(`Header row not found in ${location} sheet`);
    return [];
  }
  
  const orderIdIndex = findColumnIndex(headers, 'Order ID (if Market Place)');
  const sellerNameIndex = findColumnIndex(headers, 'Seller Name');
  
  if (orderIdIndex === -1) {
    Logger.log(`Order ID (if Market Place) column not found in ${location} sheet`);
    return [];
  }
  
  if (sellerNameIndex === -1) {
    Logger.log(`Seller Name column not found in ${location} sheet`);
    return [];
  }
  
  const dropOffs = [];
  
  // Process each row (skip header row and any empty rows above it)
  for (let i = headerRowIndex + 1; i < values.length; i++) {
    const orderId = values[i][orderIdIndex];
    const sellerName = values[i][sellerNameIndex];
    
    // Skip rows with no order ID or seller name, or placeholder values
    if (!orderId || !sellerName || 
        orderId.toString().trim() === '' || 
        sellerName.toString().trim() === '' ||
        orderId.toString().includes('--')) {
      continue;
    }
    
    dropOffs.push({
      orderId: orderId.toString().trim(),
      sellerName: sellerName.toString().trim(),
      location: location,
      rowIndex: i
    });
  }
  
  Logger.log(`Found ${dropOffs.length} valid drop-offs in ${location}`);
  return dropOffs;
}

/**
 * Update dashboard with drop-off information
 */
function updateDashboardWithDropOffs(sheet, dashboardData, dropOffs) {
  let updatedCount = 0;
  
  // Group drop-offs by compound key (orderId|sellerName) to handle duplicates
  const dropOffsByKey = new Map();
  dropOffs.forEach(dropOff => {
    const key = `${dropOff.orderId}|${dropOff.sellerName}`;
    if (!dropOffsByKey.has(key)) {
      dropOffsByKey.set(key, []);
    }
    dropOffsByKey.get(key).push(dropOff);
  });
  
  // Process each unique order ID + seller combination
  dropOffsByKey.forEach((orderDropOffs, key) => {
    const dashboardRowIndex = dashboardData.idToRowMap.get(key);
    
    if (dashboardRowIndex !== undefined) {
      // Use the first drop-off location if there are multiple
      const dropOff = orderDropOffs[0];
      
      // Update Start Location
      const startLocationCell = sheet.getRange(dashboardRowIndex + 1, dashboardData.columnIndices.startLocation + 1);
      startLocationCell.setValue(dropOff.location);
      
      // Update Dropped Off to TRUE
      const droppedOffCell = sheet.getRange(dashboardRowIndex + 1, dashboardData.columnIndices.droppedOff + 1);
      droppedOffCell.setValue(true);
      
      updatedCount++;
      
      Logger.log(`Updated Order ID ${dropOff.orderId} for seller ${dropOff.sellerName} with location ${dropOff.location}`);
      
      if (orderDropOffs.length > 1) {
        Logger.log(`Multiple drop-offs found for Order ID ${dropOff.orderId} + Seller ${dropOff.sellerName}, used location: ${dropOff.location}`);
      }
    } else {
      const [orderId, sellerName] = key.split('|');
      Logger.log(`Order ID ${orderId} + Seller ${sellerName} combination not found in dashboard`);
    }
  });
  
  Logger.log(`Updated ${updatedCount} dashboard records`);
}

/**
 * Helper function to find column index by header name (case-insensitive, handles extra spaces)
 */
function findColumnIndex(headers, columnName) {
  const normalizedColumnName = columnName.toLowerCase().trim();
  
  for (let i = 0; i < headers.length; i++) {
    const headerValue = headers[i];
    if (headerValue && headerValue.toString().toLowerCase().trim() === normalizedColumnName) {
      return i;
    }
  }
  
  // If exact match not found, try partial matching
  for (let i = 0; i < headers.length; i++) {
    const headerValue = headers[i];
    if (headerValue && headerValue.toString().toLowerCase().trim().includes(normalizedColumnName)) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Test function to verify the merge works correctly
 */
function testMergeSellerDropOffData() {
  try {
    Logger.log('Starting test merge...');
    mergeSellerDropOffData();
    Logger.log('Test merge completed successfully');
  } catch (error) {
    Logger.log('Test merge failed: ' + error.toString());
  }
}

/**
 * Function to get merge statistics without making changes
 */
function getMergeStatistics() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    const dashboardSheet = findSheetByPartialName(spreadsheet, 'Dashboard');
    const medfordSheet = findSheetByPartialName(spreadsheet, 'Medford');
    const kingstonSheet = findSheetByPartialName(spreadsheet, 'Kingston');
    
    const dashboardData = getDashboardData(dashboardSheet);
    const medfordDropOffs = getDropOffData(medfordSheet, 'Medford');
    const kingstonDropOffs = getDropOffData(kingstonSheet, 'Kingston');
    
    const allDropOffs = [...medfordDropOffs, ...kingstonDropOffs];
    
    let matchCount = 0;
    let noMatchCount = 0;
    
    allDropOffs.forEach(dropOff => {
      const key = `${dropOff.orderId}|${dropOff.sellerName}`;
      if (dashboardData.idToRowMap.has(key)) {
        matchCount++;
      } else {
        noMatchCount++;
      }
    });
    
    Logger.log('=== MERGE STATISTICS ===');
    Logger.log(`Dashboard records: ${dashboardData.dataRange.getNumRows() - dashboardData.headerRowIndex - 1}`);
    Logger.log(`Medford drop-offs: ${medfordDropOffs.length}`);
    Logger.log(`Kingston drop-offs: ${kingstonDropOffs.length}`);
    Logger.log(`Total drop-offs: ${allDropOffs.length}`);
    Logger.log(`Matching records: ${matchCount}`);
    Logger.log(`Non-matching records: ${noMatchCount}`);
    
    return {
      dashboardRecords: dashboardData.dataRange.getNumRows() - dashboardData.headerRowIndex - 1,
      medfordDropOffs: medfordDropOffs.length,
      kingstonDropOffs: kingstonDropOffs.length,
      totalDropOffs: allDropOffs.length,
      matchingRecords: matchCount,
      nonMatchingRecords: noMatchCount
    };
    
  } catch (error) {
    Logger.log('Error getting merge statistics: ' + error.toString());
    throw error;
  }
}
