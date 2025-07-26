/**
 * ENHANCED Transfer System with Marketplace + Virtual Orders
 * Handles both Marketplace orders and Virtual orders from drop-off sheets
 * Checks BOTH Dashboard AND Completed Pick Up / Ship for duplicates
 */

function transferAllOrdersComplete() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log("🚀 Starting enhanced transfer system for Marketplace + Virtual orders...");
    
    // Transfer marketplace orders
    const marketplaceResult = transferMarketplaceOrdersComplete(spreadsheet);
    
    // Transfer virtual orders
    const virtualResult = transferVirtualOrdersComplete(spreadsheet);
    
    const totalResult = {
      success: marketplaceResult.success && virtualResult.success,
      marketplaceRecords: marketplaceResult.recordsTransferred || 0,
      virtualRecords: virtualResult.recordsTransferred || 0,
      totalRecords: (marketplaceResult.recordsTransferred || 0) + (virtualResult.recordsTransferred || 0),
      message: `Transfer completed: ${marketplaceResult.recordsTransferred || 0} marketplace + ${virtualResult.recordsTransferred || 0} virtual = ${(marketplaceResult.recordsTransferred || 0) + (virtualResult.recordsTransferred || 0)} total records`
    };
    
    Logger.log("📋 Enhanced transfer completed:", JSON.stringify(totalResult, null, 2));
    return totalResult;
    
  } catch (error) {
    Logger.log(`❌ Error in transferAllOrdersComplete: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: `Enhanced transfer failed: ${error.message}`
    };
  }
}

/**
 * Transfer Marketplace orders (existing functionality)
 */
function transferMarketplaceOrdersComplete(spreadsheet) {
  try {
    Logger.log("📦 Processing Marketplace orders...");
    
    // Find required sheets
    const sourceSheetName = findSourceSheetComplete(spreadsheet);
    const dashboardSheetName = findDashboardSheetComplete(spreadsheet);
    const completedSheetName = findCompletedSheetComplete(spreadsheet);
    
    const sourceSheet = sourceSheetName ? spreadsheet.getSheetByName(sourceSheetName) : null;
    const dashboardSheet = dashboardSheetName ? spreadsheet.getSheetByName(dashboardSheetName) : null;
    const completedSheet = completedSheetName ? spreadsheet.getSheetByName(completedSheetName) : null;
    
    if (!sourceSheet) {
      throw new Error(`Marketplace source sheet not found. Available: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
    }
    
    if (!dashboardSheet) {
      throw new Error(`Dashboard sheet not found. Available: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
    }
    
    Logger.log(`✅ Marketplace Source: "${sourceSheetName}"`);
    Logger.log(`✅ Dashboard: "${dashboardSheetName}"`);
    Logger.log(`✅ Completed: "${completedSheetName || 'NOT FOUND'}"`);
    
    // Get source data
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length === 0) {
      throw new Error("Marketplace source sheet is empty");
    }
    
    const sourceHeaders = sourceData[0];
    const sourceRows = sourceData.slice(1);
    
    Logger.log(`📊 Marketplace source has ${sourceRows.length} data rows`);
    
    // Find source column indices
    const sourceColumns = {
      orderId: findColumnIndexComplete(sourceHeaders, ["Order ID"]),
      seller: findColumnIndexComplete(sourceHeaders, ["Seller"]),
      buyer: findColumnIndexComplete(sourceHeaders, ["Buyer"]),
      pickupLocation: findColumnIndexComplete(sourceHeaders, ["Pickup Location"]),
      orderedDate: findColumnIndexComplete(sourceHeaders, ["Ordered Date"])
    };
    
    // Validate source columns
    const missingColumns = [];
    Object.entries(sourceColumns).forEach(([key, index]) => {
      if (index === -1) {
        missingColumns.push(key);
      }
    });
    
    if (missingColumns.length > 0) {
      throw new Error(`Missing marketplace columns: ${missingColumns.join(', ')}`);
    }
    
    // Get existing pairs and dashboard structure
    const allExistingPairs = getAllExistingOrderIdSellerPairs(dashboardSheet, completedSheet);
    const dashboardStructure = findDashboardStructure(dashboardSheet);
    
    // Filter new records
    const newRecords = sourceRows.filter(row => {
      const orderId = row[sourceColumns.orderId];
      const seller = row[sourceColumns.seller];
      
      if (!orderId || orderId.toString().trim() === "" || !seller || seller.toString().trim() === "") {
        return false;
      }
      
      const pairKey = `${orderId.toString().trim()}|${seller.toString().trim()}`;
      const isDuplicate = allExistingPairs.has(pairKey);
      
      if (isDuplicate) {
        Logger.log(`⏭️ Skipping marketplace duplicate: ${pairKey}`);
      }
      
      return !isDuplicate;
    });
    
    Logger.log(`📊 Marketplace Analysis: ${sourceRows.length} source, ${newRecords.length} new`);
    
    // Add new records
    let recordsAdded = 0;
    let currentRow = findNextAvailableRowComplete(dashboardSheet, dashboardStructure.dataStartRow, dashboardStructure.columns.id);
    
    for (const record of newRecords) {
      const orderId = record[sourceColumns.orderId];
      const seller = record[sourceColumns.seller] || "";
      const buyer = record[sourceColumns.buyer] || "";
      const pickupLocation = record[sourceColumns.pickupLocation] || "";
      const orderedDate = record[sourceColumns.orderedDate] || "";
      
      Logger.log(`➕ Adding marketplace ${orderId} to Dashboard row ${currentRow}`);
      
      try {
        // Insert basic data
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.id + 1).setValue(orderId);
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.seller + 1).setValue(seller);
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.customer + 1).setValue(buyer);
        
        if (dashboardStructure.columns.orderMethod !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.orderMethod + 1).setValue("Marketplace");
        }
        if (dashboardStructure.columns.instorePickup !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.instorePickup + 1).setValue("Instore Pick");
        }
        if (dashboardStructure.columns.endLocation !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.endLocation + 1).setValue(pickupLocation);
        }
        if (dashboardStructure.columns.dateCreated !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.dateCreated + 1).setValue(orderedDate);
        }
        
        // Set checkbox fields for marketplace orders
        setCheckboxFieldsForOrderType(dashboardSheet, currentRow, dashboardStructure.columns, "Marketplace");
        
        recordsAdded++;
        currentRow = findNextAvailableRowComplete(dashboardSheet, currentRow + 1, dashboardStructure.columns.id);
        
      } catch (error) {
        Logger.log(`❌ Failed to add marketplace ${orderId}: ${error.message}`);
      }
    }
    
    Logger.log(`✅ Added ${recordsAdded} marketplace records to Dashboard`);
    
    return {
      success: true,
      recordsTransferred: recordsAdded,
      message: `Successfully transferred ${recordsAdded} marketplace records`
    };
    
  } catch (error) {
    Logger.log(`❌ Marketplace transfer error: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: `Marketplace transfer failed: ${error.message}`
    };
  }
}

/**
 * Transfer Virtual orders from drop-off sheets (NEW FUNCTIONALITY)
 */
function transferVirtualOrdersComplete(spreadsheet) {
  try {
    Logger.log("🔄 Processing Virtual orders...");
    
    // Find drop-off sheets
    const kingstonSheet = findSheetByPartialName(spreadsheet, 'Kingston');
    const medfordSheet = findSheetByPartialName(spreadsheet, 'Medford');
    const dashboardSheet = findSheetByPartialName(spreadsheet, 'Dashboard');
    const completedSheet = findCompletedSheetComplete(spreadsheet) ? spreadsheet.getSheetByName(findCompletedSheetComplete(spreadsheet)) : null;
    
    if (!dashboardSheet) {
      throw new Error('Dashboard sheet not found for virtual orders');
    }
    
    Logger.log(`✅ Kingston: "${kingstonSheet ? kingstonSheet.getName() : 'NOT FOUND'}"`);
    Logger.log(`✅ Medford: "${medfordSheet ? medfordSheet.getName() : 'NOT FOUND'}"`);
    Logger.log(`✅ Dashboard: "${dashboardSheet.getName()}"`);
    
    // Collect all virtual orders from both drop-off sheets
    const virtualOrders = [];
    
    if (kingstonSheet) {
      const kingstonVirtual = getVirtualOrdersFromSheet(kingstonSheet, 'Kingston');
      virtualOrders.push(...kingstonVirtual);
      Logger.log(`📦 Found ${kingstonVirtual.length} virtual orders in Kingston`);
    }
    
    if (medfordSheet) {
      const medfordVirtual = getVirtualOrdersFromSheet(medfordSheet, 'Medford');
      virtualOrders.push(...medfordVirtual);
      Logger.log(`📦 Found ${medfordVirtual.length} virtual orders in Medford`);
    }
    
    Logger.log(`📊 Total virtual orders found: ${virtualOrders.length}`);
    
    if (virtualOrders.length === 0) {
      return {
        success: true,
        recordsTransferred: 0,
        message: "No virtual orders found to transfer"
      };
    }
    
    // Get existing pairs and dashboard structure
    const allExistingPairs = getAllExistingOrderIdSellerPairs(dashboardSheet, completedSheet);
    const dashboardStructure = findDashboardStructure(dashboardSheet);
    
    // Filter new virtual orders (check for duplicates)
    const newVirtualOrders = virtualOrders.filter(order => {
      const pairKey = `${order.orderId.toString().trim()}|${order.sellerName.toString().trim()}`;
      const isDuplicate = allExistingPairs.has(pairKey);
      
      if (isDuplicate) {
        Logger.log(`⏭️ Skipping virtual duplicate: ${pairKey}`);
      }
      
      return !isDuplicate;
    });
    
    Logger.log(`📊 Virtual Analysis: ${virtualOrders.length} found, ${newVirtualOrders.length} new`);
    
    // Add new virtual orders to Dashboard
    let recordsAdded = 0;
    let currentRow = findNextAvailableRowComplete(dashboardSheet, dashboardStructure.dataStartRow, dashboardStructure.columns.id);
    
    for (const order of newVirtualOrders) {
      Logger.log(`➕ Adding virtual ${order.orderId} to Dashboard row ${currentRow}`);
      
      try {
        // Insert basic data
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.id + 1).setValue(order.orderId);
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.seller + 1).setValue(order.sellerName);
        dashboardSheet.getRange(currentRow, dashboardStructure.columns.customer + 1).setValue(order.customerName);
        
        if (dashboardStructure.columns.orderMethod !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.orderMethod + 1).setValue("Virtual");
        }
        if (dashboardStructure.columns.instorePickup !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.instorePickup + 1).setValue("Instore Pick");
        }
        if (dashboardStructure.columns.endLocation !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.endLocation + 1).setValue(order.finalDestination);
        }
        if (dashboardStructure.columns.startLocation !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.startLocation + 1).setValue(order.dropOffLocation);
        }
        if (dashboardStructure.columns.dateCreated !== -1) {
          dashboardSheet.getRange(currentRow, dashboardStructure.columns.dateCreated + 1).setValue(order.dateOfDropOff);
        }
        
        // Set checkbox fields for virtual orders
        setCheckboxFieldsForOrderType(dashboardSheet, currentRow, dashboardStructure.columns, "Virtual");
        
        recordsAdded++;
        currentRow = findNextAvailableRowComplete(dashboardSheet, currentRow + 1, dashboardStructure.columns.id);
        
      } catch (error) {
        Logger.log(`❌ Failed to add virtual ${order.orderId}: ${error.message}`);
      }
    }
    
    Logger.log(`✅ Added ${recordsAdded} virtual records to Dashboard`);
    
    return {
      success: true,
      recordsTransferred: recordsAdded,
      message: `Successfully transferred ${recordsAdded} virtual records`
    };
    
  } catch (error) {
    Logger.log(`❌ Virtual transfer error: ${error.message}`);
    return {
      success: false,
      error: error.message,
      message: `Virtual transfer failed: ${error.message}`
    };
  }
}

/**
 * Extract virtual orders from a drop-off sheet
 */
function getVirtualOrdersFromSheet(sheet, location) {
  const virtualOrders = [];
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log(`No data in ${location} sheet`);
      return virtualOrders;
    }
    
    // Find header row
    let headerRowIndex = -1;
    let headers = null;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Look for "Order ID (If Virtual)" or "Type of Order" to identify header row
      for (let j = 0; j < row.length; j++) {
        if (row[j] && (row[j].toString().includes('Order ID (If Virtual)') || row[j].toString().includes('Type of Order'))) {
          headerRowIndex = i;
          headers = row;
          break;
        }
      }
      if (headerRowIndex !== -1) break;
    }
    
    if (headerRowIndex === -1) {
      Logger.log(`Header row not found in ${location} sheet`);
      return virtualOrders;
    }
    
    // Find column indices
    const virtualOrderIdIndex = findColumnIndexComplete(headers, ['Order ID (If Virtual)']);
    const sellerNameIndex = findColumnIndexComplete(headers, ['Seller Name']);
    const customerNameIndex = findColumnIndexComplete(headers, ['Customer Name']);
    const finalDestinationIndex = findColumnIndexComplete(headers, ['Final Destination']);
    const dropOffLocationIndex = findColumnIndexComplete(headers, ['Drop Off Location']);
    const dateOfDropOffIndex = findColumnIndexComplete(headers, ['Date of Drop Off']);
    const typeOfOrderIndex = findColumnIndexComplete(headers, ['Type of Order', 'Type of order']);
    
    if (virtualOrderIdIndex === -1 || typeOfOrderIndex === -1) {
      Logger.log(`Required columns not found in ${location} sheet`);
      return virtualOrders;
    }
    
    Logger.log(`${location} columns: Virtual ID=${virtualOrderIdIndex}, Type=${typeOfOrderIndex}, Seller=${sellerNameIndex}`);
    
    // Process each row
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      const virtualOrderId = row[virtualOrderIdIndex];
      const typeOfOrder = row[typeOfOrderIndex];
      const sellerName = row[sellerNameIndex];
      const customerName = row[customerNameIndex];
      const finalDestination = row[finalDestinationIndex];
      const dropOffLocation = row[dropOffLocationIndex];
      const dateOfDropOff = row[dateOfDropOffIndex];
      
      // Only process rows where Type of Order is "Virtual" and has a Virtual Order ID
      if (virtualOrderId && 
          virtualOrderId.toString().trim() !== '' && 
          typeOfOrder && 
          typeOfOrder.toString().toLowerCase().trim() === 'virtual') {
        
        virtualOrders.push({
          orderId: virtualOrderId.toString().trim(),
          sellerName: sellerName ? sellerName.toString().trim() : '',
          customerName: customerName ? customerName.toString().trim() : '',
          finalDestination: finalDestination ? finalDestination.toString().trim() : '',
          dropOffLocation: dropOffLocation ? dropOffLocation.toString().trim() : location,
          dateOfDropOff: dateOfDropOff || '',
          location: location,
          rowIndex: i
        });
        
        Logger.log(`📦 Found virtual order: ID=${virtualOrderId}, Seller=${sellerName}, Customer=${customerName}`);
      }
    }
    
  } catch (error) {
    Logger.log(`Error processing ${location} sheet: ${error.message}`);
  }
  
  Logger.log(`📊 ${location} virtual orders extracted: ${virtualOrders.length}`);
  return virtualOrders;
}

/**
 * Updated Dashboard structure finder to include Start Location
 */
function findDashboardStructure(sheet) {
  const data = sheet.getDataRange().getValues();
  let headerRowIndex = -1;
  let headers = null;
  
  // Find header row by looking for "ID" column
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
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
    throw new Error('Could not find header row with ID column in Dashboard');
  }
  
  const columns = {
    id: findColumnIndexComplete(headers, ["ID"]),
    seller: findColumnIndexComplete(headers, ["Seller"]),
    customer: findColumnIndexComplete(headers, ["Customer"]),
    orderMethod: findColumnIndexComplete(headers, ["Order Method"]),
    instorePickup: findColumnIndexComplete(headers, ["Instore Pick Up OR Ship"]),
    startLocation: findColumnIndexComplete(headers, ["Start Location"]),  // Added for virtual orders
    endLocation: findColumnIndexComplete(headers, ["End Location"]),
    dateCreated: findColumnIndexComplete(headers, ["Date Created"]),
    // Checkbox fields
    created: findColumnIndexComplete(headers, ["Created"]),
    paid: findColumnIndexComplete(headers, ["Paid"]),
    droppedOff: findColumnIndexComplete(headers, ["Dropped Off"]),
    transported: findColumnIndexComplete(headers, ["Transported"]),
    readyForPickup: findColumnIndexComplete(headers, ["Ready For (Pickup or Ship)"]),
    pickedUpOrShipped: findColumnIndexComplete(headers, ["Picked Up Or Shiped"])
  };
  
  return {
    headerRowIndex: headerRowIndex,
    dataStartRow: headerRowIndex + 2,
    headers: headers,
    columns: columns
  };
}

/**
 * Set checkbox fields based on order type
 */
function setCheckboxFieldsForOrderType(sheet, row, columns, orderType) {
  try {
    let checkboxFields;
    
    if (orderType === "Marketplace") {
      // Marketplace orders: Created and Paid = TRUE
      checkboxFields = [
        { column: columns.created, defaultValue: true, name: "Created" },
        { column: columns.paid, defaultValue: true, name: "Paid" },
        { column: columns.droppedOff, defaultValue: false, name: "Dropped Off" },
        { column: columns.transported, defaultValue: false, name: "Transported" },
        { column: columns.readyForPickup, defaultValue: false, name: "Ready For Pickup" },
        { column: columns.pickedUpOrShipped, defaultValue: false, name: "Picked Up Or Shipped" }
      ];
    } else if (orderType === "Virtual") {
      // Virtual orders: Created, Paid, and Dropped Off = TRUE (since they're already dropped off)
      checkboxFields = [
        { column: columns.created, defaultValue: true, name: "Created" },
        { column: columns.paid, defaultValue: true, name: "Paid" },
        { column: columns.droppedOff, defaultValue: true, name: "Dropped Off" },  // TRUE for virtual
        { column: columns.transported, defaultValue: false, name: "Transported" },
        { column: columns.readyForPickup, defaultValue: false, name: "Ready For Pickup" },
        { column: columns.pickedUpOrShipped, defaultValue: false, name: "Picked Up Or Shipped" }
      ];
    }
    
    checkboxFields.forEach(field => {
      if (field.column !== -1) {
        const cell = sheet.getRange(row, field.column + 1);
        cell.insertCheckboxes();
        cell.setValue(field.defaultValue);
        Logger.log(`  ✅ Set ${field.name} checkbox to ${field.defaultValue} (${orderType})`);
      }
    });
    
  } catch (error) {
    Logger.log(`⚠️ Error setting ${orderType} checkboxes: ${error.message}`);
  }
}

/**
 * Helper function to find sheet by partial name
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

// ========== EXISTING HELPER FUNCTIONS (unchanged) ==========

function getAllExistingOrderIdSellerPairs(dashboardSheet, completedSheet) {
  const allPairs = new Set();
  
  // Get pairs from Dashboard
  const dashboardStructure = findDashboardStructure(dashboardSheet);
  if (dashboardStructure.columns.id !== -1 && dashboardStructure.columns.seller !== -1) {
    const dashboardPairs = getOrderIdSellerPairsFromSheet(dashboardSheet, dashboardStructure.dataStartRow, dashboardStructure.columns.id, dashboardStructure.columns.seller);
    Logger.log(`📊 Dashboard existing Order ID + Seller pairs: ${dashboardPairs.size}`);
    dashboardPairs.forEach(pair => allPairs.add(pair));
  }
  
  // Get pairs from Completed sheet (if it exists)
  if (completedSheet) {
    const completedStructure = findCompletedStructure(completedSheet);
    if (completedStructure.columns.id !== -1 && completedStructure.columns.seller !== -1) {
      const completedPairs = getOrderIdSellerPairsFromSheet(completedSheet, completedStructure.dataStartRow, completedStructure.columns.id, completedStructure.columns.seller);
      Logger.log(`📊 Completed existing Order ID + Seller pairs: ${completedPairs.size}`);
      completedPairs.forEach(pair => allPairs.add(pair));
    }
  }
  
  return allPairs;
}

function getOrderIdSellerPairsFromSheet(sheet, startRow, idColumnIndex, sellerColumnIndex) {
  const pairs = new Set();
  
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow >= startRow) {
      const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, Math.max(idColumnIndex, sellerColumnIndex) + 1);
      const dataValues = dataRange.getValues();
      
      dataValues.forEach(row => {
        const id = row[idColumnIndex];
        const seller = row[sellerColumnIndex];
        
        if (id && seller && id.toString().trim() !== "" && seller.toString().trim() !== "") {
          const pairKey = `${id.toString().trim()}|${seller.toString().trim()}`;
          pairs.add(pairKey);
        }
      });
    }
  } catch (error) {
    Logger.log(`⚠️ Error reading pairs from ${sheet.getName()}: ${error.message}`);
  }
  
  return pairs;
}

function findCompletedStructure(sheet) {
  const data = sheet.getDataRange().getValues();
  let headerRowIndex = -1;
  let headers = null;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
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
    headerRowIndex = 0;
    headers = data[0] || [];
  }
  
  const columns = {
    id: findColumnIndexComplete(headers, ["ID"]),
    seller: findColumnIndexComplete(headers, ["Seller"])
  };
  
  return {
    headerRowIndex: headerRowIndex,
    dataStartRow: headerRowIndex + 2,
    headers: headers,
    columns: columns
  };
}

function findColumnIndexComplete(headers, possibleNames) {
  for (const name of possibleNames) {
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (header && header.toString().trim().toLowerCase() === name.toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}

function findNextAvailableRowComplete(sheet, startRow, idColumnIndex) {
  try {
    for (let row = startRow; row <= startRow + 100; row++) {
      try {
        const idCell = sheet.getRange(row, idColumnIndex + 1).getValue();
        if (!idCell || idCell.toString().trim() === "") {
          return row;
        }
      } catch (rangeError) {
        return row;
      }
    }
    return sheet.getLastRow() + 1;
  } catch (error) {
    Logger.log(`❌ Error finding available row: ${error.message}`);
    return startRow;
  }
}

function findSourceSheetComplete(spreadsheet) {
  const possibleNames = [
    "4GV Marketplavce (All Automated)",
    "4GV Marketplavce All Automated",
    "4GV Marketplace (All Automated)",
    "4GV Marketplace All Automated"
  ];
  
  for (const name of possibleNames) {
    if (spreadsheet.getSheetByName(name)) {
      return name;
    }
  }
  
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName.includes("4GV") && sheetName.includes("Automated")) {
      return sheetName;
    }
  }
  
  return null;
}

function findDashboardSheetComplete(spreadsheet) {
  const possibleNames = [
    "Dashboard",
    "Pick Up _ Ship _ Transport System  Dashboard",
    "Pick Up _ Ship _ Transport System Dashboard"
  ];
  
  for (const name of possibleNames) {
    if (spreadsheet.getSheetByName(name)) {
      return name;
    }
  }
  
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName.toLowerCase().includes("dashboard")) {
      return sheetName;
    }
  }
  
  return null;
}

function findCompletedSheetComplete(spreadsheet) {
  const possibleNames = [
    "Completed Pick Up / Ship",
    "Completed Pick Up/Ship",
    "Completed Pickup Ship",
    "Completed",
    "Pick Up Ship Completed"
  ];
  
  for (const name of possibleNames) {
    if (spreadsheet.getSheetByName(name)) {
      return name;
    }
  }
  
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName.toLowerCase().includes("completed")) {
      return sheetName;
    }
  }
  
  return null;
}

/**
 * MAIN EXECUTION FUNCTIONS
 */

// Run this for the complete transfer (both marketplace and virtual)
function runEnhancedTransferSystem() {
  Logger.log("🚀 Starting enhanced transfer system...");
  const result = transferAllOrdersComplete();
  Logger.log("📋 Enhanced transfer completed:", JSON.stringify(result, null, 2));
  return result;
}

// Run this for marketplace orders only
function runMarketplaceOnlyTransfer() {
  Logger.log("📦 Starting marketplace-only transfer...");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = transferMarketplaceOrdersComplete(spreadsheet);
  Logger.log("📋 Marketplace transfer completed:", JSON.stringify(result, null, 2));
  return result;
}

// Run this for virtual orders only
function runVirtualOnlyTransfer() {
  Logger.log("🔄 Starting virtual-only transfer...");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const result = transferVirtualOrdersComplete(spreadsheet);
  Logger.log("📋 Virtual transfer completed:", JSON.stringify(result, null, 2));
  return result;
}

/**
 * DEBUG FUNCTIONS for testing and troubleshooting
 */

// Debug function to show what virtual orders are found
function debugVirtualOrders() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log("=== VIRTUAL ORDERS DEBUG ===");
    
    const kingstonSheet = findSheetByPartialName(spreadsheet, 'Kingston');
    const medfordSheet = findSheetByPartialName(spreadsheet, 'Medford');
    
    if (kingstonSheet) {
      Logger.log(`\n📍 KINGSTON SHEET: "${kingstonSheet.getName()}"`);
      const kingstonVirtual = getVirtualOrdersFromSheet(kingstonSheet, 'Kingston');
      Logger.log(`Found ${kingstonVirtual.length} virtual orders in Kingston:`);
      kingstonVirtual.forEach((order, index) => {
        Logger.log(`  ${index + 1}. ID: ${order.orderId} | Seller: ${order.sellerName} | Customer: ${order.customerName}`);
      });
    } else {
      Logger.log("❌ Kingston sheet not found");
    }
    
    if (medfordSheet) {
      Logger.log(`\n📍 MEDFORD SHEET: "${medfordSheet.getName()}"`);
      const medfordVirtual = getVirtualOrdersFromSheet(medfordSheet, 'Medford');
      Logger.log(`Found ${medfordVirtual.length} virtual orders in Medford:`);
      medfordVirtual.forEach((order, index) => {
        Logger.log(`  ${index + 1}. ID: ${order.orderId} | Seller: ${order.sellerName} | Customer: ${order.customerName}`);
      });
    } else {
      Logger.log("❌ Medford sheet not found");
    }
    
    Logger.log("================================");
    
  } catch (error) {
    Logger.log(`❌ Debug error: ${error.message}`);
  }
}

// Debug function to show existing Order ID + Seller pairs
function debugExistingPairs() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = findSheetByPartialName(spreadsheet, 'Dashboard');
    const completedSheet = findCompletedSheetComplete(spreadsheet) ? spreadsheet.getSheetByName(findCompletedSheetComplete(spreadsheet)) : null;
    
    Logger.log("=== EXISTING PAIRS DEBUG ===");
    
    if (dashboardSheet) {
      const allPairs = getAllExistingOrderIdSellerPairs(dashboardSheet, completedSheet);
      Logger.log(`Total existing Order ID + Seller pairs: ${allPairs.size}`);
      
      Logger.log("\nFirst 20 existing pairs:");
      Array.from(allPairs).slice(0, 20).forEach((pair, index) => {
        Logger.log(`  ${index + 1}. ${pair}`);
      });
      
      if (allPairs.size > 20) {
        Logger.log(`  ... and ${allPairs.size - 20} more pairs`);
      }
    } else {
      Logger.log("❌ Dashboard sheet not found");
    }
    
    Logger.log("==============================");
    
  } catch (error) {
    Logger.log(`❌ Debug error: ${error.message}`);
  }
}

// Debug function to show dashboard structure
function debugDashboardStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = findSheetByPartialName(spreadsheet, 'Dashboard');
    
    Logger.log("=== DASHBOARD STRUCTURE DEBUG ===");
    
    if (dashboardSheet) {
      const structure = findDashboardStructure(dashboardSheet);
      
      Logger.log(`Header row: ${structure.headerRowIndex + 1}`);
      Logger.log(`Data starts at row: ${structure.dataStartRow}`);
      Logger.log("\nColumn mapping:");
      
      Object.entries(structure.columns).forEach(([key, index]) => {
        if (index !== -1) {
          Logger.log(`  ${key}: column ${index + 1} ("${structure.headers[index]}")`);
        } else {
          Logger.log(`  ${key}: NOT FOUND`);
        }
      });
    } else {
      Logger.log("❌ Dashboard sheet not found");
    }
    
    Logger.log("===================================");
    
  } catch (error) {
    Logger.log(`❌ Debug error: ${error.message}`);
  }
}

// Test function to check what sheets are available
function debugAvailableSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    Logger.log("=== AVAILABLE SHEETS DEBUG ===");
    Logger.log(`Total sheets: ${sheets.length}`);
    
    sheets.forEach((sheet, index) => {
      const name = sheet.getName();
      const rows = sheet.getLastRow();
      const cols = sheet.getLastColumn();
      Logger.log(`  ${index + 1}. "${name}" (${rows} rows, ${cols} columns)`);
    });
    
    Logger.log("\nLooking for key sheets:");
    Logger.log(`  4GV Marketplace: ${findSourceSheetComplete(spreadsheet) || 'NOT FOUND'}`);
    Logger.log(`  Dashboard: ${findDashboardSheetComplete(spreadsheet) || 'NOT FOUND'}`);
    Logger.log(`  Completed: ${findCompletedSheetComplete(spreadsheet) || 'NOT FOUND'}`);
    Logger.log(`  Kingston: ${findSheetByPartialName(spreadsheet, 'Kingston')?.getName() || 'NOT FOUND'}`);
    Logger.log(`  Medford: ${findSheetByPartialName(spreadsheet, 'Medford')?.getName() || 'NOT FOUND'}`);
    
    Logger.log("===============================");
    
  } catch (error) {
    Logger.log(`❌ Debug error: ${error.message}`);
  }
}
