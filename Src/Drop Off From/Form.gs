function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
}

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('FormPopup')
    .setWidth(800)
    .setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(html, 'Order Form');
}

function getTodayAndNextTuesday() {
  const today = new Date();
  let nextTuesday = new Date(today);
  nextTuesday.setDate(today.getDate() + ((9 - today.getDay()) % 7 || 7));
  return {
    today: today.toDateString(),
    nextTuesday: nextTuesday.toDateString(),
    dropLocation: "Kingston"
  };
}

// New function to get seller and order data from the CSV sheet
function getSellerOrderData() {
  try {
    // Get the sheet that contains your seller/order data
    const sellerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pick Up Orders');
    
    if (!sellerSheet) {
      console.error('Pick Up Orders sheet not found');
      console.log('Available sheets:', SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()));
      return { sellers: [], orders: [] };
    }
    
    const dataRange = sellerSheet.getDataRange();
    const data = dataRange.getValues();
    console.log('Data rows found:', data.length);
    
    if (data.length === 0) {
      console.error('No data found in sheet');
      return { sellers: [], orders: [] };
    }
    
    const headers = data[0];
    console.log('Headers found:', headers);
    
    // Find column indices
    const sellerIndex = headers.indexOf('Seller');
    const orderIdIndex = headers.indexOf('Order ID');
    const buyerIndex = headers.indexOf('Buyer');
    
    console.log('Column indices - Seller:', sellerIndex, 'Order ID:', orderIdIndex, 'Buyer:', buyerIndex);
    
    if (sellerIndex === -1 || orderIdIndex === -1) {
      console.error('Required columns not found. Expected: Seller, Order ID');
      return { sellers: [], orders: [] };
    }
    
    const sellers = new Set();
    const orders = [];
    
    // Skip header row and process data
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const seller = row[sellerIndex];
      const orderId = row[orderIdIndex];
      const buyer = buyerIndex !== -1 ? row[buyerIndex] : '';
      
      if (seller && orderId) {
        sellers.add(seller);
        orders.push({
          seller: seller,
          orderId: orderId,
          buyer: buyer,
          displayText: `${orderId} - ${seller}${buyer ? ` (Buyer: ${buyer})` : ''}`
        });
      }
    }
    
    console.log('Processed sellers:', Array.from(sellers).length);
    console.log('Processed orders:', orders.length);
    
    return {
      sellers: Array.from(sellers).sort(),
      orders: orders.sort((a, b) => a.displayText.localeCompare(b.displayText))
    };
  } catch (error) {
    console.error('Error fetching seller/order data:', error);
    return { sellers: [], orders: [] };
  }
}

// New function to get orders by seller
function getOrdersBySeller(sellerName) {
  try {
    const sellerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pick Up Orders');
    
    if (!sellerSheet) {
      return [];
    }
    
    const dataRange = sellerSheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];
    
    const sellerIndex = headers.indexOf('Seller');
    const orderIdIndex = headers.indexOf('Order ID');
    const buyerIndex = headers.indexOf('Buyer');
    
    if (sellerIndex === -1 || orderIdIndex === -1) {
      return [];
    }
    
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const seller = row[sellerIndex];
      const orderId = row[orderIdIndex];
      const buyer = buyerIndex !== -1 ? row[buyerIndex] : '';
      
      if (seller === sellerName && orderId) {
        orders.push({
          seller: seller,
          orderId: orderId,
          buyer: buyer,
          displayText: `${orderId}${buyer ? ` (Buyer: ${buyer})` : ''}`
        });
      }
    }
    
    return orders.sort((a, b) => a.displayText.localeCompare(b.displayText));
  } catch (error) {
    console.error('Error fetching orders for seller:', error);
    return [];
  }
}

function saveMarketplaceOrder(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { today, nextTuesday, dropLocation } = getTodayAndNextTuesday();
  const headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Filter out empty entries
  const validData = data.filter(item => item && item.trim() !== '');
  
  validData.forEach(item => {
    const rowData = new Array(headers.length).fill('');
    
    // Find column indices and set values
    const orderIdIndex = headers.indexOf('Order ID (if Market Place)');
    const dropOffLocationIndex = headers.indexOf('Drop Off Location');
    const dateOfDropOffIndex = headers.indexOf('Date of Drop Off');
    const expectedTransportIndex = headers.indexOf('Expected Transport date (If Needed)');
    
    if (orderIdIndex !== -1) rowData[orderIdIndex] = item;
    if (dropOffLocationIndex !== -1) rowData[dropOffLocationIndex] = dropLocation;
    if (dateOfDropOffIndex !== -1) rowData[dateOfDropOffIndex] = today;
    if (expectedTransportIndex !== -1) rowData[expectedTransportIndex] = nextTuesday;
    
    sheet.insertRows(6, 1);
    sheet.getRange(6, 1, 1, rowData.length).setValues([rowData]);
  });
}

// Enhanced function to save marketplace orders with seller info
function saveMarketplaceOrderWithSeller(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { today, nextTuesday, dropLocation } = getTodayAndNextTuesday();
  const headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Filter out empty entries
  const validData = data.filter(item => item.orderId && item.orderId.trim() !== '');
  
  validData.forEach(item => {
    const rowData = new Array(headers.length).fill('');
    
    // Find column indices and set values
    const orderIdIndex = headers.indexOf('Order ID (if Market Place)');
    const sellerNameIndex = headers.indexOf('Seller Name');
    const customerNameIndex = headers.indexOf('Customer Name');
    const dropOffLocationIndex = headers.indexOf('Drop Off Location');
    const dateOfDropOffIndex = headers.indexOf('Date of Drop Off');
    const expectedTransportIndex = headers.indexOf('Expected Transport date (If Needed)');
    
    if (orderIdIndex !== -1) rowData[orderIdIndex] = item.orderId;
    if (sellerNameIndex !== -1) rowData[sellerNameIndex] = item.seller || '';
    if (customerNameIndex !== -1) rowData[customerNameIndex] = item.buyer || '';
    if (dropOffLocationIndex !== -1) rowData[dropOffLocationIndex] = dropLocation;
    if (dateOfDropOffIndex !== -1) rowData[dateOfDropOffIndex] = today;
    if (expectedTransportIndex !== -1) rowData[expectedTransportIndex] = nextTuesday;
    
    sheet.insertRows(6, 1);
    sheet.getRange(6, 1, 1, rowData.length).setValues([rowData]);
  });
}

function saveVirtualOrder(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const { today, nextTuesday, dropLocation } = getTodayAndNextTuesday();
  const headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Log received data for debugging
  console.log('Received data:', JSON.stringify(data));

  // Filter out empty entries and remove duplicates
  const validData = data.filter((entry, index, self) => {
    // Check if entry has content
    const hasContent = (entry.customerName && entry.customerName.trim() !== '') || 
                      (entry.sellerName && entry.sellerName.trim() !== '');
    
    if (!hasContent) return false;
    
    // Check for duplicates - only keep first occurrence
    const isDuplicate = self.findIndex(e => 
      e.customerName === entry.customerName && 
      e.sellerName === entry.sellerName && 
      e.endLocation === entry.endLocation
    ) !== index;
    
    return !isDuplicate;
  });

  console.log('Valid data after filtering:', JSON.stringify(validData));

  validData.forEach(entry => {
    const rowData = new Array(headers.length).fill('');
    
    // Find column indices and set values
    const customerNameIndex = headers.indexOf('Customer Name');
    const sellerNameIndex = headers.indexOf('Seller Name');
    const finalDestinationIndex = headers.indexOf('Final Destination');
    const dropOffLocationIndex = headers.indexOf('Drop Off Location');
    const dateOfDropOffIndex = headers.indexOf('Date of Drop Off');
    const expectedTransportIndex = headers.indexOf('Expected Transport date (If Needed)');
    
    if (customerNameIndex !== -1) rowData[customerNameIndex] = entry.customerName || '';
    if (sellerNameIndex !== -1) rowData[sellerNameIndex] = entry.sellerName || '';
    if (finalDestinationIndex !== -1) rowData[finalDestinationIndex] = entry.endLocation || '';
    if (dropOffLocationIndex !== -1) rowData[dropOffLocationIndex] = dropLocation;
    if (dateOfDropOffIndex !== -1) rowData[dateOfDropOffIndex] = today;
    if (expectedTransportIndex !== -1) rowData[expectedTransportIndex] = nextTuesday;
    
    sheet.insertRows(6, 1);
    sheet.getRange(6, 1, 1, rowData.length).setValues([rowData]);
  });
}
