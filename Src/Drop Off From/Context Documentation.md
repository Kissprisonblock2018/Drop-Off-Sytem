# Context Documentation - Order Management System

## Project Overview
This is a Google Apps Script-based Order Management System that provides a web interface for managing marketplace and virtual marketplace orders. The system consists of three main components that work together to handle order processing, data synchronization, and user interface interactions.

## File Structure

### 1. `Form.gs` - Main Form Logic
**Purpose**: Core Google Apps Script backend that handles form operations and data management.

**Key Functions**:
- `onEdit(e)` - Event handler for spreadsheet edits
- `showForm()` - Displays the HTML form popup (800x1000px modal dialog)
- `getTodayAndNextTuesday()` - Calculates dates and returns drop location "Kingston"
- `getSellerOrderData()` - Fetches seller and order data from 'Pick Up Orders' sheet
- `getOrdersBySeller(sellerName)` - Filters orders by specific seller
- `saveMarketplaceOrder(data)` - Saves basic marketplace orders (legacy function)
- `saveMarketplaceOrderWithSeller(data)` - Enhanced function that saves marketplace orders with seller/buyer info
- `saveVirtualOrder(data)` - Saves virtual marketplace orders with customer/seller/destination data

**Data Flow**:
- Reads from 'Pick Up Orders' sheet (columns: Seller, Order ID, Buyer)
- Writes to active sheet starting at row 6
- Maps data to columns: Order ID, Seller Name, Customer Name, Drop Off Location, Date of Drop Off, Expected Transport date

### 2. `FormPopup.html` - User Interface
**Purpose**: Modern, responsive HTML form with JavaScript for order entry.

**Design Features**:
- Lato font family with gradient backgrounds
- Green (#58b73a) and black color scheme
- Responsive design with mobile support
- Animated interactions (fade-in/fade-out effects)
- Modal popup interface (800x1000px)

**Form Types**:
1. **Marketplace Orders**:
   - Seller selection dropdown (populated from backend)
   - Order selection dropdown (filtered by seller)
   - Auto-populated fields: Order ID, Seller Name, Customer/Buyer Name
   - Add/remove multiple entries

2. **Virtual Marketplace Orders**:
   - Manual entry fields: Customer Name, Seller Name
   - Final Destination dropdown: Medford, Kingston, Shipped
   - Add/remove multiple entries

**JavaScript Functions**:
- `showMarketplace()` / `showVirtual()` - Switch between form types
- `updateOrdersBySelectedSeller()` - Dynamic filtering of orders by seller
- `submitMarketplace()` / `submitVirtual()` - Form submission with validation
- Dynamic field management (add/remove entries)

### 3. `OpenOrders.gs` - Data Synchronization
**Purpose**: Automated sync system that pulls marketplace orders from external Dashboard and populates the Pick Up Orders sheet.

**Configuration**:
```javascript
const CONFIG = {
  DASHBOARD_SHEET_URL: 'https://docs.google.com/spreadsheets/d/1MWCzWUDfjKr2mGlQcIgKF0uqFoRZsg8QUiGlGiCmFaM/',
  DASHBOARD_SHEET_NAME: 'Dashboard',
  PICKUP_ORDERS_SHEET_NAME: 'Pick Up Orders',
  HEADER_ROW: 12, // Headers start at row 12
}
```

**Key Features**:
- Runs every 4 hours (configurable trigger)
- Filters orders where: Order Method = "Marketplace" AND Dropped Off = false
- Prevents duplicate entries
- Maps Dashboard columns to Pick Up Orders columns:
  - ID → Order ID
  - Seller → Seller  
  - Customer → Buyer
  - End Location → Pickup Location

**Main Functions**:
- `syncMarketplaceOrders()` - Main sync function
- `getDashboardData()` - Fetches data from external Dashboard sheet
- `filterOrders()` - Applies business logic filters
- `setupAutoTrigger()` - Creates 4-hour recurring trigger
- `testConfiguration()` - Debugging and validation function

## Data Schema

### Pick Up Orders Sheet Structure
| Column | Purpose | Source |
|--------|---------|---------|
| Seller | Vendor name | Dashboard.Seller |
| Order ID | Unique identifier | Dashboard.ID |
| Buyer | Customer name | Dashboard.Customer |
| Pickup Location | End destination | Dashboard."End Location" |
| Last Synced | Sync timestamp | Auto-generated |

### Form Output Sheet Structure
| Column | Purpose | Form Type |
|--------|---------|-----------|
| Order ID (if Market Place) | Marketplace order ID | Marketplace |
| Seller Name | Vendor name | Both |
| Customer Name | Buyer name | Both |
| Drop Off Location | Always "Kingston" | Both |
| Date of Drop Off | Current date | Both |
| Expected Transport date | Next Tuesday | Both |
| Final Destination | End location | Virtual only |

## Workflow

### 1. Data Sync Process
1. **Automated Sync**: `OpenOrders.gs` runs every 4 hours
2. **Data Fetching**: Connects to external Dashboard sheet
3. **Filtering**: Selects marketplace orders not yet dropped off
4. **Deduplication**: Checks against existing Pick Up Orders
5. **Population**: Adds new orders to Pick Up Orders sheet

### 2. Order Entry Process
1. **Form Launch**: User triggers `showForm()` from spreadsheet
2. **Data Loading**: Form loads seller/order data from Pick Up Orders sheet
3. **Order Selection**: User selects orders via dropdown or manual entry
4. **Validation**: JavaScript validates required fields
5. **Submission**: Data saved to active sheet starting at row 6

### 3. Integration Points
- **Pick Up Orders Sheet**: Central data store populated by sync, read by form
- **Active Sheet**: Destination for form submissions
- **External Dashboard**: Source of truth for marketplace orders

## Technical Notes

### Error Handling
- Comprehensive logging in sync operations
- Graceful degradation when seller data unavailable
- Validation prevents empty submissions

### Performance Considerations
- Efficient filtering to avoid duplicate processing
- Batch operations for multiple order entries
- Minimal API calls through caching

### Maintenance Functions
- `testConfiguration()` - Validates connectivity and data structure
- `setupAutoTrigger()` / `removeAutoTrigger()` - Manages automation
- Console logging throughout for debugging

## Dependencies
- Google Apps Script runtime
- Google Sheets API (implicit)
- External Dashboard sheet access permissions
- HTML Service for popup interface

## Security Notes
- Requires access to external Google Sheet
- Uses Apps Script's built-in authentication
- Form runs in sandboxed HTML environment
- Data validation on both client and server sides

## Future Enhancement Areas
- Add order status tracking
- Implement bulk order operations
- Add email notifications for sync failures
- Enhance mobile responsiveness
- Add order search/filtering capabilities
