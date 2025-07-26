# Order Tracking System - Agent Guide

## System Overview

This is a comprehensive Google Apps Script-based order management and tracking system built for Google Sheets. The system manages orders from creation through pickup/shipping, handling both marketplace and virtual orders with transport coordination.

## Core Components

### 1. Complete Order Management (`Complete Order.gs`)
**Primary Purpose**: Handles order completion workflow when items are picked up or shipped.

**Key Features**:
- **Manual Checkbox Trigger**: When a user manually checks "Picked Up Or Shiped" in Management Log, shows confirmation dialog
- **Two-Step Confirmation**: 
  1. Confirms pickup/ship action
  2. Asks whether to process single item or all items for that customer
- **Critical Data Flow**: Dashboard → Completed Pick Up/Ship → Remove from Management Log
- **Checkbox Management**: Ensures "Picked Up Or Shiped" is ALWAYS set to TRUE in final records

**Important Functions**:
- `onEdit()`: Monitors Management Log for manual checkbox changes
- `processPickupShip()`: Processes single item pickup/ship
- `processAllCustomerItems()`: Bulk processes all items for a customer
- `updateDashboardPickupStatus()`: Critical function that sets pickup status to TRUE

### 2. Customer Order Analysis (`Same Customer.gs`)
**Primary Purpose**: Analyzes customer orders and determines readiness for pickup based on transport status.

**Key Features**:
- **Customer Grouping**: Counts orders per customer
- **Transport Status Analysis**: Checks if all customer orders are transported
- **Management Log Integration**: Applies overrides from Management Log
- **Auto-Population**: Automatically adds transported items to Management Log

**Important Functions**:
- `mergeDataAndUpdateReadyStatus()`: Main function that processes everything
- `analyzeCustomerOrders()`: Counts and analyzes customer order status
- `updateManagementLogFromDashboard()`: Syncs transported items to Management Log
- `getManagementLogOverrides()`: Gets override flags from Management Log

### 3. Order Transfer System (`Transfer Orders.gs`)
**Primary Purpose**: Transfers new orders from source sheets into the Dashboard system.

**Key Features**:
- **Dual Source Handling**: Processes both Marketplace and Virtual orders
- **Duplicate Prevention**: Checks Dashboard AND Completed sheets for existing orders
- **Smart Column Mapping**: Handles different sheet structures and naming variations
- **Checkbox Setup**: Automatically configures checkboxes based on order type

**Order Types**:
- **Marketplace Orders**: From "4GV Marketplavce (All Automated)" sheet
  - Sets: Created=TRUE, Paid=TRUE, others=FALSE
- **Virtual Orders**: From Kingston/Medford drop-off sheets  
  - Sets: Created=TRUE, Paid=TRUE, Dropped Off=TRUE, others=FALSE

**Important Functions**:
- `transferAllOrdersComplete()`: Main transfer function for both types
- `transferMarketplaceOrdersComplete()`: Handles marketplace orders
- `transferVirtualOrdersComplete()`: Handles virtual orders from drop-off sheets

### 4. Manual Trigger Controller (`Update Manuel Trigger`)
**Primary Purpose**: Central controller providing individual and comprehensive system execution.

**Key Features**:
- **Run All Functions**: Executes complete system in proper sequence
- **Individual Controllers**: Run specific functions independently
- **Error Handling**: Comprehensive error reporting and recovery
- **UI Integration**: Custom menu system for easy access

**Execution Sequence**:
1. Update Dashboard (transfer new orders)
2. Update Drop Offs (merge seller data)
3. Update Transports (manage transport workflow)
4. Update Customers (analyze groupings)
5. Update Ready for Pickup (final status with overrides)

### 5. Marketplace Integration (`Update Marketplace orders.gs`)
**Primary Purpose**: Integrates seller drop-off data with the main Dashboard.

**Key Features**:
- **Multi-Location Support**: Handles Kingston and Medford drop-offs
- **Data Matching**: Matches orders by Order ID + Seller combination
- **Location Assignment**: Updates Start Location based on drop-off location
- **Status Updates**: Marks items as "Dropped Off" when found in drop-off sheets

### 6. Transport Management (`Update Transport.gs`)
**Primary Purpose**: Complete transport workflow management system.

**Key Features**:
- **Transport Need Analysis**: Compares Start Location vs End Location
- **Transport Workflow**: Dashboard → Transport → Completed Transports
- **Status Automation**: Auto-marks transported when no transport needed
- **Completion Processing**: Moves completed transports and updates Dashboard

**Transport Logic**:
- **Same Location**: No transport needed → Mark as Transported=TRUE
- **Different Locations**: Transport needed → Add to Transport sheet
- **Completed Transport**: Move to Completed Transports + Update Dashboard

## Data Flow Architecture

### Primary Sheets Structure

1. **Dashboard** (Main tracking sheet)
   - Headers at Row 12
   - Columns: ID, Seller, Customer, Order Method, Start Location, End Location
   - Checkboxes: Created, Paid, Dropped Off, Transported, Ready For (Pickup or Ship), Picked Up Or Shiped

2. **Management Log** (Pickup/Ship management)
   - Columns: ID, Seller, Customer, Picked Up Or Shiped, Entire Order In Store Overide
   - Auto-populated from Dashboard when items are transported
   - Provides override mechanism for ready-for-pickup status

3. **Transport** (Active transports)
   - Columns: ID, Seller, Customer, Transported
   - Items needing transport between locations

4. **Completed Pick Up / Ship** (Final destination)
   - Final resting place for completed orders
   - Includes completion date

5. **Source Sheets**:
   - **"4GV Marketplavce (All Automated)"**: Marketplace orders
   - **Kingston/Medford Sheets**: Virtual order drop-offs

### Critical Data Relationships

**Order ID + Seller + Customer** = Unique record identifier across all sheets

**Status Progression**:
```
New Order → Dashboard → [Transport if needed] → Management Log → Completed Pick Up/Ship
```

**Checkbox Dependencies**:
- `Dropped Off` → enables transport analysis
- `Transported` → enables ready-for-pickup analysis  
- `Ready For (Pickup or Ship)` → enables Management Log population
- `Picked Up Or Shiped` → triggers completion workflow

## Key System Rules

### 1. Checkbox Management
- All boolean fields MUST use proper checkbox formatting
- "Picked Up Or Shiped" must ALWAYS be TRUE in completed records
- Multiple verification methods ensure checkbox values are set correctly

### 2. Customer Grouping Logic
- Orders are grouped by exact customer name match
- "Ready For (Pickup or Ship)" is TRUE only when ALL customer orders are transported
- Management Log overrides can force individual items to ready status

### 3. Transport Logic
- Start Location = End Location → No transport needed, mark Transported=TRUE
- Start Location ≠ End Location → Add to Transport sheet
- Completed transports update Dashboard Transported status

### 4. Duplicate Prevention
- System checks Dashboard AND Completed sheets before adding new orders
- Uses compound key: "OrderID|Seller|Customer"

### 5. Error Handling
- Graceful handling of missing sheets
- Column name variations supported
- Comprehensive logging for debugging

## Common Operations

### To Process a Single Pickup/Ship:
1. User checks "Picked Up Or Shiped" in Management Log
2. System shows confirmation dialog
3. Choose "NO" for single item processing
4. System moves item through completion workflow

### To Process All Customer Items:
1. User checks any item for a customer in Management Log
2. System shows confirmation dialog  
3. Choose "YES" for bulk processing
4. System processes all items for that customer

### To Run Complete System Update:
```javascript
runAllOrderManagementFunctions()
```

### To Transfer New Orders Only:
```javascript
transferAllOrdersComplete()
```

### To Update Transport System Only:
```javascript
runCompleteTransportSystem()
```

## Sheet Naming Conventions

The system handles various naming patterns:
- "Dashboard" or anything containing "dashboard"
- "Management Log", "Magement Log", "Management", "Magement"
- "Transport" or anything containing "transport"
- "Completed Pick Up / Ship" or variations
- "Kingston" or anything containing "kingston"
- "Medford" or anything containing "medford"

## Debugging and Maintenance

### Debug Functions Available:
- `debugPickedUpStatus()`: Check pickup status across sheets
- `debugTransportSystem()`: Analyze transport system status
- `checkSystemStatus()`: Overall system health check
- `getMergeStatistics()`: Get merge operation statistics

### Utility Functions:
- `setupOrderManagementSystem()`: One-time system setup
- `fixPickedUpCheckboxesInCompletedSheet()`: Fix checkbox formatting
- `setupManagementLogCheckboxes()`: Convert text to checkboxes

### Manual Triggers:
- Custom menu created via `onOpen()` function
- Individual function buttons available
- "Run All" option executes complete sequence

## Critical Success Factors

1. **Proper Sheet Structure**: Headers must be in correct positions
2. **Checkbox Formatting**: All boolean columns must use checkboxes
3. **Data Consistency**: Order ID + Seller + Customer combinations must be unique
4. **Sequence Execution**: Functions should run in proper order for best results
5. **Error Monitoring**: Check logs regularly for any processing issues

## Performance Notes

- System processes hundreds of records efficiently
- Bulk operations preferred over individual record processing
- Checkbox verification includes multiple retry attempts
- Smart caching reduces repeated sheet reads
- Graceful degradation when optional sheets are missing

This system provides comprehensive order lifecycle management from initial order through final pickup/shipping, with robust error handling, duplicate prevention, and status tracking throughout the process.
