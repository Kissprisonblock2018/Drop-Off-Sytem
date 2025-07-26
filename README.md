**4GoodVibes Marketplace Order Management System**

A Google Sheets-based order fulfillment system for managing crafts marketplace orders and facebook virtual event sales across multiple locations.

**Overview:**
This system automates the transparency and tracking to complete order lifecycle from placement to pickup, handling both marketplace orders (handled on our site) and virtual event orders (handled by sellers). It manages multi-location logistics between Kingston and Medford locations with automated process across location and simple and easy to use data collection points for sellers. 

**Features:** 
1. Dual Order Types: Marketplace orders (handled on our site) and virtual event orders (handled by sellers)
2. Multi-Location Management: Automated transport coordination between Kingston and Medford
3. Seller Drop-Off System: Web form for sellers to register item that have been dropped off in a simple and easy to use form
4. Automatic data filling: Automatically fills in data across all different parts of the process to ensure the most information with the least amount of entries
5. Management Overrides: Manual controls for special cases and simple check boxes

6. Automated Workflows: Complete order lifecycle automation with Google Apps Script

**Project Structure:
**Data/ Directory
Contains CSV files representing the current state of your Google Sheets:

Drop Off From/: Seller drop-off logs and pickup order data
Tracking System/: Main dashboard, transport logs, and completed orders

These files show the expected sheet structure and sample data format

**Src/ Directory**
Contains all Google Apps Script code:

Document 2
Drop Off From/: Seller Entires / Pick Up Orders

Doccument 1
Tracking System/: Dashboard / Management Log / Transport / Seller Drop Off Kingston / Seller Drop Off Medford / 4GV Marketplavce (All Automated) / Completed Transports / Completed Pick Up / Ship


**Order Lifecycle:**
Order Processing
Marketplace orders automatically sync from scraped data and filled into the dashboard 
Virtual orders are added when sellers drop off items


Seller Drop-Offs
Sellers use the web form to register deliveries
Choose order type: Marketplace (select existing order) or Virtual (new entry)
System automatically updates order status and location


Transport Management
System identifies items needing transport between locations
Mark items as transported when moved
Dashboard automatically updates pickup readiness


Customer Notifications
System checks if all customer orders are ready
Management can override for early pickup approval
Ready customers can be contacted for pickup


Order Completion
Mark orders as picked up through Management Log
Items automatically move to completed orders
System maintains complete audit trail


**Key Functions for Pick Up / Ship System**
Run these from the Apps Script editor or set up as triggers:
javascript// Complete system update (run daily)

runAllOrderManagementFunctions()
// Individual components
runUpdateDashboard()        // Adds New Marketplace orders
runUpdateDropOffs()         // Sync new Drop Offs
runUpdateTransports()       // Manage logistics  
runUpdateReadyForPickup()   // Check customer readiness to see who can be notified 

Customer


