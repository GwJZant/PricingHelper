PricingHelper
Created by: David Barnes
Requirements: ImportExcel, SqlServer

*******************************************************************************
* To use this tool you need to install the ImportExcel and SqlServer modules. *
* All you need to do is run the below commands in Powershell.                 *
* It does not require Administrator credentials.                              *
*                                                                             *
* Install-Module -Name ImportExcel -Scope CurrentUser -Force                  *     
* Install-Module -Name SqlServer -Scope CurrentUser -Force                    *
*******************************************************************************

This is a tool meant to save time setting up the spreadsheets we use during the pricing update process.

Normally we need to do a search in Celerant, export that search to a .csv, copy that .csv's data into a Google Sheet, then delete unnecessary columns. After that, we need to run a report, export that to a .csv, convert datatype in the barcode column to Number, remove decimals, then copy that into the Google Sheet on a separate sheet. Then we need to make a third sheet formatted for the Celerant product import which contains info copied from Sheet1. Then we make a Sheet4 with more data copied from Sheet1 and add some columns to capture inventory by store.

With this tool, the aim is to generate an Excel document with 4 sheets already built out that we just need to import into Google Sheets. The only data that the user should have to input is information that requires manual lookups like an updated price. The rest of the spreadsheet is done the exact same way every single time so we're wasting our time doing that process manually over and over.

For deciding which cost values to use when there are multiple vendors, the tool creates a formula that looks at all the values and picks the minimum price. If that behavior is not desired, you can copy and paste the value you want into that cell or you can change the formula then drag it over any/all columns you want.

This will save time at multiple steps of the manual process:
	Obtaining Sheet1 product-level data from Celerant search
	Opening the product page for each distinct product to find the old cost per vendor
	Running Export for min-max adjustment report to fill out Sheet2
	Hooking up VLOOKUP on Sheet1 to get Primary Barcode from Sheet2
	Copying data from Sheet1 to Sheet3
	Copying data from Sheet1 to Sheet4
	Hooking up VLOOKUP on Sheet4 for each store to get inventory from Sheet2

The following functionality exists:
	Sheet1 (Main sheet for recording part costs, calculating new margins and new retail values, calculating retail price differences)
	Sheet2 (Barcode lookup, TECHNICALLY not necessary anymore since Barcode Lookup is filled in automatically but leaving it in so we can still check our work)
	Sheet3 (Product import, should come together quickly since nothing is dynamically generated)
	Sheet4 (Final price change page that is used for making the document managers will use)

TODO:
	Add Part Numbers to a column at the end of Sheet1.

HOW TO RUN:
Double-click PricingHelper.bat

HOW TO INSTALL MODULES:
Open PowerShell by searching for it with the Start menu or shift + right-click the whitespace of any folder's window and click "Open PowerShell window here"
Copy and paste the Install-Module commands into PowerShell. If the modules are already installed, they will be ignored.