PricingHelper

This is a tool meant to save time setting up the spreadsheets we use during the pricing update process.

Normally we need to do a search in Celerant, export that search to a .csv, copy that .csv's data into a Google Sheet, then delete unnecessary columns. After that, we need to run a report, export that to a .csv, convert datatype in the barcode column to Number, remove decimals, then copy that into the Google Sheet on a separate sheet. Then we need to make a third sheet formatted for the Celerant product import which contains info copied from Sheet1. Then we make a Sheet4 with more data copied from Sheet4 and add some columns to capture inventory by store.

With this tool, the aim is to generate an Excel document with 4 sheets already built out that we just need to import into Google Sheets. The only data that the user should have to input is information that requires manual lookups like an updated price. The rest of the spreadsheet is done the exact same way every single time so we're wasting our time doing that process manually over and over.

In a perfect world, the user should only need to fill out new "Cost" values on Sheet1 and the entire rest of the document should be filled out.

WIP