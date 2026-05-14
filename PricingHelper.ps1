# Load the configuration file
$configFile = "$PSScriptRoot\config.json"

# Load SQL file
$sqlFilePath = "$PSScriptRoot\Queries\GetProductsByBrand.sql"

# Output filename formatting
# Check if the folder exists; if not, create it
if (-not (Test-Path -Path "$PSScriptRoot\Output")) {
    # -Force ensures it creates the entire tree if multiple levels are missing
    New-Item -ItemType Directory -Path "$PSScriptRoot\Output" -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$excelPath = "$PSScriptRoot\Output\Pricing_$timestamp.xlsx"

# Hash objects used for data storage
$allProducts = @{} # A has table to store products
$productValues = @{}
$productVendorPartNums = @{}
$allVendors = @() # A list to store all vendor shortnames

# Load the config file and ensure it exists
if (Test-Path $configFile) {
    # Read the file and convert JSON into a PowerShell Object
    $config = Get-Content -Path $configFile -Raw | ConvertFrom-Json
} else {
    Write-Host "ERROR: config.json not found!" -ForegroundColor Red
    exit
}

# Function to format columns
function Format-PricingSheet {
	param($Worksheet)
	
	# Get the number of columns
	$maxCol = $Worksheet.Dimension.Columns
	
	for ($i = 1; $i -le $maxCol; $i++) {
		$headerValue = $Worksheet.Cells[1, $i].Value
		
		switch -wildcard ($headerValue.ToString()) {
			"*Margin*" {
				$Worksheet.Column($i).Style.Numberformat.Format = "0.00%"
				break
			}
			{$_ -like "*Retail*" -or $_ -like "*Old*" -or $_ -like "*New*" -or $_ -like "*Cost*" -or $_ -like "*Price*"} {
				$Worksheet.Column($i).Style.Numberformat.Format = "$#,##0.00"
				break
			}
		}
	}
}

# Menu
Write-Host "How do you want to collect products?" -ForegroundColor Cyan
Write-Host "1. By Brand (Pet Supplies , Apparel, + Footwear)" -ForegroundColor Cyan
Write-Host "2. By Vendor (Accessories)" -ForegroundColor Cyan
Write-Host "3. Exit" -ForegroundColor Cyan
$selection = Read-Host "Enter your selection: "

if ($selection -eq "1") {
	$inputBrand = Read-Host "Enter Brand (As it appears in Celerant)"

	$sqlVars = "Brand='$inputBrand'"
} elseif ($selection -eq "2") {
	$inputVendor = Read-Host "Enter Vendor Name (As it appears in Celerant)"

	$sqlVars = "Vendor='$inputVendor'"
	$sqlFilePath = "$PSScriptRoot\Queries\GetProductsByVendor.sql"
} else {
	exit
}

# Connect to Celerant Database
Write-Host "Connecting to Celerant Database..." -ForegroundColor Cyan

# Set up SQL parameters
$sqlParams = @{
    ServerInstance         = $config.Celerant.ServerInstance
    Database               = $config.Celerant.Database
	InputFile              = $sqlFilePath
	Variable               = $sqlVars
	Username               = $config.Celerant.Username
	Password               = $config.Celerant.Password
	Encrypt                = "Mandatory"
	TrustServerCertificate = $true
}

try {
    $productData = Invoke-Sqlcmd @sqlParams
	
	foreach ($row in $productData) {
		$style        = $row.Style
		$storeId      = $row.Store
		
		# Creating a hash table for new products
		if (-not $allProducts.ContainsKey($style)) {
			$allProducts[$style] = @{}
		}
		
		# Look at the vendors for this product
		$thisStylesVendors = $row | Select-Object -ExpandProperty Vendor -Unique
		
		# For each of the vendors for this product, add them to our master list of vendors if not already present
		foreach ($v in $thisStylesVendors) {
			if ($v -notin $allVendors) {
				$allVendors += $v
			}
		}
		
		# Write data for all 4 stores to hash table
		$allProducts[$style][$storeId] = $row
	}
	
	$rowCounter = 2
	
	$reportData = foreach ($product in $allProducts.Keys) {
		# This is a table of all the product data for each of the 4 stores
		$storeProducts = $allProducts[$product]
		$currRow = $rowCounter
		
		# Total available
		$totalAvailable = 0
		foreach ($storeRecord in $storeProducts.Values) {
			$totalAvailable += $storeRecord.Available
		}
		
		$columnData = [ordered]@{
			Brand                     = $storeProducts[1].Brand
			Product                   = $product
			"Barcode Lookup"          = [string]$storeProducts[1].Barcode_Lookup + " "
			UPC                       = [string]$storeProducts[1].UPC + " "
			Cost                      = 0
			Price                     = 0
		}
		
		$productVendorComment = ""
		
		foreach ($vendor in $allVendors) {
			$vendorRecord = $storeProducts.Values | Where-Object { $_.Vendor -eq $vendor} | Select-Object -First 1
			
			if ($vendorRecord) {
				$columnData["$vendor Old"] = $vendorRecord.Part_Cost
			}
			else {
				$columnData["$vendor Old"] = ""
			}
			
			if ($vendorRecord) {
				$productVendorComment += $vendor + ": " + $vendorRecord.Part_Num + "`n"
			}
		
			$columnData["$vendor New"] = ""
		}
		
		if (-not $productVendorPartNums.ContainsKey($product)) {
			$productVendorPartNums[$product] = @{
				Comment = $productVendorComment
				Row = $currRow
			}
		}
		
		# Gather some column locations for later
		$keys = $columnData.Keys | ForEach-Object { $_ }
		$vendorOldLetters = foreach ($vendor in $allVendors) {
			$index = [array]::IndexOf($keys, "$vendor Old") + 1
			[char](64 + $index)
		}
		
		$vendorNewLetters = foreach ($vendor in $allVendors) {
			$index = [array]::IndexOf($keys, "$vendor New") + 1
			[char](64 + $index)
		}
		
		$minCostRef = ($allVendors | ForEach-Object { "[@[$_ Old]]" }) -join ","
		
		$columnData["Old Retail"]              = $storeProducts[1].Old_Retail
		
		# Grab Old Retail location
		$keys = $columnData.Keys | ForEach-Object { $_ }
		$oldRetailIndex = [array]::IndexOf($keys, "Old Retail") + 1
		$oldRetailLetter = [char](64 + $oldRetailIndex)
		
		# Old Margin
		$oldCostsMinRange = "MIN(" + (($vendorOldLetters | ForEach-Object { "$_$currRow" }) -join ",") + ")"
		$newCostsMinRange = "MIN(" + (($vendorNewLetters | ForEach-Object { "$_$currRow" }) -join ",") + ")"
		$columnData["Old Margin"]              = "=($($oldRetailLetter)$currRow - $oldCostsMinRange) / $($oldRetailLetter)$currRow"
		
		# Update the keys
		$keys = $columnData.Keys | ForEach-Object { $_ }
		$oldMarginIndex = [array]::IndexOf($keys, "Old Margin") + 1
		$oldMarginLetter = [char](64 + $oldMarginIndex)
		$columnData["Auto Retail"]             = "=$newCostsMinRange / (1 - $($oldMarginLetter)$currRow)"
		
		$keys = $columnData.Keys | ForEach-Object { $_ }
		$autoRetailIndex = [array]::IndexOf($keys, "Auto Retail") + 1
		$autoRetailLetter = [char](64 + $autoRetailIndex)
		
		# Round to the ten-cent by default
		$retailCentsOffset = 1
		if ($storeProducts[1].Last_Price -ge 12) {
			$retailCentsOffset = 0 #round to dollar if cost is over 12
		}
		
		$columnData["New Retail"]              = "=(ROUNDUP($($autoRetailLetter)$currRow, $retailCentsOffset))-0.01"
		
		# Grab New Retail location
		$keys = $columnData.Keys | ForEach-Object { $_ }
		$newRetailIndex = [array]::IndexOf($keys, "New Retail") + 1
		$newRetailLetter = [char](64 + $newRetailIndex)
		
		$columnData["New Margin"]              = "=($($newRetailLetter)$currRow - $newCostsMinRange) / $($newRetailLetter)$currRow"
		$columnData["New Retail - Old Retail"] = "=$($newRetailLetter)$currRow - $($oldRetailLetter)$currRow"
		$columnData["Description 1"]           = $storeProducts[1].Description_1
		$columnData["Description 2"]           = $storeProducts[1].Description_2
		$columnData["Department"]              = $storeProducts[1].Department
		$columnData["Type"]                    = $storeProducts[1].Type
		$columnData["Season"]                  = $storeProducts[1].Season
		$columnData["Promo"]                   = $storeProducts[1].Promo
		$columnData["Available"]               = $totalAvailable
		$columnData["Last Cost"]               = $storeProducts[1].Last_Cost
		$columnData["Last Price"]              = $storeProducts[1].Last_Price
		
		# Update some prior values
		$columnData["Cost"]                    = "=$newCostsMinRange"
		$columnData["Price"]                   = "=$($newRetailLetter)$currRow"
		
		# Keep track of values needed for Sheets 3 and 4
		if (-not $productValues.ContainsKey($product)) {
			$productValues[$product] = @{
				Cost = "=Sheet1!E$currRow"
				Price = "=Sheet1!F$currRow"
				OldRetail = "=Sheet1!$($oldRetailLetter)$currRow"
				NewRetail = "=Sheet1!$($newRetailLetter)$currRow"
				NewMinusOldRetail = "=Sheet1!$($newRetailLetter)$currRow - Sheet1!$($oldRetailLetter)$currRow"
			}
		}
		
		$rowCounter++
		
		# Create a custom object for your Excel row
		[PSCustomObject]$columnData
	}
	
	$reportData2 = foreach ($product in $allProducts.Keys) {
		# This is a table of all the product data for each of the 4 stores
		$storeProducts = $allProducts[$product]
		
		foreach ($storeRecord in $storeProducts.Values) {
			
			$columnData2 = [ordered]@{
				StoreProduct              = [string]$storeRecord.Store + $product
				"Inventory Store"         = $storeRecord.Store
				Brand                     = $storeRecord.Brand
				PRODUCT                   = $product
				"Primary Barcode"         = [string]$storeRecord.Barcode_Lookup + " "
				"On Hand"                 = $storeRecord.Available
			}
			
			# Create a custom object for your Excel row
			[PSCustomObject]$columnData2
		}		
	}
	
	$reportData3 = foreach ($product in $allProducts.Keys) {
		# This is a table of all the product data for each of the 4 stores
		$storeProducts = $allProducts[$product]
		
		$columnData3 = [ordered]@{
			Brand                     = $storeProducts[1].Brand
			Style                     = $($product) + [char]8203
			"Barcode Lookup"          = $storeProducts[1].Barcode_Lookup
			Cost                      = $productValues[$product].Cost
			Price                     = $productValues[$product].Price
			Description1              = $storeProducts[1].Description_1
		}
		
		# Create a custom object for your Excel row
		[PSCustomObject]$columnData3	
	}
	
	$reportData4 = foreach ($product in $allProducts.Keys) {
		# This is a table of all the product data for each of the 4 stores
		$storeProducts = $allProducts[$product]
		
		$columnData4 = [ordered]@{
			Brand                     = $storeProducts[1].Brand
			Style                     = $($product) + [char]8203
			Allentown                 = $storeProducts[1].Available
			Saucon                    = $storeProducts[2].Available
			Forks                     = $storeProducts[3].Available
			Trex                      = $storeProducts[4].Available
			"Old Retail"              = $productValues[$product].OldRetail
			"New Retail"              = $productValues[$product].NewRetail
			"New Retail - Old Retail" = $productValues[$product].NewMinusOldRetail
			"Description 1"           = $storeProducts[1].Description_1
			Department                = $storeProducts[1].Department
			Type                      = $storeProducts[1].Type
			SEASON                    = $storeProducts[1].Season
			PROMO                     = $storeProducts[1].Promo
		}
		
		# Create a custom object for your Excel row
		[PSCustomObject]$columnData4
	}
	
	# Generate the Excel File
	$excelPackage = $reportData | Export-Excel -Path $excelPath -WorksheetName "Sheet1" -TableStyle Medium2 -TableName "PricingTable" -AutoSize -BoldTopRow -PassThru
	Format-PricingSheet -Worksheet $excelPackage.Workbook.Worksheets["Sheet1"]
	
	foreach ($item in $productVendorPartNums.Values) {
		Write-Host $item.Row
		$cell = $excelPackage.Workbook.Worksheets["Sheet1"].Cells["B$($item.Row)"]
		
		
		# Check if a comment already exists, then add/set it
		if ($null -eq $cell.Comment) {
			$cell.AddComment("$($item.Comment)", "System")
		}
	}
	
	# Format Column B (Product) as Text so text alignment matches
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet1"] -Range "B:B" -NumberFormat "@"
	
	# Format Column C (Barcode Lookup) as Text so the E+11 goes away
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet1"] -Range "C:C" -NumberFormat "@"
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet1"] -Range "D:D" -NumberFormat "@"
	
	$excelPackage = $reportData2 | Sort-Object -Property "Inventory Store", PRODUCT | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Sheet2" -TableStyle Medium9 -AutoSize -BoldTopRow -PassThru
	Format-PricingSheet -Worksheet $excelPackage.Workbook.Worksheets["Sheet2"]
	
	$excelPackage = $reportData3 | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Sheet3" -TableStyle Medium9 -AutoSize -BoldTopRow -PassThru
	Format-PricingSheet -Worksheet $excelPackage.Workbook.Worksheets["Sheet3"]
	
	# Format Column B (Product) as Text so text alignment matches
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet3"] -Range "B:B" -NumberFormat "@"
	
	# Format Column C (Barcode Lookup) as Text so the E+11 goes away
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet3"] -Range "C:C" -NumberFormat "@"
	
	$excelPackage = $reportData4 | Export-Excel -ExcelPackage $excelPackage -WorksheetName "Sheet4" -TableStyle Medium9 -AutoSize -BoldTopRow -PassThru
	Format-PricingSheet -Worksheet $excelPackage.Workbook.Worksheets["Sheet4"]
	
	# Format Column B (Product) as Text so text alignment matches
	Set-ExcelRange -Worksheet $excelPackage.Workbook.Worksheets["Sheet4"] -Range "B:B" -NumberFormat "@"
	
	# Save and Close
	Close-ExcelPackage $excelPackage
		
} catch {
    Write-Host "Error: $_" -ForegroundColor DarkRed
}
