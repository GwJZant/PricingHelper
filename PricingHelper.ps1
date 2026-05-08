# Load the configuration file
$configFile = "$PSScriptRoot\config.json"

# Load SQL file
$sqlFilePath = "$PSScriptRoot\Queries\GetProducts.sql"

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
$allVendors = @() # A list to store all vendor shortnames

# Load the config file and ensure it exists
if (Test-Path $configFile) {
    # Read the file and convert JSON into a PowerShell Object
    $config = Get-Content -Path $configFile -Raw | ConvertFrom-Json
} else {
    Write-Host "ERROR: config.json not found!" -ForegroundColor Red
    exit
}

# Connect to Celerant Database
Write-Host "Connecting to Celerant Database..." -ForegroundColor Cyan

# Ask for Username and Password in the console
$passSecure = Read-Host "Enter $($config.Celerant.Username) SQL Password" -AsSecureString

# Convert to plain text for the SQL Driver
$passPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passSecure)
)

# Ask for Username and Password in the console
$inputBrand = Read-Host "Enter Brand (As it appears in Celerant)"

$sqlVars = "Brand=$inputBrand"

# Set up SQL parameters
$sqlParams = @{
    ServerInstance         = $config.Celerant.ServerInstance
    Database               = $config.Celerant.Database
	InputFile              = $sqlFilePath
	Variable               = $sqlVars
	Username               = $config.Celerant.Username
	Password               = $passPlain
	Encrypt                = "Mandatory"
	TrustServerCertificate = $true
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
	
	# Format Column B (Product) as Text so text alignment matches
	Set-ExcelRange -Worksheet $Worksheet -Range "B:B" -NumberFormat "@"
	
	# Format Column C (Barcode Lookup) as Text so the E+11 goes away
	Set-ExcelRange -Worksheet $Worksheet -Range "C:C" -NumberFormat "@"
}

try {
    $productData = Invoke-Sqlcmd @sqlParams
	
	foreach ($row in $productData) {
		$style        = $row.Style
		$storeId      = $row.Store
		
		# Creating a has table for new products
		if (-not $allProducts.ContainsKey($style)) {
			$allProducts[$style] = @{}
		}
		
		# Write data for all 4 stores to hash table
		$allProducts[$style][$storeId] = $row
	}
	
	$rowCounter = 2
	
	$reportData = foreach ($product in $allProducts.Keys) {
		# This is a table of all the product data for each of the 4 stores
		$storeProducts = $allProducts[$product]
		
		# Look at the vendors for this product
		$thisStylesVendors = $storeProducts.Values | Select-Object -ExpandProperty Vendor -Unique
		
		# For each of the vendors for this product, add them to our master list of vendors if not already present
		foreach ($v in $thisStylesVendors) {
			if ($v -notin $allVendors) {
				$allVendors += $v
			}
		}
		
		# Write-Host "Brand: $($storeProducts[1].Brand) Style: $($storeProducts[1].Style) Primary Barcode: $($storeProducts[1].Barcode_Lookup)"
		
		# Total available
		$totalAvailable = 0
		foreach ($storeRecord in $storeProducts.Values) {
			$totalAvailable += $storeRecord.Available
		}
		
		$columnData = [ordered]@{
			Brand                     = $storeProducts[1].Brand
			Product                   = $product
			"Barcode Lookup"          = [string]$storeProducts[1].Barcode_Lookup + " "
			Cost                      = 0
			Price                     = 0
		}
		
		foreach ($vendor in $allVendors) {
			$vendorRecord = $storeProducts.Values | Where-Object { $_.Vendor -eq $vendor} | Select-Object -First 1
			
			if ($vendorRecord) {
				$columnData["$vendor Old"] = $vendorRecord.Part_Cost
			}
			else {
				$columnData["$vendor Old"] = ""
			}
			
			$columnData["$vendor New"] = ""
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
		$currRow = $rowCounter
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
		if ($storeProducts[1].Last_Cost -ge 12) {
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
		
		$rowCounter++
		
		# Create a custom object for your Excel row
		[PSCustomObject]$columnData
	}
	
	# Generate the Excel File
	$excelPackage = $reportData | Export-Excel -Path $excelPath -WorksheetName "Sheet One" -TableStyle Medium2 -TableName "PricingTable" -AutoSize -BoldTopRow -PassThru
	Format-PricingSheet -Worksheet $excelPackage.Workbook.Worksheets["Sheet One"]
	
	# Save and Close
	Close-ExcelPackage $excelPackage
		
} catch {
    Write-Host "SQL Error: $_" -ForegroundColor DarkRed
}
