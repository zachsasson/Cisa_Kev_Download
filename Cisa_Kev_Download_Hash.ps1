# Install ImportExcel module if not installed
Install-Module -Name ImportExcel -Force -SkipPublisherCheck -RequiredVersion 7.8.6

# Import Excel module
Import-Module ImportExcel

if (Test-Path $hashFilePath) {
    #Read previous hash from file
    $previousHash = Get-Content -Path $hashFilePath
}
else {
    #If file doesn't exist, create new value for $previousHash
    $previousHash = ""
}

# Grab CSV from cisa.gov kev list and store in variable
$infile = (Invoke-WebRequest -Uri "https://www.cisa.gov/sites/default/files/csv/known_exploited_vulnerabilities.csv").Content | ConvertFrom-Csv

$hashFilePath = "C:\Scripts\PowerShell\Cisa_Kev_Download\hash.txt"

$hasher = [System.Security.Cryptography.SHA256]::Create()

$hashbytes = $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($infile))

$currentHash = [System.BitConverter]::ToString($hashbytes) -replace '-'

Write-Host "Hash of the content is: $hash"

if ($currentHash -eq $previousHash) {
    Write-Host "No Changes Detected, Exiting."
}
else {
    Write-Host "Changes detected. Proceeding with rest of script."

    # Create a new Excel package
    $excelPackage = New-Object OfficeOpenXml.ExcelPackage

    # Create tabs for each unique vendor project
    foreach ($category in $infile | Select-Object -Property vendorProject -Unique) {
        try {
            if ($category.vendorProject -and $category.vendorProject -ne "") {
                # Check if worksheet with the same name already exists
                $worksheetName = $category.vendorProject
                $worksheetIndex = 1
                while ($excelPackage.Workbook.Worksheets[$worksheetName]) {
                    $worksheetName = "$($category.vendorProject)_$worksheetIndex"
                    $worksheetIndex++
                }

                # Filter data based on the current category
                $filteredData = $infile | Where-Object { $_.vendorProject -eq $category.vendorProject }

                if ($filteredData) {
                    # Sort the filtered data by dateAdded in descending order
                    $filteredData = $filteredData | Sort-Object -Property { [datetime]$_."dateAdded" } -Descending

                    # Create a new worksheet for each category
                    $worksheet = $excelPackage.Workbook.Worksheets.Add($worksheetName)

                    # Set headers
                    $headers = $filteredData[0].PSObject.Properties.Name
                    for ($col = 1; $col -le $headers.Count; $col++) {
                        $worksheet.Cells[1, $col].Value = $headers[$col - 1]
                    }

                    # Populate data
                    for ($row = 2; $row -le $filteredData.Count + 1; $row++) {
                        $rowData = $filteredData[$row - 2].PSObject.Properties.Value
                        for ($col = 1; $col -le $rowData.Count; $col++) {
                            $worksheet.Cells[$row, $col].Value = $rowData[$col - 1]
                        }
                    }
                }
                else {
                    Write-Host "No data found for category '$($category.vendorProject)'."
                }
            }
            else {
                Write-Host "Skipping category with an empty name."
            }
        }
        catch {
            Write-Host "Error setting worksheet name for '$($category.vendorProject)': $_"
        }
    }

    # Save the workbook
    $outFile = "C:\Scripts\PowerShell\Cisa_Kev_Download\filteredKev.xlsx"
    $excelPackage.SaveAs($outFile)

    #Save current hash to file for future reference
    $currentHash | Set-Content -Path $hashFilePath

}


