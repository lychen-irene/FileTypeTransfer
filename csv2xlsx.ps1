param(
    [Parameter(Mandatory=$false)]
    [switch]$DeleteOriginal
)


$SourcePath = Read-Host "Enter the source path (or press Enter for current directory)"
if ([string]::IsNullOrWhiteSpace($SourcePath)) { $SourcePath = "." }

# Set destination path to source path if not specified
$DestinationPath = Read-Host "Enter the destination path (or press Enter for current directory)"
if ([string]::IsNullOrWhiteSpace($DestinationPath)) { $DestinationPath = $SourcePath }



# Function to convert CSV to XLSX
function Convert-CsvToXlsx {
    param(
        [string]$CsvPath,
        [string]$XlsxPath
    )
    
    try {
        # Create Excel application object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Open the CSV file
        $workbook = $excel.Workbooks.Open($CsvPath)
        
        # Save as XLSX format (xlOpenXMLWorkbook = 51)
        $workbook.SaveAs($XlsxPath, 51)
        
        # Close workbook and quit Excel
        $workbook.Close()
        $excel.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "✓ Converted: $CsvPath -> $XlsxPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Error converting $CsvPath : $($_.Exception.Message)" -ForegroundColor Red
        
        # Clean up Excel process if it's still running
        if ($excel) {
            try {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            catch { }
        }
        return $false
    }
}

# Main script execution
Write-Host "CSV to XLSX Converter" -ForegroundColor Cyan
Write-Host "Source Path: $SourcePath" -ForegroundColor Yellow
Write-Host "Destination Path: $DestinationPath" -ForegroundColor Yellow
Write-Host ""

# Check if source path exists
if (-not (Test-Path $SourcePath)) {
    Write-Host "Error: Source path '$SourcePath' does not exist!" -ForegroundColor Red
    exit 1
}

# Create destination directory if it doesn't exist
if (-not (Test-Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    Write-Host "Created destination directory: $DestinationPath" -ForegroundColor Yellow
}

# Get all CSV files from source path
$csvFiles = Get-ChildItem -Path $SourcePath -Filter "*.csv" -File

if ($csvFiles.Count -eq 0) {
    Write-Host "No CSV files found in '$SourcePath'" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($csvFiles.Count) CSV file(s) to convert..." -ForegroundColor Cyan
Write-Host ""

# Convert each CSV file
$successCount = 0
$failCount = 0

foreach ($csvFile in $csvFiles) {
    $csvFullPath = $csvFile.FullName
    $xlsxFileName = [System.IO.Path]::ChangeExtension($csvFile.Name, ".xlsx")
    $xlsxFullPath = Join-Path $DestinationPath $xlsxFileName
    
    # Skip if XLSX file already exists (optional - remove this check if you want to overwrite)
    if (Test-Path $xlsxFullPath) {
        Write-Host "⚠ Skipping $($csvFile.Name) - XLSX already exists" -ForegroundColor Yellow
        continue
    }
    
    # Convert the file
    if (Convert-CsvToXlsx -CsvPath $csvFullPath -XlsxPath $xlsxFullPath) {
        $successCount++
        
        # Delete original CSV file if requested
        if ($DeleteOriginal) {
            try {
                Remove-Item $csvFullPath -Force
                Write-Host "  → Deleted original CSV file" -ForegroundColor Gray
            }
            catch {
                Write-Host "  → Warning: Could not delete original CSV file: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
    else {
        $failCount++
    }
}

# Summary
Write-Host ""
Write-Host "Conversion Summary:" -ForegroundColor Cyan
Write-Host "✓ Successfully converted: $successCount files" -ForegroundColor Green
if ($failCount -gt 0) {
    Write-Host "✗ Failed conversions: $failCount files" -ForegroundColor Red
}

# Force garbage collection to clean up any remaining Excel processes
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Done!" -ForegroundColor Cyan