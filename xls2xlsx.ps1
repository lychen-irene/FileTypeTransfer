param(
    [switch]$DeleteOriginal = $false
)

$SourcePath = Read-Host "Enter the source path (or press Enter for current directory)"
if ([string]::IsNullOrWhiteSpace($SourcePath)) { $SourcePath = "." }

# Set destination path to source path if not specified
$DestinationPath = Read-Host "Enter the destination path (or press Enter for current directory)"
if ([string]::IsNullOrWhiteSpace($DestinationPath)) { $DestinationPath = $SourcePath }

Write-Host "Source Path: $SourcePath" -ForegroundColor Yellow
Write-Host "Destination Path: $DestinationPath" -ForegroundColor Yellow

# Ensure paths exist
if (-not (Test-Path $SourcePath)) {
    Write-Error "Source path '$SourcePath' does not exist."
    exit 1
}

if (-not (Test-Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    Write-Host "Created destination directory: $DestinationPath" -ForegroundColor Green
}

# Excel constants
$xlOpenXMLWorkbook = 51  # XLSX format
$xlExcel8 = 56          # XLS format

try {
    # Create Excel application object
    Write-Host "Starting Excel application..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Get XLS files
    $searchParams = @{
        Path = $SourcePath
        Filter = "*.xls"
    }
    
    if ($Recurse) {
        $searchParams.Recurse = $true
    }
    
    $xlsFiles = Get-ChildItem @searchParams | Where-Object { $_.Extension -eq ".xls" }
    
    if ($xlsFiles.Count -eq 0) {
        Write-Host "No XLS files found in '$SourcePath'" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($xlsFiles.Count) XLS file(s) to convert" -ForegroundColor Green
    
    $convertedCount = 0
    $errorCount = 0
    
    foreach ($file in $xlsFiles) {
        try {
            Write-Host "Converting: $($file.Name)" -ForegroundColor Cyan
            
            # Open the XLS file
            $workbook = $excel.Workbooks.Open($file.FullName)
            
            # Generate output filename
            $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".xlsx"
            $outputPath = Join-Path $DestinationPath $outputFileName
            
            # Check if output file already exists
            if (Test-Path $outputPath) {
                $response = Read-Host "File '$outputFileName' already exists. Overwrite? (y/N)"
                if ($response -notmatch '^[Yy]') {
                    Write-Host "Skipped: $($file.Name)" -ForegroundColor Yellow
                    $workbook.Close($false)
                    continue
                }
            }
            
            # Save as XLSX
            $workbook.SaveAs($outputPath, $xlOpenXMLWorkbook)
            $workbook.Close($false)
            
            Write-Host "✓ Converted: $($file.Name) → $outputFileName" -ForegroundColor Green
            $convertedCount++
            
            # Delete original if requested
            if ($DeleteOriginal) {
                Remove-Item $file.FullName -Force
                Write-Host "  Deleted original file" -ForegroundColor Gray
            }
            
        }
        catch {
            Write-Error "Failed to convert $($file.Name): $($_.Exception.Message)"
            $errorCount++
            
            # Try to close workbook if it's still open
            try {
                if ($workbook) {
                    $workbook.Close($false)
                }
            }
            catch {
                # Ignore cleanup errors
            }
        }
    }
    
    # Summary
    Write-Host "`n=== Conversion Summary ===" -ForegroundColor Magenta
    Write-Host "Successfully converted: $convertedCount files" -ForegroundColor Green
    if ($errorCount -gt 0) {
        Write-Host "Failed conversions: $errorCount files" -ForegroundColor Red
    }
    
}
catch {
    Write-Error "Excel application error: $($_.Exception.Message)"
}
finally {
    # Clean up Excel application
    if ($excel) {
        try {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Write-Host "Excel application closed" -ForegroundColor Gray
        }
        catch {
            Write-Warning "Could not properly close Excel application"
        }
    }
}

# Force garbage collection to ensure Excel is fully released
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Conversion process completed!" -ForegroundColor Green