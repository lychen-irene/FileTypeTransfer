param(
    [Parameter(Mandatory=$true, HelpMessage="Path to the XLSX file or directory containing XLSX files")]
    [string]$InputPath,
    
    [Parameter(Mandatory=$false, HelpMessage="Output directory for PDF files (optional)")]
    [string]$OutputPath = "",
    
    [Parameter(Mandatory=$false, HelpMessage="Process all XLSX files in directory")]
    [switch]$Batch
)

function Convert-XlsxToPdf {
    param(
        [string]$XlsxPath,
        [string]$PdfPath
    )
    
    $excel = $null
    $workbook = $null
    
    try {
        # Ensure paths are absolute
        $XlsxPath = (Resolve-Path $XlsxPath).Path
        $PdfPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfPath)
        
        Write-Host "Converting: $XlsxPath" -ForegroundColor Green
        
        # Create Excel application object with additional settings
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $excel.Interactive = $false
        
        # Open the workbook with specific parameters to handle encoding issues
        $workbook = $excel.Workbooks.Open(
            $XlsxPath,          # Filename
            0,                  # UpdateLinks
            $true,              # ReadOnly
            [Type]::Missing,    # Format
            [Type]::Missing,    # Password
            [Type]::Missing,    # WriteResPassword
            $true,              # IgnoreReadOnlyRecommended
            [Type]::Missing,    # Origin
            [Type]::Missing,    # Delimiter
            $false,             # Editable
            $false,             # Notify
            [Type]::Missing,    # Converter
            $false              # AddToMru
        )
        
        # Wait a moment for Excel to fully load the workbook
        Start-Sleep -Milliseconds 1000
        
        # Export as PDF with explicit parameters
        # xlTypePDF = 0, xlQualityStandard = 0, IncludeDocProps = true, IgnorePrintAreas = false
        $workbook.ExportAsFixedFormat(
            [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF,  # Type
            $PdfPath,                                                        # Filename
            [Microsoft.Office.Interop.Excel.XlFixedFormatQuality]::xlQualityStandard, # Quality
            $true,                                                          # IncludeDocProps
            $false,                                                         # IgnorePrintAreas
            [Type]::Missing,                                               # From
            [Type]::Missing,                                               # To
            $false,                                                        # OpenAfterPublish
            [Type]::Missing                                                # BitmapMissingFonts
        )
        
        Write-Host "Successfully converted to: $PdfPath" -ForegroundColor Cyan
        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $errorMsg += " Inner: " + $_.Exception.InnerException.Message
        }
        
        Write-Host "Failed to convert $XlsxPath" -ForegroundColor Red
        Write-Host "Error: $errorMsg" -ForegroundColor Red
        
        # Try alternative conversion method using SaveAs2 if ExportAsFixedFormat fails
        if ($workbook -and $errorMsg -match "範圍|range") {
            try {
                Write-Host "Attempting alternative conversion method..." -ForegroundColor Yellow
                # Use SaveAs2 method as fallback
                $workbook.SaveAs2($PdfPath, 57) # 57 = xlTypePDF
                Write-Host "Successfully converted using alternative method: $PdfPath" -ForegroundColor Cyan
                return $true
            }
            catch {
                Write-Host "Alternative method also failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        return $false
    }
    finally {
        # Clean up COM objects
        try {
            if ($workbook) {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
        catch {
            Write-Host "Warning: Error during cleanup: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}

function Test-ExcelInstalled {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Main script execution
Write-Host "XLSX to PDF Converter" -ForegroundColor Yellow
Write-Host "===================" -ForegroundColor Yellow

# Check if Excel is installed
if (-not (Test-ExcelInstalled)) {
    Write-Error "Microsoft Excel is not installed or not accessible. This script requires Excel to be installed."
    exit 1
}

# Validate input path
if (-not (Test-Path $InputPath)) {
    Write-Error "Input path does not exist: $InputPath"
    exit 1
}

$inputItem = Get-Item $InputPath
$filesToProcess = @()

if ($inputItem.PSIsContainer) {
    # Directory provided
    if ($Batch) {
        $filesToProcess = Get-ChildItem -Path $InputPath -Filter "*.xlsx" -File
        if ($filesToProcess.Count -eq 0) {
            Write-Warning "No XLSX files found in directory: $InputPath"
            exit 0
        }
        Write-Host "Found $($filesToProcess.Count) XLSX file(s) to process" -ForegroundColor Green
    } else {
        Write-Error "Directory provided but -Batch switch not specified. Use -Batch to process all XLSX files in directory."
        exit 1
    }
} else {
    # Single file provided
    if ($inputItem.Extension -ne ".xlsx") {
        Write-Error "Input file must have .xlsx extension"
        exit 1
    }
    $filesToProcess = @($inputItem)
}

# Set output directory
if ([string]::IsNullOrEmpty($OutputPath)) {
    if ($inputItem.PSIsContainer) {
        $OutputPath = $InputPath
    } else {
        $OutputPath = $inputItem.Directory.FullName
    }
} else {
    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
}

Write-Host "Output directory: $OutputPath" -ForegroundColor Green

# Process files
$successCount = 0
$totalFiles = $filesToProcess.Count

foreach ($file in $filesToProcess) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $pdfPath = Join-Path $OutputPath "$baseName.pdf"
    
    # Check if PDF already exists
    if (Test-Path $pdfPath) {
        $response = Read-Host "PDF already exists: $pdfPath. Overwrite? (y/N)"
        if ($response -notmatch '^[Yy]') {
            Write-Host "Skipping: $($file.Name)" -ForegroundColor Yellow
            continue
        }
    }
    
    if (Convert-XlsxToPdf -XlsxPath $file.FullName -PdfPath $pdfPath) {
        $successCount++
    }
    
    # Add delay to prevent COM issues and allow proper cleanup
    Start-Sleep -Milliseconds 2000
    
    # Force additional cleanup between files
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`nConversion Summary:" -ForegroundColor Yellow
Write-Host "Total files: $totalFiles" -ForegroundColor White
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Failed: $($totalFiles - $successCount)" -ForegroundColor Red

if ($successCount -eq $totalFiles) {
    Write-Host "`nAll files converted successfully!" -ForegroundColor Green
} elseif ($successCount -gt 0) {
    Write-Host "`nSome files converted successfully." -ForegroundColor Yellow
} else {
    Write-Host "`nNo files were converted." -ForegroundColor Red
    exit 1
}