# Enable error handling
$ErrorActionPreference = "Stop"

# Log function for debugging
function Write-Log {
    param($Message)
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
}

Write-Log "Script started"

try {
    # Get and validate input parameters
    if ($args.Count -lt 2) {
        throw "Required parameters missing. Usage: script.ps1 <excel-path> <pdf-path>"
    }

    $excelFilePath = $args[0]
    $pdfFilePath = $args[1]

    Write-Log "Excel File: $excelFilePath"
    Write-Log "PDF Target: $pdfFilePath"

    # Verify Excel file exists
    if (-not (Test-Path $excelFilePath)) {
        throw "Excel file not found: $excelFilePath"
    }

    # Create PDF directory if it doesn't exist
    $pdfDirectory = Split-Path -Parent -Path $pdfFilePath
    if (-not (Test-Path $pdfDirectory)) {
        New-Item -ItemType Directory -Path $pdfDirectory -Force | Out-Null
        Write-Log "Created directory: $pdfDirectory"
    }

    Write-Log "Creating Excel COM object"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Log "Opening workbook"
    $workbook = $excel.Workbooks.Open($excelFilePath)
    
    # Adjust alignment for all worksheets
    foreach ($worksheet in $workbook.Worksheets) {
        Write-Log "Configuring worksheet: $($worksheet.Name)"
        
        # Get the used range
        $usedRange = $worksheet.UsedRange
        
        # Set vertical alignment to top for used range
        $usedRange.VerticalAlignment = -4160  # xlTop
        
        # Configure page setup
        $worksheet.PageSetup.CenterHorizontally = $true
        $worksheet.PageSetup.CenterVertically = $false
        
        # Adjust margins (in points)
        $worksheet.PageSetup.TopMargin = 20
        $worksheet.PageSetup.BottomMargin = 20
        $worksheet.PageSetup.LeftMargin = 20
        $worksheet.PageSetup.RightMargin = 20
        
        # Get the used range boundaries
        $startCell = $usedRange.Cells(1, 1).Address($false, $false)
        $endCell = $usedRange.Cells($usedRange.Rows.Count, $usedRange.Columns.Count).Address($false, $false)
        
        # Set print area
        $printArea = "${startCell}:${endCell}"
        Write-Log "Setting print area to: $printArea"
        $worksheet.PageSetup.PrintArea = $printArea
    }
    
    Write-Log "Converting to PDF"
    $workbook.ExportAsFixedFormat(0, $pdfFilePath)
    
    Write-Log "PDF creation successful"
} catch {
    Write-Log "Error occurred: $_"
    Write-Log "Exception details: $($_.Exception)"
    throw
} finally {
    if ($workbook) {
        Write-Log "Closing workbook"
        $workbook.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        Write-Log "Quitting Excel"
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Write-Log "Cleanup complete"
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}