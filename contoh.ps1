# Create a new Excel Application COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Make Excel invisible

# Define the path to the Excel file and the desired output PDF file
#$excelFilePath = "C:\x\exceltopdf\sample.xlsx"  # Change this to your Excel file path
#$pdfFilePath = "C:\x\exceltopdf\file.pdf"      # Change this to your desired PDF output path
$excelFilePath = $args[0]
$pdfFilePath = $args[1]

try {
    # Open the Excel file
    $workbook = $excel.Workbooks.Open($excelFilePath)
    
    # Loop through each sheet if needed
    foreach ($sheet in $workbook.Sheets) {
        # Save the workbook as PDF using the sheet's PageSetup settings
        $sheet.ExportAsFixedFormat(0, $pdfFilePath, $false, $true)  # 0 for PDF, $false to not include document properties, $true for landscape
        
        # You might need to set the PDF filename for each sheet separately if required
        # Example:
        # $pdfFilePath = "C:\path\to\your\" + $sheet.Name + ".pdf"
    }
    
    Write-Host "PDF created successfully: $pdfFilePath"
} catch {
    Write-Host "An error occurred: $_"
} finally {
    # Close the workbook and Excel application
    $workbook.Close($false)  # Don't save changes
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}