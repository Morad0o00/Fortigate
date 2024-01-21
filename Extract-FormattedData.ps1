# Install ImportExcel module if not installed
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

Add-Type -AssemblyName System.Windows.Forms

# Create OpenFileDialog object
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select File"
$fileDialog.Filter = "Text Files (*.txt)|*.txt|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

# Show the dialog and check if the user clicked OK
$result = $fileDialog.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    # User selected a file, get the file path
    $filePath = $fileDialog.FileName

    # Prompt for password if the file is password-protected
    $password = Read-Host -Prompt "Enter password for the file (press Enter if not password-protected)"

    # Read the content of the file based on file format
    if ($filePath -match '\.txt$') {
        $fileContent = Get-Content -Path $filePath -Raw
    } elseif ($filePath -match '\.xlsx$') {
        # Import each sheet in the Excel workbook without considering column headers
        $excelPackage = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $filePath
        $sheets = $excelPackage.Workbook.Worksheets

        $fileContent = foreach ($sheet in $sheets) {
            if ($password) {
                $sheetContent = Import-Excel -Path $filePath -NoHeader -Password $password -Worksheet $sheet.Name | ForEach-Object { $_.PSObject.Properties.Value -join ' ' }
            } else {
                $sheetContent = Import-Excel -Path $filePath -NoHeader -Worksheet $sheet.Name | ForEach-Object { $_.PSObject.Properties.Value -join ' ' }
            }
            $sheetContent
        }
    } elseif ($filePath -match '\.csv$') {
        # Import CSV file and get the content
        $fileContent = Import-Csv -Path $filePath | ForEach-Object { $_.PSObject.Properties.Value -join ' ' }
    } else {
        Write-Host "Unsupported file format. Please select a TXT, XLSX, or CSV file."
        return
    }

    # Remove specific symbols ([ and ]) before extracting IP addresses
    $cleanedContent = $fileContent -replace '\[|\]'

    # Extract IP addresses using regex
    $ipAddresses = [Regex]::Matches($cleanedContent, '\b(?:\d{1,3}\.){3}\d{1,3}\b') | ForEach-Object {
        $_.Value
    }

    # Create a folder if it doesn't exist
    $outputFolder = Join-Path -Path (Get-Location) -ChildPath "FG_Formatted_IPs"
    if (-not (Test-Path $outputFolder -PathType Container)) {
        New-Item -Path $outputFolder -ItemType Directory | Out-Null
    }

    # Display the extracted IPs on the screen
    Write-Host "Extracted IPs:"
    $ipAddresses

    # Write the extracted IPs to a text file with double quotes and spaces in horizontal format
    $doubleQuotedOutputFilePath = Join-Path -Path $outputFolder -ChildPath "DoubleQuotedIPs.txt"
    $horizontalDoubleQuotedIPs = $ipAddresses -join '" "'
    $horizontalDoubleQuotedIPs = "`"$horizontalDoubleQuotedIPs`""
    $horizontalDoubleQuotedIPs | Out-File -FilePath $doubleQuotedOutputFilePath -Encoding UTF8

    Write-Host "Double-quoted IPs saved to '$doubleQuotedOutputFilePath'"

    # Write the extracted IPs to a text file in Fortigate format
    $fortigateFormattedOutputFilePath = Join-Path -Path $outputFolder -ChildPath "FortigateFormatted.txt"
    $ipAddresses | ForEach-Object { "edit $_`nset subnet $_/32`nnext" } | Out-File -FilePath $fortigateFormattedOutputFilePath -Encoding UTF8

    Write-Host "Fortigate-formatted IPs saved to '$fortigateFormattedOutputFilePath'"
} else {
    Write-Host "User canceled the file selection."
}
