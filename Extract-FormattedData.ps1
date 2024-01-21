# Display Powered By Morad Osama
Write-Host "Powered By Morad Osama"

# Create OpenFileDialog object with Multiselect enabled
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select Files"
$fileDialog.Filter = "Text Files (*.txt)|*.txt|Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
$fileDialog.Multiselect = $true

# Show the dialog and check if the user clicked OK
$result = $fileDialog.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    # User selected files, get the file paths
    $filePaths = $fileDialog.FileNames

    foreach ($filePath in $filePaths) {
        # Process each file as needed
        Write-Host "Processing file: $filePath"

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
            continue
        }

        # Remove specific symbols ([ and ]) before extracting IP addresses
        $cleanedContent = $fileContent -replace '\[|\]'

        # Extract IP addresses using regex
        $ipAddresses = [Regex]::Matches($cleanedContent, '\b(?:\d{1,3}\.){3}\d{1,3}\b') | ForEach-Object {
            $_.Value
        } | Select-Object -Unique

        # Extract emails using regex
        $emailPattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        $emails = [Regex]::Matches($cleanedContent, $emailPattern) | ForEach-Object {
            $_.Value
        } | Select-Object -Unique

        # Extract MD5 hashes using regex
        $md5Pattern = '\b([A-Fa-f0-9]{32})\b'
        $md5Hashes = [Regex]::Matches($cleanedContent, $md5Pattern) | ForEach-Object {
            $_.Value
        } | Select-Object -Unique

        # Extract SHA-256 hashes using regex
        $sha256Pattern = '\b([A-Fa-f0-9]{64})\b'
        $sha256Hashes = [Regex]::Matches($cleanedContent, $sha256Pattern) | ForEach-Object {
            $_.Value
        } | Select-Object -Unique

        # Extract SHA-1 hashes using regex
        $sha1Pattern = '\b([A-Fa-f0-9]{40})\b'
        $sha1Hashes = [Regex]::Matches($cleanedContent, $sha1Pattern) | ForEach-Object {
            $_.Value
        } | Select-Object -Unique

        # Extract URLs using regex
        $urlPattern = '\bhxxps?://[^\s/]+'
        $urls = [Regex]::Matches($cleanedContent, $urlPattern) | ForEach-Object {
            $_.Value -replace 'hxxps?://'
        } | Select-Object -Unique

        # Extract domain names from URLs and remove anything after TLD (Top-Level Domain)
        $domains = $urls | ForEach-Object {
            $uri = [uri]$_
            if ($uri.IsAbsoluteUri) {
                # Remove anything after TLD (Top-Level Domain) and the forward slash /
                $uri.Host -replace '\..*$', '' -replace '/.*$', ''
            }
        } | Select-Object -Unique

        # Create a folder if it doesn't exist
        $outputFolder = Join-Path -Path (Get-Location) -ChildPath "FG_Formatted_IPs"
        if (-not (Test-Path $outputFolder -PathType Container)) {
            New-Item -Path $outputFolder -ItemType Directory | Out-Null
        }

        # Display the extracted IPs, emails, hashes, and URLs on the screen
        Write-Host "Extracted IPs:"
        $ipAddresses

        Write-Host "Extracted Emails:"
        $emails

        Write-Host "Extracted MD5 Hashes:"
        $md5Hashes

        Write-Host "Extracted SHA-256 Hashes:"
        $sha256Hashes

        Write-Host "Extracted SHA-1 Hashes:"
        $sha1Hashes

        Write-Host "Extracted URLs:"
        $urls

        Write-Host "Extracted Domain Names:"
        $domains

        # Write the extracted IPs to a text file with double quotes and spaces in horizontal format
        $doubleQuotedOutputFilePath = Join-Path -Path $outputFolder -ChildPath "DoubleQuotedIPs.txt"
        $horizontalDoubleQuotedIPs = $ipAddresses -join '" "'
        $horizontalDoubleQuotedIPs = "`"$horizontalDoubleQuotedIPs`""
        $horizontalDoubleQuotedIPs | Out-File -FilePath $doubleQuotedOutputFilePath -Encoding UTF8 -Append

        Write-Host "Double-quoted IPs saved to '$doubleQuotedOutputFilePath'"

        # Write the extracted IPs to a text file in Fortigate format
        $fortigateFormattedOutputFilePath = Join-Path -Path $outputFolder -ChildPath "FortigateFormatted.txt"
        $ipAddresses | ForEach-Object { "edit $_`nset subnet $_/32`nnext" } | Out-File -FilePath $fortigateFormattedOutputFilePath -Encoding UTF8 -Append

        Write-Host "Fortigate-formatted IPs saved to '$fortigateFormattedOutputFilePath'"

        # Write the extracted emails to a text file
        $emailsOutputFilePath = Join-Path -Path $outputFolder -ChildPath "ExtractedEmails.txt"
        $emails -join "`r`n" | Out-File -FilePath $emailsOutputFilePath -Encoding UTF8 -Append

        Write-Host "Extracted Emails saved to '$emailsOutputFilePath'"

        # Write the extracted MD5 hashes to a text file
        $md5HashesOutputFilePath = Join-Path -Path $outputFolder -ChildPath "ExtractedMD5Hashes.txt"
        $md5Hashes -join "`r`n" | Out-File -FilePath $md5HashesOutputFilePath -Encoding UTF8 -Append

        Write-Host "Extracted MD5 Hashes saved to '$md5HashesOutputFilePath'"

        # Write the extracted SHA-256 hashes to a text file
        $sha256HashesOutputFilePath = Join-Path -Path $outputFolder -ChildPath "ExtractedSHA256Hashes.txt"
        $sha256Hashes -join "`r`n" | Out-File -FilePath $sha256HashesOutputFilePath -Encoding UTF8 -Append

        Write-Host "Extracted SHA-256 Hashes saved to '$sha256HashesOutputFilePath'"

        # Write the extracted SHA-1 hashes to a text file
        $sha1HashesOutputFilePath = Join-Path -Path $outputFolder -ChildPath "ExtractedSHA1Hashes.txt"
        $sha1Hashes -join "`r`n" | Out-File -FilePath $sha1HashesOutputFilePath -Encoding UTF8 -Append

        Write-Host "Extracted SHA-1 Hashes saved to '$sha1HashesOutputFilePath'"

        # Write the extracted URLs to a text file
        $urlsOutputFilePath = Join-Path -Path $outputFolder -ChildPath "ExtractedURLs.txt"
        $urls -join "`r`n" | Out-File -FilePath $urlsOutputFilePath -Encoding UTF8 -Append

        Write-Host "Extracted URLs saved to '$urlsOutputFilePath'"

        # Write the extracted URLs to a text file in Fortigate format
        $fortigateURLsOutputFilePath = Join-Path -Path $outputFolder -ChildPath "FortigateFormattedURLs.txt"
        $urls | ForEach-Object { "edit $_`nset type fqdn`nset fqdn $_`nnext" -replace ',' } | Out-File -FilePath $fortigateURLsOutputFilePath -Encoding UTF8 -Append

        Write-Host "Fortigate-formatted URLs saved to '$fortigateURLsOutputFilePath'"

        # Write the extracted URLs to a text file with double quotes
        $doubleQuotedURLsOutputFilePath = Join-Path -Path $outputFolder -ChildPath "DoubleQuotedURLs.txt"
        $doubleQuotedURLs = $urls -join '" "' | ForEach-Object { "`"$_`"" }
        $doubleQuotedURLs | Out-File -FilePath $doubleQuotedURLsOutputFilePath -Encoding UTF8 -Append

        Write-Host "Double-quoted URLs saved to '$doubleQuotedURLsOutputFilePath'"
    }
} else {
    Write-Host "User canceled the file selection."
}
