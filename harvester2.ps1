# Save to excel file
$excelPath = "$env:USERPROFILE\Desktop\harvester.xlsx"
 
 
# Get active network connections
$connections = netstat -ano | Where-Object { $_ -match "LISTENING|ESTABLISHED" }

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "ActivePorts"

# Write headers
$sheet.Cells.Item(1, 1) = "Local Address"
$sheet.Cells.Item(1, 2) = "Local Hostname"
$sheet.Cells.Item(1, 3) = "Remote Address"
$sheet.Cells.Item(1, 4) = "Remote Hostname"
$sheet.Cells.Item(1, 5) = "State"

$row = 2

foreach ($line in $connections) {
    $parts = $line -split "\s+"
    if ($parts.Length -ge 5) {
        $localAddress = $parts[3]
        $remoteAddress = $parts[2]

        # Extract IPs
        $localIP = ($localAddress -split ":")[0]
        $remoteIP = ($remoteAddress -split ":")[0]

        # Resolve hostnames
        try {
            $localHost = [System.Net.Dns]::GetHostEntry($localIP).HostName
        } catch {
            $localHost = "Unresolved"
        }

        try {
            $remoteHost = [System.Net.Dns]::GetHostEntry($remoteIP).HostName
        } catch {
            $remoteHost = "Unresolved"
        }

        # Write to Excel
        $sheet.Cells.Item($row, 1) = $localAddress
        $sheet.Cells.Item($row, 2) = $localHost
        $sheet.Cells.Item($row, 3) = $remoteAddress
        $sheet.Cells.Item($row, 4) = $remoteHost
        $sheet.Cells.Item($row, 5) = $parts[4]

        $row++
    }
}

# Save Excel file
$workbook.SaveAs($excelPath)
$workbook.Close($true)
$excel.Quit()






# Get local file shares
$shares = Get-SmbShare

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "LocalShares"

# Write headers
$headers = @("Name", "Path", "Description")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Write share data
$row = 2
foreach ($share in $shares) {
    $sheet.Cells.Item($row, 1).Value2 = $share.Name
    $sheet.Cells.Item($row, 2).Value2 = $share.Path
    $sheet.Cells.Item($row, 3).Value2 = $share.Description
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()





# Get environment variables as key-value pairs
$envVars = Get-ChildItem Env:

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "EnvVariables"

# Write headers
$sheet.Cells.Item(1, 1).Value2 = "Variable"
$sheet.Cells.Item(1, 2).Value2 = "Value"

# Write environment variables
$row = 2
foreach ($env in $envVars) {
    $sheet.Cells.Item($row, 1).Value2 = $env.Name
    $sheet.Cells.Item($row, 2).Value2 = $env.Value
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()







# Get mapped network drives (DriveType 4 = Network)
$mappedDrives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 4 }

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "MappedDrives"

# Write headers
$headers = @("Drive Letter", "Provider Name", "Volume Name")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Write drive data
$row = 2
foreach ($drive in $mappedDrives) {
    $sheet.Cells.Item($row, 1).Value2 = $drive.DeviceID
    $sheet.Cells.Item($row, 2).Value2 = $drive.ProviderName
    $sheet.Cells.Item($row, 3).Value2 = $drive.VolumeName
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()




# Get network adapter configurations with IP addresses
$adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "NetworkIPs"

# Write headers
$headers = @("Adapter Description", "MAC Address", "IP Address", "Subnet Mask", "Default Gateway")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Write adapter data
$row = 2
foreach ($adapter in $adapters) {
    $sheet.Cells.Item($row, 1).Value2 = $adapter.Description
    $sheet.Cells.Item($row, 2).Value2 = $adapter.MACAddress
    $sheet.Cells.Item($row, 3).Value2 = $adapter.IPAddress -join ", "
    $sheet.Cells.Item($row, 4).Value2 = $adapter.IPSubnet -join ", "
    $sheet.Cells.Item($row, 5).Value2 = $adapter.DefaultIPGateway -join ", "
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()




# Get DNS cache output
$dnsCache = ipconfig /displaydns

# Initialize record container
$records = @()
$currentRecord = @{}

foreach ($line in $dnsCache) {
    $line = $line.Trim()
    if ($line -eq "") {
        if ($currentRecord.Count -gt 0) {
            $records += [PSCustomObject]$currentRecord
            $currentRecord = @{}
        }
    } elseif ($line -match "^Record Name\s+\.+\s+(.*)$") {
        $currentRecord["RecordName"] = $matches[1]
        $currentRecord["RecordName"] = ($currentRecord["RecordName"] -split ":")[-1]
    } elseif ($line -match "^Record Type\s+\.+\s+(.*)$") {
        $currentRecord["RecordType"] = $matches[1]
        $currentRecord["RecordType"] = ($currentRecord["RecordType"] -split ":")[-1]
    } elseif ($line -match "^Time To Live\s+\.+\s+(.*)$") {
        $currentRecord["TTL"] = $matches[1]
        $currentRecord["TTL"] = ($currentRecord["TTL"] -split ":")[-1]
    } elseif ($line -match "^Data Length\s+\.+\s+(.*)$") {
        $currentRecord["DataLength"] = $matches[1]
        $currentRecord["DataLength"] = ($currentRecord["DataLength"] -split ":")[-1]
    } elseif ($line -match "^Section\s+\.+\s+(.*)$") {
        $currentRecord["Section"] = $matches[1]
        $currentRecord["Section"] = ($currentRecord["Section"] -split ":")[-1]
    } elseif ($line -match "^PTR Record\s+\.+\s+(.*)$") {
        $currentRecord["PTRRecord"] = $matches[1]
        $currentRecord["PTRRecord"] = ($currentRecord["PTRRecord"] -split ":")[-1]
    }
}

# Final flush in case last record wasn't followed by empty line
if ($currentRecord.Count -gt 0) {
    $records += [PSCustomObject]$currentRecord
}

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "DNSCache"

# Write headers
$headers = @("Record Name", "Record Type", "TTL", "Data Length", "Section", "PTR Record")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Write records
$row = 2
foreach ($record in $records) {
    $sheet.Cells.Item($row, 1).Value2 = $record.RecordName
    $sheet.Cells.Item($row, 2).Value2 = $record.RecordType
    $sheet.Cells.Item($row, 3).Value2 = $record.TTL
    $sheet.Cells.Item($row, 4).Value2 = $record.DataLength
    $sheet.Cells.Item($row, 5).Value2 = $record.Section
    $sheet.Cells.Item($row, 6).Value2 = $record.PTRRecord
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()





# Get all local user accounts
$localUsers = Get-LocalUser

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)

# Add new worksheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = "LocalUsers"

# Write headers
$headers = @("Name", "Full Name", "Enabled", "Description", "Last Logon")
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Write user data
$row = 2
foreach ($user in $localUsers) {
    $sheet.Cells.Item($row, 1).Value2 = $user.Name
    $sheet.Cells.Item($row, 2).Value2 = $user.FullName
    if ($user.Enabled){
     $sheet.Cells.Item($row, 3).Value2 = "True"
    } else {
     $sheet.Cells.Item($row, 3).Value2 = "False"
    }
    $sheet.Cells.Item($row, 4).Value2 = $user.Description
    $sheet.Cells.Item($row, 5).Value2 = $user.LastLogon
    $row++
}

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()




#----------------------- mail



# === CONFIGURATION ===
$NewSheetName = "MailAccounts"

# === Extract Outlook account info (via COM) ===
$outlook = New-Object -ComObject Outlook.Application
$session = $outlook.Session
$accounts = @()

foreach ($acc in $session.Accounts) {
    $accounts += [PSCustomObject]@{
        DisplayName = $acc.DisplayName
        SmtpAddress = $acc.SmtpAddress
        AccountType = $acc.AccountType  # 0 = Exchange, 1 = IMAP, 2 = POP3, etc
        UserName    = $acc.UserName
    }
}

# Release Outlook COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($session) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

if ($accounts.Count -eq 0) {
    #Write-Host "No Outlook accounts found."
    exit
}

# === Open existing Excel file and add a new sheet ===
if (-not (Test-Path $ExcelPath)) {
    #Write-Host "Excel file not found: $ExcelPath"
    exit
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($ExcelPath)

# Check for existing sheet with same name
try {
    $sheet = $workbook.Sheets.Item($NewSheetName)
    #Write-Host "Sheet '$NewSheetName' already exists. Overwriting."
    $sheet.Delete()
} catch {
    #Write-Host "Sheet '$NewSheetName' does not exist. Creating new."
}

# Add new sheet
$sheet = $workbook.Sheets.Add()
$sheet.Name = $NewSheetName

# Write headers
$headers = @("DisplayName", "SmtpAddress", "AccountType", "UserName")
for ($col = 0; $col -lt $headers.Count; $col++) {
    $sheet.Cells.Item(1, $col + 1) = $headers[$col]
}

# Write data
$row = 2
foreach ($acc in $accounts) {
    $sheet.Cells.Item($row, 1) = $acc.DisplayName
    $sheet.Cells.Item($row, 2) = $acc.SmtpAddress
    $sheet.Cells.Item($row, 3) = $acc.AccountType
    $sheet.Cells.Item($row, 4) = $acc.UserName
    $row++
}

# Autofit columns
$sheet.Columns.AutoFit()

# Save and clean up
$workbook.Save()
$workbook.Close($true)
$excel.Quit()




#------------------------mail




# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()



# Confirmation
Write-Host "saved to $excelPath"

