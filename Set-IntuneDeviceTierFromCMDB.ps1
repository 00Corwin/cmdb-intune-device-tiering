<#
.SYNOPSIS
    Maps ServiceNow CMDB server tiering data to Intune Device Categories.
 
.DESCRIPTION
    Imports a CSV export from ServiceNow CMDB and uses Microsoft Graph to map
    devices into Intune Device Categories such as Tier-0 through Tier-5.
 
    This can be used to support staged Defender/Intune policy rollout,
    update rings, dynamic Entra ID groups and blast-radius reduction.
 
.CSV REQUIREMENTS
    Required columns:
        Name
        Tier
 
    Optional column:
        Serial Number
 
.EXAMPLE
    .\Set-IntuneDeviceTierFromCMDB.ps1 -CsvPath ".\cmdb_servers.csv" -WhatIf
 
.EXAMPLE
    .\Set-IntuneDeviceTierFromCMDB.ps1 -CsvPath ".\cmdb_servers.csv" -CreateMissingCategories
 
.NOTES
    Required Graph permissions:
        DeviceManagementManagedDevices.ReadWrite.All
        DeviceManagementConfiguration.ReadWrite.All
#>
 
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
 
    [switch]$CreateMissingCategories,
 
    [string]$DefaultTier = "Tier-Unclassified",
 
    [string]$ReportPath = ".\IntuneTieringReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)
 
function Normalise-Tier {
    param(
        [string]$TierValue,
        [string]$FallbackTier
    )
 
    if ([string]::IsNullOrWhiteSpace($TierValue)) {
        return $FallbackTier
    }
 
    $clean = $TierValue.Trim()
 
    switch -Regex ($clean) {
        '^(Tier[-_\s]?)?0$|^T0$' { return "Tier-0" }
        '^(Tier[-_\s]?)?1$|^T1$' { return "Tier-1" }
        '^(Tier[-_\s]?)?2$|^T2$' { return "Tier-2" }
        '^(Tier[-_\s]?)?3$|^T3$' { return "Tier-3" }
        '^(Tier[-_\s]?)?4$|^T4$' { return "Tier-4" }
        '^(Tier[-_\s]?)?5$|^T5$' { return "Tier-5" }
        default { return $null }
    }
}
 
function Get-GraphCollection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )
 
    $results = @()
    $nextUri = $Uri
 
    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri
        if ($response.value) {
            $results += $response.value
        }
        $nextUri = $response.'@odata.nextLink'
    }
 
    return $results
}
 
function Get-CsvProperty {
    param(
        [object]$Row,
        [string[]]$PossibleNames
    )
 
    foreach ($name in $PossibleNames) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            return $Row.$name
        }
    }
 
    return $null
}
 
# Validate CSV path
if (-not (Test-Path $CsvPath)) {
    throw "CSV path not found: $CsvPath"
}
 
# Microsoft Graph module check
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Microsoft Graph PowerShell SDK is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
}
 
Import-Module Microsoft.Graph.Authentication
 
$scopes = @(
    "DeviceManagementManagedDevices.ReadWrite.All",
    "DeviceManagementConfiguration.ReadWrite.All"
)
 
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes | Out-Null
 
$graphBase = "https://graph.microsoft.com/beta"
 
Write-Host "Importing CMDB CSV: $CsvPath" -ForegroundColor Cyan
$cmdbRows = Import-Csv -Path $CsvPath
 
if (-not $cmdbRows) {
    throw "CSV contained no rows."
}
 
$firstRowProperties = $cmdbRows[0].PSObject.Properties.Name
 
if ($firstRowProperties -notcontains "Name") {
    throw "CSV must contain a 'Name' column."
}
 
if ($firstRowProperties -notcontains "Tier") {
    throw "CSV must contain a 'Tier' column."
}
 
Write-Host "Retrieving Intune device categories..." -ForegroundColor Cyan
$categories = Get-GraphCollection -Uri "$graphBase/deviceManagement/deviceCategories"
 
$categoryByName = @{}
foreach ($category in $categories) {
    $categoryByName[$category.displayName.ToLower()] = $category
}
 
Write-Host "Retrieving Intune managed devices..." -ForegroundColor Cyan
$managedDevices = Get-GraphCollection -Uri "$graphBase/deviceManagement/managedDevices?`$select=id,deviceName,serialNumber,operatingSystem,deviceCategoryDisplayName"
 
# Build lookup tables
$devicesByName = @{}
$devicesBySerial = @{}
 
foreach ($device in $managedDevices) {
    if (-not [string]::IsNullOrWhiteSpace($device.deviceName)) {
        $key = $device.deviceName.ToLower()
        if (-not $devicesByName.ContainsKey($key)) {
            $devicesByName[$key] = @()
        }
        $devicesByName[$key] += $device
    }
 
    if (-not [string]::IsNullOrWhiteSpace($device.serialNumber)) {
        $serialKey = $device.serialNumber.ToLower()
        if (-not $devicesBySerial.ContainsKey($serialKey)) {
            $devicesBySerial[$serialKey] = @()
        }
        $devicesBySerial[$serialKey] += $device
    }
}
 
$report = @()
 
foreach ($row in $cmdbRows) {
    $csvName = Get-CsvProperty -Row $row -PossibleNames @("Name", "Device Name", "ComputerName", "Computer Name")
    $csvTier = Get-CsvProperty -Row $row -PossibleNames @("Tier", "Device Tier", "Server Tier")
    $csvSerial = Get-CsvProperty -Row $row -PossibleNames @("Serial Number", "SerialNumber", "Serial")
 
    $desiredTier = Normalise-Tier -TierValue $csvTier -FallbackTier $DefaultTier
 
    $result = [ordered]@{
        CsvName              = $csvName
        CsvSerial            = $csvSerial
        CsvTier              = $csvTier
        DesiredTier          = $desiredTier
        IntuneDeviceName     = $null
        IntuneSerial         = $null
        OperatingSystem      = $null
        CurrentCategory      = $null
        MatchMethod          = $null
        Action               = $null
        Status               = $null
        Message              = $null
    }
 
    if ([string]::IsNullOrWhiteSpace($csvName)) {
        $result.Action = "Skipped"
        $result.Status = "Failed"
        $result.Message = "CSV row has no device name."
        $report += [pscustomobject]$result
        continue
    }
 
    if ([string]::IsNullOrWhiteSpace($desiredTier)) {
        $result.Action = "Skipped"
        $result.Status = "Failed"
        $result.Message = "Tier value '$csvTier' could not be mapped to Tier-0 through Tier-5."
        $report += [pscustomobject]$result
        continue
    }
 
    # Ensure device category exists
    $categoryKey = $desiredTier.ToLower()
 
    if (-not $categoryByName.ContainsKey($categoryKey)) {
        if ($CreateMissingCategories) {
            if ($PSCmdlet.ShouldProcess($desiredTier, "Create Intune Device Category")) {
                try {
                    $body = @{
                        displayName = $desiredTier
                        description = "Created by CMDB tiering automation on $(Get-Date -Format 'yyyy-MM-dd')"
                    } | ConvertTo-Json
 
                    $newCategory = Invoke-MgGraphRequest `
                        -Method POST `
                        -Uri "$graphBase/deviceManagement/deviceCategories" `
                        -Body $body `
                        -ContentType "application/json"
 
                    $categoryByName[$categoryKey] = $newCategory
                }
                catch {
                    $result.Action = "CreateCategory"
                    $result.Status = "Failed"
                    $result.Message = "Failed to create category '$desiredTier': $($_.Exception.Message)"
                    $report += [pscustomobject]$result
                    continue
                }
            }
            else {
                $result.Action = "CreateCategory"
                $result.Status = "WhatIf"
                $result.Message = "Would create missing category '$desiredTier'."
                $report += [pscustomobject]$result
                continue
            }
        }
        else {
            $result.Action = "Skipped"
            $result.Status = "Failed"
            $result.Message = "Device category '$desiredTier' does not exist. Re-run with -CreateMissingCategories if appropriate."
            $report += [pscustomobject]$result
            continue
        }
    }
 
    $targetCategory = $categoryByName[$categoryKey]
 
    # Match Intune managed device by name first, then serial number
    $matchedDevices = @()
    $nameKey = $csvName.ToLower()
 
    if ($devicesByName.ContainsKey($nameKey)) {
        $matchedDevices = $devicesByName[$nameKey]
        $result.MatchMethod = "Name"
    }
    elseif (-not [string]::IsNullOrWhiteSpace($csvSerial)) {
        $serialKey = $csvSerial.ToLower()
        if ($devicesBySerial.ContainsKey($serialKey)) {
            $matchedDevices = $devicesBySerial[$serialKey]
            $result.MatchMethod = "Serial"
        }
    }
 
    if (-not $matchedDevices -or $matchedDevices.Count -eq 0) {
        $result.Action = "Skipped"
        $result.Status = "Failed"
        $result.Message = "No matching Intune managed device found."
        $report += [pscustomobject]$result
        continue
    }
 
    if ($matchedDevices.Count -gt 1) {
        $result.Action = "Skipped"
        $result.Status = "Failed"
        $result.Message = "Multiple Intune devices matched. Manual review required."
        $report += [pscustomobject]$result
        continue
    }
 
    $device = $matchedDevices[0]
 
    $result.IntuneDeviceName = $device.deviceName
    $result.IntuneSerial = $device.serialNumber
    $result.OperatingSystem = $device.operatingSystem
    $result.CurrentCategory = $device.deviceCategoryDisplayName
 
    if ($device.deviceCategoryDisplayName -eq $desiredTier) {
        $result.Action = "NoChange"
        $result.Status = "Success"
        $result.Message = "Device already assigned to '$desiredTier'."
        $report += [pscustomobject]$result
        continue
    }
 
    $actionDescription = "Set Intune Device Category from '$($device.deviceCategoryDisplayName)' to '$desiredTier'"
 
    if ($PSCmdlet.ShouldProcess($device.deviceName, $actionDescription)) {
        try {
            $body = @{
                deviceCategoryId = $targetCategory.id
            } | ConvertTo-Json
 
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "$graphBase/deviceManagement/managedDevices/$($device.id)/setDeviceCategory" `
                -Body $body `
                -ContentType "application/json" | Out-Null
 
            $result.Action = "Updated"
            $result.Status = "Success"
            $result.Message = $actionDescription
        }
        catch {
            $result.Action = "Update"
            $result.Status = "Failed"
            $result.Message = $_.Exception.Message
        }
    }
    else {
        $result.Action = "Update"
        $result.Status = "WhatIf"
        $result.Message = "Would $actionDescription"
    }
 
    $report += [pscustomobject]$result
}
 
$report | Export-Csv -Path $ReportPath -NoTypeInformation
 
Write-Host ""
Write-Host "Completed." -ForegroundColor Green
Write-Host "Report written to: $ReportPath" -ForegroundColor Green
 
$summary = $report | Group-Object Status, Action | Select-Object Name, Count
$summary | Format-Table -AutoSize
