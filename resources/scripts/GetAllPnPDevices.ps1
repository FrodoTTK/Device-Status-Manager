# GetAllPnPDevices.ps1
# Get the execution directory of the current script
$currentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define output file name
$outputFile = Join-Path $currentDirectory "..\PnPDevices.json"

# Execute pnputil command and capture output
$pnputilOutput = & pnputil /enum-devices

# Initialize an empty device information list
$devices = @()
$currentDevice = @{}

# Analyze the output of pnputil
foreach ($line in $pnputilOutput) {
    # Match Instance ID line
    if ($line -match '^Instance ID:\s*(.+)') {
        if ($currentDevice.Count -gt 0) {
            $devices += [PSCustomObject]$currentDevice
            $currentDevice = @{}
        }
        $currentDevice['InstanceID'] = $matches[1]
    }
    elseif ($line -match '^Device Description:\s*(.+)') {
        $currentDevice['DeviceDescription'] = $matches[1]
    }
    elseif ($line -match '^Class Name:\s*(.+)') {
        $currentDevice['ClassName'] = $matches[1]
    }
    elseif ($line -match '^Class GUID:\s*(.+)') {
        $currentDevice['ClassGUID'] = $matches[1]
    }
    elseif ($line -match '^Manufacturer Name:\s*(.+)') {
        $currentDevice['ManufacturerName'] = $matches[1]
    }
    elseif ($line -match '^Status:\s*(.+)') {
        $currentDevice['Status'] = $matches[1]
    }
    elseif ($line -match '^Driver Name:\s*(.+)') {
        $currentDevice['DriverName'] = $matches[1]
    }
}

# Add the last device information
if ($currentDevice.Count -gt 0) {
    $devices += [PSCustomObject]$currentDevice
}

# Format device information as JSON and save it to a file
$devices | ConvertTo-Json | Set-Content -Path $outputFile
