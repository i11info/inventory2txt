# 00_collect_info.ps1
# This script collects system, hardware, and software information
# and saves a report with a filename that includes the current date, device name,
# motherboard manufacturer, motherboard model, and motherboard serial number.

# Function to get motherboard information
function Get-MotherboardInfo {
    $Motherboard = Get-WmiObject -Class Win32_BaseBoard
    return $Motherboard
}

# Retrieve motherboard information
$MotherboardInfo = Get-MotherboardInfo
$MotherboardManufacturer = $MotherboardInfo.Manufacturer
$MotherboardModel = $MotherboardInfo.Product
$MotherboardSerialNumber = $MotherboardInfo.SerialNumber

# Define the output filename using current date, device name, and motherboard details
$DeviceName = $env:COMPUTERNAME
$Date = Get-Date -Format "yyyy-MM-dd_HH-mm"
$SanitizedManufacturer = $MotherboardManufacturer -replace '[^\w]', '-'
$SanitizedModel = $MotherboardModel -replace '[^\w]', '-'
$SanitizedSerialNumber = $MotherboardSerialNumber -replace '[^\w]', '-'
$OutputFileName = "$DeviceName`_$Date`_$SanitizedManufacturer`_$SanitizedModel`_$SanitizedSerialNumber.txt"
$OutputFile = Join-Path -Path $PSScriptRoot -ChildPath $OutputFileName

# Create an array to accumulate report content
$ReportContent = @()

# Header
$ReportContent += "========== SYSTEM REPORT =========="
$ReportContent += "Device Name: $DeviceName"
$ReportContent += "Report generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm")"
$ReportContent += "Output File: $OutputFileName"
$ReportContent += "===================================`n"

# System Information (Selected OS-level details)
# Only include data not repeated elsewhere in the report.
$ReportContent += "===== System Information ====="
$SysInfo = Get-ComputerInfo -ErrorAction SilentlyContinue | 
           Select-Object CsManufacturer, CsModel, CsSystemType, CsDomain, TimeZone
$ReportContent += ($SysInfo | Format-List | Out-String)

# BIOS Information (includes serial numbers)
$ReportContent += "`n===== BIOS Information ====="
$ReportContent += (Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue | Out-String)

# Motherboard Information
$ReportContent += "`n===== Motherboard Information ====="
$ReportContent += (Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue | Out-String)

# Processor Information - output line by line
$ReportContent += "`n===== Processor Information ====="
$ReportContent += (Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Format-List | Out-String)

# Memory Information (Compact Format with Memory Type)
$ReportContent += "`n===== Physical Memory ====="

# Function to map SMBIOSMemoryType to readable memory type
function Get-MemoryType {
    param (
        [uint32]$SMBIOSMemoryType
    )
    switch ($SMBIOSMemoryType) {
        20 { return "DDR" }
        21 { return "DDR2" }
        22 { return "DDR2 FB-DIMM" }
        24 { return "DDR3" }
        26 { return "DDR4" }
        34 { return "DDR5" }
        default { return "Unknown" }
    }
}

$PhysicalMemory = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue |
                  Where-Object { $_.Capacity -and $_.Capacity -ne 0 } |
                  Select-Object @{Name='BankLabel';Expression={$_.BankLabel}},
                                @{Name='Capacity(GB)';Expression={[math]::Round($_.Capacity/1GB, 2)}},
                                @{Name='Speed(MHz)';Expression={$_.Speed}},
                                @{Name='Manufacturer';Expression={$_.Manufacturer}},
                                @{Name='SerialNumber';Expression={$_.SerialNumber}},
                                @{Name='MemoryType';Expression={Get-MemoryType $_.SMBIOSMemoryType}}

if ($PhysicalMemory) {
    $ReportContent += ($PhysicalMemory | Format-Table -AutoSize | Out-String)
} else {
    $ReportContent += "No physical memory information available or all entries have empty values."
}

# Disk Drive Information
$ReportContent += "`n===== Disk Drive Information =====`n"

$ReportContent += (Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue |
    Sort-Object DeviceID |
    Select-Object DeviceID, Caption, SerialNumber,
                  @{Name="Size(GB)"; Expression={[math]::round($_.Size / 1GB, 2)}} |
    Out-String)

# Retrieve all disks and sort them by Number in ascending order
$disks = Get-Disk | Sort-Object Number

# Initialize an array for table data
$tableData = @()

# Iterate through each disk
foreach ($disk in $disks) {
    # Retrieve the disk caption (model) and name
    $diskInfo = Get-CimInstance Win32_DiskDrive | Where-Object { $_.DeviceID -eq $disk.DeviceID }
    
    # Retrieve partitions associated with the current disk
    $partitions = Get-Partition -DiskNumber $disk.Number
    
    # Filter out reserved partitions (those without a drive letter)
    $usablePartitions = $partitions | Where-Object { $_.DriveLetter }

    if ($usablePartitions) {
        foreach ($partition in $usablePartitions) {
            # Retrieve the volume associated with the partition
            $volume = Get-Volume -Partition $partition
            
            # Add data to the table array
            $tableData += [PSCustomObject]@{
                'Disk Number' = $disk.Number
                'Drive Letter' = $partition.DriveLetter
                'Volume Label' = $volume.FileSystemLabel
                'Offset (GB)' = [math]::round($partition.Offset / 1GB, 2)
                'Size (GB)' = [math]::round($partition.Size / 1GB, 2)
                'Free Space (GB)' = [math]::round($volume.SizeRemaining / 1GB, 2)
            }
        }
    } else {
        $tableData += [PSCustomObject]@{
            'Disk Number' = $disk.Number
            'Drive Letter' = 'N/A'
            'Volume Label' = 'No usable partitions'
            'Offset (GB)' = 'N/A'
            'Size (GB)' = 'N/A'
            'Free Space (GB)' = 'N/A'
        }
    }
}

# Convert table data to formatted output and append to report
$ReportContent += $tableData | Format-Table -AutoSize | Out-String

# Network Adapter Information with error handling
$ReportContent += "`n===== Network Adapter Information ====="
try {
    $ReportContent += (Get-NetAdapter -ErrorAction SilentlyContinue | Out-String)
} catch {
    $ReportContent += "Error retrieving network adapter information: $_`n"
}

# GPU Information - output each field on a separate line
$ReportContent += "`n===== GPU Information ====="
$GPUInfo = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
           Select-Object Caption, Description, Name, AdapterRAM, DriverDate, DriverVersion, VideoModeDescription
$ReportContent += ($GPUInfo | Format-List | Out-String)

# Local User Accounts Information Section
$ReportContent += "`n===== Local User Accounts ====="
try {
    $users = Get-LocalUser -ErrorAction SilentlyContinue | Select-Object Name, Enabled, Description
    if ($users) {
        $ReportContent += ($users | Format-Table -AutoSize | Out-String)
    }
    else {
        $ReportContent += "No local user accounts found."
    }
} catch {
    $ReportContent += "Error retrieving local user accounts: $_"
}

# Windows Version and OS Details (customized to include only available data)
$ReportContent += "`n===== Windows Version and OS Details ====="
$OSDetailLines = @()
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    if ($os.Caption) {
        $OSDetailLines += "Product Name: $($os.Caption)"
    }
    if ($os.Version) {
        $OSDetailLines += "Version: $($os.Version)"
    }
    if ($os.BuildNumber) {
        $OSDetailLines += "OS Build: $($os.BuildNumber)"
    }
    if ($os.InstallDate) {
        $InstallDate = $os.InstallDate
        if ($InstallDate -is [DateTime]) {
            $OSDetailLines += "Installed on: $($InstallDate.ToString('yyyy-MM-dd'))"
        } else {
            $OSDetailLines += "Installed on: $InstallDate"
        }
    }
} catch {
    $OSDetailLines += "Error retrieving OS details: $_"
}
$ReportContent += ($OSDetailLines -join "`n")

# License Information Section
$ReportContent += "`n===== License Information =====`n"
function ConvertTo-ProductKey {
    param (
        [byte[]]$DigitalProductId
    )
    $Key = ""
    $Chars = "BCDFGHJKMPQRTVWXY2346789".ToCharArray()
    $KeyStartIndex = 52
    $KeyEndIndex = $KeyStartIndex + 15
    $Digits = $DigitalProductId[$KeyStartIndex..$KeyEndIndex]
    for ($i = 24; $i -ge 0; $i--) {
        $Current = 0
        for ($j = 14; $j -ge 0; $j--) {
            $Current = $Current * 256 -bxor $Digits[$j]
            $Digits[$j] = [math]::Floor($Current / 24)
            $Current = $Current % 24
        }
        $Key = $Chars[$Current] + $Key
        if (($i % 5) -eq 0 -and $i -ne 0) {
            $Key = "-" + $Key
        }
    }
    return $Key
}

function Get-WindowsProductKey {
    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $regValue = Get-ItemProperty -Path $registryPath -ErrorAction SilentlyContinue
    if ($regValue -and $regValue.DigitalProductId) {
         try {
              $productKey = ConvertTo-ProductKey $regValue.DigitalProductId
              if ($productKey -and $productKey -notmatch "Unknown") {
                 return $productKey
              }
              else {
                 $partial = ($regValue.DigitalProductId[52..66] | ForEach-Object { $_.ToString("X2") }) -join ""
                 return "Partial Product Key (hex): $partial"
              }
         } catch {
              $partial = ($regValue.DigitalProductId[52..66] | ForEach-Object { $_.ToString("X2") }) -join ""
              return "Partial Product Key (hex): $partial"
         }
    }
    else {
         return "No Windows Product Key found."
    }
}

$windowsKey = Get-WindowsProductKey
$ReportContent += "Windows Product Key: $windowsKey"

# Combined Installed Software (64-bit and 32-bit)
$ReportContent += "`n===== Installed Software (All) ====="
$regPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$softwareList = @()

function Convert-DateFormat {
    param (
        [string]$dateString
    )
    if ($dateString -match '^\d{1,2}/\d{1,2}/\d{4}$') {
        try {
            $date = [datetime]::ParseExact($dateString, 'M/d/yyyy', $null)
            return $date.ToString('yyyyMMdd')
        } catch {
            return ''
        }
    }
    else {
        return $dateString
    }
}

foreach ($path in $regPaths) {
    $softwareList += Get-ItemProperty $path | Select-Object @{Name="DisplayName"; Expression={$_.DisplayName}},
                                                            @{Name="DisplayVersion"; Expression={$_.DisplayVersion}},
                                                            @{Name="Publisher"; Expression={$_.Publisher}},
                                                            @{Name="InstallDate"; Expression={Convert-DateFormat $_.InstallDate}}
}

$softwareList = $softwareList | ForEach-Object {
    $_.DisplayName = if ($_.DisplayName.Length -gt 61) { $_.DisplayName.Substring(0, 61) + "..." } else { $_.DisplayName }
    $_.DisplayVersion = if ($_.DisplayVersion.Length -gt 14) { $_.DisplayVersion.Substring(0, 14) + "..." } else { $_.DisplayVersion }
    $_.Publisher = if ($_.Publisher.Length -gt 21) { $_.Publisher.Substring(0, 21) + "..." } else { $_.Publisher }
    $_
}

$softwareList = $softwareList | Where-Object { -not [string]::IsNullOrEmpty($_.DisplayName) } | Sort-Object DisplayName
$ReportContent += ($softwareList | Format-Table DisplayName, DisplayVersion, Publisher, InstallDate -AutoSize | Out-String)

# Installed Windows Updates (Hotfixes) - now the last section
$ReportContent += "`n===== Installed Windows Updates ====="
$ReportContent += (Get-HotFix -ErrorAction SilentlyContinue | Out-String)

# Finalize the report and write it to file
# Join all content with newline
$FinalReport = $ReportContent -join "`n"
# Replace two or more consecutive newline characters with just two newline characters (one empty line)
$FinalReport = $FinalReport -replace "(\r?\n\s*){2,}", "`n`n"

$FinalReport | Out-File -FilePath $OutputFile -Encoding utf8
Write-Host "System report generated at: $OutputFile"
