@echo off
title Cau hinh Windows Server 24/7
set "BAT_PATH=%~f0"

net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -NoProfile -Command "Start-Process -FilePath $env:BAT_PATH -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=[IO.File]::ReadAllText($env:BAT_PATH); $ps=($c -split ('###PS_' + 'START###'))[1]; Set-Content -Path (Join-Path $env:TEMP 'ServerMode_run.ps1') -Value $ps -Encoding UTF8"
powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\ServerMode_run.ps1"
del "%TEMP%\ServerMode_run.ps1" >nul 2>&1

echo.
pause
exit /b

###PS_START###
$ErrorActionPreference = "SilentlyContinue"

function Print-Status {
    param(
        [string]$Criteria,
        [ValidateSet("OK", "CHUA_OK", "FAIL")]
        [string]$Status,
        [string]$Note = ""
    )
    switch ($Status) {
        "OK" {
            Write-Host "[OK]      " -ForegroundColor Green -NoNewline
            Write-Host $Criteria
        }
        "CHUA_OK" {
            Write-Host "[CHUA OK] " -ForegroundColor Yellow -NoNewline
            Write-Host "$Criteria -> $Note"
        }
        "FAIL" {
            Write-Host "[FAIL]    " -ForegroundColor Red -NoNewline
            Write-Host "$Criteria -> $Note"
        }
    }
}

function Get-PowerCfgACValue {
    param([string]$SubGroup, [string]$Setting)
    $output = powercfg /query SCHEME_CURRENT $SubGroup $Setting 2>$null
    $hexMatches = $output | Select-String -Pattern "0x[0-9a-fA-F]{8}"
    if ($hexMatches -and $hexMatches.Count -ge 1) {
        return $hexMatches[0].Matches[0].Value
    }
    return $null
}

Write-Host "=== DANG CAU HINH CHE DO SERVER 24/7 CHO WINDOWS ===" -ForegroundColor Cyan

# 1. Vo hieu hoa che do Ngu (Sleep = Never)
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /setactive SCHEME_CURRENT
$sleepVal = Get-PowerCfgACValue -SubGroup "SUB_SLEEP" -Setting "STANDBYIDLE"
if ($sleepVal -eq "0x00000000") {
    Print-Status -Criteria "1. Vo hieu hoa Sleep (Khong bao gio tu ngu)" -Status "OK"
} else {
    Print-Status -Criteria "1. Vo hieu hoa Sleep" -Status "FAIL" -Note "Khong the dat gia tri Sleep = 0"
}

# 2. Vo hieu hoa Ngu dong (Hibernate) va Fast Startup
powercfg /hibernate off
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled" -Value 0 -Type DWord
$hiberVal = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "HibernateEnabled").HibernateEnabled
$fastBootVal = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name "HiberbootEnabled").HiberbootEnabled
if ($hiberVal -eq 0 -and $fastBootVal -eq 0) {
    Print-Status -Criteria "2. Tat Hibernate & Fast Startup (Reset lam moi hoan toan)" -Status "OK"
} else {
    Print-Status -Criteria "2. Tat Hibernate & Fast Startup" -Status "FAIL" -Note "Kiem tra lai quyen Admin"
}

# 3. Khong bao gio tat o cung (Hard Disk Timeout = 0)
powercfg /change disk-timeout-ac 0
powercfg /change disk-timeout-dc 0
powercfg /setactive SCHEME_CURRENT
$diskVal = Get-PowerCfgACValue -SubGroup "SUB_DISK" -Setting "DISKIDLE"
if (($diskVal -eq "0x00000000") -or ($null -eq $diskVal)) {
    Print-Status -Criteria "3. Giu o cung hoat dong lien tuc (Khong tu ngat dien)" -Status "OK"
} else {
    Print-Status -Criteria "3. Giu o cung hoat dong lien tuc" -Status "FAIL" -Note "Khong the dat Disk Timeout = 0"
}

# 4. Tu dong khoi dong lai khi he thong gap loi (BSOD Auto-Restart)
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -Name "AutoReboot" -Value 1 -Type DWord
$autoRebootVal = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -Name "AutoReboot").AutoReboot
if ($autoRebootVal -eq 1) {
    Print-Status -Criteria "4. Tu dong Reset khi Windows gap loi he thong" -Status "OK"
} else {
    Print-Status -Criteria "4. Tu dong Reset khi gap loi" -Status "FAIL" -Note "Loi ghi Registry CrashControl"
}

# 5. Chan Windows Update tu y khoi dong lai ngang xuong
$wuPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
if (-not (Test-Path $wuPath)) { New-Item -Path $wuPath -Force | Out-Null }
Set-ItemProperty -Path $wuPath -Name "NoAutoRebootWithLoggedOnUsers" -Value 1 -Type DWord
$wuVal = (Get-ItemProperty -Path $wuPath -Name "NoAutoRebootWithLoggedOnUsers").NoAutoRebootWithLoggedOnUsers
if ($wuVal -eq 1) {
    Print-Status -Criteria "5. Chan Windows Update tu dong Reset ngang xuong" -Status "OK"
} else {
    Print-Status -Criteria "5. Chan Windows Update tu dong Reset" -Status "FAIL" -Note "Loi ghi Policy WindowsUpdate"
}

# 6. Luon luon ket noi mang (PCIe, USB, Standby Network & Registry Card mang)
powercfg /setacvalueindex SCHEME_CURRENT SUB_PCIEXPRESS ASPM 0 2>$null
powercfg /setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0 2>$null
powercfg /setacvalueindex SCHEME_CURRENT SUB_NONE F15576E8-98B7-4186-B944-EAFA664402D9 1 2>$null
powercfg /setactive SCHEME_CURRENT

$activeNics = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
$nicRegOk = $true
foreach ($nic in $activeNics) {
    $pnpId = $nic.PnPDeviceID
    Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable | Where-Object { $_.InstanceName -like "*$pnpId*" } | ForEach-Object {
        Set-CimInstance -InputObject $_ -Property @{ Enable = $false }
    }
    $nicReg = Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}" |
              Get-ItemProperty | Where-Object { $_.NetCfgInstanceId -eq $nic.InterfaceGuid }
    if ($nicReg) {
        Set-ItemProperty -Path $nicReg.PSPath -Name "PnPCapabilities" -Value 24 -Type DWord
        $verifyVal = (Get-ItemProperty -Path $nicReg.PSPath -Name "PnPCapabilities").PnPCapabilities
        if ($verifyVal -ne 24) { $nicRegOk = $false }
    }
}
$usbVal = Get-PowerCfgACValue -SubGroup "2a737441-1930-4402-8d77-b2bebba308a3" -Setting "48e6b7a6-50f5-4782-a5d4-53bb8f07e226"
$usbOk = ($usbVal -eq "0x00000000") -or ($null -eq $usbVal)
if ($nicRegOk -and $usbOk) {
    Print-Status -Criteria "6. Luon luon ket noi mang (Card mang, USB & che do cho)" -Status "OK"
} else {
    Print-Status -Criteria "6. Luon luon ket noi mang" -Status "FAIL" -Note "Khong the khoa tiet kiem dien card mang"
}

# 7. Tu dong bat may khi co dien lai (BIOS AC Power Recovery)
$manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
$biosOk = $false
$biosNote = "Mainboard ($manufacturer) khong ho tro chinh BIOS tu Windows. Can vao BIOS bat thu cong."

if ($manufacturer -match "Dell") {
    Write-Host "   -> Dang kiem tra giao tiep BIOS Dell (co the mat 10-15 giay lan dau)..." -ForegroundColor DarkGray
    if (-not (Get-Module -ListAvailable -Name DellBIOSProvider)) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
        Install-Module -Name DellBIOSProvider -Force -AllowClobber | Out-Null
    }
    Import-Module DellBIOSProvider 2>$null
    if (Test-Path "DellSmbios:\PowerManagement\ACBehavior") {
        Set-Item -Path "DellSmbios:\PowerManagement\ACBehavior" -Value "PowerOn" 2>$null
        $acVal = (Get-Item -Path "DellSmbios:\PowerManagement\ACBehavior").CurrentValue
        if ($acVal -eq "PowerOn") { $biosOk = $true }
        else { $biosNote = "May Dell co dat mat khau BIOS hoac khong ho tro ACBehavior qua WMI." }
    } elseif (Test-Path "DellSmbios:\PowerManagement\ACRecovery") {
        Set-Item -Path "DellSmbios:\PowerManagement\ACRecovery" -Value "PowerOn" 2>$null
        $acVal = (Get-Item -Path "DellSmbios:\PowerManagement\ACRecovery").CurrentValue
        if ($acVal -eq "PowerOn") { $biosOk = $true }
    } else {
        $biosNote = "Khong tim thay muc ACBehavior trong DellSmbios. Can vao BIOS (F2) chinh thu cong."
    }
}
elseif ($manufacturer -match "Lenovo") {
    $setBios = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_SetBiosSetting 2>$null
    $saveBios = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_SaveBiosSettings 2>$null
    if ($setBios -and $saveBios) {
        Invoke-CimMethod -InputObject $setBios -MethodName SetBiosSetting -Arguments @{ parameter = "AfterPowerLoss,PowerOn" } | Out-Null
        Invoke-CimMethod -InputObject $saveBios -MethodName SaveBiosSettings | Out-Null
        $checkLenovo = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_BiosSetting | Where-Object { $_.CurrentSetting -like "AfterPowerLoss,PowerOn*" }
        if ($checkLenovo) { $biosOk = $true }
    }
}
elseif ($manufacturer -match "HP|Hewlett-Packard") {
    $hpBios = Get-CimInstance -Namespace root\hp\instrumentedBIOS -ClassName HP_BIOSSettingInterface 2>$null
    if ($hpBios) {
        Invoke-CimMethod -InputObject $hpBios -MethodName SetBIOSSetting -Arguments @{ Name = "After Power Loss"; Value = "Power On" } | Out-Null
        $checkHp = Get-CimInstance -Namespace root\hp\instrumentedBIOS -ClassName HP_BIOSSetting | Where-Object { $_.Name -eq "After Power Loss" -and $_.Value -like "*Power On*" }
        if ($checkHp) { $biosOk = $true }
    }
}

if ($biosOk) {
    Print-Status -Criteria "7. BIOS AC Recovery (Tu bat may sau khi mat dien)" -Status "OK"
} else {
    Print-Status -Criteria "7. BIOS AC Recovery (Tu bat may sau khi mat dien)" -Status "CHUA_OK" -Note $biosNote
}

# 8. Bat Wake-on-LAN (WoL) tren Windows de danh thuc may tu dien thoai
$wolOk = $false
$wolNotes = @()
$wolTargets = @()

$wolNics = Get-NetAdapter | Where-Object { $_.Status -eq "Up" -and $_.InterfaceDescription -notmatch "Tailscale|Loopback|Virtual|Hyper-V|VMware|TAP|VPN" }
if (-not $wolNics) {
    $wolNics = $activeNics
}

foreach ($nic in $wolNics) {
    $pnpId = $nic.PnPDeviceID

    # Bat Power Management WoL qua NetAdapterPowerManagement (chuan Windows)
    Set-NetAdapterPowerManagement -Name $nic.Name -WakeOnMagicPacket Enabled -ErrorAction SilentlyContinue

    # Cho phep thiet bi danh thuc may tinh (Allow this device to wake the computer)
    Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceWakeEnable | Where-Object { $_.InstanceName -like "*$pnpId*" } | ForEach-Object {
        Set-CimInstance -InputObject $_ -Property @{ Enable = $true }
    }
    Get-CimInstance -Namespace root\wmi -ClassName MSNdis_DeviceWakeOnMagicPacketOnly | Where-Object { $_.InstanceName -like "*$pnpId*" } | ForEach-Object {
        Set-CimInstance -InputObject $_ -Property @{ EnableWakeOnMagicPacketOnly = $true }
    }

    # Bat cac thuoc tinh nang cao lien quan den Wake-on-LAN (Realtek, Intel, v.v.)
    $advProps = Get-NetAdapterAdvancedProperty -Name $nic.Name -ErrorAction SilentlyContinue
    foreach ($propName in @("*WakeOnMagicPacket", "WakeOnMagicPacket", "ShutdownWakeOnLan", "*ShutdownWakeOnLan", "S5WakeOnLan")) {
        $matchProp = $advProps | Where-Object { $_.RegistryKeyword -eq $propName }
        if ($matchProp) {
            Set-NetAdapterAdvancedProperty -Name $nic.Name -RegistryKeyword $propName -RegistryValue 1 -ErrorAction SilentlyContinue
        }
    }

    # Kiem tra lai trang thai WoL
    $pm = Get-NetAdapterPowerManagement -Name $nic.Name -ErrorAction SilentlyContinue
    if ($pm -and $pm.WakeOnMagicPacket -eq "Enabled") {
        $wolOk = $true
    }

    $ipv4 = (Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1).IPAddress
    $isWifi = ($nic.MediaType -match "802\.11|Wireless" -or $nic.InterfaceDescription -match "Wi-Fi|Wireless|WLAN")
    $wolTargets += [PSCustomObject]@{
        Name = $nic.Name
        Desc = $nic.InterfaceDescription
        Mac  = $nic.MacAddress
        IP   = $ipv4
        Type = $(if ($isWifi) { "Wi-Fi (khong on dinh cho WoL)" } else { "Day LAN (chuan WoL)" })
    }
}

if ($wolOk) {
    Print-Status -Criteria "8. Bat Wake-on-LAN (Magic Packet) tren card mang" -Status "OK"
} else {
    Print-Status -Criteria "8. Bat Wake-on-LAN (Magic Packet) tren card mang" -Status "CHUA_OK" -Note "Card mang (hoac Wi-Fi) co the khong ho tro WoL qua Windows"
}

# 9. Bat Wake-on-LAN & Tat Deep Sleep trong BIOS (Dell/Lenovo/HP)
$biosWolOk = $false
$biosWolNote = "Mainboard ($manufacturer) can vao BIOS bat 'Wake on LAN / Power On by PCI-E' va tat 'ErP Ready'."

if ($manufacturer -match "Dell") {
    if (Test-Path "DellSmbios:\PowerManagement\DeepSleepCtrl") {
        Set-Item -Path "DellSmbios:\PowerManagement\DeepSleepCtrl" -Value "Disabled" 2>$null
    }
    if (Test-Path "DellSmbios:\PowerManagement\WakeOnLAN") {
        Set-Item -Path "DellSmbios:\PowerManagement\WakeOnLAN" -Value "LANOnly" 2>$null
        $wolBiosVal = (Get-Item -Path "DellSmbios:\PowerManagement\WakeOnLAN").CurrentValue
        if ($wolBiosVal -match "LAN") { $biosWolOk = $true }
        else { $biosWolNote = "Khong the bat WakeOnLAN tren Dell BIOS (co the co mat khau BIOS)." }
    }
}
elseif ($manufacturer -match "Lenovo") {
    $setBios = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_SetBiosSetting 2>$null
    $saveBios = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_SaveBiosSettings 2>$null
    if ($setBios -and $saveBios) {
        Invoke-CimMethod -InputObject $setBios -MethodName SetBiosSetting -Arguments @{ parameter = "WakeOnLAN,Primary" } | Out-Null
        Invoke-CimMethod -InputObject $saveBios -MethodName SaveBiosSettings | Out-Null
        $checkLenovoWol = Get-CimInstance -Namespace root\wmi -ClassName Lenovo_BiosSetting | Where-Object { $_.CurrentSetting -like "WakeOnLAN,Primary*" -or $_.CurrentSetting -like "WakeOnLAN,Enable*" }
        if ($checkLenovoWol) { $biosWolOk = $true }
    }
}
elseif ($manufacturer -match "HP|Hewlett-Packard") {
    $hpBios = Get-CimInstance -Namespace root\hp\instrumentedBIOS -ClassName HP_BIOSSettingInterface 2>$null
    if ($hpBios) {
        Invoke-CimMethod -InputObject $hpBios -MethodName SetBIOSSetting -Arguments @{ Name = "Wake On LAN"; Value = "Follow Boot Order" } | Out-Null
        $checkHpWol = Get-CimInstance -Namespace root\hp\instrumentedBIOS -ClassName HP_BIOSSetting | Where-Object { $_.Name -eq "Wake On LAN" -and $_.Value -notmatch "Disabled" }
        if ($checkHpWol) { $biosWolOk = $true }
    }
}

if ($biosWolOk) {
    Print-Status -Criteria "9. BIOS Wake-on-LAN (Nhan lenh bat may khi da tat)" -Status "OK"
} else {
    Print-Status -Criteria "9. BIOS Wake-on-LAN (Nhan lenh bat may khi da tat)" -Status "CHUA_OK" -Note $biosWolNote
}

# 10. Cau hinh Tailscale chay ngam (Unattended Mode - khong ngat ket noi khi logout / man hinh khoa)
$tsCmd = Get-Command "tailscale" -ErrorAction SilentlyContinue
$tsExe = if ($tsCmd) { $tsCmd.Source } elseif (Test-Path "${env:ProgramFiles}\Tailscale\tailscale.exe") { "${env:ProgramFiles}\Tailscale\tailscale.exe" } else { $null }

if ($tsExe) {
    $tsOut = & $tsExe set --unattended 2>&1
    if ($LASTEXITCODE -eq 0) {
        Print-Status -Criteria "10. Tailscale Unattended Mode (Duy tri ket noi khi logout/lock)" -Status "OK"
    } else {
        Print-Status -Criteria "10. Tailscale Unattended Mode" -Status "FAIL" -Note "$tsOut"
    }
} else {
    Print-Status -Criteria "10. Tailscale Unattended Mode" -Status "CHUA_OK" -Note "Khong tim thay Tailscale da cai dat tren he thong"
}

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "THONG TIN DE BAT MAY TU XA BANG DIEN THOAI:" -ForegroundColor Yellow
foreach ($t in $wolTargets) {
    Write-Host (" - Card: {0} [{1}]" -f $t.Name, $t.Type)
    Write-Host ("   MAC Address : {0}" -f $t.Mac) -ForegroundColor Green
    Write-Host ("   IP LAN      : {0}" -f $t.IP)
}
$tsIp = (Get-NetIPAddress -InterfaceAlias "*Tailscale*" -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1).IPAddress
if ($tsIp) {
    Write-Host (" - IP Tailscale: {0} (dung de remote khi may da bat)" -f $tsIp) -ForegroundColor Cyan
}
Write-Host ""
Write-Host "HUONG DAN BAT TU DIEN THOAI:" -ForegroundColor Yellow
Write-Host " 1. Cung Wi-Fi nha: Dung app Wake-on-LAN nhap MAC Address o tren (nen cam day LAN)."
Write-Host " 2. Tu xa qua 4G/5G: Dung O cam thong minh (ket hop muc 7 AC Recovery)"
Write-Host "    hoac gui WoL qua Router nha / thiet bi luon bat trong mang LAN."
Write-Host "=====================================================" -ForegroundColor Cyan