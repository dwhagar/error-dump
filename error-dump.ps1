<#
  Diagnostic bundle creator
  Outputs:
    System errors (24 h)
    System Services
    System Drivers
    System Information
    Windows Update History
    Installed Programs
    Network Configuration
    Reliability Report
    manifest.txt
    DiagnosticBundle.zip
#>

# -------- Paths --------
$desk = [Environment]::ExpandEnvironmentVariables(
          (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders').Desktop)
Set-Location $desk
$stamp = Get-Date -Format 'yyyy-MM-dd HH-mm'

# -------- 1. System errors (24 h) --------
$since  = (Get-Date).AddHours(-24)
$logs   = Get-WinEvent -ListLog * -ErrorAction SilentlyContinue |
          Where-Object { $_.IsEnabled -and ($_.LogType -in 'Administrative','Operational') } |
          Select-Object -ExpandProperty LogName
$events = foreach ($log in $logs) {
             Get-WinEvent -FilterHashtable @{LogName=$log;Level=2,3;StartTime=$since} -ErrorAction SilentlyContinue
          }
$events |
  Select-Object TimeCreated,LogName,ProviderName,Id,LevelDisplayName,Message |
  Export-Clixml -Path "System errors $stamp.xml" -Depth 1

# -------- 2. System services --------
Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
  Export-Clixml -Path 'System Services.xml' -Depth 1

# -------- 3. System drivers --------
Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
  Export-Clixml -Path 'System Drivers.xml' -Depth 1

# -------- 4. System information --------
[pscustomobject]@{
  BIOS       = Get-CimInstance Win32_BIOS              -ErrorAction SilentlyContinue
  Computer   = Get-CimInstance Win32_ComputerSystem    -ErrorAction SilentlyContinue
  Processor  = Get-CimInstance Win32_Processor         -ErrorAction SilentlyContinue
  BaseBoard  = Get-CimInstance Win32_BaseBoard         -ErrorAction SilentlyContinue
  OS         = Get-CimInstance Win32_OperatingSystem   -ErrorAction SilentlyContinue
} | Export-Clixml -Path 'System Information.xml' -Depth 1

# -------- 5. Windows Update history --------
Get-CimInstance Win32_QuickFixEngineering -ErrorAction SilentlyContinue |
  Export-Clixml -Path 'Windows Update History.xml' -Depth 1

# -------- 6. Installed programs --------
$uninst = @(
  'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
  'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
  'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
Get-ItemProperty $uninst -ErrorAction SilentlyContinue |
  Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |
  Export-Clixml -Path 'Installed Programs.xml' -Depth 1

# -------- 7. Network configuration --------
Get-NetIPConfiguration -ErrorAction SilentlyContinue |
  Export-Clixml -Path 'Network Configuration.xml' -Depth 1

# -------- 8. Reliability report --------
Get-WinEvent -LogName 'Microsoft-Windows-Reliability/Operational' -ErrorAction SilentlyContinue |
  Where-Object { $_.TimeCreated -ge $since } |
  Export-Clixml -Path 'Reliability Report.xml' -Depth 1

# -------- 9. Manifest --------
@"
System errors $stamp.xml    : System log errors and warnings, last 24 h
System Services.xml         : Service status, start type, account
System Drivers.xml          : Installed drivers, version, signing
System Information.xml      : BIOS, motherboard, CPU, OS build
Windows Update History.xml  : Installed and failed updates
Installed Programs.xml      : Programs from uninstall registry keys
Network Configuration.xml   : IP, gateway, DNS, interface details
Reliability Report.xml      : Reliability Monitor events, last 24 h
instructions.txt            : AI analysis and output instructions
manifest.txt                : This file
"@ | Set-Content -Encoding UTF8 'manifest.txt'

# -------- 10. Create output folder and move files --------
$folder = "Diagnostic Bundle"
if (-not (Test-Path $folder)) {
    New-Item -ItemType Directory -Path $folder | Out-Null
}

$files = @(
  "System errors $stamp.xml",
  'System Services.xml',
  'System Drivers.xml',
  'System Information.xml',
  'Windows Update History.xml',
  'Installed Programs.xml',
  'Network Configuration.xml',
  'Reliability Report.xml',
  'manifest.txt',
  'instructions.txt'
)

# -------- 10. Instructions for AI-based analysis --------
$prompt = @"
Read each XML file in DiagnosticBundle individually without using Python.  
1. Start with **System errors $stamp.xml** (24-hour System log); list every error or warning with Event ID, ProviderName, LogName, and timestamp.  
2. For any listed entry, consult supporting files only as needed:  
   • **System Services.xml** - service name, status, start-type, account  
   • **System Drivers.xml** - driver file, version, device class, signing state  
   • **Windows Update History.xml** - patches linked to the component  
   • **Installed Programs.xml** - conflicting or outdated software  
   • **System Information.xml** - BIOS, motherboard, CPU, OS build  
   • **Network Configuration.xml** - adapter, IP, DNS, gateway, VLAN  
   • **Reliability Report.xml** - crash and install timeline  
3. Match each error to the most relevant driver, service, update, program, or hardware element.  
4. Rank problems by combined severity, frequency, and recency.  
5. Output a rendered markdown table: Severity | Event ID | Probable cause | Possible fix
"@

$instructionsPath = Join-Path $folder 'instructions.txt'
$prompt | Set-Content -Encoding UTF8 $instructionsPath

foreach ($file in $files) {
    Move-Item -Path $file -Destination $folder -Force
}

Write-Host ""
Write-Host "Diagnostic data saved to '$folder'."
Write-Host ""
Write-Host "The detailed analysis prompt is in '$instructionsPath'."
Write-Host ""
Write-Host "To analyze:"
Write-Host "1. Open ChatGPT (or another capable LLM)."
Write-Host "2. Upload all XML files and 'instructions.txt' (no ZIP)."
Write-Host "3. Paste the line below, which tells the AI to read the file and follow it:"
Write-Host ""
Write-Host "------------------------------------------------------------"
Write-Host "Read the contents of instructions.txt and execute the diagnostic prompt found there."
Write-Host "------------------------------------------------------------"
Write-Host ""