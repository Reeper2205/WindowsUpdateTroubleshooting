# Run as Administrator required
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    Write-Warning "Please run as Administrator!"
    Break
}

$endpoints = @(
    "fe3.delivery.mp.microsoft.com"
    "download.windowsupdate.com"
    "update.microsoft.com"
    "windowsupdate.microsoft.com"
    "catalog.update.microsoft.com"
    "ntservicepack.microsoft.com"
    "go.microsoft.com"
    "dl.delivery.mp.microsoft.com"
)

# Create timestamp for log file name
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = "C:\Temp\WindowsUpdate_Troubleshoot_$timestamp.log"

# Start transcript to capture all output
Start-Transcript -Path $logFile -Append

Write-Host "`n=== System Information ===" -ForegroundColor Green
systeminfo | Select-String "OS Name","OS Version","System Boot Time"

Write-Host "`n=== Testing Windows Update Endpoints ===" -ForegroundColor Green
foreach ($endpoint in $endpoints) {
    Write-Host "`nTesting connection to: $endpoint" 
    Test-NetConnection -ComputerName $endpoint -Port 443
    Resolve-DnsName $endpoint
}

Write-Host "`n=== Windows Update Service Status ===" -ForegroundColor Green
Get-Service -Name wuauserv, bits, cryptsvc, TrustedInstaller | Format-Table -AutoSize

Write-Host "`n=== Generating Windows Update Log ===" -ForegroundColor Green
Get-WindowsUpdateLog

Write-Host "`n=== Running SFC Scan ===" -ForegroundColor Green
sfc /scannow

Write-Host "`n=== Running DISM Health Restoration ===" -ForegroundColor Green
DISM /Online /Cleanup-Image /RestoreHealth

Write-Host "`n=== Resetting Windows Update Components ===" -ForegroundColor Green
# Stop relevant services
Write-Host "Stopping services..."
Stop-Service -Name BITS, wuauserv, cryptsvc, msiserver -Force

# Delete qmgr*.dat files
Write-Host "Removing BITS queue..."
Remove-Item "$env:ALLUSERSPROFILE\Application Data\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue
Remove-Item "$env:ALLUSERSPROFILE\Microsoft\Network\Downloader\qmgr*.dat" -ErrorAction SilentlyContinue

# Rename Software Distribution and Catroot2 folders
Write-Host "Renaming Software Distribution and Catroot2 folders..."
Rename-Item -Path "$env:SystemRoot\SoftwareDistribution" -NewName "SoftwareDistribution.old" -ErrorAction SilentlyContinue
Rename-Item -Path "$env:SystemRoot\System32\Catroot2" -NewName "Catroot2.old" -ErrorAction SilentlyContinue

# Reset Windows Update policies
Write-Host "Resetting Windows Update policies..."
& "$env:SystemRoot\System32\rundll32.exe" pnpclean.dll,RunDLL_PnpClean /DRIVERS /MAXCLEAN
netsh winsock reset
netsh winhttp reset proxy

# Start services again
Write-Host "Starting services..."
Start-Service -Name BITS
Start-Service -Name wuauserv
Start-Service -Name cryptsvc
Start-Service -Name msiserver

Write-Host "`n=== Checking Windows Update Service Status After Reset ===" -ForegroundColor Green
Get-Service -Name wuauserv, bits, cryptsvc, TrustedInstaller | Format-Table -AutoSize

Write-Host "`n=== Checking Pending Reboot Status ===" -ForegroundColor Green
$pendingRebootTests = @(
    @{
        Name = 'RebootPending'
        Test = { Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' }
    }
    @{
        Name = 'RebootRequired'
        Test = { Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' }
    }
    @{
        Name = 'PendingFileRename'
        Test = { Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations' }
    }
)

foreach ($test in $pendingRebootTests) {
    $result = Invoke-Command -ScriptBlock $test.Test
    Write-Host "$($test.Name): $result"
}

Write-Host "`n=== Troubleshooting Complete ===" -ForegroundColor Green
Write-Host "Log file saved to: $logFile"

# Stop transcript
Stop-Transcript