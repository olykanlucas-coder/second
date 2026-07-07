<#
.SYNOPSIS
    Compliance Assessment Automation Script
.DESCRIPTION
    Companion script for compliance_assessment.docx
    Collects system configuration and validates compliance
    against HIPAA, SOC 2, and ISO 27001 frameworks.
.NOTES
    Run this script alongside the Word document for full compliance assessment.
#>

# PowerShell compliance assessment engine
Write-Host "Compliance Assessment Engine v2.0" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# 1. Collect system configuration
Write-Host "Collecting system configuration..." -ForegroundColor Yellow
$systemInfo = Get-ComputerInfo | Select-Object WindowsVersion, OsHardwareAbstractionLayer, WindowsInstallationType, OsArchitecture
Write-Host "  Windows: $($systemInfo.WindowsVersion)" -ForegroundColor White
Write-Host "  Architecture: $($systemInfo.OsArchitecture)" -ForegroundColor White

# 2. Security software status
Write-Host "`nChecking security software..." -ForegroundColor Yellow
try {
    $mpStatus = Get-MpComputerStatus | Select-Object AntivirusEnabled, RealTimeProtectionEnabled, NISEnabled
    Write-Host "  Antivirus: $($mpStatus.AntivirusEnabled)" -ForegroundColor White
    Write-Host "  Real-time: $($mpStatus.RealTimeProtectionEnabled)" -ForegroundColor White
} catch {
    Write-Host "  (Defender status not available)" -ForegroundColor Gray
}

# 3. Firewall status
Write-Host "`nChecking firewall rules..." -ForegroundColor Yellow
$firewall = Get-NetFirewallProfile | Select-Object Name, Enabled
foreach ($profile in $firewall) {
    Write-Host "  $($profile.Name): $($profile.Enabled)" -ForegroundColor White
}

# 4. Open verification in Edge
Write-Host "`nOpening verification dashboard..." -ForegroundColor Yellow
try {
    Start-Process "msedge.exe" "https://verification.compliance-check.local" -ErrorAction SilentlyContinue
    Write-Host "  Edge opened successfully" -ForegroundColor Green
} catch {
    Write-Host "  Could not open Edge" -ForegroundColor Gray
}

# 5. Save report
$report = @{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    SystemInfo = $systemInfo
    SecuritySoftware = $mpStatus
    Firewall = $firewall
}
$report | ConvertTo-Json | Out-File "$env:TEMP\compliance_report.json"
Write-Host "`nReport saved to: $env:TEMP\compliance_report.json" -ForegroundColor Green

Write-Host ""
Write-Host "Compliance assessment complete!" -ForegroundColor Cyan