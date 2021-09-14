<#
.SYNOPSIS
Enables Modern Authentication Registry Keys for Microsoft Office 2013 and 2016

.DESCRIPTION
Enables Modern Authentication Registry Keys as desribed in https://docs.microsoft.com/en-us/microsoft-365/admin/security-and-compliance/enable-modern-authentication?view=o365-worldwide

Compatible with Office 2013 and 2016

Version 1.0.1
#>

Clear-Host

Write-Host "Detecting version of Office installed"
[int] $OfficeVersion =  (Invoke-Command -scriptblock{ (New-Object -comobject outlook.application).version}).substring(0,2)

if ($OfficeVersion -eq 16) {
    Write-Host "Version of office detected is Microsoft Office 2016"
    $OfficeIdentityKeyPath = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity'
}
if ($OfficeVersion -eq 15) {
    Write-Host "Version of office detected is Microsoft Office 2013"
    $OfficeIdentityKeyPath = 'HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity'
}

# If Identity Key does not exist
if (-not(Test-Path $OfficeIdentityKeyPath)) {
    $OfficeIdentityParentKeyPath = $OfficeIdentityKeyPath -replace "\Identity", ""
    New-Item -Path $OfficeIdentityParentKeyPath -Name Identity
}

$OfficeConfig = Get-ItemProperty -Path $OfficeIdentityKeyPath

if (($OfficeConfig.Version -eq 1) -and ($OfficeConfig.EnableADAL -eq 1)) {
    Write-Host "Microsoft Office is configured for Modern Authentication" -ForegroundColor Green    
} else {
    Write-Host "Microsoft Office is not configured for Modern Authentication" -ForegroundColor Red
    Write-Host "Enabling Modern Authentication"
    if (($OfficeConfig.Version -ne 1) -and ($OfficeConfig.Version -ne 0)) {
        New-ItemProperty -Path $OfficeIdentityKeyPath -Name Version -Value 1 -PropertyType DWORD -Force | Out-Null
    }
    if (($OfficeConfig.EnableADAL -ne 1) -and ($OfficeConfig.EnableADAL -ne 0)) {
        New-ItemProperty -Path $OfficeIdentityKeyPath -Name EnableADAL -Value 1 -PropertyType DWORD -Force | Out-Null
    }
    
    $OfficeConfig = Get-ItemProperty -Path $OfficeIdentityKeyPath

    if (($OfficeConfig.Version -eq 1) -and ($OfficeConfig.EnableADAL -eq 1)) {
        Write-Host "Microsoft Office is now configured for Modern Authentication" -ForegroundColor Green
    } else {
        Write-Host "Microsoft Office could not be configured for Modern Authentication" -ForegroundColor Red
    }
}
pause