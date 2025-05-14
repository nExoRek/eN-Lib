<#
.SYNOPSIS
            DRAFT SCRIPT
    uninstall MSIEXE app by checking uninstall key and using it to run uninstaller.
.EXAMPLE
    .\uninstall-MSIApp.ps1 -appDisplayName "MyApp"
    
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250430
        last changes
        - 250430 initialized

    #TO|DO
    - this is draft script - i'm not working actively on it
    - should have ability to list displaynames and different queries - match/exact etc
    - error handling
    - reporting to eventlog
#>

[CmdletBinding()]
param (
    #uninstall function - checking uninstall key and uses it to run uninstaller.
    [Parameter(mandatory=$true,position=0)]
        [string]$appDisplayName   
)

$ItemProperties = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName,UninstallString
$uninstallStr = $ItemProperties | Where-Object { $_.DisplayName -eq $appDisplayName } | Select-Object -ExpandProperty UninstallString
if([string]::isNullOrEmpty($uninstallStr)) {
    Write-Host "No uninstall string found for $appDisplayName"
    return -1
}
$uninstallArgs = (($uninstallStr -split ' ')[1] -replace '/I','/X') + ' /q'
Start-Process msiexec.exe -ArgumentList $uninstallArgs -NoNewWindow -PassThru
