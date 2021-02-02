<#
.SYNOPSIS
    eNLib module installation script
.DESCRIPTION
    Copies module files to PATH diectory. by default choses user directory.
    use -useGlobalModuleFolder parameter to install for all users. requires to be run
    from elevated console.

.EXAMPLE
    .\install-eNLibModule.ps1
    copies module files to user module folder
.EXAMPLE
    .\install-eNLibModule.ps1 -useGlobalModuleFolder
    copies module files to global module folder. must be run in eveleted mode.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201019
        last changes
        - 201019 initialized
#>

[CmdletBinding()]
param (
    [Parameter(mandatory=$false,position=0)]
        [switch]$useGlobalModuleFolder
)
if($useGlobalModuleFolder) {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    
    if(-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) ) {
        write-host 'you need to run console in elevated mode in order to copy module to global path' -ForegroundColor Red
        exit -1
    }
    $moduleFolder="C:\Program Files\WindowsPowerShell\Modules"

} else {
    $moduleFolder = [string]($Env:PSModulePath -split ';'|select-string $env:username)
    if( [string]::IsNullOrEmpty($moduleFolder) ) {
        write-host 'user module folder not defined. please change $PSModulePath variable to include your folder name
or use -useGlobalMOduleFolder parameter' -ForegroundColor Yellow
        write-host -ForegroundColor Red "not installed."
        exit -2
    }
}

remove-module -name eNLib -ErrorAction SilentlyContinue

$eNLibFolder="$moduleFolder\eNLib"
try {
    if(-not (test-path $eNLibFolder) ) {
        new-item -Type Directory $eNLibFolder|Out-Null 
    }
    Copy-Item -Path *.psd1,*.psm1 -Destination $eNLibFolder
} catch {
    throw
}
Get-ChildItem $eNLibFolder
write-host "files copied to $eNLibFolder." -ForegroundColor Green

