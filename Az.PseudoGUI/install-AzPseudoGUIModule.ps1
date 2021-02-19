<#
.SYNOPSIS
    AzPseudoGUI module installation script
.DESCRIPTION
    Copies module files to PATH diectory. by default choses user directory.
    use -useGlobalModuleFolder parameter to install for all users. requires to be run
    from elevated console.

.EXAMPLE
    .\install-AzPseudoGUIModule.ps1
    copies module files to user module folder
.EXAMPLE
    .\install-AzPseudoGUIModule.ps1 -useGlobalModuleFolder
    copies module files to global module folder. must be run in eveleted mode.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 210208
        last changes
        - 210208 run using relative path fix
        - 210202 initialized
    
    #TO|DO
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
    $moduleFolder = [string]($Env:PSModulePath -split ';'|select-string $env:username)[0]
    if( [string]::IsNullOrEmpty($moduleFolder) ) {
        write-host 'user module folder not defined. please change $PSModulePath variable to include your folder name
or use -useGlobalMOduleFolder parameter' -ForegroundColor Yellow
        write-host -ForegroundColor Red "not installed."
        exit -2
    }
}

remove-module -name AzPseudoGUI -ErrorAction SilentlyContinue

$AzPseudoGUIFolder="$moduleFolder\AzPseudoGUI"
try {
    if(-not (test-path $AzPseudoGUIFolder) ) {
        new-item -Type Directory $AzPseudoGUIFolder|Out-Null 
    }
    $sourceFolder = get-item $MyInvocation.InvocationName
    Copy-Item -Path "$($sourceFolder.Directory.FullName)\*.psd1","$($sourceFolder.Directory.FullName)\*.psm1" -Destination $AzPseudoGUIFolder
} catch {
    throw
}
Get-ChildItem $AzPseudoGUIFolder
write-host "files copied to $AzPseudoGUIFolder." -ForegroundColor Green

