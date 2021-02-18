[CmdletBinding()]
param (
    [Parameter(mandatory=$false,position=0)]
        [switch]$useGlobalModuleFolder
)
    $moduleFolder = [string]($Env:PSModulePath -split ';'|select-string $env:username)[0]
    if( [string]::IsNullOrEmpty($moduleFolder) ) {
        write-host 'user module folder not defined. please change $PSModulePath variable to include your folder name
or use -useGlobalMOduleFolder parameter' -ForegroundColor Yellow
        write-host -ForegroundColor Red "not installed."
        exit -2
    }
remove-module level1module -Force -ea SilentlyContinue
remove-module level2module -Force -ea SilentlyContinue

$testFolder1="$moduleFolder\level1module"
$testFolder2="$moduleFolder\level2module"
try {
    if(-not (test-path $testFolder1) ) {
        new-item -Type Directory $testFolder1|Out-Null 
    }
    $sourceFolder = get-item $MyInvocation.InvocationName
    Copy-Item -Path "$($sourceFolder.Directory.FullName)\level1module.psd1","$($sourceFolder.Directory.FullName)\level1module.psm1" -Destination $testFolder1
} catch {
    throw
}
write-host "files copied to $testFolder1." -ForegroundColor Green
Get-ChildItem $testFolder1
try {
    if(-not (test-path $testFolder2) ) {
        new-item -Type Directory $testFolder2|Out-Null 
    }
    $sourceFolder = get-item $MyInvocation.InvocationName
    Copy-Item -Path "$($sourceFolder.Directory.FullName)\level2module.psd1","$($sourceFolder.Directory.FullName)\level2module.psm1" -Destination $testFolder2
} catch {
    throw
}
Get-ChildItem $testFolder2
write-host "files copied to $testFolder2." -ForegroundColor Green

