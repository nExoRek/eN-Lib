#requires -module eNLib, Microsoft.Online.SharePoint.PowerShell
[CmdletBinding()]
param (
    #CSV input file
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSV
    
)

$header = @('Site name','URL')
$data = load-CSV -inputCSV $inputCSV -headerIsCritical -header $header -delimiter ','

try { 
    $sposites=get-sposite 
} catch { 
    'no connection' 
}

foreach($site in $data) {
    $siteURL = ($site.url).split('/')
    $siteRealtive = "SDS"+$siteURL[-1]
    $newSiteURL = 
    #Start-SPOSiteRename -Identity $site.url -NewSiteUrl 
}