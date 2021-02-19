#requires -module eNLib, Microsoft.Online.SharePoint.PowerShell
[CmdletBinding()]
param (
    #CSV input file
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSV,
    #delimiter
    [Parameter(mandatory=$false,position=1)]
        [string]$delimiter=';'
    
)

$header = @('Site name','URL')
$data = load-CSV -inputCSV $inputCSV -headerIsCritical -header $header -delimiter $delimiter

try { 
    $sposites = get-sposite 
} catch { 
    write-log 'no connection' -type error
    exit -1
}

foreach($site in $data) {
    $siteURL = ($site.url).replace('https://','').split('/')
    $siteRealtive = "/SDS"+$siteURL[-1]
    $newSiteURL = "{0}{1}{2}" -f 'https://', ($siteURL[0..($siteURL.count - 2)] -join '/'), $siteRealtive

    $newSiteTitle = "SDS $($site.'site name')"

    write-log "site ""$($site.URL)"" will be changed to ""$newSiteURL"" and title ""$newSiteTitle""" -type info
    try {
        Start-SPOSiteRename -Identity $site.url -NewSiteUrl $newSiteURL -NewSiteTitle $newSiteTitle -Confirm:$false
    } catch {
        write-log "error renaming site $($_.exception)" -type error
    }
}
write-log "done." -type ok