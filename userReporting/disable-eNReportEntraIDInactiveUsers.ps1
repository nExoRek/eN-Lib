[CmdletBinding()]
param (
    #input file with user activity report
    [Parameter(mandatory=$true,position=0)]
        [string]$inputFile,
    #do not confirm each disable operation - deafult will ask
    [Parameter(mandatory=$false,position=1)]
        [switch]$doNotConfirmDisable,
    #period of inactivity time to disable objects
    [Parameter(mandatory=$false,position=2)]
        [int]$monthsOld = -6    
)

try {
    $entraUsers = import-csv $inputFile
} catch {
    $_
}
if($monthsOld -gt 0) { $monthsOld = -1 * $monthsOld }
$toDisable = @()

foreach($eUser in $entraUsers) {
    Write-Verbose "processing $($eUser.DisplayName);$($eUser.ADSync);$($eUser.AccountEnabled)"
    if($eUser.LastLogonDate) {
        try {
            $lld = get-date $eUser.LastLogonDate
        } catch {
            write-host "not able to detect date $($eUser.displayname) - '$($eUser.LastLogonDate)'"
            continue
        }
    } else {
        $lld = get-date "1/1/1971"
        $eUser.LastLogonDate = $lld
    }
    if( $lld -lt (get-date).AddMonths($monthsOld) -and $eUser.AccountEnabled -eq 'True') {
        $toDisable += $eUser
    }
    if(!$eUser.ADSync) { $eUser.ADSync = 'False' }
}
if($toDisable.count -lt 1) {
    return "no unused accounts detected"
}
Write-Warning "list of accounts inactive for $(-1 * $monthsOld) months:"
$toDisable | ft DisplayName,LastLogonDate,AccountEnabled,ADSync,mail | out-host

write-warning "you can exit in any moment by pressing ctrl-c"
$bodyParams = @{  
    AccountEnabled = "false"  
}
foreach($obj in $toDisable) {
    write-host  "processing '$($obj.DisplayName)':'$($obj.mail)'..."
    if(!$doNotConfirmDisable) {
        Write-Warning "do you want to disable account? [$($obj.LastLogonDate)]"
        $readHostValue = Read-Host -Prompt "Enter Y to accept"
        if($readHostValue -ne 'y') {
            write-host "skipping account"
            continue
        }
    }
    write-host "disabling $($obj.DisplayName)..."
    Update-MgUser -UserId $obj.ID -BodyParameter $bodyParams
    Write-host -ForegroundColor Green "ok."
}
write-host 'done.'