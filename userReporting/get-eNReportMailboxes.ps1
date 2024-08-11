<#
.SYNOPSIS
    draft script for Exchange stats - it will be part of several script gathering statistics about user accounts.
    to be used for reporting - useful for migration or cleanup projects.
.DESCRIPTION
    here be dragons
.EXAMPLE
    .\get-eNReportMailboxes.ps1

.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 240811
        last changes
        - added delegated permissions to understand shared mailboxes (and security check),
            dived on steps and ability to pick up the work in case of broken job
            fixed mailbox type check
            other fixes
        - 240718 UPNs (for merge) and account status (for independent reports)
        - 240717 initialized

    #TO|DO
    * proper file description
    * instead of pickUp - provide file name and work as a 'refresh' to update on particular steps
#>
#requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param (
    #skip connection - if you're already connected
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect,
    #skip delegation permissions
    [Parameter(mandatory=$false,position=1)]
        [switch]$skipDelegations,
    #stats are generated for a long time. if you broke the script run, pick up where you left it
    [Parameter(mandatory=$false,position=2)]
        [switch]$pickUp
    
)
$VerbosePreference = 'Continue'

if(!$skipConnect) {
    Disconnect-ExchangeOnline -confirm:$false
    Connect-ExchangeOnline
}

#get some domain information 
$domain = (Get-AcceptedDomain|? Default -eq $true).domainName
Write-log "connected to $domain" -type info
$outfile = "mbxstats-$domain-$(get-date -Format "yyMMdd-hhmm").csv"

$stepFiles = @(
    "step0",
    "tmp_recipients.csv",
    "tmp_UPNs.csv",
    "tmp_mbxStats.csv"
)

$lastStep = 1
$pickUpFile = 'none'
if($pickUp) {
    try {
        $pickUpFile = $stepFiles[1]
        get-item $pickUpFile -ErrorAction Stop
        $lastStep = 2
    } catch {
        write-log "$pickUpFile not found" -silent
    }
    try {
        $pickUpFile = $stepFiles[2]
        get-item $pickUpFile -ErrorAction Stop
        $lastStep = 3
    } catch {
        write-log "$pickUpFile not found" -silent
    }
    try {
        $pickUpFile = $stepFiles[3]
        get-item $pickUpFile -ErrorAction Stop
        $lastStep = 4
    } catch {
        write-log "$pickUpFile not found" -silent
    }
} 
if($lastStep -gt 1) {
    write-log "found $pickUpFile. picking up the work from here..." -type info
    $recipients = load-CSV $pickUpFile
    write-log "loaded $($recipients.count) records." -type info
} else {
    write-log "close unfished job" -silent
    Remove-Item tmp_recipients.csv -ErrorAction SilentlyContinue
    Remove-Item tmp_UPNs.csv -ErrorAction SilentlyContinue
    Remove-Item tmp_mbxStats.csv -ErrorAction SilentlyContinue
}

if($lastStep -lt 2) {
    #'Recipients' is much wider, providing additional object infomration, thus starting from a getting all 'emails' in the tenant
    write-log "getting general recipients stats..." -type info
    $recipients = get-recipient |
        Select-Object Identity,userPrincipalName,enabled,DisplayName,FirstName,LastName,RecipientType,RecipientTypeDetails,`
            @{L='emails';E={$_.EmailAddresses -join ';'}},delegations, `
            WhenMailboxCreated,LastInteractionTime,LastUserActionTime,TotalItemSize,ExchangeObjectId
    #save current step
    $recipients | export-csv -nti -Encoding unicode tmp_recipients.csv
}

<#      recipient types
only UserMailbox has actual mailbox 
RecipientType                  RecipientTypeDetails
-------------                  --------------------
UserMailbox                    DiscoveryMailbox
UserMailbox                    UserMailbox
UserMailbox                    SharedMailbox
UserMailbox                    RoomMailbox
MailUser                       GuestMailUser
MailUser                       MailUser
MailContact                    MailContact
MailUniversalDistributionGroup RoomList
MailUniversalDistributionGroup GroupMailbox
MailUniversalSecurityGroup     MailUniversalSecurityGroup
MailUniversalDistributionGroup MailUniversalDistributionGroup
#>

if($lastStep -lt 3) {
    #some parameters make sens only for mailboxes - filter out non-mailbox enabled exchange objects and get the identity UPN for them
    write-log "getting UPNs from mailboxes..." -type info
    $recipients |? RecipientType -match 'userMailbox'| %{
        $mbx = Get-mailbox -identity $_.ExchangeObjectId
        $_.userPrincipalName = $mbx.userPrincipalName
        $_.enabled = $mbx.enabled
    }
    #save current step
    $recipients | export-csv -nti -Encoding unicode tmp_UPNs.csv
}

if($lastStep -lt 4) {
    #to know more about activity on a mailbox, get some last usage and basic size stats
    write-log "enriching mbx statistics..." -type info
    $recipients |? RecipientType -match 'userMailbox'| %{
        $stats = get-mailboxStatistics -identity $_.ExchangeObjectId
        $_.LastInteractionTime = $stats.LastInteractionTime
        $_.LastUserActionTime = $stats.LastUserActionTime
        $_.TotalItemSize = $stats.TotalItemSize
    }
    $recipients | export-csv -nti -Encoding unicode tmp_mbxStats.csv
}

if($lastStep -lt 5) {
    #especially useful for migration projects - mailbox delegations
    if(!$skipDelegations) {
        $recipients |? RecipientType -match 'userMailbox'| %{
            $permissions = Get-MailboxPermission -identity $_.ExchangeObjectId |
                ?{$_.isinherited -eq $false -and $_.user -notmatch 'NT AUTHORITY'} |
                %{"{0}:{1}" -f $_.user,$_.accessRights}
            if($permissions) {
                $_.delegations = $permissions -join "| "
            }
        }
    } else {
        Write-log "permissions check skipped." -type info
    }
}

#final results export
$recipients | Sort-Object RecipientTypeDetails,identity | Export-Csv -nti -Encoding unicode -Path $outfile

write-log "clean up..." -type info
Remove-Item tmp_recipients.csv -ErrorAction SilentlyContinue
Remove-Item tmp_UPNs.csv -ErrorAction SilentlyContinue
Remove-Item tmp_mbxStats.csv -ErrorAction SilentlyContinue

write-log "$outfile written." -type ok
