<#
.SYNOPSIS
    draft script for Exchange stats - it will be part of several script gathering statistics about user accounts.
    to be used for reporting - useful for migration or cleanup projects.
.DESCRIPTION
    here be dragons
.EXAMPLE
    .\get-eNMailboxInfo.ps1

    
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 240718
        last changes
        - 240718 UPNs (for merge) and account status (for independent reports)
        - 240717 initialized

    #TO|DO
    * a lot - this is just a starter
#>
#requires -Modules ExchangeOnlineManagement
[CmdletBinding()]
param (
    #skip connection - if you're already connected
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect
    
)
$VerbosePreference = 'Continue'

if(!$skipConnect) {
    Disconnect-ExchangeOnline -force
    Connect-ExchangeOnline
}

$domain = (Get-AcceptedDomain|? Default -eq $true).domainName
Write-Verbose "connected to $domain"
$outfile = "mbxstats-$domain-$(get-date -Format "yyMMdd-hhmm").csv"

write-verbose "getting general recipients stats..."
$recipients = get-recipient |
    Select-Object Identity,DisplayName,FirstName,LastName,RecipientType,RecipientTypeDetails,@{L='emails';E={$_.EmailAddresses -join ';'}},WhenMailboxCreated,`
        userPrincipalName,enabled,`
        LastInteractionTime,LastUserActionTime,TotalItemSize,ExchangeObjectId

write-log "getting UPNs from mailboxes..." -type info
$recipients |? RecipientTypeDetails -match 'mailbox'| %{
    $mbx = Get-mailbox -identity $_.ExchangeObjectId
    $_.userPrincipalName = $mbx.userPrincipalName
    $_.enabled = $mbx.enabled
}

write-log "enriching mbx statistics..." -type info
$recipients |? RecipientTypeDetails -match 'mailbox'| %{
    $stats = get-mailboxStatistics -identity $_.ExchangeObjectId
    $_.LastInteractionTime = $stats.LastInteractionTime
    $_.LastUserActionTime = $stats.LastUserActionTime
    $_.TotalItemSize = $stats.TotalItemSize
}
$recipients | Sort-Object RecipientTypeDetails,identity | Export-Csv -nti -Encoding unicode -Path $outfile
write-host "$outfile written." -ForegroundColor Green
