<#
.SYNOPSIS
    Prepares a report on AD objects with a focus on activity time - when the object has authenticated.
    Allows to prepare report for User and Computer objects. 
.DESCRIPTION
    Search-ADAccount commandlet is useful for quick ad-hoc queried, but it does not return all required object attributes 
    for proper reporting. This script is gathering much more information and is a part of a wider project allowing to
    create aggregated object reporting to support migrations, clean up or regular audits.

    requires to be run As Administrator as running in less priviledged context is not returing some values - e.g. 'enabled'
    status is sometimes returnes, sometimes not. 
.EXAMPLE
    .\get-eNReportADObjects.ps1
    
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.LINK
    http://www.selfadsi.org/ads-attributes/user-userAccountControl.htm
.NOTES
    nExoR ::))o-
    version 240718
        last changes
        - 240718 initiated as a wider project eNReport
        - 240519 initialized

    #TO|DO
    - resultpagesize - not managed. for now only for environments under 2k objects
#>
#requires -module ActiveDirectory
[CmdletBinding()]
param (
    #Parameter description
    [Parameter(mandatory=$false,position=0)]
        [validateSet('User','Computer')]
        [string]$objectType='User',
    #days of inactivity. 0 to make a full list
    [Parameter(mandatory=$false,position=1)]
        [int]$DaysInactive = 0 #by default make a full report   
)
$VerbosePreference = 'Continue'

#check for admin priviledges. there is this strange bug [or feature (; ] that if you run console without
#admin, some account do report 'enabled' attribute, some are not. so it's suggested to run as admin.
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if(-not $isAdmin) {
    Write-Warning "It's recommended to run script as administrator for full attribute visibility"
}

Write-Verbose "searching '$objectType' objects inactive for $DaysInactive days"

#http://www.selfadsi.org/ads-attributes/user-userAccountControl.htm
#USER does not require password 
#$UF_PASSWD_NOTREQD = "(userAccountControl:1.2.840.113556.1.4.803:=32)"
#USER password does not expire
#$UF_DONT_EXPIRE_PASSWD = "(userAccountControl:1.2.840.113556.1.4.803:=65536)"
[regex]$rxParentOU = 'CN=.*?,(.*$)'
$exportCSVFile = "AD{0}s-{1}-{2}.csv" -f $objectType,(Get-ADDomain).DNSRoot,(get-date -Format "yyMMdd-hhmm")


$DaysInactiveStr = (get-date).addDays(-$DaysInactive)
if($objectType -eq 'User') {
    $inactiveObjects = get-ADuser `
        -Filter {(lastlogondate -notlike "*" -OR lastlogondate -le $DaysInactiveStr)} `
        -Properties enabled,userPrincipalName,mail,distinguishedname,givenName,surname,samaccountname,displayName,description,lastLogonDate,PasswordLastSet
    Write-Verbose "found $(($inactiveObjects|Measure-Object).count) objects"
    $inactiveObjects |
        select-object samaccountname,userPrincipalName,enabled,givenName,surname,displayName,mail,description,`
            lastLogonDate,@{L='daysInactive';E={if($_.LastLogonDate) {$lld=$_.LastLogonDate} else {$lld="1/1/1970"} ;(New-TimeSpan -End (get-date) -Start $lld).Days}},PasswordLastSet,`
            distinguishedname,@{L='parentOU';E={$rxParentOU.Match($_.distinguishedName).groups[1].value}} | 
        Sort-Object daysInactive,parentOU |
        Export-csv $exportCSVFile -NoTypeInformation -Encoding utf8
} else {
    $inactiveObjects = get-ADComputer `
        -Filter {(lastlogondate -notlike "*" -OR lastlogondate -le $DaysInactiveStr)} `
        -Properties enabled,distinguishedname,samaccountname,displayName,description,lastLogonDate,PasswordLastSet
    Write-Verbose "found $(($inactiveObjects|Measure-Object).count) objects"
    $inactiveObjects |
        select-object samaccountname,enabled,displayName,description,`
            lastLogonDate,@{L='daysInactive';E={if($_.LastLogonDate) {$lld=$_.LastLogonDate} else {$lld="1/1/1970"} ;(New-TimeSpan -End (get-date) -Start $lld).Days}},PasswordLastSet,`
            distinguishedname,@{L='parentOU';E={$rxParentOU.Match($_.distinguishedName).groups[1].value}} | 
        Sort-Object daysInactive,parentOU |
        Export-csv $exportCSVFile -NoTypeInformation -Encoding utf8
}
Write-Verbose "results saved in '$exportCSVFile'"
