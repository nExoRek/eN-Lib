<#
.SYNOPSIS
    Search for EntraID account activity dates. 
    attributes are populated only with AAD P1 or higher license.
.DESCRIPTION
    proper permissioned are required:
        - domain.read.all to get the tenant name
        - auditlog.read.all to access signinactivity
        - user.read.all for user details
    Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All"
.EXAMPLE
    .\search-eNInactiveEntraUsers.ps1

.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 240520
        last changes
        - 240520 initialized

    #TO/DO
    * pagefile for big numbers

#>
#requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Users
[CmdletBinding()]
param (
    #skip connecting [second run]
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect
    
)
$VerbosePreference = 'Continue'
if(!$skipConnect) {
    Write-Verbose "athenticate to tenant..."
    #"Domain.ReadWrite.All" comes from get-mgDomain - but is not required.
    #"email" comes from get-mgDomain - and was double-requesting the authentication without this option
    #Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All","email"
    Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Domain.Read.All","email"
}
Write-Verbose "getting connection info..."
$ctx = Get-MgContext
Write-Verbose "connected as '$($ctx.Account)'"
if($ctx.Scopes -notcontains 'User.Read.All' -or $ctx.Scopes -notcontains 'AuditLog.Read.All' -or $ctx.Scopes -notcontains 'Domain.Read.All' -or $ctx.Scopes -notcontains 'Directory.Read.All') {
    throw "you need to connect using connect-mgGraph -Scopes User.Read.All,AuditLog.Read.All,Directory.Read.All,Domain.Read.All"
} else {
}
$tenantDomain = (get-MgDomain|? isdefault).id
$exportCSVFile = "EntraUsers-{0}-{1}.csv" -f $tenantDomain,(get-date -Format "yyMMdd-hhmm")
Write-Verbose "getting user info..."
$entraUsers = Get-MgUser -Property id,displayname,givenname,surname,accountenabled,userprincipalname,mail,signInActivity,userType,OnPremisesSyncEnabled -all |
    select-object displayname,accountenabled,givenname,surname,userprincipalname,userType,mail,id,`
    @{L='Hybrid';E={$_.OnPremisesSyncEnabled}},`
    @{L='LastLogonDate';E={if($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime } else { get-date "1/1/1970"} }},`
    @{L='LastNILogonDate';E={if($_.SignInActivity.LastNonInteractiveSignInDateTime) { $_.SignInActivity.LastNonInteractiveSignInDateTime } else { get-date "1/1/1970"} }},`
    @{L='licenses';E={(Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber -join ','}}
#adding field with lower of the two lastlogondate inactivity times. 
$entraUsers = $entraUsers | Select-Object *,`
    @{L='daysInactive';E={((New-TimeSpan -End (get-date) -Start $_.LastLogonDate).Days,(New-TimeSpan -End (get-date) -Start $_.LastNILogonDate).Days | Measure-Object -Minimum).Minimum}} |
        Sort-Object daysInactive,DisplayName 

Write-Verbose "found $($entraUsers.count) users."
$entraUsers | export-csv -nti $exportCSVFile -Encoding unicode
Write-Verbose "results saved in '$exportCSVFile'." 
