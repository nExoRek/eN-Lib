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
    version 240627
        last changes
        - 240627 MFA - for now only general status, AADP1 error handling
        - 240520 initialized

    #TO/DO
    * pagefile for big numbers
    * add 'extended MFA info' option
    * add 'administrative roles'
    * add ability to enable/disable parts of the report - MFA, licenses, admin roles, activity

#>
#requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Users
[CmdletBinding()]
param (
    #skip connecting [second run]
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect
    
)
$VerbosePreference = 'Continue'

Function Get-MFAMethods {
    <#
      .SYNOPSIS
        Get the MFA status of the user
    #>
    param(
      [Parameter(Mandatory = $true)] $userId
    )
    process{
      # Create MFA details object
      $mfaMethods  = [PSCustomObject][Ordered]@{
        status            = "-"
        authApp           = "-"
        phoneAuth         = "-"
        fido              = "-"
        helloForBusiness  = "-"
        helloForBusinessCount = 0
        emailAuth         = "-"
        tempPass          = "-"
        passwordLess      = "-"
        softwareAuth      = "-"
        authDevice        = ""
        authPhoneNr       = "-"
        SSPREmail         = "-"
      }
      # Get MFA details for each user
      try {
        [array]$mfaData = Get-MgUserAuthenticationMethod -UserId $userId -ErrorAction Stop
      } catch {
        $mfaMethods.status = 'error'
        return $mfaMethods
      }
      ForEach ($method in $mfaData) {
          Switch ($method.AdditionalProperties["@odata.type"]) {
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  { 
              # Microsoft Authenticator App
              $mfaMethods.authApp = $true
              $mfaMethods.authDevice += $method.AdditionalProperties["displayName"] 
              $mfaMethods.status = "enabled"
            } 
            "#microsoft.graph.phoneAuthenticationMethod"                  { 
              # Phone authentication
              $mfaMethods.phoneAuth = $true
              $mfaMethods.authPhoneNr = $method.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
              $mfaMethods.status = "enabled"
            } 
            "#microsoft.graph.fido2AuthenticationMethod"                   { 
              # FIDO2 key
              $mfaMethods.fido = $true
              $fifoDetails = $method.AdditionalProperties["model"]
              $mfaMethods.status = "enabled"
            } 
            "#microsoft.graph.passwordAuthenticationMethod"                { 
              # Password
              # When only the password is set, then MFA is disabled.
              if ($mfaMethods.status -ne "enabled") {$mfaMethods.status = "disabled"}
            }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
              # Windows Hello
              $mfaMethods.helloForBusiness = $true
              $helloForBusinessDetails = $method.AdditionalProperties["displayName"]
              $mfaMethods.status = "enabled"
              $mfaMethods.helloForBusinessCount++
            } 
            "#microsoft.graph.emailAuthenticationMethod"                   { 
              # Email Authentication
              $mfaMethods.emailAuth =  $true
              $mfaMethods.SSPREmail = $method.AdditionalProperties["emailAddress"] 
              $mfaMethods.status = "enabled"
            }               
            "microsoft.graph.temporaryAccessPassAuthenticationMethod"    { 
              # Temporary Access pass
              $mfaMethods.tempPass = $true
              $tempPassDetails = $method.AdditionalProperties["lifetimeInMinutes"]
              $mfaMethods.status = "enabled"
            }
            "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { 
              # Passwordless
              $mfaMethods.passwordLess = $true
              $passwordLessDetails = $method.AdditionalProperties["displayName"]
              $mfaMethods.status = "enabled"
            }
            "#microsoft.graph.softwareOathAuthenticationMethod" { 
              # ThirdPartyAuthenticator
              $mfaMethods.softwareAuth = $true
              $mfaMethods.status = "enabled"
            }
          }
      }
      #Write-Verbose "$userID -> $($mfaMethods.status)"
      Return $mfaMethods
    }
}

if(!$skipConnect) {
    Write-Verbose "athenticate to tenant..."
    #"Domain.ReadWrite.All" comes from get-mgDomain - but is not required.
    #"email" comes from get-mgDomain - and was double-requesting the authentication without this option
    #Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All","email"
    Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Domain.Read.All","email","UserAuthenticationMethod.Read.All"
}
Write-Verbose "getting connection info..."
$ctx = Get-MgContext
Write-Verbose "connected as '$($ctx.Account)'"
#if($ctx.Scopes -notcontains 'User.Read.All' -or $ctx.Scopes -notcontains 'AuditLog.Read.All' -or $ctx.Scopes -notcontains 'Domain.Read.All' -or $ctx.Scopes -notcontains 'Directory.Read.All') {
if($ctx.Scopes -notcontains 'User.Read.All' -or $ctx.Scopes -notcontains 'AuditLog.Read.All' -or $ctx.Scopes -notcontains 'Domain.Read.All' -or $ctx.Scopes -notcontains 'UserAuthenticationMethod.Read.All') {
    throw "you need to connect using connect-mgGraph -Scopes User.Read.All,AuditLog.Read.All,Domain.Read.All","UserAuthenticationMethod.Read.All"
} else {
}
$tenantDomain = (get-MgDomain|? isdefault).id
$exportCSVFile = "EntraUsers-{0}-{1}.csv" -f $tenantDomain,(get-date -Format "yyMMdd-hhmm")
[System.Collections.ArrayList]$userQuery = @('id','displayname','givenname','surname','accountenabled','userprincipalname','mail','signInActivity','userType','OnPremisesSyncEnabled')
$AADP1 = $true

Write-Verbose "getting user info..."
try {
    $entraUsers = Get-MgUser -ErrorAction Stop -Property $userQuery -all 
} catch {
    if($_.exception.hresult -eq -2146233088) {
        write-host "sorry.. it seems that you do not have a AAD P1 license - you need to purchase trial or at least single AAD P1 to have audit logging enabled. last logon information will not be available." -ForegroundColor Red
        $userQuery.remove('signInActivity')
        $AADP1 = $false
    } else {
        write-host -ForegroundColor Red $_.exception.message
        return $_.exception.hresult
    }
}
if(!$AADP1) {
    try {
        $entraUsers = Get-MgUser -ErrorAction Stop -Property $userQuery -all 
    } catch {
        write-host -ForegroundColor Red $_.exception.message
        return $_.exception.hresult
    }
}
Write-Verbose "getting the MFA info on accounts..."
$EntraUsers = $EntraUsers | Select-Object *,MFAStatus 
$EntraUsers | %{ $_.MFAStatus = (Get-MFAMethods $_.id).status }

Write-Verbose "getting License info and some final output tuning..."
if($AADP1) {
$entraUsers = $entraUsers |
    select-object displayname,userType,accountenabled,givenname,surname,userprincipalname,mail,MFAStatus,`
        @{L='Hybrid';E={$_.OnPremisesSyncEnabled}},`
        @{L='LastLogonDate';E={if($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime } else { get-date "1/1/1970"} }},`
        @{L='LastNILogonDate';E={if($_.SignInActivity.LastNonInteractiveSignInDateTime) { $_.SignInActivity.LastNonInteractiveSignInDateTime } else { get-date "1/1/1970"} }},`
        @{L='licenses';E={(Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber -join ','}},id 
    #adding field with lower of the two lastlogondate inactivity times. 
    $entraUsers = $entraUsers | Select-Object *,`
        @{L='daysInactive';E={((New-TimeSpan -End (get-date) -Start $_.LastLogonDate).Days,(New-TimeSpan -End (get-date) -Start $_.LastNILogonDate).Days | Measure-Object -Minimum).Minimum}} |
            Sort-Object daysInactive,DisplayName 
} else {
    $entraUsers = $entraUsers |
        select-object displayname,userType,accountenabled,givenname,surname,userprincipalname,mail,MFAStatus,`
        @{L='Hybrid';E={$_.OnPremisesSyncEnabled}},`
        @{L='LastLogonDate';E={'NO AADP1'}},`
        @{L='LastNILogonDate';E={'NO AADP1'}},`
        @{L='licenses';E={(Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber -join ','}},id,`
        @{L='daysInactive';E={'NO AADP1'}}
}

Write-Verbose "found $($entraUsers.count) users."
$entraUsers | Sort-Object UserType,AccountEnabled,daysInactive,DisplayName | export-csv -nti $exportCSVFile -Encoding unicode
Write-Verbose "results saved in '$exportCSVFile'." 
