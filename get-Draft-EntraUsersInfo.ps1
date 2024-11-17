
[CmdletBinding()]
param (
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect
    
)
if(!$skipConnect) {
    write-host "athenticate to tenant..."
    #user.ReadWrite.All will be later required to block inactive users 
    #"Domain.ReadWrite.All" comes from get-mgDomain - but is not required.
    #"email" comes from get-mgDomain - and was double-requesting the authentication without this option
    #Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All","email"
    Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","Domain.Read.All","email","UserAuthenticationMethod.Read.All" -NoWelcome
}

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
      Return $mfaMethods
    }
}

$tenantDomain = (get-MgDomain|? isdefault).id
Write-Host "connected to $tenantDomain" -ForegroundColor Yellow
$exportFile = "$tenantDomain.EntraUsers.csv"
write-host "getting user info..."
$EntraUsers = Get-MgUser -Property id,displayname,givenname,surname,accountenabled,userprincipalname,mail,userType,OnPremisesSyncEnabled -all 
$EntraUsers = $EntraUsers | select-object displayname,accountenabled,@{L='ADSync';E={$_.OnPremisesSyncEnabled}},MFAStatus, `
    givenname,surname,userprincipalname,userType,mail, `
    @{L='licenses';E={(Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber -join ','}},id 
    
$EntraUsers | %{ $_.MFAStatus = (Get-MFAMethods $_.id).status }
    
$EntraUsers | Sort-Object UserType,AccountEnabled,DisplayName |
        export-csv -nti $exportFile -Encoding utf8

Write-Host "$exportFile created." -ForegroundColor Green