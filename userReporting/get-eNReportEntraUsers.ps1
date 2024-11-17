<#
.SYNOPSIS
    Reporting script, allowing to prepare aggregated information on the user accounts: 
     - general user information
     - MFA is checking extended attributes on the account so it will work for per-user and Conditional Access
     - AD Roles
     - last logon times (attributes are populated only with AAD P1 or higher license)
    As a part of a wider project, may be combined with AD and Exchange Online, giving better overview on hybrid identity.
.DESCRIPTION
    proper permissioned are required:
        - domain.read.all to get the tenant name
        - auditlog.read.all to access signInActivity
        - user.read.all for user details
        - "Directory.Read.All" - general read permissions
    Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All"
.EXAMPLE
    .\get-eNReportEntraUsers.ps1

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
        - 240718 initiated as a more generalized project, service plans display names check up, segmentation
        - 240627 MFA - for now only general status, AADP1 error handling
        - 240520 initialized

    #TO/DO
    * pagefile for big numbers
    * add 'extended MFA info' option
    * is it possible to check Conditional Access policies enforcing MFA?
    * add 'administrative roles'
    * to validate: if MFA will be visible as not enabled when from CA and not configured - I assume yes, but requires verification
#>
#requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Users
[CmdletBinding()]
param (
    #skip connecting [second run]
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect,
    #skip checking MFA status
    [Parameter(mandatory=$false,position=1)]
        [switch]$skipMFACheck,
    #skip getting user licenses information
    [Parameter(mandatory=$false,position=2)]
        [switch]$skipLicenseCheck,
    #automatically generate XLSX report using eNLib 
    [Parameter(mandatory=$false,position=3)]
        [switch]$xlsxReport
    
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

function convert-SKUCodeToDisplayName {
  param([string]$SKUname)

  $ServicePlan = $spInfo | Where-Object { $_.psobject.Properties.value -contains $SKUname }
  if($ServicePlan) {
      if($ServicePlan -is [array]) { $ServicePlan = $ServicePlan[0] }
      $property = ($ServicePlan.psobject.Properties| Where-Object value -eq $SKUname).name
      switch($property) {
          'Service_Plan_Name' {
              return $ServicePlan.'Service_Plans_Included_Friendly_Names'
          }
          'Service_Plans_Included_Friendly_Names' {
              return $ServicePlan.'Service_Plan_Name'
          }
          'Product_Display_Name' {
              return $ServicePlan.'String_Id'
          }
          'String_Id' {
              return $ServicePlan.'Product_Display_Name'
          }
          default { return $null }
      }
  } else {
      return $SKUname
  }
}

if(!$skipConnect) {
    try {
        Disconnect-MgGraph -ErrorAction Stop
    } catch {
        write-host 'testing error'
        write-verbose $_.Exception
        $_.ErrorDetails
    }
    Write-Verbose "athenticate to tenant..."
    #"Domain.ReadWrite.All" comes from get-mgDomain - but is not required.
    #"email" comes from get-mgDomain - and was double-requesting the authentication without this option
    #Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All","email"
    try {
        Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Domain.Read.All","email","UserAuthenticationMethod.Read.All"
    } catch {
        throw "error connecting. $($_.Exception)"
        return
    }
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
    Write-Verbose "found $($entraUsers.count) users."
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
        Write-Verbose "found $($entraUsers.count) users."
    } catch {
        write-host -ForegroundColor Red $_.exception.message
        return $_.exception.hresult
    }
}

if(!$skipMFACheck) {
    Write-Verbose "getting the per-user MFA info on accounts..."
    $EntraUsers = $EntraUsers | Select-Object *,MFAStatus 
    $EntraUsers | %{ $_.MFAStatus = (Get-MFAMethods $_.id).status }
} else {
    Write-Verbose "skipping the per-user MFA check..."
}

Write-Verbose "some output tuning..."
if($AADP1) {
$entraUsers = $entraUsers |
    select-object displayname,userType,accountenabled,givenname,surname,userprincipalname,mail,MFAStatus,`
        @{L='Hybrid';E={if($_.OnPremisesSyncEnabled) {$_.OnPremisesSyncEnabled} else {"FALSE"} }},`
        @{L='LastLogonDate';E={if($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime } else { get-date "1/1/1970"} }},`
        @{L='LastNILogonDate';E={if($_.SignInActivity.LastNonInteractiveSignInDateTime) { $_.SignInActivity.LastNonInteractiveSignInDateTime } else { get-date "1/1/1970"} }},`
        licenses,id 
    #adding field with lower of the two lastlogondate inactivity times. 
    $entraUsers = $entraUsers | Select-Object *,`
        @{L='daysInactive';E={((New-TimeSpan -End (get-date) -Start $_.LastLogonDate).Days,(New-TimeSpan -End (get-date) -Start $_.LastNILogonDate).Days | Measure-Object -Minimum).Minimum}} |
            Sort-Object daysInactive,DisplayName 
} else {
    $entraUsers = $entraUsers |
        select-object displayname,userType,accountenabled,givenname,surname,userprincipalname,mail,MFAStatus,`
        @{L='Hybrid';E={if($_.OnPremisesSyncEnabled) {$_.OnPremisesSyncEnabled} else {"FALSE"} }},`
        @{L='LastLogonDate';E={'NO AADP1'}},`
        @{L='LastNILogonDate';E={'NO AADP1'}},`
        licenses,id,`
        @{L='daysInactive';E={'NO AADP1'}}
}

if(!$skipLicenseCheck) {
    Write-Verbose "getting License info..."
    $spFile = ".\servicePlans.csv"

    if(!(test-path $spFile)) {
        Write-Verbose "file containing plans list not found - downloading..."
        [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
        Invoke-WebRequest $url -OutFile $spFile
    } 
    $spInfo = import-csv $spFile -Delimiter ','

    $entraUsers | %{ 
    $userLicenses = @()
    foreach($sku in (Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber ) {
        $userLicenses += convert-SKUCodeToDisplayName -SKUName $sku
    }
    $_.licenses = $userLicenses -join ";"
    }
} else {
    Write-Verbose "skipping license check..."
}
$entraUsers | Sort-Object UserType,AccountEnabled,daysInactive,DisplayName | export-csv -nti $exportCSVFile -Encoding unicode

if($xlsxReport) {
    write-host "creating xls report"
    $xlsFile = $exportCSVFile.Substring(0,20)
    Rename-Item $exportCSVFile "$xlsFile.csv"
    &(convert-CSV2XLS "$xlsFile.csv" -XLSfileName "$xlsFile.xlsx")
} else {
    Write-Verbose "results saved in '$exportCSVFile'." 
}
