<#
.SYNOPSIS
    Short description
.DESCRIPTION
    here be dragons

.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250125
        last changes
        - 250125 initialized

    #TO|DO
#>
function get-eNADPrivilegedUsers {
    <#
    .SYNOPSIS
        get all priviliedged users in domain
    .DESCRIPTION
        temporary script to be integrated with the rest of the ad report tool. the plan is to leave it separately for full reporting
        and some basic capability included in a general report.
    .EXAMPLE
        .\get-eNADPriviledgedUsers.ps1
        
        create audit file.
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 240124
            last changes
            - 240124 init
    #>
    [CmdletBinding()]
    param ( )
    ##### PREPERATION #####
    $PSDefaultParameterValues = @{ 
        "Export-CSV:noTypeInformation"=$true
        "Export-CSV:encoding" = "UTF8"
    }
    $VerbosePreference = 'Continue'

    write-verbose "preparing basic objects..."

    $domainInfoObject = Get-ADDomain
    $forestInfoObject = Get-ADForest #ReplicaDirectoryServers multivalue
    $RootDSE = Get-ADRootDSE

    $dnsRoot = $domainInfoObject.dnsroot
    $domainSID = $domainInfoObject.domainSID

    #https://docs.microsoft.com/en-us/windows/security/identity-protection/access-control/active-directory-security-groups
    $wellKnownSids = @{
        CACHEABLE_PRINCIPALS_GROUP           = "S-1-5-32-571"
        NON_CACHEABLE_PRINCIPALS_GROUP       = "S-1-5-32-572"
        DEVICE_OWNERS                        = "S-1-5-32-583"
        'Power Users'                        = "S-1-5-32-547"
        "RAS Servers"                        = "S-1-5-32-553"
        "RDS Management Servers"             = "S-1-5-32-577"
        "Remote Desktop Users"               = "S-1-5-32-555"
        "Administrators"                     = "S-1-5-32-544"
        "Remote Management Users"            = "S-1-5-32-580"
        "Storage Replica Administrators"     = "S-1-5-32-582"
        "Windows Authorization Access Group" = "S-1-5-32-560"
        "System Managed Accounts Group"      = "S-1-5-32-581"
        "Backup Operators"                   = "S-1-5-32-551"
        "Network Configuration Operators"    = "S-1-5-32-556"
        "Terminal Server License Servers"    = "S-1-5-32-561"
        "Hyper-V Administrators"             = "S-1-5-32-578"
        "IIS_IUSRS"                          = "S-1-5-32-568"
        "Account Operators"                  = "S-1-5-32-548"
        "RDS Remote Access Servers"          = "S-1-5-32-575"
        "Print Operators"                    = "S-1-5-32-550"
        "Access Control Assistance Operators" = "S-1-5-32-579"
        "Incoming Forest Trust Builders"     = "S-1-5-32-557"
        "Server Operators"                   = "S-1-5-32-549"
        "Distributed COM Users"              = "S-1-5-32-562"
        "Certificate Service DCOM Access"    = "S-1-5-32-574"
        "Performance Monitor Users"          = "S-1-5-32-558"
        "Performance Log Users"              = "S-1-5-32-559"
        "Pre-Windows 2000 Compatible Access" = "S-1-5-32-554"
        "Event Log Readers"                  = "S-1-5-32-573"
        "Users"                              = "S-1-5-32-545"
        "Replicator"                         = "S-1-5-32-552"
        "Cryptographic Operators"            = "S-1-5-32-569"
        "RDS Endpoint Servers"               = "S-1-5-32-576"
        "Guests"                             = "S-1-5-32-546"

        "Enterprise Read-only Domain Controllers" = "$domainSID-498"
        "Domain Admins"                      = "$domainSID-512"
        "Domain Users"                       = "$domainSID-513"
        "Domain Guests"                      = "$domainSID-514"
        "Domain Computers"                   = "$domainSID-515"
        "Domain Controllers"                 = "$domainSID-516"
        "Cert Publishers"                    = "$domainSID-517"
        "Schema Admins"                      = "$domainSID-518"
        "Enterprise Admins"                  = "$domainSID-519"
        "Group Policy Creator Owners"        = "$domainSID-520"
        "Read-only Domain Controllers"       = "$domainSID-521"
        "Cloneable Domain Controllers"       = "$domainSID-522"
        CDC_RESERVED                         = "$domainSID-524"
        "PROTECTED USERS"                    = "$domainSID-525"
        "Key Admins"                         = "$domainSID-526"
        "Enterprise Key Admins"              = "$domainSID-527"
    }
    $dynamicSIDgroups = @(
        "DnsAdmins",
            #EXCHANGE
        "Organization Management",
        "Recipient Management",
        "View-Only Organization Management",
        "Public Folder Management",
        "UM Management",
        "Help Desk",
        "Records Management",
        "Discovery Management",
        "Server Management",
        "Delegated Setup",
        "Hygiene Management",
        "Compliance Management",
        "Security Reader",
        "Security Administrator",
        "Exchange Servers",
        "Exchange Trusted Subsystem",
        "Managed Availability Servers",
        "Exchange Windows Permissions",
        "ExchangeLegacyInterop",
        "Exchange Install Domain Servers"
    )

    #ADMIN GORUPS MEMBERSHIP 
    write-verbose "gather privileged users..."
    $reportPrivilegedGroupMembers = "privilegedGroupMembers.csv"
    foreach($group in $wellKnownSids.keys) {
        if($group -eq 'domain users' -or $group -eq 'domain computers') {
            #these are not privileged, and taken care of in different part of the script
            continue
        }
        try {
            $oGrpName = (Get-ADGroup -Identity $wellKnownSids[$group] -ErrorAction SilentlyContinue).name
        } catch {
            Write-Verbose "$($wellKnownSids[$group]) not found." 
            continue
        }
        $oGrpMembers = Get-ADGroupMember -Identity $wellKnownSids[$group]
        write-verbose "$($wellKnownSids[$group]) group name: $oGrpName"
        write-verbose "number of members: $($oGrpMembers.count)" 
        $oGrpMembers | Select-Object @{L='groupname';E={$oGrpName}},@{L='memberName';E={$_.name}},distinguishedName,objectClass | Export-Csv -Path $reportPrivilegedGroupMembers -Append
    }
    foreach($group in $dynamicSIDgroups) {
        try {
            $oGrpMembers = Get-ADGroupMember -Identity $group -ErrorAction Stop
        } catch {
            write-verbose "$group not found."
            continue
        }
        Write-Verbose "group name: $group"
        Write-Verbose "number of members: $($oGrpMembers.count)" 
        $oGrpMembers | Select-Object @{L='groupname';E={$group}},@{L='memberName';E={$_.name}},distinguishedName,objectClass | Export-Csv -Path $reportPrivilegedGroupMembers -Append
    }
    write-host "admin group membership saved as $reportPrivilegedGroupMembers"
}
function get-eNEntraIDPrivilegedUsers {
    <#
    .SYNOPSIS
        Simple auditing script allowing to get the list of all users assgined to any Entra ID Role. 
    .DESCRIPTION
        This one is already using mgGraph.
    .EXAMPLE
        get full report on all roles that have any memebers

        .\get-eNEntraIDAdmins.ps1
    .EXAMPLE
        get full report sorted by a user name, script will not try to connect assuming you're already authenticated with a proper permissions

        .\get-eNEntraIDAdmins.ps1 -sortBy user -skipConnect

    .INPUTS
        None.
    .OUTPUTS
        csv report file.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241029
            last changes
            - 241029 initialized

        #TO|DO
        - add MFA status
        - is that hybrid account
    #>
    [CmdletBinding()]
    param (
        #show (detail) information on members and their roles or just a list of role members
        [Parameter(mandatory=$false,position=0)]
            [string][validateSet('user','role')]$sortBy='user',
        #assume you're already connected with mgGraph to skip authentication
        [Parameter(mandatory=$false,position=1)]
            [switch]$skipConnect,
        #export CSV file delimiter
        [Parameter(mandatory=$false,position=2)]
            [string][validateSet(',',';','default')]$delimiter='default'
        
    )

    $tenantDomain = (Get-MgOrganization).VerifiedDomains | ? IsDefault | Select-Object -ExpandProperty name
    $outFile=$tenantDomain + "_AdminReport_"+(get-date -Format 'yyMMdd')+'.csv'

    if(!$skipConnect) {
        try {
            Disconnect-MgGraph -ErrorAction Stop
        } catch {
            write-verbose $_.Exception
            $_.ErrorDetails
        }
        Write-Verbose "athenticate to tenant..."
        try {
            Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","RoleManagement.Read.Directory"
        } catch {
            throw "error connecting. $($_.Exception)"
            return
        }
    }
    Write-Host "getting roles and members (may take up to 5 min)..."
    $EntraRoles = Get-MgDirectoryRole

    $RoleMemebersList = `
    foreach($role in $EntraRoles) {
        $rDN=$role.DisplayName
        $rID=$role.Id
        $rMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $rID
        foreach($member in $rMembers) {
            $eUser = get-mgUser -UserId $member.Id -Property Id,userPrincipalName,AccountEnabled
            $eUser | Select-Object @{L='RoleName';E={$rDN}},`
                @{L='roleID';E={$rID}},`
                @{L='userID';E={$eUser.id}},`
                @{L='userName';E={$eUser.userPrincipalName}},`
                @{L='enabled';E={$eUser.AccountEnabled}}
                #@{L='MFA';E={'not yet implemented'}}
        }
    } 

    #unsupported in PS 5.1
    #$sortedMemebersList = ($sortBy -eq 'Role') ? ($RoleMemebersList | Sort-Object RoleName) : ($RoleMemebersList | Select-Object userName,userID,enabled,RoleName,roleID | Sort-Object userName)
    if($sortBy -eq 'Role') {
        $sortedMemebersList = $RoleMemebersList | Sort-Object RoleName
    } else {
        $sortedMemebersList = $RoleMemebersList | Select-Object userName,userID,enabled,RoleName,roleID | Sort-Object userName
    }

    $exportParam = @{
        NoTypeInformation = $true
        Encoding = 'UTF8'
        Path = $outFile
    }
    if($delimiter -ne 'default') { 
        $exportParam.add('delimiter',$delimiter)
    }
    $sortedMemebersList|export-csv @exportParam

    Write-Host -ForegroundColor Green "exported to .\$outFile.`ndone."

}
function get-eNReportADObjects {
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

    $wellKnownAdminSids = @("S-1-5-32-547","S-1-5-32-553","S-1-5-32-577","S-1-5-32-544","S-1-5-32-582","S-1-5-32-560","S-1-5-32-581","S-1-5-32-551",`
        "S-1-5-32-556","S-1-5-32-561","S-1-5-32-578","S-1-5-32-548","S-1-5-32-575","S-1-5-32-550","S-1-5-32-579","S-1-5-32-557","S-1-5-32-549","S-1-5-32-573","S-1-5-32-569","S-1-5-32-576",`
        "$domainSID-498","$domainSID-512","$domainSID-516","$domainSID-517","$domainSID-518","$domainSID-519","$domainSID-520","$domainSID-521","$domainSID-522","$domainSID-525","$domainSID-526","$domainSID-527")
    $dynamicAdminSIDgroups = @(
        "DnsAdmins",
            #EXCHANGE
        "Organization Management",
        "Recipient Management",
        "View-Only Organization Management",
        "Public Folder Management",
        "UM Management",
        "Help Desk",
        "Records Management",
        "Discovery Management",
        "Server Management",
        "Delegated Setup",
        "Hygiene Management",
        "Compliance Management",
        "Security Reader",
        "Security Administrator",
        "Exchange Servers",
        "Exchange Trusted Subsystem",
        "Managed Availability Servers",
        "Exchange Windows Permissions",
        "ExchangeLegacyInterop",
        "Exchange Install Domain Servers"
    )

    #check for admin priviledges. there is this strange bug [or feature (; ] that if you run console without
    #admin, some account do report 'enabled' attribute, some are not. so it's suggested to run as admin.
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
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
        $ADObjects = get-ADuser `
            -Filter {(lastlogondate -notlike "*" -OR lastlogondate -le $DaysInactiveStr)} `
            -Properties enabled,userPrincipalName,mail,distinguishedname,givenName,surname,samaccountname,displayName,description,lastLogonDate,PasswordLastSet
        Write-Verbose "found $(($ADObjects|Measure-Object).count) objects"
        $ADObjects = $ADObjects | select-object samaccountname,userPrincipalName,enabled,givenName,surname,displayName,mail,description,`
            lastLogonDate,@{L='daysInactive';E={if($_.LastLogonDate) {$lld=$_.LastLogonDate} else {$lld="1/1/1970"} ;(New-TimeSpan -End (get-date) -Start $lld).Days}},PasswordLastSet,`
            distinguishedname,@{L='parentOU';E={$rxParentOU.Match($_.distinguishedName).groups[1].value}}, @{L='isAdmin';E={$false}}
        #add check if user belongs to any privileged group
        foreach($ADuser in $ADObjects) {
            foreach($group in $wellKnownAdminSids.keys) {
                if($ADuser.memberOf -contains $wellKnownAdminSids[$group]) {
                    $ADuser.isAdmin = $true
                    break
                }
            }
            foreach($group in $dynamicAdminSIDgroups) {
                if($ADuser.memberOf -contains $group) {
                    $ADuser.isAdmin = $true
                    break
                }
            }
        }
        #final sorting and export
        $ADObjects | Sort-Object daysInactive,parentOU | Export-csv $exportCSVFile -NoTypeInformation -Encoding utf8
    } else {
        $ADObjects = get-ADComputer `
            -Filter {(lastlogondate -notlike "*" -OR lastlogondate -le $DaysInactiveStr)} `
            -Properties enabled,distinguishedname,samaccountname,displayName,description,lastLogonDate,PasswordLastSet
        Write-Verbose "found $(($ADObjects|Measure-Object).count) objects"
        $ADObjects |
            select-object samaccountname,enabled,displayName,description,`
                lastLogonDate,@{L='daysInactive';E={if($_.LastLogonDate) {$lld=$_.LastLogonDate} else {$lld="1/1/1970"} ;(New-TimeSpan -End (get-date) -Start $lld).Days}},PasswordLastSet,`
                distinguishedname,@{L='parentOU';E={$rxParentOU.Match($_.distinguishedName).groups[1].value}} | 
            Sort-Object daysInactive,parentOU |
            Export-csv $exportCSVFile -NoTypeInformation -Encoding utf8
    }
    Write-Verbose "results saved in '$exportCSVFile'"
}
function get-eNReportEntraUsers {
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
            #write-host 'testing error'
            write-verbose $_.Exception
            $_.ErrorDetails
        }
        Write-Verbose "athenticate to tenant..."
        #"Domain.ReadWrite.All" comes from get-mgDomain - but is not required.
        #"email" comes from get-mgDomain - and was double-requesting the authentication without this option
        #Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Directory.Read.All","Domain.Read.All","email"
        try {
            Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All","Domain.Read.All","UserAuthenticationMethod.Read.All","email","profile" -ErrorAction Stop
        } catch {
            throw "error connecting. $($_.Exception)"
            return
        }
    }
    Write-Verbose "getting connection info..."
    $ctx = Get-MgContext
    Write-Verbose "connected as '$($ctx.Account)'"
    if($ctx.Scopes -notcontains 'User.Read.All' -or $ctx.Scopes -notcontains 'AuditLog.Read.All' -or $ctx.Scopes -notcontains 'Domain.Read.All' -or $ctx.Scopes -notcontains 'UserAuthenticationMethod.Read.All') {
        throw "you need to connect using connect-mgGraph -Scopes User.Read.All,AuditLog.Read.All,Domain.Read.All","UserAuthenticationMethod.Read.All"
    } else {
    }
    try {
        $tenantDomain = (get-MgDomain -ErrorAction Stop | ? isdefault).id
    } catch {
        throw "error getting tenant information. $($_.Exception)"
    }
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
}
function get-eNReportMailboxes {
    <#
    .SYNOPSIS
        script for Exchange stats for reciepients and mailboxes. it is a part of a wider 'eNReport' package and may be used as a part of 
        general account audit or separately.
        script is useful for reporting supporting migration or cleanup projects.
    .DESCRIPTION
        script by default is making all type of checks: finds actual user UPN, gets detailed mailbox statistics and checks for delegated permissions.
        you can disable certain steps by using switches.
    .EXAMPLE
        .\get-eNReportMailboxes.ps1 

        by default it will ask you to authenticate with a web browser and then will get all mailboxes and recipient in the tenant and provide some basic stats.
    .EXAMPLE
        .\get-eNReportMailboxes.ps1 -skipConnect -inputFile .\tmp_recipients.csv -skipUPNs -skipMbxStats -skipDelegations

        this is a refresher for a chosen mailboxes from a previous run. it will skip connection to EXO using current session, load data from a file, skip UPN check, 
        mailbox statistics and permissions check.
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241220
            last changes
            - 241220 cleanup option, a bit of description
            - 240811 added delegated permissions to understand shared mailboxes (and security check),
                dived on steps, data refresh
                fixed mailbox type check
                other fixes
            - 240718 UPNs (for merge) and account status (for independent reports)
            - 240717 initialized

        #TO|DO
        * proper file description
    #>
    [CmdletBinding()]
    param (
        #skip connection - if you're already connected
        [Parameter(mandatory=$false,position=0)]
            [switch]$skipConnect,
        #load existing file with recipient list
        [Parameter(mandatory=$false,position=1)]
            [string]$inputFile,
        #skip UPNs
        [Parameter(mandatory=$false,position=2)]
            [switch]$skipUPNs,
        #skip mailbox statistics
        [Parameter(mandatory=$false,position=3)]
            [switch]$skipMbxStats,
        #skip delegation permissions
        [Parameter(mandatory=$false,position=4)]
            [switch]$skipDelegations,
        #do not remove partial tmp files (for debug)
        [Parameter(mandatory=$false,position=5)]
            [switch]$noCleanup
        
    )
    $VerbosePreference = 'Continue'

    if(!$skipConnect) {
        Disconnect-ExchangeOnline -confirm:$false -ErrorAction SilentlyContinue
        try {
            Connect-ExchangeOnline -ErrorAction Stop
        } catch {
            write-log "not connected to Exchange Online." -type error
            write-log $_.Exception -type error
            return
        }
    }

    #get some domain information 
    $domain = (Get-AcceptedDomain|? Default -eq $true).domainName
    Write-log "connected to $domain" -type info
    $outfile = "mbxStats-$domain-$(get-date -Format "yyMMdd-hhmm").csv"

    #'Recipients' is much wider, providing additional object infomration, thus starting from a getting all 'emails' in the tenant
    #load from file...
    if($inputFile) {
        try {
            write-log "loading $inputFile..." -type info
            $recipients = load-CSV $inputFile #header enformcement
        } catch {
            write-log "can't load data from $inputFile" -type error
            write-log $_.Exception -silent
            return
        }
    } else { #read from EXO
        write-log "getting general recipients stats..." -type info
        $recipients = get-recipient |
            Select-Object Identity,userPrincipalName,PrimarySmtpAddress,enabled,DisplayName,FirstName,LastName,RecipientType,RecipientTypeDetails,`
                @{L='emails';E={$_.EmailAddresses -join ';'}},delegations, ForwardingAddress, ForwardingSmtpAddress, `
                WhenMailboxCreated,LastInteractionTime,LastUserActionTime,TotalItemSize,ExchangeObjectId
        #save current step
        $recipients | export-csv -nti -Encoding unicode tmp_recipients.csv
    }
    $numberOfRecords = $recipients.count
    write-log "loaded $numberOfRecords records." -type info

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

    if(!$skipUPNs) {
        #some parameters make sens only for mailboxes - filter out non-mailbox enabled exchange objects and get the identity UPN for them
        write-log "getting UPNs from mailboxes..." -type info
        $recipients |? RecipientType -match 'userMailbox'| %{
            $mbx = Get-mailbox -identity $_.ExchangeObjectId
            $_.userPrincipalName = $mbx.userPrincipalName
            $_.enabled = -not $mbx.AccountDisabled
            $_.ForwardingAddress = $mbx.ForwardingAddress
            $_.ForwardingSmtpAddress = $mbx.ForwardingSmtpAddress
        }
        #save current step
        $recipients | export-csv -nti -Encoding unicode tmp_UPNs.csv
    } else {
        write-log "UPN check skipped." -type info
    }

    if(!$skipMbxStats) {
        #to know more about activity on a mailbox, get some last usage and basic size stats
        write-log "enriching mbx statistics..." -type info
        $recipients |? RecipientType -match 'userMailbox'| %{
            $stats = get-mailboxStatistics -identity $_.ExchangeObjectId
            $_.LastInteractionTime = $stats.LastInteractionTime
            $_.LastUserActionTime = $stats.LastUserActionTime
            $_.TotalItemSize = $stats.TotalItemSize
        }
        $recipients | export-csv -nti -Encoding unicode tmp_mbxStats.csv
    } else {
        write-log "mailbox stats skipped." -type info
    }

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

    #final results export
    $recipients | Sort-Object RecipientTypeDetails,identity | Export-Csv -nti -Encoding unicode -Path $outfile

    If(-not $noCleanup) {
        write-log "clean up..." -type info
        Remove-Item tmp_recipients.csv -ErrorAction SilentlyContinue
        Remove-Item tmp_UPNs.csv -ErrorAction SilentlyContinue
        Remove-Item tmp_mbxStats.csv -ErrorAction SilentlyContinue
    } else {
        write-log "partial files kept. look for 'tmp_*.csv' files." -type info
    }

    write-log "$outfile written." -type ok

}
function join-eNReportHybridUsersInfo {
    <#
    .SYNOPSIS
        merge AD, Entra ID and Exchange reports into single user activity report generated by 3 other script from
        the same package.

        this script is to cobine reports from 3 sources into a single view: AD, Entra ID and Exchange Online.
        such a view grants ability to decide on the preparations for migration - which objects require to be 
        amended, which are synced or how they will merge during the sync. 

    .DESCRIPTION
        using outputs from get-eNReportADObjects.ps1 (AD), get-eNReportEntraUsers.ps1 (EntraID) and get-eNReportMailboxes.ps1 (EXO)
        
        the most difficult part is to merge the outputs matching the objects. there is no 'Table' type in PowerShell
        allowing to lookup and update records. I wrote a simple equivalent based on MetaVerse concept (aka DB Table). 
        MetaVerse is a global table containing all data from all sources and allows to lookup and update entries.

        let's assume scenario that you are preparing for enabling Cloud Sync. If there is a AD user:
        sn: Surname1
        gn: Givenname1
        displayName: Givenname1 Surname1
        mail: givenname1.surname1@company.com
        UPN: gsurname1@comapny.local

        and equivalent user in EntraID:
        sn: Changed-Surname
        gn: Givennam1
        displayName: Changed-Surname Givenname1
        mail: givenname1.changed-surname@company.com
        UPN: givenname1.surname1@company.com

        it's very difficult to findout pairs to verify how to amend/fix user object. analysing is quite time extensive. 
        this script allows you to create a unified view matching on different attributes. you may create several reports
        (aka views) by matching by different attributes or 'any' match allowing to find matches on different attributes 
        - e.g. on example above AD.mail - match EntraID.UPN . 

        MATCHING
        EXO objects are easy to match as every recipient has an EID object so there is no confusion.
        actual challenge is with matching AD and EID objects - especially when there is no actual hybrid sync. Users
        may have duplicates, different names, parcial information etc. that is why the script is trying to use different 
        set of attributes to find a match even if they are not really on sync.

        *****
        although other functions from the package are independed, this one is using eNLib. no one is going to use this
        script anyways, and it's so much easier for me to reuse these functions. actually I had to extend some lib functions
        so only the newest eNLib version is compatible. 

        short instruction:
        - generate outputs from desired systems (AD, EntraID, EXO)
        - combine the reports with the script into a single view

    .EXAMPLE
        .\join-eNReportHybridUsersInfo.ps1 -inputCSVAD .\ADUsers-company.local-241111-1026.csv -inputCSVEntraID .\EntraUsers-company.com-241111-1028.csv

        combines a report made from a local 'company.local' domain with a EntraID information for 'company.com' tenant. 
    .INPUTS
        CSV report from other scripts
    .OUTPUTS
        merged report
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241223
            last changes
            - 241223 matching for guest accounts, better AD-EID matching (dupes handling)
            - 241220 'any' fixed, lots of changes to matching and sorting, export only for chosen files... 
            - 241210 mutliple fixes to output, daysinactive, dupe detection. dupes are still not matched entirely properly.. that will require some additional logic
            - 241126 massive logic fixes. tested on 3 sources... still lots to be done but starting to work properly
            - 241112 whole logic changed - MetaVerse functions added and whole process is using MV to operate on data
            - 240718 initiated as a bigger project, extended with Exchange checking
            - 240627 add displayname as matching attribute. forceHybrid is for now default and parameter doesn't do anything
            - 240520 initialized

        #TO|DO
        ** dups handling - this is difficult one, how to create a proper logic to match...
        ** BUILD SCHEMA - currently it's static
        * ability to choose between static and dynamic schema... or simply intorduce 'views' known from DBs - output 
            should be a 'view' from entire MetaVerse while now it's the same
        * edge scenarios - eg. the same UPN on both sides, but account is not hybrid; maybe some other i did not expect?
        * change hybrid user detection / currently matching is ONLY in forced hybrid... which should not be the case
        * change time values to represent the same 'never' value
        * what is 'identity' attribute? it's not being populated
        *Idea so it works exactly like MV - all tables are kept separately until the very export. each table should be expanded with a reference column
        pointing to an object from other table. then implement 'view' or 'export' that is creating one single file with different options
        such as 'only matched', 'all', etc. 
        * auto-fix UPN suffix for soft matching (domain.local to domain.com) - to enforce pseudo-hybrid matching
        * some fileds are non-mandatory while executing - such as EXO delegations - but mandatory here. should allow for flexibility
        
    #>
    [CmdletBinding()]
    param (
        #output file from AD
        [Parameter(mandatory=$false,position=0)]
            [string]$inputCSVAD,
        #output file from Entra ID
        [Parameter(mandatory=$false,position=1)]
            [string]$inputCSVEntraID,
        #output file from Exchange Online 
        [Parameter(mandatory=$false,position=2)]
            [string]$inputCSVEXO,
        #force match for non-hybrid users - low accuracy... key attribute to match the users, default userPrincipalName
        [Parameter(mandatory=$false,position=3)]
            [validateSet('userPrincipalName','mail','displayName','all','hybridOnly')]
            [string]$matchBy = 'all',
        #open file after conversion
        [Parameter(mandatory=$false,position=4)]
            [alias('run')]
            [switch]$openOnConversion = $true
        
    )
    $VerbosePreference = 'Continue'
    # Function to update information from different data sources
    function Update-MetaverseData {
        param (
            #metaverse object to work on
            [Parameter(Mandatory,Position = 0)]
                [hashtable]$mv,
            #key object ID, 
            [Parameter(Mandatory,Position = 1)]
                [int]$objectID,
            #object with new values
            [Parameter(Mandatory,Position = 2)]
                [PSobject]$dataSource
        )

        if(-not $mv.ContainsKey($objectID)) {
            # If the objectID with a given ID does not exist in the metaverse - thow an error
            throw -1
        }

        # Merge attributes for the specified person
        foreach ($propertyName in ( ($dataSource.psobject.Properties | ? memberType -eq 'NoteProperty')).name) {
            $mv[$objectID][$propertyName] = $dataSource.$propertyName
        }
        Write-debug "metaverse object $objectID has been updated"
    }

    function Add-MetaverseData {
        param (
            #metaverse object to work on
            [Parameter(Mandatory,Position = 0)]
                [hashtable]$mv,
            #object with new values
            [Parameter(Mandatory,Position = 1)]
                [PSObject]$dataSource
        )

        function new-objectID {
            $newID = 0
            if($mv.count -eq 0) { return 0 } #mv is empty - initialize
            foreach($mvOID in $mv.Keys) {
                if($mvOID -gt $newID) { $newID = $mvOID }
            }
            return ($newID + 1)
        }

        $newID = new-objectID
        $mv[$newID] = @{} #initialise a new entry
        #FIX change to externally defined object schema
        $newEntry = @{
            "AD_samaccountname"="";"AD_userPrincipalName"="";"AD_enabled"="";"AD_givenName"="";"AD_surname"="";"AD_displayName"="";"AD_mail"="";"AD_description"="";"AD_lastLogonDate"="";"AD_daysInactive"=23000;"AD_PasswordLastSet"="";"AD_distinguishedname"="";"AD_parentOU"="";
            "DisplayName"="";"UserType"="";"AccountEnabled"="";"GivenName"="";"Surname"="";"UserPrincipalName"="";"Mail"="";"MFAStatus"="";"Hybrid"="";"LastLogonDate"="";"LastNILogonDate"="";"licenses"="";"Id"="";"daysInactive"=23000;
            "EXO_Identity"="";"EXO_DisplayName"="";"EXO_FirstName"="";"EXO_LastName"="";"EXO_RecipientType"="";"EXO_RecipientTypeDetails"="";"EXO_emails"="";"EXO_WhenMailboxCreated"="";"EXO_userPrincipalName"="";"EXO_enabled"="";"EXO_delegations"="";"EXO_LastInteractionTime"="";"EXO_LastUserActionTime"="";"EXO_TotalItemSize"="";"EXO_ExchangeObjectId"=""
        } 
        # prepare new entry rewriting object property values to hashtable 
        foreach ($propertyName in ( ($dataSource.psobject.Properties | ? memberType -eq 'NoteProperty')).name) {
            
            #TODO - add update of chosen attributes only, not the whole object
            $newEntry.$propertyName = $dataSource.$propertyName
        }
        $mv[$newID] = $newEntry
        Write-debug "metaverse object ID $newID has been added to MV table"
    }

    # Function to search the metaverse for a specific key-value match
    function Search-MetaverseData {
        <#
        .SYNOPSIS
            Search the Metaverse table
        .DESCRIPTION
            here be dragons
        .EXAMPLE
            Search-MetaverseData -mv $myMetaVerse -......
        
            
        .INPUTS
            None.
        .OUTPUTS
            returns an object containing objectID (key), attribute with matched value and the value itself
            @{
                mvID = $mvKey
                elementProperty = $elementKey
                elementValue = $mvElement[$elementKey]
            }
        .LINK
            https://w-files.pl
        .NOTES
            nExoR ::))o-
            version 241106
                last changes
                - 241106 initialized
        
            #TO|DO
            - description
            - different types of varaibles [int/string]
            - lookup for substring and whole words
        #>
        [CmdletBinding(DefaultParameterSetName = 'any')]
        param (
            #metaverse object to search thru
            [parameter(Mandatory,position=0)]
                [validateNotNullOrEmpty()]
                [hashtable]$mv,
            #look for value on ANY column (super-soft match)
            [Parameter(mandatory=$false,position=1,ParameterSetName = 'any')]
                [switch]$anyColumn = $true,
            #name of the stored object parameter to use in search. 
            [parameter(position=1,ParameterSetName = 'single')]
                [string]$columnName,
            #substring to search for
            [parameter(Mandatory,position=1,ParameterSetName = 'any')]
            [parameter(Mandatory,position=2,ParameterSetName = 'single')]
                [string]$lookupValue,
            #pass hashtable to be used for search
            [Parameter(Mandatory,position=1, ParameterSetName = 'byObject')]
                [PSObject]$lookupTable
        )

        if($PSCmdlet.ParameterSetName -eq 'single') {
            $lookupTable = @{ 
                $columnName = $lookupValue 
            }
        }

        $foundMatches = @()
        foreach ($mvKey in $mv.Keys) {
            $mvElement = $mv[$mvKey]

            if($PSCmdlet.ParameterSetName -eq 'any') {
                foreach($lookupMVColumn in $mvElement.Keys) {
                    if ($mvElement[$lookupMVColumn] -eq $lookupvalue) {
                        $returnedResult = @{
                            mvID = $mvKey
                            elementProperty = $lookupMVColumn
                            elementValue = $mvElement[$lookupMVColumn]
                        }
                        [array]$foundMatches += $returnedResult
                    }
                }
            } else {            
                foreach($lookupColumn in $lookupTable.Keys) {
                    if(-not $mvElement.ContainsKey($lookupColumn)) { #key exists check
                        Write-Debug "WARNING. column not found: $lookupColumn"
                        continue
                    } 
                    $lookupValue = $lookupTable[$lookupColumn]
                    if([string]::isNullOrEmpty($lookupValue)) { #lookup value must not be null
                        #maybe some warning info here?
                        continue 
                    }
                    if ($mvElement[$lookupColumn] -eq $lookupvalue) {
                        $returnedResult = @{
                            mvID = $mvKey
                            elementProperty = $lookupColumn
                            elementValue = $mvElement[$lookupColumn]
                        }
                        [array]$foundMatches += $returnedResult
                        #FIX - it should just add a mach, but do not allow to make a dupe. for now - first match exist
                        #return $foundMatches
                    } 
                }
            }
        }

        return $foundMatches
    }


    #$VerbosePreference = 'Continue'
    $exportCSVFile = "CombinedReport-{0}.csv" -f (get-date -Format "yyMMdd-hhmm")
    #this headers are to enforce strict header check during import. it could be safely minimized leaving only part of the columns ... but then the final export will have empty values
    $headerEntraID = @('id','displayname','givenname','surname','accountenabled','userprincipalname','mail','userType','Hybrid','givenname','surname','userprincipalname','userType','mail','daysInactive')
    $headerAD = @('samaccountname','userPrincipalName','enabled','givenName','surname','displayName','mail','description','daysInactive')
    $headerEXO =  @('RecipientType','RecipientTypeDetails','emails','delegations','WhenMailboxCreated','LastInteractionTime','LastUserActionTime','TotalItemSize','ExchangeObjectId')

    #report should always have all the fields - metafile should be a static schema
    $metaverseUserInfo = @{}

    #region reading inputs
    Write-log "loading CSV files.." -type info
    $reports = 0
    if($inputCSVEntraID) {
        $EntraIDData = load-CSV $inputCSVEntraID `
            -header $headerEntraID `
            -headerIsCritical
        $reports++
        if([string]::isNullOrEmpty($EntraIDData)) {
            exit
        }
    }
    if($inputCSVAD) {
        $ADData = load-CSV $inputCSVAD `
            -header $headerAD `
            -headerIsCritical `
            -prefix 'AD_'
        $reports++
        if([string]::isNullOrEmpty($ADData)) {
            exit
        }
    }
    if($inputCSVEXO) {
        $EXOData = load-CSV $inputCSVEXO `
            -header $headerEXO `
            -headerIsCritical `
            -prefix 'EXO_'
        $reports++
        if([string]::isNullOrEmpty($EXOData)) {
            exit
        }
    }
    if($reports -lt 2) {
        Write-Log "at least two reports are required for merge" -type error
        return
    }
    #endregion

    #region start from populating EntraID
    if($EntraIDData) {
        Write-Verbose "filling EntraID user info..."
        foreach($entraIDEntry in $EntraIDData) {
            Add-MetaverseData -mv $metaverseUserInfo -dataSource $entraIDEntry
        }
    }
    #endregion

    #region populate AD data
    if($ADData) {
        Write-Verbose "adding AD user info..."
        foreach($ADuser in $ADData) {
            #check if user already exists from Entra source
            $matchedEID = $false
            if($EntraIDData) {
    #if 'hybrid' flag - check onpremisessid to match 

                [array]$entraFound = Search-MetaverseData -mv $metaverseUserInfo -lookupTable @{ 
                    userPrincipalName = $ADuser."AD_userPrincipalName"
                    displayName       = $ADuser."AD_displayName"
                    mail              = $ADuser."AD_mail"
                }
                #match may be on several attributes for the same object or for several different objects (mvIDs)
                #so I'm checking how many unique IDs are found
                if(($entraFound | Select-Object mvID -Unique).count -gt 1) {
                    write-verbose "AD: $($entraFound[0].elementValue): duplicate found on $($entraFound.elementProperty -join ',') attributes."
                    if($entraFound|? elementproperty -eq 'userPrincipalName') { #difficult to choose, but UPN matching is imho the strongest. then mail. displyname is rather a 'soft match'and may have many duplicates
                        $matchedEID = $true
                        Update-MetaverseData -mv $metaverseUserInfo -dataSource $ADuser -objectID ($entraFound |? elementProperty -eq 'userPrincipalName').mvID
                    }elseif($entraFound|? elementProperty -eq 'mail') {
                        #DUPE RISK - with guest accounts
                        $matchedEID = $true
                        Update-MetaverseData -mv $metaverseUserInfo -dataSource $ADuser -objectID ($entraFound |? elementProperty -eq 'mail').mvID
                    }
                } 

                if(($entraFound | Select-Object mvID -Unique).count -eq 1) { 
                    $matchedEID = $true
                    Write-debug 'matched-adding'
                    Update-MetaverseData -mv $metaverseUserInfo -dataSource $ADuser -objectID $entraFound[0].mvID
                }
            }
            if(-not $matchedEID) {
                Write-debug 'non-ad-adding'
                Add-MetaverseData -mv $metaverseUserInfo -dataSource $ADuser
            }
        }
    }
    #endregion

    #region populate EXO data
    Write-Verbose "adding EXO mailboxes info..."
    foreach($recipient in $EXOData) {
        $userFound = $false
    #    if($recipient.EXO_userPrincipalName) { #only mailboxes have UPNs - user mailboxes
            [array]$exoFound = Search-MetaverseData -mv $metaverseUserInfo -lookupTable @{ 
                userPrincipalName = $recipient."EXO_userPrincipalName"      #regular user
                mail = $recipient."EXO_PrimarySMTPAddress"                  #guest users
                AD_userPrincipalName = $recipient."EXO_userPrincipalName"   #ad synced users
            }
            if($exoFound.Count -gt 0) {
                $userFound = $true
                if(($exoFound | Select-Object mvID -Unique).Count -gt 1) {
                    write-verbose "EXO: $($exoFound[0].elementValue): duplicate records for EXO matching"
                }
                #in case duplicate was found - it will overwrite the first found entry - this must be improved
                Update-MetaverseData -mv $metaverseUserInfo -dataSource $recipient -objectID $exoFound[0].mvID
            } 
    #    } else { #match guest accounts

    #    }
        if(!$userFound) {
            Add-MetaverseData -mv $metaverseUserInfo -dataSource $recipient
        }
    }
    #endregion

    #export all results, extending with Hybrid_daysInactive attribute being lower of the comparison between EID and AD
    #select is enforced as I want the parameters provided in a given order
    $finalHeader = @()
    if($EntraIDData) { 
        $finalHeader += @("DisplayName","UserType","AccountEnabled","GivenName","Surname","UserPrincipalName","Mail","MFAStatus","Hybrid","LastLogonDate","LastNILogonDate","licenses","Id","daysInactive") 
    }
    if($ADData) { 
        $finalHeader += @("AD_samaccountname","AD_userPrincipalName","AD_enabled","AD_givenName","AD_surname","AD_displayName","AD_mail","AD_description","AD_lastLogonDate","AD_daysInactive","AD_PasswordLastSet","AD_distinguishedname","AD_parentOU")
    }
    if($EXOData) { 
        $finalHeader += @("EXO_PrimarySMTPAddress","EXO_DisplayName","EXO_FirstName","EXO_LastName","EXO_RecipientType","EXO_RecipientTypeDetails","EXO_emails","EXO_delegations","EXO_ForwardingAddress", "EXO_ForwardingSmtpAddress","EXO_WhenMailboxCreated","EXO_userPrincipalName","EXO_enabled","EXO_Identity","EXO_LastInteractionTime","EXO_LastUserActionTime","EXO_TotalItemSize","EXO_ExchangeObjectId")
    }

    $metaverseUserInfo.Keys | %{ 
        $metaverseUserInfo[$_] |
            Select-Object $finalHeader |
            Select-Object *,@{L='Hybrid_daysInactive';E={($_.daysInactive,$_.AD_daysInactive|Measure-Object -Minimum).minimum}} |
            Sort-Object Hybrid_daysInactive,displayName,AD_displayName,EXO_DisplayName -Descending
    } | Export-Csv -Encoding unicode -NoTypeInformation $exportCSVFile

    Write-Log "merged report saved to '$exportCSVFile'." -type ok
    if($openOnConversion) {
        $params = @{
            CSVfileName = $exportCSVFile
            openOnConversion = $true
        }
        write-log "converting..."
        &(convert-CSV2XLS @params)
    }
    write-log "done." -type ok

}