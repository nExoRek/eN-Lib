<#
.SYNOPSIS
    set of functions for auditing and reporting on accounts in AD, EID and EXO mailboxes. abilty to generate provileged users report,
    merge data to have a big picture on the accounts for migrations, cleanups or regular audits.
    privileged accounts reports, service plans loookup, MFA report and other functions are included. 

    eNLib is required for CSV-XLS conversions.
.DESCRIPTION
    module is in early stages - there is some mess with functions and reporting, lack of unification. some parameters and behaviour may change before 
    mature version is ready.

    module contains several auditing functions and each of them may be useful to gather some interesting data. yet there are three main functions to 
    generate reports from three different sources about the identities:
        - get-eNReportADObjects - for AD objects
        - get-eNReportEntraUsers - for EntraID
        - get-eNReportEXOMailboxes - for Exchange Online mailboxes
    after getting results from two or three of above sources, there is a function to merge data and generate a combined report:
        - join-eNReportHybridUserInfo

    join-eNReportHybridUsersInfo -inputCSVAD .\ADUsers-w-files.pl-250124-1239.csv -inputCSVEntraID .\EntraUsers-w-files.pl-250124-0254.csv -inputCSVEXO .\mbxstats-w-files.pl-250124-0256.csv
    command will join all three reports and generate a final one, containing combined data. 

    Analysing data in combined report
    All the rest is setting up proper filters in Excel file. Below some hints and explanations to columns and file structure (assuming all 3 sources were used). Because of vast number of 
    scenarios and queries it is impossible to describe all combinations. Below are hints and suggestions – we need to define some set of default queries to be reported leaving some space 
    for creativity for extra information.
    Many columns have value as a confirmation if matching was proper, there are no discrepancies in naming or if you need to use value for further investigation – all names, display names 
    and IDs. These columns may be hidden when creating some final report to minimize complexity of the view.

    General
    •	Columns with no prefix comes from EntraID
    •	Columns with AD_ prefix comes from AD
    •	Columns with EXO_ prefix comes from Exchange Online
    •	Value ‘23000’ in ‘daysInactive’ is filled by script for empty values for easiness of sorting later in Excel and basically means ‘no value’
    •	Values similar to ‘1970-01-01’ or ‘1600-12-31’ or ‘20112’ comes from Microsoft way or providing timestamps in systems and are equivalent of my ‘23000’
        meaning that timestamp has never been set (never used)
    •	Matching the names is set to automatic – meaning that it doesn’t matter which scenario is valid for customer, it will try to find corresponding object 
        between AD and EID (Exchange mailboxes will always have EID user). Script it trying to match by UserPrincipalName, email and displayName . If any of the attributes 
        does not match, you will find the same user twice (for AD and EID) 
    
    EntraID Columns
    UserType:	there are two types – guest and member. It’s a main filter to use dependently on type of account for review.
        It’s good to take a look on guest accounts in the tenant to see if there are any anomalies – e.g. unexpectedly big amount of guest may be a signal of oversharing, 
        accounts not used for a long time could be cleaned out. 
        When filter is enabled for member accounts it will allow to review all user-related queries such as unused accounts, licenses, mailbox sizes, accounts that are 
        not synced etc.
    AccountEnabled:	good filter to use in combination with licensing and activity – e.g. ‘disabled accounts with licenses’ are potentially good way to optimise licenses 
        assignment and ‘enabled account not used for <number> of days’ is a good way to detect unused accounts. Similarly ‘enabled accounts with MFA status disabled’ allow to fish out unsecure accounts.
    UserPrincipalName:	useful to detect incorrect UPNs, especially in tenants with numerous domain suffixes configured
    MFAStatus:	main column allowing to fish unsecured accounts – good to combine with AccountEnabled. Mind that EAM MFA is undetectable (Microsoft bug I reported to support.
        EAM is in preview, only us and FNTC has it configured as for the date of this document).
    LastLogonData, LastNILogonDate, daysInactive:	there are two types of logon dates reported – Interactive and Non-Interactive. Dates are useful for some heavy 
        troubleshooting when trying to establish what is going on with the account. In regular report both columns may be hidden and ‘daysInactive’ is calculated value of days
        the account reported any kind of activity on any of the logon type. Similar fields exist for AD_ allowing to detect situations such as:
        ‘account not used in AD but is synchronized, and used in EID, so it can not be disabled’
        ‘account is used in AD but not in EID, so maybe license is not necessary’
    Licenses:	all assigned licenses – allows to quickly prepare license report, good to combine with ‘AccountEnabled’
    Hybrid:	‘TRUE’ means that account is synchronized from AD. Allows to detect improper synchronization – accounts that exist on both sides but are not synced.
    
    AD Columns
    AD_UserPrincipalName:	allows to detect incorrect UPN values – good during preparation to migration, to fix UPNs to tenant domain
    AD_enabled:	great in combination with other columns allowing to query:
        unused accounts but enabled
        improper location of disabled accounts 
        account enabled on one side (AD/EID) 
    AD_lastLogonDate, AD_daysInactive:	actual date and calculated value in number of days till now for activity queries
    AD_passwordLastSet:	for queries allowing to understand when password has been set for the user for the last time. Most interesting is empty value as it means
        that account has never been used.
    AD_parentOU	useful for sorting view by the location and to quickly detect location anomalies for accounts. Useful for general cleanup and in scenarios where 
        synchronization is filtered to particular OUs – moving account out of sync scope will unsync account.
    
    Exchange Online columns
    EXO_RecipientTypeDetails:	by default report is showing all recipients, not only mailboxes. It may be used to filter view to:
        see how many contacts or groups are in the tenant
        check resources (RoomMailbox,EquipementMailbox)
        limit view only to mailboxes (UserMailbox)
        detect all mailboxes that are configured as shared on are user mailboxes but used as shared – combination of UserMailbox and SharedMailbox views
    EXO_emails:	all aliases on the mailbox – very useful in migration or synchronization projects
    EXO_delegations:	useful for investigations on shared mailboxes. E.g. UserMailbox with numerous delegations is probably a SharedMailbox which may allow for 
        conversion and removal of the license. Other way around – SharedMailbox with no delegations may mean that mailbox is unused.
    EXO_forwardingAddress, EXOforwardingSMTPAddress:	may help detect leakage of corporate emails – if all emails are forwarded on external address. Currently 
        it is reported also by Office Defender, but good to check periodically.
    EXO_enabled:	this is a very difficult as there are numerous scenarios when account status and mailbox status may differ. Can’t explain these anomalies 
        at the moment, but these are interesting scenarios that may be helpful in rare investigations.
    EXO_lastInteractionTime, EXO_lastUserActionTime:	there is no simple definition of ‘unused mailbox’ as some mailboxes may be archives, forwarders or very rarely
        used in certain situations (e.g. some event once a year). In combination with other columns may be useful in investigations during migration projects to detect unused mailboxes but these pretty much always require consulting with customer.
    
    Other
    Hybrid_daysInactive:	is lower value of daysInactive from EID and AD_daysInactive from AD. Allows to quickly filter totally unused accounts.

.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250329
        last changes
        - 250329 service plan info removed -> new library eNGBL created
        - 250209 fixed join-report, MFAreport extended to get full info from two commandlets, other fixes
        - 250206 included get-eNServicePlanInfo, module definition amendmend, MFA report function added
        - 250203 isAdmin for EID, some optmization for MFA check
        - 250131 isAdmin for AD added... not sure if join function will handle it... 
        - 250125 initialized

    #TO|DO
    * application permissions for EID
    * AD permissions crawler to detect non-standard delegations
    * code optimization
    * ent-size tenant queries (currently unsupported)
    * PS version check functions to replace missing #requires
    * join must handle all attributes by default (no static list)
    * unify parameters and behaviour:
        - auto excel conversion and opening for all functions 
        - same connect/skip experience
        - output file naming convention
        - progress bar for all functions
        - re-think write-log and unify
    * is it possible to check Conditional Access policies enforcing MFA?        
#>
Function Get-MFAMethods {
<#
.SYNOPSIS
    internal function for this module - get the details on configured MFA methods of the single user
.NOTES
    nExoR ::))o-
    version 250209
        last changes
        - 250209 module-verified 
#>
    param(
        [Parameter(Mandatory, Position = 0)] 
            [string]$userId,
        [parameter(Mandatory = $false, Position = 1)]
            [switch]$onlyStatus
    )

    process {
        # Create MFA details object
        $mfaMethods  = [PSCustomObject][Ordered]@{
            MFAstatus = "disabled"
            softwareAuth = $false
            authApp = $false
            authDevice = ""
            phoneAuth = $false
            authPhoneNr = ""
            fido = $false
            fidoDetails = ""
            helloForBusiness = $false
            helloForBusinessDetails = ""
            emailAuth = $false
            SSPREmail = ""
            tempPass = $false
            tempPassDetails = ""
            passwordLess = $false
            passwordLessDetails = ""
        }

        write-debug "MFA Methods - Get-MgUserAuthenticationMethod"
        try {
            [array]$mfaData = Get-MgUserAuthenticationMethod -UserId $userId -ErrorAction Stop
        } catch {
            $mfaMethods.MFAstatus = 'error'
            return $mfaMethods
        }
        if($onlyStatus) {
            Write-Debug "MFAMethods - only status"
            if($mfaData[0].AdditionalProperties["@odata.type"] -eq "#microsoft.graph.passwordAuthenticationMethod" -and $mfaData.Count -eq 1) {
                return "disabled"
            } elseif($mfaData.Count -gt 1) {
                return "enabled"
            } else {
                return "error"
            }
        }        
        ForEach ($method in $mfaData) {
            Switch ($method.AdditionalProperties["@odata.type"]) {
<#                "#microsoft.graph.passwordAuthenticationMethod" { 
                    # Password
                    # When only the password is set, then MFA is disabled.
                    if ($mfaMethods.MFAstatus -ne "enabled") {$mfaMethods.MFAstatus = "disabled"}
                }
#>
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
                    # Microsoft Authenticator App
                    $mfaMethods.authApp = $true
                    $mfaMethods.authDevice += "[{0}]" -f $method.AdditionalProperties["displayName"] 
                    $mfaMethods.MFAstatus = "enabled"
                } 
                "#microsoft.graph.phoneAuthenticationMethod" { 
                    # Phone authentication
                    $mfaMethods.phoneAuth = $true
                    $mfaMethods.authPhoneNr += "[{0}]" -f ($method.AdditionalProperties["phoneType", "phoneNumber"] -join ' ')
                    $mfaMethods.MFAstatus = "enabled"
                } 
                "#microsoft.graph.fido2AuthenticationMethod" { 
                    # FIDO2 key
                    $mfaMethods.fido = $true
                    $mfaMethods.fidoDetails += "[{0}]" -f $method.AdditionalProperties["model"]
                    $mfaMethods.MFAstatus = "enabled"
                } 
                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
                    # Windows Hello
                    $mfaMethods.helloForBusiness = $true
                    $mfaMethods.helloForBusinessDetails += "[{0}]" -f $method.AdditionalProperties["displayName"]
                    $mfaMethods.MFAstatus = "enabled"
                } 
                "#microsoft.graph.emailAuthenticationMethod"                   { 
                    # Email Authentication
                    $mfaMethods.emailAuth = $true
                    $mfaMethods.SSPREmail += "[{0}]" -f $method.AdditionalProperties["emailAddress"] 
                    $mfaMethods.MFAstatus = "enabled"
                }               
                "microsoft.graph.temporaryAccessPassAuthenticationMethod"    { 
                    # Temporary Access pass
                    $mfaMethods.tempPass = $true
                    $mfaMethods.tempPassDetails += "[{0}]" -f $method.AdditionalProperties["lifetimeInMinutes"]
                    $mfaMethods.MFAstatus = "enabled"
                }
                "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { 
                    # Passwordless
                    $mfaMethods.passwordLess = $true
                    $mfaMethods.passwordLessDetails += "[{0}]" -f $method.AdditionalProperties["displayName"]
                    $mfaMethods.MFAstatus = "enabled"
                }
                "#microsoft.graph.softwareOathAuthenticationMethod" { 
                    # ThirdPartyAuthenticator
                    $mfaMethods.softwareAuth = $true
                    $mfaMethods.MFAstatus = "enabled"
                }
            }
        }
        Return $mfaMethods
    }
}
function get-MFAReport {
<#
.SYNOPSIS
    get the MFA status of a particular user or all users in the tenant
.DESCRIPTION
    get-MgReportAuthenticationMethodUserRegistrationDetail is providing a nice report but lacking some actual details and works only for 
    enabled users. 
    get-MgUserAuthenticationMethod is not providing default 2FA configured...
    so the only way to have everything is to combine both methods.

    this function is a wrapper for aforementioned functions and using internal get-MFAMethods. 
.EXAMPLE
    get-eNAuditorMFAReport

    prepares a report for all users in a tenant
.EXAMPLE
    get-eNAuditorMFAReport -userId 12de9a48-99d0-4ce5-be38-0cc79c876c33

    prepares a report for a user with provided objectID
.EXAMPLE
    get-eNAuditorMFAReport -userPrincipalName nexor@w-files.pl

    prepares a report for a user with a UPN nexor@w-files.pl
.NOTES
    nExoR ::))o-
    version 250211
        last changes
        - 250211 as it turned out, reportdetails is not working for accounts lacking license, not disabled.. that had conseqences.
        - 250209 extended report from both commandlets
#>
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        #username provided as ID
        [Parameter(mandatory=$false,position=0,ParameterSetName='uID')]
            [string]$userId,    
        #username provided as UPN
        [Parameter(mandatory=$false,position=0,ParameterSetName='upn')]
            [string]$userPrincipalName,
        #no username - check for all users
        [Parameter(mandatory=$false,position=0,ParameterSetName='default')]
            [switch]$all = $true,
        #extended MFA information
        [Parameter(mandatory=$false,position=1,ParameterSetName='uID')]
        [Parameter(mandatory=$false,position=1,ParameterSetName='upn')]
        [Parameter(mandatory=$false,position=1,ParameterSetName='default')]
            [switch]$extendedMFAInformation,
        #automatically convert to Excel and open
        [Parameter(mandatory=$false,position=2,ParameterSetName='default')]
            [switch]$run
    )

    $VerbosePreference = "continue"
    try {
        $ctx = Get-MgContext -ErrorAction Stop
    } catch {
        $_.Exception
        return
    }
    Write-Verbose "connected as $($ctx.account)."
    if(-not $ctx) {
        try { 
            Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All","openid","profile" -ErrorAction Stop
        } catch {
            $_.Exception
            return
        }
    }

    if($PSCmdlet.ParameterSetName -eq 'uID') { #single check - by uID
        $EIDuser += Get-MgUser -userId $userId -Property accountEnabled,userPrincipalName,Id,LicenseAssignmentStates -ErrorAction Stop
        if(-not $EIDUser) {
            Write-Verbose "no user(s) found."
            return
        }
    }
    if($PSCmdlet.ParameterSetName -eq 'upn') { #single check - by UPN
        Write-Debug "checking for $userPrincipalName"
        $EIDuser += Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'" -Property accountEnabled,userPrincipalName,Id,LicenseAssignmentStates -ErrorAction Stop
        if(-not $EIDUser) {
            Write-Verbose "no '$userPrincipalName' found."
            return
        }
    }
    if($PSCmdlet.ParameterSetName -eq 'upn' -or $PSCmdlet.ParameterSetName -eq 'uID') {
        Write-Verbose "getting basic details for '$($EIDuser.userPrincipalName)'"
        if($EIDuser.LicenseAssignmentStates) {
            try {
                $mfaStatus = Get-MgReportAuthenticationMethodUserRegistrationDetail -Filter "userPrincipalName eq '$($EIDuser.userPrincipalName)'" -ErrorAction SilentlyContinue    
            } catch {
                Write-Verbose $_.Exception
                return
            }
        } else {
            $mfaStatus = [PSCustomObject]@{
                UserDisplayName = $EIDuser.userPrincipalName
                UserPrincipalName = $EIDuser.userPrincipalName
                Id = $EIDuser.Id
                AccountEnabled = $EIDUser.AccountEnabled
                IsAdmin = ''
                IsMfaCapable = ''
                IsMfaRegistered = ''
                IsPasswordlessCapable = ''
                IsSsprCapable = ''
                IsSsprEnabled = ''
                IsSsprRegistered = ''
                LastUpdatedDateTime = ''
                MethodsRegistered = ''
                IsSystemPreferredAuthenticationMethodEnabled = ''
                SystemPreferredAuthenticationMethods = ''
                UserPreferredMethodForSecondaryAuthentication = ''
                AdditionalProperties = 0
            }
        }
        $mfaDetails = Get-MFAMethods -userId $EIDuser.Id
        if($extendedMFAInformation) {
            foreach($prop in $mfaDetails.PSObject.Properties) {
                $mfaStatus | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
            }
        }
        return $mfaStatus | select-object @{L='AccountEnabled';E={$EIDuser.AccountEnabled}},*
    }

    #default : all users report 
    $outFile = "eNMFAReport-"+(get-date -Format 'yyMMdd-HHmmss')+'.csv'
    $MFAReport = @()
    try {
        $EIDUsers = Get-MgUser -Filter "usertype eq 'member'" -All -Property userPrincipalName,Id,AccountEnabled,LicenseAssignmentStates -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose $_.Exception
        return
    }
    $nrOfEIDUsers = $EIDUsers.count
    Write-Verbose "$nrOfEIDUsers member users found. gathering MFA status..."
    $current = 0
    foreach($EIDuser in $EIDUsers) {
        write-progress -activity "getting MFA status" -status "processing $($EIDuser.userPrincipalName)" -percentComplete (($current/$nrOfEIDUsers)*100)
        $current++
        $mfaStatus = $null
        #Get-MgReportAuthenticationMethodUserRegistrationDetail doesn't work for unlicensed accounts
        if(-not [string]::isNullOrEmpty($EIDuser.LicenseAssignmentStates)) {
            $mfaStatus = Get-MgReportAuthenticationMethodUserRegistrationDetail -Filter "userPrincipalName eq '$($EIDuser.userPrincipalName)'" -ErrorAction SilentlyContinue | `
                Select-Object UserDisplayName,UserPrincipalName,Id,@{L='AccountEnabled';E={ $EIDUser.AccountEnabled}},IsAdmin, `
                    IsMfaCapable,IsMfaRegistered,IsPasswordlessCapable,IsSsprCapable,IsSsprEnabled,IsSsprRegistered,LastUpdatedDateTime, `
                    @{L='MethodsRegistered';E={$_.MethodsRegistered -join ','}},IsSystemPreferredAuthenticationMethodEnabled, `
                    @{L='SystemPreferredAuthenticationMethods';E={$_.SystemPreferredAuthenticationMethods -join ','}}, `
                    UserPreferredMethodForSecondaryAuthentication, @{L='AdditionalProperties';E={$_.AdditionalProperties.Count}}
        }
        if([string]::isNullOrEmpty($mfaStatus)) {
            $mfaStatus = [PSCustomObject]@{
                UserDisplayName = $EIDuser.userPrincipalName
                UserPrincipalName = $EIDuser.userPrincipalName
                Id = $EIDuser.Id
                AccountEnabled = $EIDUser.AccountEnabled
                IsAdmin = ''
                IsMfaCapable = ''
                IsMfaRegistered = ''
                IsPasswordlessCapable = ''
                IsSsprCapable = ''
                IsSsprEnabled = ''
                IsSsprRegistered = ''
                LastUpdatedDateTime = ''
                MethodsRegistered = ''
                IsSystemPreferredAuthenticationMethodEnabled = ''
                SystemPreferredAuthenticationMethods = ''
                UserPreferredMethodForSecondaryAuthentication = ''
                AdditionalProperties = 0
            }
            #continue
        }
        if($extendedMFAInformation) {
            $mfaDetails = Get-MFAMethods -userId $EIDuser.Id
            foreach($prop in $mfaDetails.PSObject.Properties) {
                $mfaStatus | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
            }
        }
        $MFAReport += $mfaStatus
    }

    $MFAReport | Sort-Object UserDisplayName | Export-Csv -Path $outFile -NoTypeInformation
    Write-Verbose "results saved as $outFile."
    if($run) {
        $xlsFile = convert-CSV2XLS -CSVfileName $outFile
        Start-Process $xlsFile
    }
    Write-Host 'done.' -ForegroundColor Green
}
function get-ADPrivilegedUsers {
<#
.SYNOPSIS
    get all priviliedged users in AD domain.
.DESCRIPTION
    script is checking all well known SIDs for prviledged groups in AD. if there are permissions assigned via non-standard role
    it will not be included.. there is a permission crawler script to detect non-standard permissions but not yet included in this build.
.EXAMPLE
    .\get-eNAuditorADprivililegedUsers.ps1
    
    creates report file.
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

    #TO/DO 
    - tuning after putting to module such as output file name, parameters, etc. haven't been using this script for a while...
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
function get-EntraIDPrivilegedUsers {
<#
.SYNOPSIS
    auditing script allowing to get the list of all users assgined to any Entra ID Role including PIM roles.
.DESCRIPTION
    script is queyring all EID roles to look for the members and if AAD P1 license is available, checks
    fot the PIM roles and their members. 
    outputs the report in CSV format.
.EXAMPLE
    get full report on all roles that have any memebers

    .\get-eNAuitorEntraIDPrivilegedUsers.ps1
.EXAMPLE
    get full report sorted by a user name, script will not try to connect assuming you're already authenticated with a proper permissions

    .\get-eNAuitorEntraIDPrivilegedUsers.ps1 -skipConnect

.INPUTS
    None.
.OUTPUTS
    csv report file.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250331
        last changes
        - 250331 PIM
        - 241029 initialized

    #TO|DO
#>
    [CmdletBinding()]
    param (
        #assume you're already connected with mgGraph to skip authentication
        [Parameter(mandatory=$false,position=0)]
            [switch]$skipConnect,
        #export CSV file delimiter
        [Parameter(mandatory=$false,position=1)]
            [string][validateSet(',',';','default')]$delimiter='default'
        
    )

    function get-userInfo {
        param(
            [parameter(Mandatory)]
                [string]$objId,
            [parameter(Mandatory)]
                [string]$roleName,
            [parameter(Mandatory)]
                [string]$rID,
            [parameter(Mandatory=$false)]
                [string]$DirectoryScopeID = "/",
            [parameter(Mandatory=$false)]
                [string]$Type = 'Direct Member',
            [parameter(Mandatory=$false)]
                [string]$StartDateTime = 'N/A',
            [parameter(Mandatory=$false)]
                [string]$EndDateTime = 'Permanent'
        )

        if(!$tmpPIMusers.ContainsKey($objId)) {
            try {
                #accountEnabled is not passed via additionalProperties
                $eidObject = Get-MgDirectoryObjectById -Ids $objId -ErrorAction SilentlyContinue
                $idType = $eidObject['@odata.type']
                if($eidObject['@odata.type'] -eq "#microsoft.graph.group") {
                    $tmpPIMusers.Add($objId,@{
                            displayName = $eidObject.AdditionalProperties.displayName
                            userPrincipalName = ""
                            accountEnabled = ""
                            identityType = "Group"
                        })
                } else { #user
                    $eUser = Get-MgUser -UserId $objId -Property AccountEnabled #not available via additionalProperties
                    $tmpPIMusers.Add($objId,@{
                        displayName = $eidObject.AdditionalProperties.displayName
                        userPrincipalName = $eidObject.AdditionalProperties.userPrincipalName
                        accountEnabled = $eUser.AccountEnabled
                        identityType = "User"
                    })
                }
            } catch {
                $tmpPIMusers.Add($objId,@{
                    displayName = $_.exception
                    userPrincipalName = 'err'
                    accountEnabled = 'err'
                    identityType = 'err'
                })
            }
        }
        $scopeName = ""
        $scopeType = ""
        if($DirectoryScopeID -ne "/") {
            $scopeObj = Get-MgDirectoryObjectById -Ids $DirectoryScopeID.Replace('/','') -ErrorAction SilentlyContinue
            if($scopeObj) {
                $scopeName = $scopeObj.AdditionalProperties.displayName
                $scopeType = $scopeObj.AdditionalProperties['@odata.type']
            }
        }

        return [PSCustomObject][ordered]@{ 
            ID = $objId
            identityType = $tmpPIMusers[$objId].identityType
            IdentityName = $tmpPIMusers[$objId].displayName
            userPrincipalName = $tmpPIMusers[$objId].userPrincipalName
            enabled = $tmpPIMusers[$objId].accountEnabled
            RoleName = $roleName
            roleID = $rID
            DirectoryScopeID = $DirectoryScopeID
            scopeType = $scopeType
            scopeName = $scopeName
            Type = $Type
            StartDateTime  = $StartDateTime
            EndDateTime    = $EndDateTime
        }

    }

    <#
    try {
        import-module Microsoft.Graph.Authentication -ErrorAction Stop
    } catch {
        throw "error importing mgGraph module. $($_.Exception)"
        return
    }
    #>
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
    } else {
        if($null -eq (Get-MgContext)) {
            throw "not connected. please connect to mgGraph first."
        }
    }
    $VerbosePreference = 'Continue'

    $tenantDomain = (Get-MgOrganization).VerifiedDomains | ? IsDefault | Select-Object -ExpandProperty name
    $outFile = "EntraIDPrivileged-{0}-{1}.csv" -f $tenantDomain,(get-date -Format 'yyMMdd')
    $PIMSKUs = @("AAD_PREMIUM_P2", "ENTERPRISEPREMIUM")
    $tmpPIMusers = @{} #for query speed optimisation - cache for already queried users
    $RoleMemebersList = @() #array to store results

    Write-Verbose "getting roles and members (may take up to 5 min)..."
    $EntraRoles = Get-MgDirectoryRole

    foreach($role in $EntraRoles) {
        $rID=$role.Id
        $rMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $rID
        foreach($member in $rMembers) {
            $RoleMemebersList += get-userInfo -objId $member.Id -roleName $role.DisplayName -rID $rID 
        }
    } 

    #region PIM roles    

    #checking available SKUs 
    $SKUs = Get-MgSubscribedSku
    $servicePlans = $SKUs.ServicePlans | Select-Object -ExpandProperty ServicePlanName -Unique
    $hasPIM = $servicePlans | Where-Object { $_ -in $PIMSKUs }

    if (!$hasPIM) {
        Write-Verbose "No PIM SKUs found. skipping PIM roles."
    } else {
        Write-Verbose "PIM SKU found. checking PIM roles."

        # Get eligible assignments
        $eligible = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -All -ExpandProperty Principal,RoleDefinition
        #get role names 
        #$roles = Get-MgRoleManagementDirectoryRoleDefinition -All

        if($eligible.Count -lt 1) {
            Write-verbose "No eligible assignments found."
        } else {
            Write-verbose  "Found $($eligible.Count) eligible assignments."
            foreach($member in $eligible) {
                If($member.Principal.AdditionalProperties.'@odata.type' -ne "#microsoft.graph.user"){
                    #STOP
                    write-verbose "THIS IS NOT A USER:"
                    $member.Principal.additionalProperties|out-host
                }
                $RoleMemebersList += get-userInfo -objId $member.Principal.Id -roleName $member.RoleDefinition.DisplayName -rID $member.RoleDefinition.Id -DirectoryScopeID $member.DirectoryScopeId -Type 'Eligible' -StartDateTime $member.StartDateTime -EndDateTime $member.EndDateTime
            }
        }
    }
    #unsupported in PS 5.1
    #$sortedMemebersList = ($sortBy -eq 'Role') ? ($RoleMemebersList | Sort-Object RoleName) : ($RoleMemebersList | Select-Object userName,userID,enabled,RoleName,roleID | Sort-Object userName)
    $sortedMemebersList = $RoleMemebersList | Sort-Object 'roleName'

    $exportParam = @{
        NoTypeInformation = $true
        Encoding = 'UTF8'
        Path = $outFile
    }
    if($delimiter -ne 'default') {
        $exportParam.Add('Delimiter',$delimiter)
    } 
    $sortedMemebersList | export-csv @exportParam

    Write-Host -ForegroundColor Green "exported to .\$outFile.`ndone."
}
function get-ReportADObjects {
<#
.SYNOPSIS
    Prepares a report on AD objects with a focus on activity time - when the object has authenticated.
    Allows to prepare report for User and Computer objects. 
.DESCRIPTION
    Search-ADAccount commandlet is useful for quick ad-hoc queried, but it does not return all required object attributes 
    for proper reporting. This script is gathering much more information and is a part of a wider project allowing to
    create aggregated object reporting to support migrations, clean up or regular audits.

    requires to be run As Administrator as running in less privilileged context is not returing some values - e.g. 'enabled'
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
    version 250131
        last changes
        - 250131 added isAdmin check - that required to also add 'memberOf' field.
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
    #check for admin priviledges. there is this strange bug [or feature (; ] that if you run console without
    #admin, some account do report 'enabled' attribute, some are not. so it's suggested to run as admin.
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if(-not $isAdmin) {
        Write-Warning "It's recommended to run script as administrator for full attribute visibility"
    }

    #can't add requires as it would count for a whole module... I don't want that.
    $ADmodulePresent =  get-module ActiveDirectory -ListAvailable
    if($null -eq $ADmodulePresent) { 
        Write-Error "ActiveDirectory module not present. please install RSAT tools. you sure it's DC?"
        return
    } 
    try {
        $domainSID = (Get-ADDomain).domainSID
    } catch {
        Write-Error "error getting domain SID. are you sure you're connected to the domain?"
        return
    }
    $wellKnownAdminSids = @("S-1-5-32-547","S-1-5-32-553","S-1-5-32-577","S-1-5-32-544","S-1-5-32-582","S-1-5-32-560","S-1-5-32-581","S-1-5-32-551",`
        "S-1-5-32-556","S-1-5-32-561","S-1-5-32-578","S-1-5-32-548","S-1-5-32-575","S-1-5-32-550","S-1-5-32-579","S-1-5-32-557","S-1-5-32-549","S-1-5-32-573","S-1-5-32-569","S-1-5-32-576",`
        "$domainSID-498","$domainSID-512","$domainSID-516","$domainSID-517","$domainSID-518","$domainSID-519","$domainSID-520","$domainSID-521","$domainSID-522","$domainSID-525","$domainSID-526","$domainSID-527")
    #these are dynamic, possible to query but too niche to make an effort. sorry.
    $adminGroupNames = @("DnsAdmins","Organization Management","Recipient Management","View-Only Organization Management","Public Folder Management",`
        "UM Management","Help Desk","Records Management","Discovery Management","Server Management","Delegated Setup","Hygiene Management","Compliance Management",`
        "Security Reader","Security Administrator","Exchange Servers","Exchange Trusted Subsystem","Managed Availability Servers","Exchange Windows Permissions",`
        "ExchangeLegacyInterop","Exchange Install Domain Servers"
    )
    foreach($sid in $wellKnownAdminSids) { $adminGroupNames += (Get-ADObject -Filter "ObjectSID -eq '$sid'").distinguishedname }

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
            -Properties enabled,userPrincipalName,mail,distinguishedname,givenName,surname,samaccountname,displayName,description,lastLogonDate,PasswordLastSet,memberOf
        Write-Verbose "found $(($ADObjects|Measure-Object).count) objects"
        $ADObjects = $ADObjects | select-object samaccountname,userPrincipalName,enabled,givenName,surname,displayName,mail,description,`
            lastLogonDate,@{L='daysInactive';E={if($_.LastLogonDate) {$lld=$_.LastLogonDate} else {$lld="1/1/1970"} ;(New-TimeSpan -End (get-date) -Start $lld).Days}},PasswordLastSet,`
            distinguishedname,@{L='parentOU';E={$rxParentOU.Match($_.distinguishedName).groups[1].value}}, @{L='isAdmin';E={$false}},@{L="memberOf";E={$_.memberOf -join ';'}}
        #add check if user belongs to any privileged group
        foreach($ADuser in $ADObjects) {
            foreach($membership in ($ADuser.memberOf -split ';')) {
                if($adminGroupNames -contains $membership) {
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
function get-ReportEntraUsers {
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
    version 250403
        last changes
        - 250403 error handling improvement
        - 250218 missing isAdmin attribute on non-EIDP1 
        - 250209 servicePlans created/saved in temp folder
        - 250203 isAdmin for EID, some optmization for MFA check, additional parameters and attributes, some optimisations
        - 240718 initiated as a more generalized project, service plans display names check up, segmentation
        - 240627 MFA - for now only general status, AADP1 error handling
        - 240520 initialized

    #TO/DO
    * pagefile for big numbers

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
        [Parameter(mandatory=$false,position=3)]
            [switch]$skipIsAdminCheck,
        #automatically generate XLSX report using eNLib 
        [Parameter(mandatory=$false,position=4)]
            [switch]$xlsxReport
        
    )
    $VerbosePreference = 'Continue'

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
            Connect-MgGraph -Scopes "RoleManagement.Read.Directory","User.Read.All","AuditLog.Read.All","Domain.Read.All","UserAuthenticationMethod.Read.All","email","profile","openid" -ErrorAction Stop
        } catch {
            throw "error connecting. $($_.Exception)"
            return
        }
    }
    Write-Verbose "getting connection info..."
    $ctx = Get-MgContext
    Write-Verbose "connected as '$($ctx.Account)'"
    if($ctx.Scopes -notcontains 'User.Read.All' -or $ctx.Scopes -notcontains 'AuditLog.Read.All' -or $ctx.Scopes -notcontains 'Domain.Read.All' -or $ctx.Scopes -notcontains 'UserAuthenticationMethod.Read.All' `
        -or $ctx.Scopes -notcontains 'RoleManagement.Read.Directory') {
        throw "you need to connect using connect-mgGraph -Scopes User.Read.All,AuditLog.Read.All,Domain.Read.All,UserAuthenticationMethod.Read.All,RoleManagement.Read.Directory"
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
        Write-Verbose "getting the MFA info on accounts..."
        $EntraUsers = $EntraUsers | Select-Object *,@{L='MFAStatus';E={ Get-MFAMethods $_.id -onlyStatus }}
    } else {
        Write-Verbose "skipping the MFA check..."
    }

    Write-Verbose "some output tuning..."
    if($AADP1) {
        $entraUsers = $entraUsers |
            select-object displayname,userType,accountenabled,isAdmin,givenname,surname,userprincipalname,mail,MFAStatus,`
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
            select-object displayname,userType,accountenabled,isAdmin,givenname,surname,userprincipalname,mail,MFAStatus,`
            @{L='Hybrid';E={if($_.OnPremisesSyncEnabled) {$_.OnPremisesSyncEnabled} else {"FALSE"} }},`
            @{L='LastLogonDate';E={'NO AADP1'}},`
            @{L='LastNILogonDate';E={'NO AADP1'}},`
            licenses,id,`
            @{L='daysInactive';E={'NO AADP1'}}
    }

    if(!$skipIsAdminCheck) {
        #get all privilileged user IDs
        $pids = Get-MgRoleManagementDirectoryRoleAssignment | select-object -ExpandProperty principalId -Unique
        foreach($eidU in $entraUsers) {
            if($pids -contains $eidU.id) {
                $eidU.isAdmin = $true
            } else {
                $eidU.isAdmin = $false
            }
        }
    }
    if(!$skipLicenseCheck) {
        Write-Verbose "getting License info..."
        $TempFolder = [System.IO.Path]::GetTempPath()
        $spFile = "$TempFolder\servicePlans.csv"
        $plansFile = $true
    
        if(!(test-path $spFile)) {
            Write-Verbose "file containing plans list not found - downloading..."
            [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
            try {
                Invoke-WebRequest $url -OutFile $spFile
            } catch {
                Write-Error "unable to download plans definition file. display names will not be accessible"
                $_.Exception
                $plansFile = $false
            }
        } 
        if($plansFile) {
            $spInfo = import-csv $spFile -Delimiter ','

            $entraUsers | %{ 
                $userLicenses = @()
                foreach($sku in (Get-MgUserLicenseDetail -UserId $_.id).SkuPartNumber ) {
                    $userLicenses += convert-SKUCodeToDisplayName -SKUName $sku
                }
                $_.licenses = $userLicenses -join ";"
            }
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
function get-ReportMailboxes {
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
    $VerbosePreference = 'Continue'

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
function join-ReportHybridUsersInfo {
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
    version 250209
        last changes
        - 250209 properties in reports will now be dynamically added - if report contains more attributes, they will be added to the final output
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
            return
        }
    }
    if($inputCSVAD) {
        $ADData = load-CSV $inputCSVAD `
            -header $headerAD `
            -headerIsCritical `
            -prefix 'AD_'
        $reports++
        if([string]::isNullOrEmpty($ADData)) {
            return
        }
    }
    if($inputCSVEXO) {
        $EXOData = load-CSV $inputCSVEXO `
            -header $headerEXO `
            -headerIsCritical `
            -prefix 'EXO_'
        $reports++
        if([string]::isNullOrEmpty($EXOData)) {
            return
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
                    if($entraFound|? elementproperty -eq 'userPrincipalName') { #difficult to choose, but UPN matching is the strongest. then mail. displyname is rather a 'soft match'and may have many duplicates
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
    $finalHeader = New-Object System.Collections.ArrayList
    if($EntraIDData) { 
        foreach($el in @("DisplayName","UserType","AccountEnabled","GivenName","Surname","UserPrincipalName","Mail","MFAStatus","Hybrid","LastLogonDate","LastNILogonDate","licenses","Id","daysInactive") ) {
            $finalHeader.Add($el) | Out-Null
        }
        foreach($prop in $EntraIDData[0].psobject.Properties) {
            if($prop.Name -notin $finalHeader) {
                $finalHeader.Add($prop.Name) | Out-Null
            }
        }
    }
    if($ADData) { 
        foreach($el in @("AD_samaccountname","AD_userPrincipalName","AD_enabled","AD_givenName","AD_surname","AD_displayName","AD_mail","AD_description","AD_lastLogonDate","AD_daysInactive","AD_PasswordLastSet","AD_distinguishedname","AD_parentOU") ) {
            $finalHeader.Add($el) | Out-Null
        }
        foreach($prop in $ADData[0].psobject.Properties) {
            if($prop.Name -notin $finalHeader) {
                $finalHeader.Add($prop.Name) | Out-Null
            }
        }
    }
    if($EXOData) { 
        foreach($el in @("EXO_PrimarySMTPAddress","EXO_DisplayName","EXO_FirstName","EXO_LastName","EXO_RecipientType","EXO_RecipientTypeDetails","EXO_emails","EXO_delegations","EXO_ForwardingAddress", "EXO_ForwardingSmtpAddress","EXO_WhenMailboxCreated","EXO_userPrincipalName","EXO_enabled","EXO_Identity","EXO_LastInteractionTime","EXO_LastUserActionTime","EXO_TotalItemSize","EXO_ExchangeObjectId") ) {
            $finalHeader.Add($el) | Out-Null
        }
        foreach($prop in $EXOData[0].psobject.Properties) {
            if($prop.Name -notin $finalHeader) {
                $finalHeader.Add($prop.Name) | Out-Null
            }
        }
    }

    $finalResults = $metaverseUserInfo.Keys | %{ 
        $metaverseUserInfo[$_] |
            Select-Object $finalHeader |
            Select-Object *,@{L='Hybrid_daysInactive';E={($_.daysInactive,$_.AD_daysInactive|Measure-Object -Minimum).minimum}} 
    } #,Hybrid_daysInactive,displayName,AD_displayName,EXO_DisplayName
    $finalResults | 
        Sort-Object { 
            $flag = if([string]::isNullOrEmpty($_.isAdmin) ) { 'Z' } else { 'A' }
            $flag2 = if([string]::isNullOrEmpty($_.AD_isAdmin) ) { 'Z' } else { 'A' }
            $p1 = $_.isAdmin
            $p2 = $_.AD_isAdmin
            $p3 = $_.Hybrid_daysInactive
            @($flag,$flag2,$p1,$p2,$p3)
        } | 
        Export-Csv -Encoding unicode -NoTypeInformation $exportCSVFile

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