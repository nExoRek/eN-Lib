[CmdletBinding()]
Param(
    #skip connecting  alsready - already authenticated
    [Parameter(mandatory=$false,position=0)]
    [switch]$skipConnect
)

function Parse-AppPermissions {
    Param(
    #App role assignment object
    [Parameter(Mandatory=$true)]$appRoleAssignments)

    $appCount = 0
    $calendarAppCount = 0 
    $contactsAppCount = 0 
    $mailsAppCount = 0 
    $riskyAppCount = 0 
    $directoryAppCount = 0 
    $filesAppCount = 0 
    $sitesAppCount = 0 

    foreach ($appRoleAssignment in $appRoleAssignments) {
        $resID = $appRoleAssignment.ResourceDisplayName
        $roleID = (Get-ServicePrincipalRoleById $appRoleAssignment.resourceId).appRoles | Where-Object {$_.id -eq $appRoleAssignment.appRoleId} | Select-Object -ExpandProperty Value
        if (!$roleID) { $roleID = "Orphaned ($($appRoleAssignment.appRoleId))" }
        $OAuthAppPerm["[" + $resID + "]"] += $("," + $roleID)
        
        if ($roleID) {
            $scopes = $roleID.Split(" ")
            $calendarScopes = $scopes | Where-Object { $_ -like "*Calendars*" }
            if ($calendarScopes) {
                $calendarAppCount += $calendarScopes.Count
            }
            $contactsScopes = $scopes | Where-Object { $_ -like "*Contacts*" }
            if ($contactsScopes) {
                $contactsAppCount += $contactsScopes.Count
            }
            $mailsScopes = $scopes | Where-Object { $_ -like "*Mail.*" }
            if ($mailsScopes) {
                $mailsAppCount += $mailsScopes.Count
            }
            $riskyScopes = $scopes | Where-Object { $_ -like "AppRoleAssignment.ReadWrite.All" }
            if ($riskyScopes) {
                $riskyAppCount += $riskyScopes.Count
            }
            $directoryScopes = $scopes | Where-Object { $_ -like "Directory.ReadWrite*" }
            if ($directoryScopes) {
                $directoryAppCount += $directoryScopes.Count
            }
            $filesScopes = $scopes | Where-Object { $_ -like "Files*" }
            if ($filesScopes) {
                $filesAppCount += $filesScopes.Count
            }
            $sitesScopes = $scopes | Where-Object { $_ -like "Sites*" }
            if ($sitesScopes) {
                $sitesAppCount += $sitesScopes.Count
            }    
            $appCount++ # Count every delegation
        }
    }
    # Return counts after processing all items
    return $appCount, $calendarAppCount, $contactsAppCount, $mailsAppCount, $riskyAppCount, $directoryAppCount, $filesAppCount, $sitesAppCount
}

function Parse-DelegatePermissions {

    Param(
    #oauth2PermissionGrants object
    [Parameter(Mandatory=$true)]$oauth2PermissionGrants)

    $delegationCount = 0
    $calendarDelegationCount = 0 # Initialize the calendar delegation count
    $contactsDelegationCount = 0 # Initialize the contacts delegation count
    $mailsDelegationCount = 0 # Initialize the mail delegation count
    $riskyDelegationCount = 0 # Initialize the risky delegation count
    $directoryDelegationCount = 0 # Initialize the directory delegation count
    $filesDelegationCount = 0 # Initialize the files delegation count
    $sitesDelegationCount = 0 # Initialize the sites delegation count

    foreach ($oauth2PermissionGrant in $oauth2PermissionGrants) {
        $resID = (Get-ServicePrincipalRoleById $oauth2PermissionGrant.ResourceId).appDisplayName
        if ($null -ne $oauth2PermissionGrant.PrincipalId) {
            $userId = "(" + (Get-UserUPNById -objectID $oauth2PermissionGrant.principalId) + ")"
        }
        else { $userId = $null }

        if ($oauth2PermissionGrant.Scope) {
            $scopes = $oauth2PermissionGrant.Scope.Split(" ")
            $calendarScopes = $scopes | Where-Object { $_ -like "*Calendars*" }
            if ($calendarScopes) {
                $calendarDelegationCount += $calendarScopes.Count
            }
            $contactsScopes = $scopes | Where-Object { $_ -like "*Contacts*" }
            if ($contactsScopes) {
                $contactsDelegationCount += $contactsScopes.Count
            }
            $mailsScopes = $scopes | Where-Object { $_ -like "*Mail.*" }
            if ($mailsScopes) {
                $mailsDelegationCount += $mailsScopes.Count
            }
            $riskyScopes = $scopes | Where-Object { $_ -like "AppRoleAssignment.ReadWrite.All" }
            if ($riskyScopes) {
                $riskyDelegationCount += $riskyScopes.Count
            }
            $directoryScopes = $scopes | Where-Object { $_ -like "Directory.ReadWrite*" }
            if ($directoryScopes) {
                $directoryDelegationCount += $directoryScopes.Count
            }
            $filesScopes = $scopes | Where-Object { $_ -like "Files*" }
            if ($filesScopes) {
                $filesDelegationCount += $filesScopes.Count
            }
            $sitesScopes = $scopes | Where-Object { $_ -like "Sitesd*" }
            if ($sitesScopes) {
                $sitesDelegationCount += $sitesScopes.Count
            }    
            $OAuthDelegatedPerm["[" + $resID + $userId + "]"] += ($scopes -join ",")
            $delegationCount++ # Count every delegation
        }
        else { $OAuthDelegatedPerm["[" + $resID + $userId + "]"] += "Orphaned scope" }
    }

    return $delegationCount, $calendarDelegationCount, $contactsDelegationCount, $mailsDelegationCount, $riskyDelegationCount, $directoryDelegationCount, $filesDelegationCount, $sitesDelegationCount
}

function Get-ServicePrincipalRoleById {

    Param(
    #Service principal object
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$spID)

    if (!$SPPerm[$spID]) {
        $SPPerm[$spID] = Get-MgServicePrincipal -ServicePrincipalId $spID -Verbose:$false -ErrorAction Stop
    }
    return $SPPerm[$spID]
}

function Get-UserUPNById {

    Param(
    #User objectID
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$objectID)
    if (!$SPusers[$objectID]) {
        $SPusers[$objectID] = (Get-MgUser -UserId $objectID -Property UserPrincipalName).UserPrincipalName
    }
    return $SPusers[$objectID]
}

Write-Verbose "Connecting to Graph API..."
#Import-Module Microsoft.Graph.Beta.Applications -Verbose:$false -ErrorAction Stop
if(-not $skipConnect) {
    try {
        Connect-MgGraph -Scopes "Directory.Read.All","Application.Read.All" -ErrorAction Stop -NoWelcome
    } catch { 
        throw $_ 
    }
}
#Make sure we include Custom security attributes in the report, if requested
$properties = "appDisplayName,appId,appOwnerOrganizationId,displayName,id,createdDateTime,AccountEnabled,passwordCredentials,keyCredentials,tokenEncryptionKeyId,verifiedPublisher,Homepage,PublisherName,tags"

#Get the list of Service principal objects within the tenant.
#Only /beta returns publisherName currently
$SPs = @()

Write-Verbose "Retrieving list of service principals..."
#AccessReview.ReadWrite.All
#Agreement.ReadWrite.All
#Application.Read.All
#Application.ReadWrite.All
#AppRoleAssignment.ReadWrite.All
#AuditLog.Read.All
#CustomSecAttributeDefinition.ReadWrite.All
#DeviceManagementApps.ReadWrite.All
#DeviceManagementConfiguration.ReadWrite.All
#DeviceManagementManagedDevices.ReadWrite.All
#DeviceManagementRBAC.ReadWrite.All
#DeviceManagementServiceConfig.ReadWrite.All
#Directory.AccessAsUser.All+Directory.Read.All
#Directory.ReadWrite.All+Domain.ReadWrite.All
#email
#EntitlementManagement.ReadWrite.All
#Files.ReadWrite.All
#Group.ReadWrite.All
#IdentityRiskyUser.Read.All
#OnPremDirectorySynchronization.ReadWrite.All
#openid
#Organization.Read.All
#Organization.ReadWrite.All
#Policy.Read.All
#Policy.ReadWrite.AuthenticationFlows
#Policy.ReadWrite.AuthenticationMethod
#Policy.ReadWrite.ConditionalAccess
#Policy.ReadWrite.CrossTenantAccess
#Policy.ReadWrite.DeviceConfiguration
#Policy.ReadWrite.FeatureRollout
#Policy.ReadWrite.PermissionGrant
#Policy.ReadWrite.TrustFramework
#profile+Reports.Read.All
#RoleAssignmentSchedule.ReadWrite.Directory
#RoleEligibilitySchedule.Read.Directory
#RoleManagement.Read.Directory
#RoleManagement.ReadWrite.Directory
#SecurityEvents.Read.All
#SharePointTenantSettings.ReadWrite.All
#Sites.ReadWrite.All
#User.Read+User.Read.All
#User.ReadWrite.All
#offline_access
#if (!$IncludeBuiltin) { $SPs = Get-MgServicePrincipal -All -Filter "tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')" -Property $properties -ErrorAction Stop -Verbose:$false }
#else { 
    $SPs = Get-MgServicePrincipal -All -Property $properties -ErrorAction Stop -Verbose:$false 
#}

#Set up some variables
$SPperm = @{} #hash-table to store data for app roles and stuff
$SPusers = @{} #hash-table to store data for users assigned delegate permissions and stuff
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;
$appsWithDelegatedAccess = 0;
$appsWithDelegatedCalendarAccess = 0;
$appsWithDelegatedContactsAccess = 0;
$appsWithDelegatedRiskyAccess = 0;
$appsWithDelegatedMailAccess = 0;
$appsWithDelegatedDirectoryReadWriteAccess = 0;
$appsWithAccess = 0;
$appsWithCalendarAccess = 0;
$appsWithContactsAccess = 0;
$appsWithMailAccess = 0;
$appsWithRiskyAccess = 0;
$appsAddedLast30Days = 0;
$appsWithReadWriteConsentCount = 0;
$appsWithDelegatedReadWriteConsentCount = 0;
$appsWithSitesAccess = 0;
$appsWithFilesAccess = 0;
$appsWithDelegatedFilesAccess = 0;
$appsWithDelegatedSitesAccess = 0;

# Calculate the current date minus 30 days
$thirtyDaysAgo = (Get-Date).AddDays(-30)

#Process the list of service principals
foreach ($SP in $SPs) {
    #Progress message
    $ActivityMessage = "Retrieving data for service principal $($SP.DisplayName). Please wait..."
    $StatusMessage = ("Processing service principal {0} of {1}: {2}" -f $count, @($SPs).count, $SP.id)
    $PercentComplete = ($count / @($SPs).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    Write-Verbose "Processing service principal $($SP.id)..."

    #Get owners info. We do not use $expand, as it returns the full set of object properties
    Write-Verbose "Retrieving owners info..."
    $owners = @()
    $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $SP.id -Property id,userPrincipalName -All -ErrorAction Stop -Verbose:$false
    if ($owners) { $owners = $owners.userPrincipalName }

    #Include information about group/directory role memberships. Cannot use /memberOf/microsoft.graph.directoryRole :(
    Write-Verbose "Retrieving group/directory role memberships..."
    $res = Get-MgServicePrincipalMemberOf -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false
    $memberOfGroups = ($res.AdditionalProperties | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.group"}).displayName -join "|" #d is Case-sensitive!
    $memberOfRoles = ($res.AdditionalProperties | Where-Object {$_.'@odata.type' -eq "#microsoft.graph.directoryRole"}).displayName -join "|" #d is Case-sensitive!

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Application Name" = $SP.displayName
        "Application Additional Name" = (&{if ($SP.appDisplayName) { $SP.appDisplayName } else { $null }}) #Apparently appDisplayName can be null
        "Publisher" = (&{if ($SP.PublisherName) { $SP.PublisherName } else { $null }})
        "Verified" = (&{if ($SP.verifiedPublisher.verifiedPublisherId) { $SP.verifiedPublisher.displayName } else { "Not verified" }})
        "Homepage" = (&{if ($SP.Homepage) { $SP.Homepage } else { $null }})
        "Created on" = (&{if ($SP.AdditionalProperties.createdDateTime) {(Get-Date($SP.AdditionalProperties.createdDateTime) -format d)} else { "N/A" }})
        "Enabled" = $SP.AccountEnabled
        "Total delegations" =  $null
        "Calendar delegations" =  $null
        "Contacts delegations" =  $null
        "Mails delegations" =  $null
        "Risky delegations" = $null
        "Sites delegations" = $null
        "Files delegations" = $null
        "Directory delegations" = $null
        "Permissions (delegate)" = $null
        "Authorized By (delegate)" = $null
        "Total apps" = $null
        "Calendar apps" = $null
        "Contacts apps" = $null
        "Mails apps" = $null
        "Directory apps" = $null
        "Risky apps" = $null
        "Files apps" = $null
        "Sites apps" = $null
        "Last modified (application)" = $null
        "Permissions (application)" = $null
        "Owners" = (&{if ($owners) { $owners -join "," } else { $null }})
        "Member of (groups)" = $memberOfGroups
        "Member of (roles)" = $memberOfRoles
        "ObjectId" = $SP.id
        "IsBuiltIn" = $SP.tags -notcontains "WindowsAzureActiveDirectoryIntegratedApp"
    }

    #Check for appRoleAssignments (application permissions)
    Write-Verbose "Retrieving application permissions..."
    try {
        $appRoleAssignments = @()
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false

        $OAuthAppPerm = @{};
        $assignedto = @();$resID = $null; $userId = $null;

        #process application permissions entries
        if (!$appRoleAssignments) {
            Write-Verbose "No application permissions to report on for SP $($SP.id), skipping..."
            $objPermissions.'Total apps' = 0
            $objPermissions.'Calendar apps' = 0
            $objPermissions.'Contacts apps' = 0
            $objPermissions.'Mails apps' = 0
            $objPermissions.'Risky apps' = 0
            $objPermissions.'Directory apps' = 0
            $objPermissions.'Files apps' = 0
            $objPermissions.'Sites apps' = 0
        }
        else {
            $objPermissions.'Last modified (application)' = (Get-Date($appRoleAssignments.CreationTimestamp | Select-Object -Unique | Sort-Object -Descending | Select-Object -First 1) -format d)
            $appsPermissionsCounts = Parse-AppPermissions $appRoleAssignments
            $objPermissions.'Permissions (application)' = (($OAuthAppPerm.GetEnumerator()  | ForEach-Object { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join "|")
            $objPermissions.'Total apps' = $appsPermissionsCounts[0]
            $objPermissions.'Calendar apps' = $appsPermissionsCounts[1]
            $objPermissions.'Contacts apps' = $appsPermissionsCounts[2]
            $objPermissions.'Mails apps' = $appsPermissionsCounts[3]
            $objPermissions.'Risky apps' = $appsPermissionsCounts[4]
            $objPermissions.'Directory apps' = $appsPermissionsCounts[5]
            $objPermissions.'Files apps' = $appsPermissionsCounts[6] 
            $objPermissions.'Sites apps' = $appsPermissionsCounts[7] 
            if ($appsPermissionsCounts[0] -gt 0) {$appsWithAccess++;}
            if ($appsPermissionsCounts[1] -gt 0) {$appsWithCalendarAccess++;}
            if ($appsPermissionsCounts[2] -gt 0) {$appsWithContactsAccess++;}
            if ($appsPermissionsCounts[3] -gt 0) {$appsWithMailAccess++;}
            if ($appsPermissionsCounts[4] -gt 0) {$appsWithRiskyAccess++;}
            if ($appsPermissionsCounts[5] -gt 0) {$appsWithDirectoryReadWriteAccess++;}
            if ($appsPermissionsCounts[6] -gt 0) {$appsWithFilesAccess++;}
            if ($appsPermissionsCounts[7] -gt 0) {$appsWithSitesAccess++;}
        }
    }
    catch { Write-Verbose "Failed to retrieve application permissions for SP $($SP.id) ..." }

    #Check for oauth2PermissionGrants (delegate permissions)
    #Use / here, as /v1.0 does not return expiryTime
    Write-Verbose "Retrieving delegate permissions..."
    try {
        $oauth2PermissionGrants = @()
        $oauth2PermissionGrants = Get-MgServicePrincipalOAuth2PermissionGrant -All -ServicePrincipalId $SP.id -ErrorAction Stop -Verbose:$false

        $OAuthDelegatedPerm = @{};
        $assignedto = @();$resID = $null; $userId = $null;

        #process delegate permissions entries
        if (!$oauth2PermissionGrants) {
            Write-Verbose "No delegate permissions to report on for SP $($SP.id), skipping..."
            $objPermissions.'Total delegations' = 0
            $objPermissions.'Calendar delegations' = 0
            $objPermissions.'Contacts delegations' = 0
            $objPermissions.'Mails delegations' = 0
            $objPermissions.'Risky delegations' = 0
            $objPermissions.'Directory delegations' = 0
            $objPermissions.'Files delegations' = 0
            $objPermissions.'Sites delegations' = 0
        }
        else {
            $delegationCounts = Parse-DelegatePermissions $oauth2PermissionGrants
            $objPermissions.'Permissions (delegate)' = (($OAuthDelegatedPerm.GetEnumerator() | ForEach-Object { "$($_.Name):$($_.Value.ToString().TrimStart(','))"}) -join "|")
            $objPermissions.'Total delegations' = $delegationCounts[0]
            $objPermissions.'Calendar delegations' = $delegationCounts[1]
            $objPermissions.'Contacts delegations' = $delegationCounts[2]
            $objPermissions.'Mails delegations' = $delegationCounts[3]
            $objPermissions.'Risky delegations' = $delegationCounts[4]
            $objPermissions.'Directory delegations' = $delegationCounts[5]
            $objPermissions.'Files delegations' = $delegationCounts[6]
            $objPermissions.'Sites delegations' = $delegationCounts[7]
            if ($delegationCounts[0] -gt 0) {$appsWithDelegatedAccess++;}
            if ($delegationCounts[1] -gt 0) {$appsWithDelegatedCalendarAccess++;}
            if ($delegationCounts[2] -gt 0) {$appsWithDelegatedContactsAccess++;}
            if ($delegationCounts[3] -gt 0) {$appsWithDelegatedMailAccess++;}
            if ($delegationCounts[4] -gt 0) {$appsWithDelegatedRiskyAccess++;}
            if ($delegationCounts[5] -gt 0) {$appsWithDelegatedDirectoryReadWriteAccess++;}
            if ($delegationCounts[6] -gt 0) {$appsWithDelegatedFilesAccess++;}
            if ($delegationCounts[7] -gt 0) {$appsWithDelegatedSitesAccess++;}


            if (($oauth2PermissionGrants.ConsentType | Select-Object -Unique) -eq "AllPrincipals") { $assignedto += "All users (admin consent)" }
            $assignedto +=  @($OAuthDelegatedPerm.Keys) | ForEach-Object {if ($_ -match "\((.*@.*)\)") {$Matches[1]}}
            $objPermissions.'Authorized By (delegate)' = (($assignedto | Select-Object -Unique) -join ",")
        }
    }
    catch { Write-Verbose "Failed to retrieve delegate permissions for SP $($SP.id) ..." }
    # Check if the app was added in the last 30 days and increment the counter if so
    if ($SP.AdditionalProperties.createdDateTime -and ((Get-Date($SP.AdditionalProperties.createdDateTime)) -gt $thirtyDaysAgo)) {
        $appsAddedLast30Days++
    }
    # Check for ReadWrite consents in both application and delegate permissions and increment the counter if found
    if ($objPermissions.'Permissions (application)' -match ".*ReadWrite.*") {
        $appsWithReadWriteConsentCount++
    }
    # Check for ReadWrite consents in both application and delegate permissions and increment the counter if found
    if ($objPermissions.'Permissions (delegate)' -match ".*ReadWrite.*") {
        $appsWithDelegatedReadWriteConsentCount++
    }
    $output.Add($objPermissions)
}

$output = $output | Sort-Object {$_."Application Name"}

$outputFiltered = $output | Where-Object {$_.'Permissions (application)' -match "AppRoleAssignment.ReadWrite.All" -or $_.'Permissions (delegate)' `
-match "AppRoleAssignment.ReadWrite.All"} | Sort-Object {$_."Application Name"}

#Export the result to CSV file
$outputFilePath = "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppInventory"
$output | Select-Object * -ExcludeProperty Number | Export-CSV -nti -Path "$outputFilePath.csv"
Write-Host "Output exported to $($PWD)\$($outputFilePath).csv"
