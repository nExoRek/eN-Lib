<# PRIVATE FUNCTIONS #>

function get-ServicePlanInfoFile {
<#
.SYNOPSIS
    license assignement information comes from a CSV file published on Microsoft site. this support function checks if the file is present
    in the local temp folder and if not, downloads it.
.DESCRIPTION
    force parameter ensures that the file is downloaded even if it exists. this is useful when the file on the server has changed.
    the file is downloaded to the local temp folder and imported as a CSV file.
.EXAMPLE
    get-ServicePlanInfoFile -force

    forces download of the file even if it exists. this is useful when the file on the server has changed.
.INPUTS
    None.
.OUTPUTS
    "$TempFolder\servicePlans.csv"
.LINK
    https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250412
        last changes
        - 250412 initialized

    #TO|DO
#>
    [CmdletBinding()]
    param (
        #force download of the file even if it exists
        [switch]$force
    )

    $TempFolder = [System.IO.Path]::GetTempPath()
    $spFile = "$TempFolder\servicePlans.csv"
    [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"

    if(!(test-path $spFile) -or $force) {
        Write-Verbose "file containing plans list not found - downloading..."
        try {
            Invoke-WebRequest $url -OutFile $spFile
        } catch {
            throw "cannot download definitions the file."
        }
    } 
    $spInfo = import-csv $spFile -Delimiter ','
    return $spInfo
}
function connect-licenseGraph {
    [cmdletbinding()]
    param(
    )

    Write-Debug ("[Scope level: {0}]" -f $MyInvocation.ScopeDepth)

    $ctx = get-mgContext
    if(-not $ctx) {
        Write-Verbose "connection not found - connecting..."
        try {
            Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All"
        } catch {
            throw "cannot connect to Microsoft Graph."
        }
        $ctx = get-mgContext
        if(-not $ctx) {
            throw "you need to be connected to continue."
        }
    } else {
        if($ctx.Scopes -notcontains 'Group.ReadWrite.All' -or $ctx.Scopes -notcontains 'Directory.ReadWrite.All') {
            Write-Verbose "connection found, but not valid for this module. trying to extend scope..."
            try {
                Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All" -force
            } catch {
                throw "cannot connect to Microsoft Graph."
            }
            #retest to manage 'cancel'
            $ctx = get-mgContext
            if(-not $ctx) {
                throw "you need to be connected to continue."
            }
            Write-Verbose "connection found and valid."
        }
    }
    Write-host "connected as $($ctx.Account)"
    return $ctx
}
function convert-toShortString {
    param( 
        $licenseInfo
    ) 

    $htLicenses = @{}

    foreach($lic in $licenseInfo) {
        if($htLicenses.ContainsKey($lic.SkuId)) {
            $htLicenses[$lic.SkuId] = "[{0}]{1}" -f $lic.groupDisplayName,$htLicenses[$lic.SkuId]
        } else {
            $htLicenses[$lic.SkuId] = "[{0}]{1}" -f $lic.groupDisplayName,$lic.skuDisplayName
        }
    }
    return $htLicenses.Values -join ','
}
function convert-toShortStringExt {
    param( 
        $licenseInfo,
        [string]$skuId
    ) 

    $licInfo = ""

    foreach($lic in $licenseInfo|? SkuId -eq $skuId) {
        $licInfo += "[{0}]" -f $lic.groupDisplayName
    }
    if($licenseInfo.disabledPlanDisplayName) {
        $licInfo += ("({0})" -f ($licenseInfo.disabledPlanDisplayName -join '|') )
    }

    return $licInfo

}
<# PUBLIC FUNCTIOND #>

function find-SKU {
    [CmdletBinding()]
    param(
        #lookup the name (internal, displayname or GUID) of the SKU and shows all other values
        [Parameter(mandatory,position=0,ValueFromPipeline)]
            [string]$id,
        #limit lookup to EXACT name and shows other values 
        [Parameter(position=1)]
            [switch]$exact,
        #force download new SKU file
        [Parameter(position=2)]
            [switch]$force    
    )

    Begin {
        $VerbosePreference = 'Continue'
        if($force) {
            $spInfo = get-ServicePlanInfoFile -force
        } else {
            $spInfo = get-ServicePlanInfoFile
        }
    }

    Process {
        if($exact) {
            $SKU = $spInfo | Where-Object { $_.Product_Display_Name -eq $id -or $_.String_Id -eq $id -or $_.GUID -eq $id }
        } else {
            $SKU = $spInfo | Where-Object { $_.Product_Display_Name -match $id -or $_.String_Id -match $id -or $_.GUID -match $id }
        }
        if($SKU) {
            return $SKU | Select-Object @{L='SKUFriendlyName';E={$_.Product_Display_Name}},@{L='SKUCodeName';E={$_.String_Id}},@{L='SKUGUID';E={$_.GUID}} -Unique
        } else {
            return $null
        }
    }
}

function find-ServicePlan {
<#
.SYNOPSIS
    function has 2 capabilities:
        - lookus up Service Plan details based on a partial name - GUID, Friendly Name or Code Name. 
        - shows all SKUs containing looked up Service Plans - very usefull if you want to check your options,
          which licenses are containing capability you are looking for
.DESCRIPTION
    by default only checkes the CSV file for a service plan. with 'showSKUs' parameter additionally shows all
    SKUs containing found plans.
.EXAMPLE
    find-ServicePlan teams

    will show all Service Plans containing 'teams' in the name
.EXAMPLE
    find-ServicePlan teams1 -showSKUs

    TEAMS1 is one of the code names. command will show all license SKUs containing this plan
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250415
        last changes
        - 250415 initialized

    #TO|DO
#>

    [CmdletBinding(DefaultParameterSetName = 'byName')]
    param(
        #finds all Service Plan details. plan name may be partial and be a code name, friendly name or GUID
        [Parameter(ParameterSetName = 'byName',mandatory,position=0)]
        [Parameter(ParameterSetName = 'SKUs',mandatory,position=0)]
            [string]$id,
        #limit lookup to EXACT the name (internal, displayname) and shows details. 
        [Parameter(ParameterSetName = 'byName',position=1)]
        [Parameter(ParameterSetName = 'SKUs',position=1)]
            [switch]$exact,
        #provides a list of all SKUs containing the Service Plan name. 
        [Parameter(ParameterSetName = 'SKUs',position=2)]
            [switch]$showSKUs
    )

    Begin {
        $VerbosePreference = 'Continue'
        $spInfo = get-ServicePlanInfoFile
    }

    Process {
        if($exact) {
            Write-debug 'exact search'
            $foundSKUs = $spInfo | Where-Object { $_.Service_Plans_Included_Friendly_Names -eq $id -or $_.Service_Plan_Name -eq $id -or $_.Service_Plan_Id -eq $id } 
        } else {
            $foundSKUs = $spInfo | Where-Object { $_.Service_Plans_Included_Friendly_Names -match $id -or $_.Service_Plan_Name -match $id -or $_.Service_Plan_Id -match $id } 
        }
        if($showSKUs) {
            return $foundSKUs | Select-Object @{L='ServicePlanFriendlyName';E={$_.Service_Plans_Included_Friendly_Names}},
                @{L='ServicePlanCodeName';E={$_.Service_Plan_Name}},
                @{L='ServicePlanGUID';E={$_.Service_Plan_Id}},
                @{L='SKUFriendlyName';E={$_.Product_Display_Name}},
                @{L='SKUCodeName';E={$_.String_Id}},
                @{L='SKUGUID';E={$_.GUID}} -Unique
        } else {
            return $foundSKUs | Select-Object @{L='ServicePlanFriendlyName';E={$_.Service_Plans_Included_Friendly_Names}},
                @{L='ServicePlanCodeName';E={$_.Service_Plan_Name}},
                @{L='ServicePlanGUID';E={$_.Service_Plan_Id}} -Unique
        }
    }
}

function show-ServicePlans {
<#
.SYNOPSIS
    display information on Service Plans included in a provided license SKU. 
.DESCRIPTION
    function allows to provide a name [string] or an SKU object from find-SKU function for pipelining.

.EXAMPLE
    show-eNGBLServicePlans -nameSKU EOP_ENTERPRISE_PREMIUM 
    
    shows friendly name of EOP_ENTERPRISE_PREMIUM and GUID. 
.EXAMPLE
    find-eNGBLSKU '365 E3'|show-eNGBLServicePlans  
    
    looks up for all licenses containing '365 e3' in their name and pass it to a function that displays all Service Plans 
.LINK
    https://w-files.pl
.LINK
    https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
.NOTES
    nExoR ::))o-
    version 250415
        last changes
        - 250415 display fixes
        - 250310 v2

    #TO|DO
#>
    [CmdletBinding(DefaultParameterSetName = 'bySKU')]
    param(
        #name of the SKU to show Service Plans for
        [Parameter(ParameterSetName = 'byName',mandatory,position=0)]
            [string]$nameSKU,
        #SKU object from find-SKU to show Service Plans for
        [Parameter(ParameterSetName = 'bySKU',mandatory,position=0,ValueFromPipeline)]
            [PSObject]$objSKU
    )

    Begin {
        $VerbosePreference = 'Continue'
        $spInfo = get-ServicePlanInfoFile
        if($PSCmdlet.ParameterSetName -eq 'byName') {
            $objSKU = find-SKU -id $nameSKU -exact | select-object -first 1 #some licenses have multiple names - strange, but true
        }
    }

    Process {
        write-verbose "SKU Friendly name: $($objSKU.SKUFriendlyName); ID: $($objSKU.SKUCodeName); GUID: $($objSKU.SKUGUID) contains following Service Plans:" | out-host
        $spInfo | 
            Where-Object {$_.Product_Display_Name -eq $objSKU.SKUFriendlyName} | 
            Select-Object @{L='ServicePlanFriendlyName';E={$_.Service_Plans_Included_Friendly_Names}},
                @{L='ServicePlanCodeName';E={$_.Service_Plan_Name}},
                @{L='ServicePlanGUID';E={$_.Service_Plan_Id}} | out-host
    }
    end {}
}

function compare-SKUs {
<#
.SYNOPSIS
    compares plans for two SKUs showing which Plans are the same and which are different.
    SKU may be provided as GUID, code name or displayname.
.DESCRIPTION
    comparison is using information from Microsoft CSV. 
    
    for some reason Microsoft is not updating it often enough and there are situations you will not find the plan you currently use. 
    if you need to make Service Plan comparisons from such a SKU, the other option is to copmpare Plan from CSV with object taken
    from get-mgSubscribedSku
.EXAMPLE
    compare-eNGBLSKUs -SKU1 'Office 365 E3' -SKU2 'Microsoft 365 E3'

    shows differences between the two licenses
.EXAMPLE
    $skus = get-mgSubscribedSku -all
    compare-eNGBLSKUs -SKU1 52ea0e27-ae73-4983-a08f-13561ebdb823 -objSKU ($skus|? SkuPartNumber -eq 'Microsoft_Teams_Enterprise_New')

    SKU1 is a 'Teams Premium (for Departments)'
    Microsoft_Teams_Enterprise_New does not exisit in Microsoft file (at the date of writing that) and I still want to understand the difference.
    so I get the object from tenant - it contains SKU Ids of all Service Plans included, so may compare the two.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250415
        last changes
        - 250415 fixes
        - 250414 initialized

    #TO|DO
    - error in handing object.. 
#>

    [CmdletBinding(DefaultParameterSetName='string')]
    param(
        [Parameter(ParameterSetName='string',Mandatory, Position=0)]
        [Parameter(ParameterSetName='object',Mandatory, Position=0)]
            [string]$SKU1,

        [Parameter(ParameterSetName='string',Mandatory, Position=1)]
            [string]$SKU2,

        [Parameter(ParameterSetName='object',Mandatory, Position=1)]
            $objSKU
    )

    $spInfo = get-ServicePlanInfoFile

    # Resolve SKU info using your existing find-SKU
    $sku1Obj = find-SKU -id $SKU1 -exact
    if (-not $sku1Obj) {
        throw "SKU-1 not found."
    }

    if($PSCmdlet.ParameterSetName -eq 'string') {
        $sku2Obj = find-SKU -id $SKU2 -exact
        if (-not $sku2Obj) {
            throw "SKU-2 not found."
        }
    } else {
        if(-not $objSKU.ServicePlans.ServicePlanId) {
            throw "wrong object for SKU2"
        }
    }

    #get Service Plan GUIDs for each SKU
    $plans1 = $spInfo|? {$_.Product_Display_Name -eq $sku1Obj.SKUFriendlyName} 
    $guids1 = $plans1.Service_Plan_Id

    if($PSCmdlet.ParameterSetName -eq 'string') {
        $plans2 = $spInfo | Where-Object { $_.GUID -eq $sku2Obj.SKUGUID }
        $guids2 = $plans2.Service_Plan_Id
    } else {
        $guids2 = $objSKU.ServicePlans.ServicePlanId
    }

    $onlyIn1 = $guids1 | Where-Object { $_ -notin $guids2 }
    $onlyIn2 = $guids2 | Where-Object { $_ -notin $guids1 }
    $inBoth  = $guids1 | Where-Object { $_ -in $guids2 }

    if($PSCmdlet.ParameterSetName -eq 'string') {
        $sku2FriendlyName = $($sku2Obj.SKUFriendlyName)
    } else {
        $sku2FriendlyName = $objSKU.SkuPartNumber
    }

    Write-Host "`nPlans in $($sku1Obj.SKUFriendlyName) but not in $sku2FriendlyName :" -ForegroundColor 'Magenta'
    if(-not $onlyIn1) {
        write-host "none."
    } else {
        $onlyIn1 | %{ find-ServicePlan $_ | Select-Object -First 1 } | Format-Table
    }
    Write-Host "`nPlans in $sku2FriendlyName but not in $($sku1Obj.SKUFriendlyName) :" -ForegroundColor 'Yellow'
    if(-not $onlyIn2) {
        write-host "none."
    } else {
        $onlyIn2 | %{ find-ServicePlan $_ | Select-Object -First 1 } | Format-Table
    }
    Write-Host "`nPlans common to both:" -ForegroundColor 'Green'
    if(-not $inBoth) {
        write-host "none."
    } else {
        $inBoth | %{ find-ServicePlan $_ | Select-Object -First 1 } | Format-Table
    }
}

function set-GroupLicense {
    <#
.SYNOPSIS
    assignes license (SKU) to a group allowing to choose disabled Service Plans. 
    uses gridview component to facilietate selection of group and SKU.
.DESCRIPTION
    here be dragons
.EXAMPLE
    set-eNGBLGroupLicense

    runs the function allowing to choose all elements from the Grid View list.
.EXAMPLE
    set-eNGBLGroupLicense -GroupID 12345678-1234-1234-1234-123456789012 -SKUId 12345678-1234-1234-1234-123456789012 

    skips searching the group and SKU using provided values and displays only Service Plans to be disabled from SKU.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250412
        last changes
        - 250412 initialized

    #TO|DO
    - disabled service plans are lacking otpion for automation...
    #>
    
#requires -module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
            [string]$GroupID,
        [Parameter(Position=1)]
            [string]$SKUId,
        [Parameter(Position=2)]
            [switch]$Force
    )

    $VerbosePreference = 'Continue'
    $ctx = connect-LicenseGraph

    if(-not $GroupID) {
        Write-debug 'grp'
        $group = get-mgGroup -Filter "securityEnabled eq true" -Property DisplayName,Id,AssignedLicenses | 
            Select-Object DisplayName,Id,@{L='licenses';E={($_.AssignedLicenses.SkuId | find-SKU).SKUFriendlyName }} | 
            Sort-Object DisplayName |
            Out-GridView -Title "Choose group" -OutputMode Single
        if($null -eq $group) {
            throw "No group selected."
        }
    } else {
        try {
            $group = get-mgGroup -GroupId $GroupID | Select-Object DisplayName,Id
        } catch {
            throw "group with ID $GroupID not found."
        }
    }
    Write-Verbose "group chosen: $($group.DisplayName):$($group.Id)"

    #$spInfo = get-ServicePlanInfoFile
    if(-not $SKUId) {
        write-debug 'skuid'
        $sku = Get-MgSubscribedSku -All| 
            Select-Object @{L='License';E={(find-SKU $_.SkuPartNumber -exact).SKUFriendlyName}},
                SkuPartNumber,
                @{L='AvailableLicenses';E={$_.PrepaidUnits.Enabled}},
                consumedUnits,SkuId | 
            Sort-Object License |
            Out-GridView -Title "Choose SKU" -OutputMode Single
        if($null -eq $sku) {
            throw "No SKU selected."
        }
    } else {
        try {
            $sku = Get-MgSubscribedSku -Id $SKUId | 
                Select-Object @{L='License';E={(find-SKU $_.SkuPartNumber -exact).SKUFriendlyName}},
                    SkuPartNumber, 
                    @{L='AvailableLicenses';E={$_.PrepaidUnits.Enabled}},
                    consumedUnits,SkuId
        } catch {
            throw "SKU with ID $SKUId not found."
        }
    }
    Write-Verbose "SKU chosen: $($sku.SkuPartNumber)"

    $disabledPlans = show-ServicePlans -id $sku.SkuPartNumber | Out-GridView -title "show plans to disable or cancel to assign full license" -OutputMode Multiple
    if($disabledPlans) {
        $licenses = @{
            AddLicenses = @(@{
                SkuId = $sku.SkuId
                DisabledPlans = $disabledPlans.ServicePlanGUID
            })
            RemoveLicenses = @()
        }
    } else {
        Write-Verbose "no plans to disable"
        $licenses = @{
            AddLicenses = @(@{
                SkuId = $sku.SkuId
            })
            RemoveLicenses = @()
        }
    }

    write-host "`nGroup '$($group.DisplayName)' will be assigned '$($sku.License)'" -ForegroundColor Magenta
    if($disabledPlans) {
        write-host "Disabled plans:" -ForegroundColor Yellow
        $disabledPlans | Format-Table
    }

    if (-not $Force) {
        $answer = Read-Host "Do you want to continue with assignment? [Y/N]"
        if ($answer -notin @('Y','y','Yes','yes')) {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }
    }
    set-MgGroupLicense -GroupId $group.Id -BodyParameter $licenses

}

function get-LicenseAssignment {
<#
.SYNOPSIS
    user license reporting function.
    available reporting options:
    - all users and their license information
    - only users with certain license assigned
    - only particular users by providing their Ids
    - standard report (compact) and extended version
    report may be returned as object to support other functions or exported to CSV file.
.DESCRIPTION
    here be dragons
.EXAMPLE
    get-eNGBLUserLicenseInfo

    prepares full list of users with their license information.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250414
        last changes
        - 250414 extended version
        - 250412 initialized

    #TO|DO
    - add detailed reporting including expluded SP - each license as a seperate column?
    - main function showing and analysing license assignment - neends options such as 'dupes', userId, process() 
#>

    [cmdletbinding()]
    param(
        #by default report is exported to CSV file. use this switch if you want to return it as object.
        [Parameter(mandatory=$false,position=0)]
            [switch]$asObject,
        #return only users with certain license assigned. you may use SKU friendly name, code name or GUID.
        [Parameter(mandatory=$false,position=1)]
            [string]$SKUId,
        #limit to certain account(s) by providing their object IDs.
        [Parameter(mandatory=$false,position=2)]
            [string[]]$userId,
        #extended report
        [Parameter(position=3)]
            [switch]$extendedReport
    )

    begin {
        function fetch-nameCache {
            param(
                [parameter(mandatory,position=0)]
                    [string]$objectID,
                [parameter(position=1)]
                    [validateSet('Group','SKU','SP')]
                    [string]$objectType = 'Group'
            )
    
            #check if object has already been queried
            if($displayNameCache.ContainsKey($objectID)) {
                Write-Debug "existing : $objectType : $($displayNameCache[$objectID])"
                return $displayNameCache[$objectID]
            } else {
                switch($objectType) {
                    'Group' {
                        try {
                            $group = Get-MgGroup -GroupId $objectID
                            #cache the group for later use
                            $global:displayNameCache[$objectID] = $group.DisplayName
                        } catch {
                            Write-Error "Group with ID $objectID not found."
                        }
                    }
                    'SKU' {
                        try {
                            $sku = find-SKU -id $objectID -exact
                            #cache the SKU for later use
                            $global:displayNameCache[$objectID] = $sku ? $sku.SKUFriendlyName : ($tenantSKUs | Where-Object { $_.SkuId -eq $objectID } | Select-Object -ExpandProperty SkuPartNumber)
                        } catch {
                            Write-Error "SKU with ID $objectID not found."
                            $_.Exception
                        }
                    }
                    'SP' {
                        try {
                            $sp = find-ServicePlan -id $objectID -exact
                            #cache the SKU for later use
                            $global:displayNameCache[$objectID] = $sp[0].ServicePlanFriendlyName
                        } catch {
                            Write-Error "Service Plan with ID $objectID not found."
                        }
                    }
                }
                if(-not $displayNameCache[$objectID]) {
                    $displayNameCache[$objectID] = $objectID
                }
                return $displayNameCache[$objectID]
            }
        }
        $allUserLicenses = @()
        $global:displayNameCache = @{}
        $ctx = connect-licenseGraph
        $domain = $ctx.Account.Split('@')[1]
        $outFile = "{0}-userLicenses-{1}.csv" -f $domain,(get-date).ToString('yyyyMMdd-HHmmss')
        $tenantSKUs = Get-MgSubscribedSku -All -ErrorAction Stop | Select-Object SkuId, SkuPartNumber,ConsumedUnits

        #check if SKUId exists in the tenant
        if($SKUId) {
            Write-Debug "looking for SKU $SKUId"
            $skuFilter = find-SKU -id $SKUId -exact
            if($null -eq $skuFilter) {
                Write-Error "SKU with ID $SKUId not found."
                return
            }
        } else {
            write-debug 'no skuid chosen'
        }

        #prepare user list
        if($userId) {
            Write-Debug "looking for user(s) $userId"
            $eIDUsersToProcess = @()
            # Retrieve all users in the tenant with required properties
            foreach($uid in $userId) {
                try {
                    $eIDUser = Get-MgUser -userId $uid -Property AssignedLicenses, LicenseAssignmentStates, DisplayName, Id, UserPrincipalName, AccountEnabled -ErrorAction Stop
                    $eIDUsersToProcess += $eIDUser
                } catch {
                    Write-Error "User with ID $uid not found."
                }
            }
        } else {
            write-debug "getting all tenant users"
            # Retrieve all users in the tenant with required properties
            $eIDUsersToProcess = Get-MgUser -All -Property AssignedLicenses, LicenseAssignmentStates, DisplayName, Id, UserPrincipalName, AccountEnabled -ErrorAction Stop
        }
        #ensure the is at least one user to process (;
        if(!$eIDUsersToProcess) {
            write-host ("there are no users having '{0}' ({1}) license assigned." -f $SKUId,$skuFilter.SKUFriendlyName)
            return
        }

        #post-filtering - only users with certain license assigned
        if($SKUId) {
            # Retrieve all users in the tenant with required properties
            $eIDUsersToProcess = $eIDUsersToProcess | Where-Object { $_.AssignedLicenses.SkuId -contains $skuFilter.SKUGUID }
            if(!$eIDUsersToProcess) {
                write-host ("there are no users having '{0}' ({1}) license assigned." -f $SKUId,$skuFilter.SKUFriendlyName)
                return
            }
        }
    }

    process {
        #expand license infrormation with human readable names
        foreach ($user in $eIDUsersToProcess) {
            # Add the user's license information to the array
            # Construct a custom object to store the user's license information
            $extendedUserInfo = [PSCustomObject]@{
                UserId = $user.Id
                UserDisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                AccountEnabled = $user.AccountEnabled
                licenseInfo = @()
            }
            if(-not [string]::IsNullOrEmpty($user.LicenseAssignmentStates) ) {
    
                # Loop through license assignment states
                foreach ($assignment in $user.LicenseAssignmentStates) {
                    $customLicensesAssignment = @{
                        AssignedByGroup = $assignment.AssignedByGroup
                        DisabledPlans = $assignment.DisabledPlans
                        Error = $assignment.Error
                        LastUpdatedDateTime = $assignment.LastUpdatedDateTime
                        SkuId = $assignment.SkuId
                        State = $assignment.State
                        groupDisplayName = ""
                        skuDisplayName = ""
                        disabledPlanDisplayName = @()
                    }
    #                fetch-nameCache -objectID $assignment.SkuId -objectType 'SKU'

                    #get SKU displayName
                    $customLicensesAssignment.skuDisplayName = fetch-nameCache -objectID $assignment.SkuId -objectType 'SKU'

                    #check group name
                    $assignedByGroup = $assignment.AssignedByGroup
                    $customLicensesAssignment.groupDisplayName = if ($null -ne $assignedByGroup) {
                        # If the license was assigned by a group, get the group name
                        fetch-nameCache -objectID $assignedByGroup -objectType 'Group'
                    } else {
                        # If the license was assigned directly by the user
                        "User"
                    }

                    #get disabled plans display names
                    if($assignment.DisabledPlans) {
                        foreach ($plan in $assignment.DisabledPlans) {
                            $planDisplayName = fetch-nameCache -objectID $plan -objectType 'SP'
                            $customLicensesAssignment.disabledPlanDisplayName += $planDisplayName
                        }
                    } 
                    $extendedUserInfo.licenseInfo += $customLicensesAssignment
                }
            }
            $allUserLicenses += $extendedUserInfo
            
        }    
    }
    end {
        #filter users - only with certain SKU assigned
        if($extendedReport) {
            $extendedUserReport = @()
            foreach($entry in $allUserLicenses) {
                $extendedEntry = [PSCustomObject]@{
                    UserDisplayName = $entry.UserDisplayName
                    UserPrincipalName = $entry.UserPrincipalName
                    AccountEnabled = $entry.AccountEnabled
                    UserId = $entry.UserId
                }
                foreach($sku in $tenantSKUs) {
                    #check if user has this licence
                    $p = $entry.licenseInfo|? skuid -eq $sku.SkuId
                    $extendedEntry | Add-Member -MemberType NoteProperty -Name $sku.SkuPartNumber -Value ($p ? (convert-toShortStringExt -licenseInfo $p -skuId $sku.SkuId) : "" ) -Force
                }
                $extendedUserReport += $extendedEntry
            }
            $allUserLicenses = $extendedUserReport
        } else {
            $allUserLicenses = $allUserLicenses | 
                Select-Object UserDisplayName,UserPrincipalName,AccountEnabled,UserId, @{L='Licenses';E={ convert-toShortString $_.licenseInfo }} 
        }
        if($asObject) { 
            return $allUserLicenses
        } 
        
        $allUserLicenses | Export-Csv -Path $outFile -NoTypeInformation -Force -Encoding UTF8
        Write-Verbose "exported to $outFile. converting to XLSX..."
        convert-CSV2XLS -CSVfileName $outFile -openOnConversion

    }
}

function move-betaUserstoGroupLicense {
    <#
.SYNOPSIS
    ::this function is in early beta - may greatly change in the future::
    find users with certain license, assign license to a chosen group, add person to a group, remove direct assignment
.DESCRIPTION
    here be dragons
.EXAMPLE
    move-eNGBLbetaUsersToGroupLicense

    run the script 
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://learn.microsoft.com/en-us/entra/identity/users/licensing-powershell-graph-examples
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250403
        last changes
        - 250403 initialized

    #TO|DO
    - missing option to disable service plans
    - function currently to migrate licenses users to GBL. 
#>

    [CmdletBinding()]
    param (
    )
    $VerbosePreference = 'SilentlyContinue'
    $ctx = connect-licenseGraph

    # Retrieve all SKUs in the tenant
    $SKUs = Get-MgSubscribedSku -All | Select-Object SkuId, SkuPartNumber
    $chosenSKU = $SKUs | Out-GridView -Title "choose license to migrate to GBL" -OutputMode Single
    if($null -eq $chosenSKU) { 
        Write-Host 'cancelled.'
        return
    }

    #prepare list of groups with their assigned licenses and allows selection 
    $allGroups = Get-MgGroup -All -Property DisplayName, GroupTypes, Description, AssignedLicenses, Id
    $GBLgroup = $allGroups | 
        Sort-Object DisplayName | 
        Select-Object DisplayName, @{L='licenses';E={$_.AssignedLicenses.skuId -join ','}},Id,Description | 
        Out-GridView -Title "choose GBL group to assign license to" -OutputMode Single
    if($null -eq $GBLgroup) {
        write-host 'cancelled.'
        return
    } 

    #checks if group already has a license - if no, assign it
    if($GBLgroup.licenses -match $chosenSKU.skuId) {
        Write-Warning ("group '{0}' already has '{1}' license assigned." -f $GBLgroup.displayName, $chosenSKU.SkuId)
    } else {
        Write-Verbose ("assigning '{0}' license to group '{1}'" -f $chosenSKU.SkuId,$GBLgroup.displayName)
        $params = @{
            AddLicenses = @(
                @{
                    SkuId = $chosenSKU.skuId
                }
            )
            # Keep the RemoveLicenses key empty as we don't need to remove any licenses
            RemoveLicenses = @()
        }
        #ERROR HANDLING - do not continue if error occurs
        <#
        | License assignment failed because service plan 4828c8ec-dc2e-4779-b502-87ac9ce28ab7 depends on the service plan(s)  
        #>
        Set-MgGroupLicense -GroupId $GBLgroup.Id -BodyParameter $params
    }    

    # Retrieve all users in the tenant with required properties
    $eIDUsersToProcess = Get-MgUser -All -Property AssignedLicenses, LicenseAssignmentStates, DisplayName, Id, UserPrincipalName 
    $licensedUsers = $eIDUsersToProcess | Where-Object {
        $_.AssignedLicenses.SkuId -contains $chosenSKU.SkuId
    }
    if(!$licensedUsers) {
        write-host ("there are no users having '{0}' ({1}) license assigned." -f $chosenSKU.SkuId,$chosenSKU.SkuPartNumber)
        return
    }

    #although users are filtered to only licensed users, we don't know assignemnt source - GBL or direct. let's get the details
    $allUserLicenses = @()
    foreach ($user in $licensedUsers) {
        $assignmentMethods = @()
        # Loop through license assignment states
        foreach ($assignment in $user.LicenseAssignmentStates|? SkuId -eq $chosenSKU.skuId) {
            $skuId = $assignment.SkuId
            $assignedByGroup = $assignment.AssignedByGroup
            $assignmentMethod = if ($null -ne $assignedByGroup) {
                # If the license was assigned by a group, get the group name
                $group = Get-MgGroup -GroupId $assignedByGroup
                if ($group) { $group.DisplayName } else { "Unknown Group" }
            } else {
                # If the license was assigned directly by the user
                "User"
            }
            $assignmentMethods += $assignmentMethod
        }
    
        # Construct a custom object to store the user's license information
        $userLicenseInfo = [PSCustomObject]@{
            UserId = $user.Id
            UserDisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            SkuId = $skuId
            SkuPartNumber = $chosenSKU
            AssignedBy = $assignmentMethods -join ','
        }

        # Add the user's license information to the array
        $allUserLicenses += $userLicenseInfo
    }

    # Export the results to a CSV file
    #$path = Join-path $env:LOCALAPPDATA ("UserLicenseAssignments_" + [string](Get-Date -UFormat %Y%m%d) + ".csv")
    $usersToMove = $allUserLicenses | 
        Sort-Object UserDisplayName | 
        Out-GridView -Title 'choose users to move to GBL group' -OutputMode Multiple
    if($null -eq $usersToMove) {
        write-host 'no users chosen.'
        return
    } 
    foreach($user in $usersToMove) {
        #add to group first
        try {
            New-MgGroupMember -GroupId $GBLgroup.Id -DirectoryObjectId $user.UserId
        } catch {
            Write-Error $_.Exception
            continue
        }
        #then remove direct license
        try {
            Set-MgUserLicense -UserId $user.UserId -RemoveLicenses $chosenSKU.SKUId -AddLicenses @() -ErrorAction Stop | Out-Null
            write-verbose "user '{0}' has a '{1}' license removed." -f $user.UserDisplayName,$chosenSKU.SKUId
        } catch {
            $_.Exception
        }
    }
    # Display the location of the CSV file
    #Write-Host "CSV file generated at: $((Get-Item $path).FullName)"
    write-host -fore Green 'done.'
}

function remove-betaUserLicenseAssignment {
<#
.SYNOPSIS
    ::this function is in early beta - may greatly change in the future::
    Helps remove duplicate direct licenses - when coming from two sources (group, direct)
.DESCRIPTION
    here be dragons
.EXAMPLE
    remove-betaUserLicenseAssignment
    
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://learn.microsoft.com/en-us/entra/identity/users/licensing-powershell-graph-examples
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250403
        last changes
        - 250403 initialized

    #TO|DO
    - needs to be deviced on two separate function: reporting function and solely removal function. then if will be
    combination of both: 
        find-dupeAssignments -skuid | remove-dupeAssignments 
        find-dupeAssignments  #displays all duplicate assignments - this might be get-tenantLicenseAssignments with parameter. 

    - to different cleanups - how to design 'clean dupes'? it's easy for all but not with a choice -> each dupe as separate row 
        - other cleanup is per skuid
        - pseudo-dupe: define sku1 and sku2 - show all accounts that have both SKUs assigned.
#>

    [CmdletBinding(DefaultParameterSetName = 'all')]
    param (
        #default - simply show all duplicate licenses
        [Parameter(ParameterSetName='all',mandatory=$false,position=0)]
            [switch]$all=$true,
        #SKUid to filter cleanup to only single SKUid
        [Parameter(ParameterSetName='bySKU',mandatory=$false,position=0)]
            [string]$SKUId,
        #limit cleanup to dupe assignments only
        [Parameter(ParameterSetName='bySKU',mandatory=$false,position=1)]
            [switch]$dupes
    )
    $VerbosePreference =  'Continue'

    $ctx = connect-licenseGraph
    Write-Verbose "getting tenant SKUs..."
    $tenantSKUs = Get-MgSubscribedSku | 
        Select-Object @{L='SKUFriendlyName';E={$fn = (find-SKU $_.SkuPartNumber -exact).SKUFriendlyName; $fn ? $fn : $_.SkuPartNumber}},
            SkuPartNumber,
            @{L='AvailableLicenses';E={$_.PrepaidUnits.Enabled}},
            consumedUnits,SkuId |
        Sort-Object SKUFriendlyName

    #which license you want to remove?
    if(-not $SKUId) {
        $chosenSKU = $tenantSKUs | Out-GridView -Title "Choose SKU" -OutputMode Single
        if($null -eq $chosenSKU) {
            throw "No SKU selected."
        }
        $SKUId = $chosenSKU.SkuId
    } else {
        $chosenSKU = $tenantSKUs | Where-Object { $_.SkuId -eq $SKUId } | Select-Object -First 1
    }

    #look for users with provided SKU assigned
    $licensesToCleanup = get-LicenseAssignment -asObject -SKUId $SKUId
    if($null -eq $licensesToCleanup) {
        write-host "no assignments with provided parameters found."
        return
    } 
    #split the view on the license we currently work on and all the rest (informationally only)
    $tmpUserLicesenses = @()
    foreach($userLic in $licensesToCleanup) {
        $groupedLicenses = $userLic.licenseInfo | Group-Object SKUId 
        if($dupes) {
            if( ($groupedLicenses | ? name -eq $SKUId).count -lt 2 ) { continue }
        } 
        $tmpUserLicesenses += $userLic | Select-Object UserDisplayName,UserPrincipalName,AccountEnabled,`
            @{L='searched license';E={ convert-toShortString ($groupedLicenses | ? name -eq $SKUId).group }}, `
            @{L='other licenses';E={ convert-toShortString ($groupedLicenses | ? name -ne $SKUId).group }}, `
            UserId
    }
    #there might be no (dupe) assignments
    if(-not $tmpUserLicesenses) {
        $str = ""
        if($dupes) { $str = 'duplicated' }
        write-host "no users with $str '$($chosenSKU.SKUFriendlyName)' license found."
        return
    }
    $licensesToCleanup = $tmpUserLicesenses | Sort-Object UserDisplayName | Out-GridView -Title "choose users for  license removal" -OutputMode Multiple
    if($null -eq $licensesToCleanup) {
        write-host "choice cancelled."
        return
    } 
    Write-Warning "DO YOU WANT TO REMOVE '$($chosenSKU.SKUFriendlyName)' LICENSE FROM BELOW USERS?"
    $licensesToCleanup | select-object UserDisplayName,AccountEnabled,'searched license' | Out-Host
    $answer = get-answerBox -OKButtonText "continue" -title 'choose users to clean' -message 'are you sure you want to remove user license assignements?' -icon Exclamation 
    if (-not $answer) { return }
    
    foreach($user in $licensesToCleanup) {
        Write-Verbose "processing user '$($user.UserDisplayName)'..."
        #remove direct license
        try {
            Set-MgUserLicense -UserId $user.UserId -RemoveLicenses $chosenSKU.SKUId -AddLicenses @{} -ErrorAction Stop | Out-Null
            write-verbose ("user '{0}' has a '{1}' license removed." -f $user.UserDisplayName,$chosenSKU.SKUId)
        } catch {
            write-error ("{0};{1}" -f $_.Exception.HResult,$_.Exception.Message)
        }
    }
    # Display the location of the CSV file
    #Write-Host "CSV file generated at: $((Get-Item $path).FullName)"
    write-host -fore Green 'done.'
}