<#
.SYNOPSIS
    find users with certain license, assign license to a chosen group, add person to a group, remove direct assignment
.DESCRIPTION
    here be dragons
.EXAMPLE
    .\tmp_lictype.ps1

    
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
#>

[CmdletBinding()]
param (
    #skip connect if you're already connected
    [Parameter(mandatory=$false,position=0)]
        [switch]$skipConnect
    
)
$VerbosePreference = 'SilentlyContinue'

if(!$skipConnect) {
    try {
        Disconnect-MgGraph -ErrorAction Stop
    } catch {
        #write-host 'testing error'
        write-verbose $_.Exception
        $_.ErrorDetails
    }
    Write-Verbose "athenticate to tenant..."
    try {
        Connect-MgGraph -Scopes "User.Read.All","Domain.Read.All","email","profile","openid" -ErrorAction Stop
    } catch {
        throw "error connecting. $($_.Exception)"
        return
    }
}
Write-Verbose "getting connection info..."
$ctx = Get-MgContext
Write-Verbose "connected as '$($ctx.Account)'"

# Retrieve all SKUs in the tenant
$SKUs = Get-MgSubscribedSku -All | Select-Object SkuId, SkuPartNumber
$chosenSKU = $SKUs | Out-GridView -Title "choose license to migrate to GBL" -OutputMode Single
if($null -eq $chosenSKU) { 
    Write-Host 'cancelled.'
    return
}

$allGroups = Get-MgGroup -All -Property DisplayName, GroupTypes, Description, AssignedLicenses, Id
$GBLgroup = $allGroups | Sort-Object DisplayName | Select-Object DisplayName, @{L='licenses';E={$_.AssignedLicenses.skuId -join ','}},Id,Description | Out-GridView -Title "choose GBL group to assign license to" -OutputMode Single
if($null -eq $GBLgroup) {
    write-host 'cancelled.'
    return
} 
#check if group already has a license - if no, assign it
if($GBLgroup.licenses -match $chosenSKU.skuId) {
    Write-Verbose ("group '{0}' already has '{1}' license assigned." -f $GBLgroup.displayName, $chosenSKU.SkuId)
} else {
    Write-Verbose ("assigning '{0}' license to group '{1}'" -f $chosenSKU.SkuId,$GBLgroup.displayName)
    $params = @{
        AddLicenses = @(
            @{
                SkuId = $chosenSKU.skuId
            }
        )
        # Keep the RemoveLicenses key empty as we don't need to remove any licenses
        RemoveLicenses = @(
        )
    }
    #ERROR HANDLING - do not continue if error occurs
    <#
       | License assignment failed because service plan 4828c8ec-dc2e-4779-b502-87ac9ce28ab7 depends on the service plan(s)  
    #>
    Set-MgGroupLicense -GroupId $GBLgroup.Id -BodyParameter $params
}    

# Retrieve all users in the tenant with required properties
$users = Get-MgUser -All -Property AssignedLicenses, LicenseAssignmentStates, DisplayName, Id, UserPrincipalName 
$licensedUsers = $users | Where-Object {
    $_.AssignedLicenses.SkuId -contains $chosenSKU.SkuId
}
if(!$licensedUsers) {
    write-host ("there are no users having '{0}' ({1}) license assigned." -f $chosenSKU.SkuId,$chosenSKU.SkuPartNumber)
    return
}

#although users are filtered we don't know if licenses are assigned by group or direct. 

# Initialize an empty array to store the user license information
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
$usersToMove = $allUserLicenses | Sort-Object UserDisplayName | Out-GridView -Title 'choose user to move to GBL group' -OutputMode Multiple
if($null -eq $usersToMove) {
    write-host 'no users to move chosen'
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
