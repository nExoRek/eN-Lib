<#
.SYNOPSIS
    Short description
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
 
# Retrieve all users in the tenant with required properties
$users = Get-MgUser -All -Property AssignedLicenses, LicenseAssignmentStates, DisplayName, Id, UserPrincipalName
 
# Initialize an empty array to store the user license information
$allUserLicenses = @()
 
foreach ($user in $users) {
    # Initialize a hash table to track all assignment methods for each license
    $licenseAssignments = @{}
 
    # Loop through license assignment states
    foreach ($assignment in $user.LicenseAssignmentStates) {
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
 
        # Ensure all assignment methods are captured
        if (-not $licenseAssignments.ContainsKey($skuId)) {
            $licenseAssignments[$skuId] = @($assignmentMethod)
        } else {
            $licenseAssignments[$skuId] += $assignmentMethod
        }
    }
 
    # Process assigned licenses
    foreach ($skuId in $licenseAssignments.Keys) {
        # Get SKU details from the pre-fetched list
        $sku = $skus | Where-Object { $_.SkuId -eq $skuId } | Select-Object -First 1
        $skuPartNumber = if ($sku) { $sku.SkuPartNumber } else { "Unknown SKU" }
 
        # Sort and join the assignment methods
        #$assignmentMethods = ($licenseAssignments[$skuId] | Sort-Object -Unique) -join ", "
 
        # Construct a custom object to store the user's license information
        $userLicenseInfo = [PSCustomObject]@{
            UserId = $user.Id
            UserDisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            SkuId = $skuId
            SkuPartNumber = $skuPartNumber
            AssignedBy = $licenseAssignments[$skuId]
        }
 
        # Add the user's license information to the array
        $allUserLicenses += $userLicenseInfo
    }
}
 
# Export the results to a CSV file
#$path = Join-path $env:LOCALAPPDATA ("UserLicenseAssignments_" + [string](Get-Date -UFormat %Y%m%d) + ".csv")
$dupes = @()
if($null -eq $allUserLicenses) {
    write-host 'cancelled.'
} else {
    foreach($user in $allUserLicenses) {
        if($user.assignedBy.count -gt 1) { #same license may be from several sources?
            $dupes += $user | select-object UserDisplayName,SkuPartNumber,@{L='assignments';E={($_.AssignedBy|sort -Unique) -join ',' }},UserId,SKUId
        }
    }
} 
if($dupes) {
    $dupes | Out-GridView -Title "select users to remove duplicate licenses" -OutputMode Multiple
    if($dupes) {
        foreach($dupe in $dupes) {
            try {
                Set-MgUserLicense -UserId $dupe.UserId -RemoveLicenses $dupe.SKUId -AddLicenses @() -ErrorAction Stop | Out-Null
                write-verbose "user '{0}' has a '{1}' license removed." -f $dupe.UserDisplayName,$dupe.SKUId
            } catch {
                $_.Exception
            }
        }
    } else {
        Write-Verbose 'cancelled.'
    }
} else {
    write-host 'no dupe licenses found'
}
# Display the location of the CSV file
#Write-Host "CSV file generated at: $((Get-Item $path).FullName)"
write-host -fore Green 'done.'
