#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$CsvOnly=$false)

function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param()

    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Double-check and try to eastablish a session
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-EXOMailbox, Get-User, Get-Group, Get-ManagementRoleAssignment -SkipLoadingFormatData -ShowBanner:$false } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    return $true
}

#Find the user matching a given DisplayName. If multiple entries are returned, use the -RoleAssignee parameter to determine the correct one. If unique entry is found, return UPN, otherwise return DisplayName
function getUPN ($user,$role) {
    $UPN = @(Get-User $user -ErrorAction SilentlyContinue | ? {(Get-ManagementRoleAssignment -Role $role -RoleAssignee $_.SamAccountName -ErrorAction SilentlyContinue)})
    if ($UPN.Count -ne 1) { return $user }
    if ($UPN) { return $UPN.UserPrincipalName }
    else { return $user }
}

#Find the group matching a given DisplayName. If multiple entries are returned, use the -RoleAssignee parameter to determine the correct one. If unique entry is found, return the email address if present, or GUID. Otherwise return DisplayName
function getGroup ($group,$role) {
    $grp = @(Get-Group $group -ErrorAction SilentlyContinue | ? {(Get-ManagementRoleAssignment -Role $role -RoleAssignee $_.SamAccountName -ErrorAction SilentlyContinue)})
    if ($grp.Count -ne 1) { return $group }
    if ($grp) {
        if ($grp.WindowsEmailAddress.ToString()) { return $grp.WindowsEmailAddress.ToString() }
        else { return $grp.Guid.Guid.ToString() }
    }
    else { return $group }
}

#Make sure we are connected to Exchange Remote PowerShell
if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }
$roleAssignments += Get-ManagementRoleAssignment

$output = @()
#Prepare the output. As we use the -getEffectiveUsers switch, we can get rid of "empty" Role groups
foreach ($ra in $RoleAssignments) {
    if ($ra.EffectiveUserName -eq "All Group Members" -and $ra.RoleAssigneeType -eq "RoleGroup") { continue } #skip empty role groups

    if ($ra.EffectiveUserName -eq "All Group Members" -and $ra.AssignmentMethod -eq "Direct") {
        if ($ra.RoleAssigneeType -ne "RoleGroup") { $principal = (getGroup $ra.RoleAssignee $ra.Role) }
        else { $principal = $ra.EffectiveUserName }
    }
    else { $principal = (getUPN $ra.EffectiveUserName $ra.Role) }

    $raobj = New-Object psobject
    $raobj | Add-Member -MemberType NoteProperty -Name AssignmentType -Value $ra.AssignmentMethod
    $raobj | Add-Member -MemberType NoteProperty -Name AssigneeName -Value $ra.RoleAssigneeName
    $raobj | Add-Member -MemberType NoteProperty -Name "Effective Assignee" -Value $principal
    $raobj | Add-Member -MemberType NoteProperty -Name AssignmentChain -Value (&{if ($ra.AssignmentChain) {$ra.AssignmentChain} else {$null}})
    $raobj | Add-Member -MemberType NoteProperty -Name AssigneeType -Value $ra.RoleAssigneeType
    $raobj | Add-Member -MemberType NoteProperty -Name AssignedRoles -Value $ra.Role
    $output += $raobj
}

$outputFiltered = $output | Where-Object {$_.'AssignedRoles' -match "ApplicationImpersonation" -or $_.'AssignedRoles' -match "Application Exchange Full Access" `
-or $_.'AssignedRoles' -match "Application Mail Full Access"} 
#Export the result to CSV file
$output | Select-Object * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.csv"
Write-Host "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.csv"
$outputFiltered | Select-Object * -ExcludeProperty Number | Export-CSV -nti -Path  "Filtered_$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.csv"
Write-Host "Output filtered for AppRoleAssignment.ReadWrite.All exported to $($PWD)\Filtered_$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.csv"
if (-not $CsvOnly) 
{ 
    Write-Host "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.xlsx"
    $output | Select-Object * -ExcludeProperty Number | Export-Excel -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.xlsx" -AutoSize -TableName "AppPermissions"
    Write-Host "Output filtered for AppRoleAssignment.ReadWrite.All exported to $($PWD)\Filtered_$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.xlsx"
    $outputFiltered | Select-Object * -ExcludeProperty Number | Export-Excel -Path "Filtered_$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ExchangeManagementRoleInventory.xlsx" -AutoSize -TableName "AppPermissions"
}