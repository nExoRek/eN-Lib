<#
.SYNOPSIS
    generates report on AAD roles memebers.
.DESCRIPTION        
    AzureAD nor AzureADPreview is capable of getting user MFA status so MSOnline module is used. by default
    if gets the list by user as it gives capability to add some additional account information such as
    MFA status, enablement status etc. sorting by the group is more robust, but with very minium information
    on the role members only.
.EXAMPLE
    .\list-AADAdmins.ps1
    
    generates list of all AAD role memebers and output to csv file
.EXAMPLE
    .\list-AADAdmins.ps1 -sortBy role
    
    generates list of all AAD role groups and its memebers and output to csv file
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201029
        last changes
        - 201029 initialized
#>
#requires -module MSOnline
[CmdletBinding()]
param (
    #show (detail) information on members and their roles or just a list of role members
    [Parameter(mandatory=$false,position=0)]
        [string][validateSet('user','role')]$sortBy='user',
    #export CSV file delimiter
    [Parameter(mandatory=$false,position=1)]
        [string][validateSet(',',';')]$delimiter=';'
)

try {
    $MSOLdomains=Get-MsolDomain -ErrorAction Stop
} catch {
    Write-Host -ForegroundColor Red "you need to connect with connect-MSOLService before running the script"
    exit -1
}
($MSOLdomains|Where-Object IsInitial -eq $true|Select-Object -ExpandProperty name) -match '((?<tenantname>.*?)\.)'|out-null
$tenantName=$Matches['tenantname']
$outFile=$tenantName+"_AdminReport_"+(get-date -Format 'yyMMdd')+'.csv'
Write-Host "getting roles and members (may take up to 5 min)..."
$AzureADRoleMembers = Get-MsolRole|ForEach-Object{
    $rn=$_.name;Get-MsolRoleMember -RoleObjectId $_.objectid|
    Where-Object RoleMemberType -eq 'user'|
    Select-Object displayname,EmailAddress,@{N='role';E={$rn}},objectid
}

if($sortBy -eq 'user') {
    write-host "checking admin accounts status..."
    $AzureADRoleMembers=$AzureADRoleMembers|Select-Object *,enabled,MFAstatus,UserPrincipalName
    foreach($admin in $AzureADRoleMembers){
    #$AzureADRoleMembers|ForEach-Object {
        $msolMember=Get-MsolUser -ObjectId $admin.objectid
        $admin.enabled=(-not $msolMember.blockCredentials)
        $admin.MFAstatus=$msolMember.StrongAuthenticationRequirements.state
        $admin.UserPrincipalName=$msolMember.UserPrincipalName
    }
    
    $sortedMemebersList = $AzureADRoleMembers|Group-Object UserPrincipalName|
        Select-Object @{N='UserPrincipalName';E={$_.group.UserPrincipalName|Select-Object -unique}},
            @{N='displayname';E={$_.group.displayname|Select-Object -unique}},
            @{N='roles';E={($_.group.role|Sort-Object) -join ','}},
            @{N="enabled";E={$_.group.enabled|Select-Object -unique}},
            @{N='MFAStatus';E={$_.group.MFAstatus|Select-Object -unique}}
} else {
    $sortedMemebersList = $AzureADRoleMembers|Group-Object role|
        Select-Object name,@{N='displayname';E={$_.group.displayname|Select-Object -unique}}
}

Write-Host "exporting results, sorted by [$sortBy]..."
$sortedMemebersList|export-csv -nti -Delimiter $delimiter -Encoding UTF8 -Path $outFile

Write-Host -ForegroundColor Green "exported to .\$outFile.`ndone."
