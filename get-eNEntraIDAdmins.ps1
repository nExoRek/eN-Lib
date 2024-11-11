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
#requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Identity.SignIns,Microsoft.Graph.Users
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

$sortedMemebersList = ($sortBy -eq 'Role') ? ($RoleMemebersList | Sort-Object RoleName) : ($RoleMemebersList | Select-Object userName,userID,enabled,RoleName,roleID | Sort-Object userName)

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
