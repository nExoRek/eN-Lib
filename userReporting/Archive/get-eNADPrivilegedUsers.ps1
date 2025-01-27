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
#requires -module ActiveDirectory
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
