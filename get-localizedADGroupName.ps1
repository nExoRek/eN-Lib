#requires -module ActiveDirectory
[CmdletBinding()]
param (
    #group name - as a value from $wellKnownSIDs hashtable
    [Parameter(mandatory=$true,position=0)]
        [string]$name
    
)
#https://docs.microsoft.com/en-us/windows/win32/secauthz/well-known-sids
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

    "Enterprise Read-only Domain Controllers" = "{domainSID}-498"
    "Domain Admins"                      = "{domainSID}-512"
    "Domain Users"                       = "{domainSID}-513"
    "Domain Guests"                      = "{domainSID}-514"
    "Domain Computers"                   = "{domainSID}-515"
    "Domain Controllers"                 = "{domainSID}-516"
    "Cert Publishers"                    = "{domainSID}-517"
    "Schema Admins"                      = "{domainSID}-518"
    "Enterprise Admins"                  = "{domainSID}-519"
    "Group Policy Creator Owners"        = "{domainSID}-520"
    "Read-only Domain Controllers"       = "{domainSID}-521"
    "Cloneable Domain Controllers"       = "{domainSID}-522"
    CDC_RESERVED                         = "{domainSID}-524"
    "PROTECTED USERS"                    = "{domainSID}-525"
    "Key Admins"                         = "{domainSID}-526"
    "Enterprise Key Admins"              = "{domainSID}-527"
}
$domainSID = (get-adDomain).domainSID
$resolveSID = $wellKnownSids[$name].replace('{domainSID}',$domainSID)
return ([System.Security.Principal.SecurityIdentifier]$resolveSID).Translate([System.Security.Principal.NTAccount]).Value
