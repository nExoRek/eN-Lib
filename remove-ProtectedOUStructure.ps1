<#
.SYNOPSIS
    removes OU tree with 'protecy object from deletion' flags.
    !!!!USE WITH CARE!!!!
.DESCRIPTION
    although official document states, that -recusive removes all subOUs even protected ones... it somehow doesn't work for me
    https://docs.microsoft.com/en-us/powershell/module/addsadministration/remove-adorganizationalunit
    
    quote:
    >Note: Specifying this parameter removes all child objects of an OU that are marked with ProtectedFromAccidentalDeletion.

    script wrote for LAB environment when i had to create and remove lots of OU structures 
.EXAMPLE
    .\remove-ProtectedOUStructure.ps1 myBestOU
    
    will search for 'myBestOU' in the domain. if name is not unique, will terminate. if OU is found,
    'ProtectedFromAccidentalDeletion' flag will be removed from this OU and all sub OUs and then 
    delte entire tree recursively
.EXAMPLE
    .\remove-ProtectedOUStructure.ps1 'OU=myBestOU,DC=w-files,DC=pl'
    
    'ProtectedFromAccidentalDeletion' flag will be removed from this OU and all sub OUs and then 
    delte entire tree recursively.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201015
#>
#REQUIRES -module ActiveDirectory
[CmdletBinding()]
param (
    #OU name or DistinguishedName
    [Parameter(mandatory=$true,position=0)]
        [string]$OUName
)
if($OUName -notcontains 'OU=') {
    $search=Get-ADObject -Filter "objectClass -eq 'organizationalUnit' -and name -eq ""$OUName"""
    if($search.count -gt 1) {
        write-host -ForegroundColor red 'inconclusive'
        $search|Select-Object distinguishedname
        exit -6
    }
}

write-host -ForegroundColor yellow "!!!THIS WILL PERMANENTLY REMOVE $OUName AND ALL SUBOUs."
Write-Host -ForegroundColor yellow "type [capital] Y to continue"
$key = [console]::ReadKey()
if ($key.keyChar -cne 'Y') {
    write-host "`nScript ended by user choice. Quitting." -ForegroundColor Yellow
    exit -1
}

write-host "`nas you wish..."
Get-ADObject -SearchBase $OUName -Filter *|Set-ADObject -ProtectedFromAccidentalDeletion $false
Remove-ADOrganizationalUnit -Identity $OUName -Recursive -Confirm:$false
write-host -ForegroundColor green 'done.'
