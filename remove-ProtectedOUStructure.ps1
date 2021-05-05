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
    .\remove-ProtectedOUStructure.ps1 myBestOU -removeAllFound
    
    will search for 'myBestOU' in the domain. if name is not unique, will enumerate thru all found objects allowing
    to choose if to delete it one by one.
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
    version 210505
    changes:
     - 210505 remove all, fixes to logic
     - 201015 initialization
#>
#REQUIRES -module ActiveDirectory
[CmdletBinding()]
param (
    #OU name or DistinguishedName
    [Parameter(mandatory=$true,position=0)]
        [string]$OUName,
    #use to delete multiple OUs with the same name under different parents.
    [Parameter(mandatory=$false,position=1)]
        [switch]$removeAllFound
)

function remove-OU {
    param([string]$OUDistinguishedName)
    
    write-host -ForegroundColor yellow "!!!THIS WILL PERMANENTLY REMOVE " -NoNewline
    write-host -BackgroundColor Black -ForegroundColor Red $OUDistinguishedName -NoNewline
    write-host -ForegroundColor yellow " AND ALL SUBOUs."
    Write-Host -ForegroundColor yellow "type [capital] Y to continue"
    $key = [console]::ReadKey()
    if ($key.keyChar -cne 'Y') {
        write-host "`nCancelled by user choice. Skipping." -ForegroundColor Yellow
        return -1
    }

    write-host "`nas you wish..."
    Get-ADObject -SearchBase $OUDistinguishedName -Filter *|Set-ADObject -ProtectedFromAccidentalDeletion $false
    Remove-ADOrganizationalUnit -Identity $OUDistinguishedName -Recursive -Confirm:$false
    write-host -ForegroundColor green 'removed.'
}

if($OUName -notmatch 'OU=') {
    $search=Get-ADObject -Filter "objectClass -eq 'organizationalUnit' -and name -eq ""$OUName"""
    write-host "found:"
    $search|Select-Object distinguishedname|Out-Host
    if($search.count -gt 1 -and -not $removeAllFound) {
        write-host -ForegroundColor red 'inconclusive'
        write-host "use 'removeAllFound' to remove multiple OUs."
        exit -6
    }
} else {
    try {
        Get-ADOrganizationalUnit -Identity $OUName -ErrorAction Stop
    } catch {
        write-host -ForegroundColor Red "$OUName not found."
        exit -5
    }
    $search=@($OUName)
}

foreach($ou in $search) {
    remove-OU -OUDistinguishedName $ou
}
write-host "`nfinished."
