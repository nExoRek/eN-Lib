<#
.SYNOPSIS
    provides immutableID for Cloud Sync
.DESCRIPTION
    used for hard-matching during Cloud Sync deployment. script search for the user by SAM or UPN, generates immutableID
    and returns an object that might be easily converted into CSV for the later push

    later may be pushed by:
    connect-MgGraph -Scopes "Directory.ReadWrite.All"
    after connecting with mgGraph to a tenat:
    get-mguser -UserId <$upn or userID> -Property userPrincipalName,displayname,mail,OnPremisesImmutableID|select userPrincipalName,displayname,mail,OnPremisesImmutableID
    Update-MgUser -UserId <$upn OR ID> -OnPremisesImmutableId <$immutableID>

.EXAMPLE
    .\get-ADUserImmutableID.ps1 nexor@w-files.pl

    checks if the user exists in the AD, if so - provides an object with sam, upn and immutableID

.EXAMPLE
    cat usersSAMlist.txt | %{ .\get-ADUserImmutableID.ps1 -sam $_ } | export-csv -nti userImmutableIDs.csv

    enumerates all entires in a text file 'usersSAMlist.txt' that should be a flat text file. 
    then for each entry looks for the user and retrun an object.
    eventually dumps the reult in the csv file to be later used to be pushed to EntraID.
.INPUTS
    None.

.OUTPUTS
    object containing sam, upn and calculated immutableID
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 241126
        last changes
        - 241126 initialized

    #TO|DO
#>
[CmdletBinding(DefaultParameterSetName = 'upn')]
param( 
    # attribute to search for the user in AD
    [Parameter(Mandatory,ParameterSetName = 'sam')]
        [validateNotNullOrEmpty()]
        [string]$sam,
    # attribute to search for the user in AD
    [Parameter(Mandatory,ParameterSetName = 'upn')]
        [validateNotNullOrEmpty()]
        [string]$upn
)

if($PSCmdlet.ParameterSetName -eq 'sam') {
    $adUser = Get-ADUser $sam 
} else {
    $adUser = get-ADUser -filter "userPrincipalName -eq '$upn'"
}
if(!$adUser) {
    write-host "$upn$sam not found." 
    return
}
write-verbose ("user '{0}' has object ID '{1}'" -f $sam,$adUser.objectGuid)
$immutableID = [Convert]::ToBase64String([guid]::New($adUser.objectGuid).ToByteArray())
write-verbose "immutableID: $immutableID"
[PSCustomObject]@{
    userPrincipalName = $adUser.userPrincipalName
    sAMAccountName = $adUser.sAMAccountName
    immutableID = $immutableID
}
