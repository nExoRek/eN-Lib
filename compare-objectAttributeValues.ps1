<#
.SYNOPSIS
    compares values on two object attributes
.DESCRIPTION
    Compare-Object can't actually compare objects but rather tables. this script intends
    to take all attribute values from object and compare them with attribute values
    against the other one. 
    both objects must be same type or share same/similar attributes to make comparison sensible. 
.EXAMPLE
    .\compare-objectAttributeValues.ps1 (get-aduser user1) (get-aduser user2)
    
    shows differences between general attributes of user1 and user2
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201016

    #TO|DO
    - check object types and add 'force' flag to skip comparison
    - separately list attribute names, that do not exist on the other object
    - compare complex values such as arrays and objects
#>

[cmdletbinding()]
param(
    [parameter(mandatory=$true,position=1)]
        [alias('obj1')]
        [psobject]$refObject,
    [parameter(mandatory=$true,position=2)]
        [alias('obj2')]
        [psobject]$compareTo
)


$refObject.Psobject.Properties|%{ 
    if($_.value -ne $compareTo.$($_.name)) { 
        write-host -NoNewline -ForegroundColor Yellow "$($_.name)"
        write-host -NoNewline " -> "
        write-host -NoNewline -ForegroundColor DarkYellow "$($_.value); "
        write-host "$($compareTo.$($_.name)) " 
    } 
}
