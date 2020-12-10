<#
.SYNOPSIS
    reads all groups from AD and puts name and location to file
.DESCRIPTION
    some simple migration support...
.EXAMPLE
    .\list-ADGroupsToCSV.ps1
    
    creates CSV file with default name with list of groups and their location.
.INPUTS
    None.
.OUTPUTS
    CSV list
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201210
        last changes
        - 201210 initialized
#>
#requires -module ActiveDirectory
[CmdletBinding()]
param (
    #output file nam
    [Parameter(mandatory=$false,position=0)]
        [string]$outputCSV="_groupList-$(get-date -format yyMMddhhmm).csv",
    #delimiter for CSV
    [Parameter(mandatory=$false,position=1)]
        [string][validateSet(',',';')]$delimiter=';'
    
)
[regex]$rxLocation="^CN=.*?,(?<loc>.*?),DC" #regular expression to get only OU location part, without domain and object
Get-ADGroup -Filter * |
    Select-Object name,GroupCategory,GroupScope,@{N="location";E={$rxLocation.Match($_.distinguishedName).groups['loc'].value}} |
    export-csv -Delimiter $delimiter -NoTypeInformation -Encoding UTF8 -Path $outputCSV
write-host -ForegroundColor green "group location dumpped to .\$outputCSV"
