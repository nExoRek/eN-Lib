<#
.SYNOPSIS
    prepare report file with AccountSku information from the tenant.
.DESCRIPTION
    it is handy extension to get-AccountSku which formats nice report table including marketing Sku names.
    you need to connect with connect-MSOLService before running the report.
    script is using eNLib support library - use 'install-module eNLib'.
.EXAMPLE
    connect-MSOLService
    .\get-AccountSkuReport.ps1
    
    generate CSV report file
.EXAMPLE
    &(convert-CSV2XLS (.\get-AccountSkuReport.ps1))

    generates the CSV report file, converts it to XLSX with eNLIB's convert function and automatically opens Excel
.INPUTS
    None.
.OUTPUTS
    CSV report file
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 211030
        last changes
        - 211030 initialized

    #TO|DO
#>
#Requires -Module MSOnline, eNLib
[CmdletBinding()]
param (
   
)

try {
    $AccountSkus = Get-MsolAccountSku 
} catch {
    write-log "error getting tenant SKUs. $($_.Exception)" -type error
    exit
}
$AccountName = $AccountSkus[0].accountName
$AccountSkus = $AccountSkus | 
    Select-Object SkuPartNumber,ActiveUnits,ConsumedUnits,LockedOutUnits,SuspendedUnits,WarningUnits,@{L='ServicePlans';E={$_.serviceStatus.ServicePlan.ServiceName -join ';'}} |
    sort-object -Descending ActiveUnits

#https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$SKUNamesFileURI = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
try {
    Invoke-WebRequest -Uri $SKUNamesFileURI -OutFile ProductNames.csv
} catch {
    write-log "unable to download Product Names file. " -type warning
    $SKUNamesFileURI = $false
}
$ProductNamesReference = @{}
if($SKUNamesFileURI) {
    $productNames = import-csv ProductNames.csv
    $productNames | Select-Object -Unique "string_ id",product_display_name|%{$ProductNamesReference.add($_."string_ id",$_.product_display_name)}
    $AccountSkus | %{ 
        if( $ProductNamesReference.ContainsKey($_.SkuPartNumber) ) {
            $_.SkuPartNumber = $ProductNamesReference[$_.SkuPartNumber] 
        } else {
            write-log "$($_.SkuPartNumber) not found" -type warning
        }
    }
}

$exportFile = ".\$AccountName-AccountSkus-$(get-date -Format yyMMddHHmm).csv"

$AccountSkus | export-csv -NoTypeInformation $exportFile
write-log "report exported" -type ok
return (get-item $exportFile)
