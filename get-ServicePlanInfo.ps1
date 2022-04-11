<#
.SYNOPSIS
    display information on Service Plans
.DESCRIPTION
    constant problems I encounter with licenses (called 'products') are:
      - does this or that license contain some service plan?
      - what is given SKU - since I have technical output and interace shows different name?
      - which service plans are included in given licence?
    this script addresses exactly these question. it downloads current SKU name listing from Microsoft doc and
    lookup the names.
    this is very simple script - if you want to refresh SKU names (e.g. new CSV appeard on the docs) simply remove 
    'servicePlans.csv' file.
.EXAMPLE
    .\get-ServicePlanInfo.ps1
    
    shows all licenses/products and service plans
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -resolveName EOP_ENTERPRISE_PREMIUM 
    
    shows friednly name of EOP_ENTERPRISE_PREMIUM. works for both - Service Plans and License names
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -showProducts 'INTUNE_O365'
    
    shows all licenses/products that include INTUNE_O365 service plan. you can use either SKU name or Firendly name
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -productName 'Microsoft 365 F3'
    
    shows all service plans included in 'Microsoft 365 F3' license. you can use either code name or friendly name.
.LINK
    https://w-files.pl
.LINK
    https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
.NOTES
    nExoR ::))o-
    version 220331
        last changes
        - 210331 MS fixed CSV ... changing column name /:
        - 220315 initialized

    #TO|DO
    - header check
#>
[CmdletBinding(DefaultParameterSetName='default')]
param (
    #resolve Service Plan or License name 
    [Parameter(ParameterSetName='resolveNames',mandatory=$true,position=0)]
        [string]$resolveName,
    #display all licenses containing given SP
    [Parameter(ParameterSetName='listLicenses',mandatory=$true,position=0)]
        [string]$showProducts,
    #show SPs for given license type
    [Parameter(ParameterSetName='listServicePlans',mandatory=$true,position=0)]
        [string]$productName    
)

function resolveNames {
    param([string]$name)

    $ServicePlan = $spInfo | Where-Object { $_.psobject.Properties.value -contains $name }
    if($ServicePlan) {
        if($ServicePlan -is [array]) { $ServicePlan = $ServicePlan[0] }
        $property = ($ServicePlan.psobject.Properties| Where-Object value -eq $name).name
        switch($property) {
            'Service_Plan_Name' {
                return $ServicePlan.'Service_Plans_Included_Friendly_Names'
            }
            'Service_Plans_Included_Friendly_Names' {
                return $ServicePlan.'Service_Plan_Name'
            }
            'Product_Display_Name' {
                return $ServicePlan.'String_Id'
            }
            'String_Id' {
                return $ServicePlan.'Product_Display_Name'
            }
            default { return $null }
        }
    } else {
        return $null
    }
}

function listLicenses {
    param([string]$name)
    return ($spInfo | Where-Object {$_.Service_Plan_Name -eq $name -or $_.Service_Plans_Included_Friendly_Names -eq $name} | Select-Object @{L="products containing $name";E={$_.Product_Display_Name}})

}

function listServicePlans {
    param([string]$name)
    return (
        $spInfo | 
        Where-Object {$_.Product_Display_Name -eq $name -or $_.'String_Id' -eq $name} | 
        Select-Object @{L='SKU';E={$_.Service_Plan_Name}},@{L='Firendly Name';E={$_.Service_Plans_Included_Friendly_Names}}
    )
}

function default {
    return $spInfo
}

$spFile = ".\servicePlans.csv"

if(!(test-path $spFile)) {
    [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    Invoke-WebRequest $url -OutFile serviceplans.csv
} 
$spInfo = import-csv $spFile -Delimiter ','

$run = "$($PSCmdlet.ParameterSetName) -name '$([string]$PSBoundParameters.Values)'"
Invoke-Expression $run
