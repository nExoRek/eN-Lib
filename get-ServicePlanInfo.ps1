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
    .\get-ServicePlanInfo.ps1 -lookupName EOP_ENTERPRISE_PREMIUM 
    
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
    version 250205
        last changes
        - 250205 overhaul - parameter names fixed and values more intiutive
        - 220331 MS fixed CSV ... changing column name /:
        - 220315 initialized

    #TO|DO
    - header check
#>
[CmdletBinding(DefaultParameterSetName='lookupName')]
param (
    #lookup the name (internal, displayname) and shows details. you can use partial name of the license
    [Parameter(ParameterSetName='lookupName',mandatory=$true,position=0)]
        [string]$lookupName,
    #display all licenses (boundles) containing given Service Plan - to quickly find boundles with searched feature
    [Parameter(ParameterSetName='findPlan',mandatory=$true,position=0)]
        [string]$findPlan,
    #show SPs for given license type
    [Parameter(ParameterSetName='listServicePlans',mandatory=$true,position=0)]
        [string]$productName    
)

function lookupName {
    param([string]$name)

    $ServicePlan = $spInfo | Where-Object { $_.Product_Display_Name -match $name -or $_.Service_Plan_Name -match $name -or $_.String_Id -match $name }
    if($ServicePlan) {
        return $ServicePlan|Select-Object Product_Display_Name,String_Id,GUID -Unique
    } else {
        return $null
    }
}

function findPlan {
    param([string]$name)
    return (
        $spInfo | 
        Where-Object {$_.Service_Plans_Included_Friendly_Names -match $name -or $_.Service_Plan_Name -match $name} |
        select-object  Service_Plan_Name,Service_Plans_Included_Friendly_Names,Service_Plan_Id,Product_Display_Name,String_Id,GUID -unique | 
        Sort-object Service_Plan_Name
    )

}

function listServicePlans {
    param([string]$name)
    return (
        $spInfo | 
        Where-Object {$_.Product_Display_Name -eq $name -or $_.String_Id -eq $name} | 
        Select-Object @{L='SKU';E={$_.Service_Plan_Name}},@{L='Firendly Name';E={$_.Service_Plans_Included_Friendly_Names}},Service_Plan_Id
    )
}

function default {
    return $spInfo
}

$spFile = ".\servicePlans.csv"

if(!(test-path $spFile)) {
    Write-Verbose "file containing plans list not found - downloading..."
    [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
    Invoke-WebRequest $url -OutFile $spFile
} 
$spInfo = import-csv $spFile -Delimiter ','

$run = "$($PSCmdlet.ParameterSetName) -name '$([string]$PSBoundParameters.Values)'"
Invoke-Expression $run
