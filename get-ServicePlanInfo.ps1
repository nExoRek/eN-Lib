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
    .\get-ServicePlanInfo.ps1 -lookupName EOP_ENTERPRISE_PREMIUM 
    
    shows friendly name of EOP_ENTERPRISE_PREMIUM. works for both - Service Plans and License names, may be partial.
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -lookupName 'Business Standard' 
    
    looks up for all licenses containing 'Business Standard' in their name. here - friendly name will match. may be partial.
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -findPlan 'INTUNE'
    
    shows all licenses/products that include any service plan containing 'INTUNE' in the name. you can use either 
    SKU name or Firendly name for plans. may be partial. 
.EXAMPLE
    .\get-ServicePlanInfo.ps1 -lookupName 'business basic'

Product_Display_Name                        String_Id                                   GUID
--------------------                        ---------                                   ----
Microsoft 365 Business Basic                O365_BUSINESS_ESSENTIALS                    3b555118-da6a-4418-894f-7df1e2096870
Microsoft 365 Business Basic                SMB_BUSINESS_ESSENTIALS                     dab7782a-93b1-4074-8bb1-0e61318bea0b
Microsoft 365 Business Basic EEA (no Teams) Microsoft_365_Business_Basic_EEA_(no_Teams) b1f3042b-a390-4b56-ab61-b88e7e767a97
    .\get-ServicePlanInfo.ps1 -productServicePlans O365_BUSINESS_ESSENTIALS

SKU                       Firendly Name                      Service_Plan_Id
---                       -------------                      ---------------
BPOS_S_TODO_1             To-Do (Plan 1)                     5e62787c-c316-451f-b873-1d05acd4d12c
EXCHANGE_S_STANDARD       EXCHANGE ONLINE (PLAN 1)           9aaf7827-d63c-4b61-89c3-182f06f82e5c
FLOW_O365_P1              FLOW FOR OFFICE 365                0f9b09cb-62d1-4ff4-9129-43f4996f83f4
FORMS_PLAN_E1             MICROSOFT FORMS (PLAN E1)          159f4cd6-e380-449f-a816-af1a9ef76344
MCOSTANDARD               SKYPE FOR BUSINESS ONLINE (PLAN 2) 0feaeb32-d00e-4d66-bd5a-43b5b83db82c
OFFICEMOBILE_SUBSCRIPTION OFFICEMOBILE_SUBSCRIPTION          c63d4d19-e8cb-460e-b37c-4d6c34603745
POWERAPPS_O365_P1         POWERAPPS FOR OFFICE 365           92f7a6f3-b89b-4bbd-8c30-809e6da5ad1c
PROJECTWORKMANAGEMENT     MICROSOFT PLANNE                   b737dad2-2f6c-4c65-90e3-ca563267e8b9
SHAREPOINTSTANDARD        SHAREPOINTSTANDARD                 c7699d2e-19aa-44de-8edf-1736da088ca1
SHAREPOINTWAC             OFFICE ONLINE                      e95bec33-7c88-4a70-8e19-b10bd9d0c014
SWAY                      SWAY                               a23b959c-7ce8-4e57-9140-b90eb88a9e97
TEAMS1                    TEAMS1                             57ff2da0-773e-42df-b2af-ffb7a2317929
YAMMER_ENTERPRISE         YAMMER_ENTERPRISE                  7547a3fe-08ee-4ccb-b430-5077c5041653
    
    productServicePlans require an exact name to limit the output, so in the first step lookUp function was used to find
    proper license name, then it was provided for productServicePlans to show all Service Plans included in the license.
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
    #show Service Plans for given license type. product name can be either SKU or friendly name but must be exact
    [Parameter(ParameterSetName='productServicePlans',mandatory=$true,position=0)]
        [string]$productServicePlans    
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

function productServicePlans {
    param([string]$name)
    return (
        $spInfo | 
        Where-Object {$_.Product_Display_Name -eq $name -or $_.String_Id -eq $name} | 
        Select-Object @{L='SKU';E={$_.Service_Plan_Name}},@{L='Firendly Name';E={$_.Service_Plans_Included_Friendly_Names}},Service_Plan_Id
    )
}

$spFile = ".\servicePlans.csv"

if(!(test-path $spFile)) {
    Write-Verbose "file containing plans list not found - downloading..."
    try {
        [System.Uri]$url = "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
        Invoke-WebRequest $url -OutFile $spFile
    } catch {
        Write-Error "cannot download definitions the file."
        return
    }
} 
$spInfo = import-csv $spFile -Delimiter ','

$run = "$($PSCmdlet.ParameterSetName) -name '$([string]$PSBoundParameters.Values)'"
Invoke-Expression $run
