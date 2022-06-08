<#
.SYNOPSIS
    search resource by providing it's name only.
.DESCRIPTION
    support script allowing to quickly locate resource and some basic informaiton on it.
.EXAMPLE
    .\search-AzureByName.ps1 somename

    searches for all resources and resource-container with the name 'somename'
.EXAMPLE
    .\search-AzureByName.ps1 somename -partial

    searches for all resources and resource-container containing 'somename' in the name
.EXAMPLE
    .\search-AzureByName.ps1 somename -partial|? type -match 'storage' 

    searches for all resources and resource-container containing 'somename' in the name and is a type of 'storageaccount'
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.LINK
    https://dev.to/omiossec/get-started-with-azure-resource-graph-and-powershell-3pmo 
.NOTES
    nExoR ::))o-
    version 220608
        last changes
        - 220608 v1
        - 220602 initialized

    #TO|DO
#>
#requires -modules eNlib,Az.ResourceGraph -Version 7
[CmdletBinding()]
param (
    #resource name
    [Parameter(mandatory=$true,position=0)]
        [string]$name,
    #use to search substring of the name 
    [Parameter(mandatory=$false,position=1)]
        [switch]$partial

)
class AzResourceGraphException : Exception {
    [string] $additionalData

    AzResourceGraphException($Message, $additionalData) : base($Message) {
        $this.additionalData = $additionalData
    }
}

$compare = $partial ? 'contains' : '=~'

try {
    (Search-AzGraph -Query "rsourceContainers 
        | where name $compare '$name' | join kind = inner (resourceContainers 
            | where type =~ 'microsoft.resources/subscriptions' 
            | project subscriptionId,subscriptionName=name) on subscriptionId
        | project type, name, subscriptionName, location, id
    " -ErrorVariable $graphError -ErrorAction SilentlyContinue).data
    if ($null -ne $graphError) {
        $errorJSON = $graphError.ErrorDetails.Message | ConvertFrom-Json
        throw [AzResourceGraphException]::new($errorJSON.error.details.code, $errorJSON.error.details.message)
    }
} catch [AzResourceGraphException] {
    Write-Log "An error on KQL query" -type error
    Write-Log $_.Exception.message
    Write-Log $_.Exception.additionalData
}
catch {
    Write-Log "An error occurred in the script" -type error
    Write-Log $_.Exception.message
}

try{
    (Search-AzGraph "resources 
        | where name $compare '$name' 
        | join kind = inner (resourceContainers 
            | where type =~ 'microsoft.resources/subscriptions' 
            | project subscriptionId,subscriptionName=name) on subscriptionId
        | project type,name,subscriptionName,resourceGroup,location
    ").data
    if ($null -ne $graphError) {
        $errorJSON = $grapherror.ErrorDetails.Message | ConvertFrom-Json
        throw [AzResourceGraphException]::new($errorJSON.error.details.code, $errorJSON.error.details.message)
    }
} catch [AzResourceGraphException] {
    Write-Log "An error on KQL query" -type error
    Write-Log $_.Exception.message
    Write-Log $_.Exception.additionalData
}
    catch {
    Write-Log "An error occurred in the script" -type error
    Write-Log $_.Exception.message
}
