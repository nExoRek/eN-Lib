<#
.SYNOPSIS
    search Azure resource by IP address.
.DESCRIPTION
    searches Azure Resources and Networks using provided IP and Mask in CIDR format.
    you can search for a single resource or all resources containing given IP-as-a-string.
.EXAMPLE
    .\search-AzureByIP.ps1 10.0.0.0

    searches for all reasources containing '10.' as a first octet
.EXAMPLE
    .\search-AzureByIP.ps1 10.10.10.10

    searches for a reasource with IP 10.10.10.10

.LINK
    https://w-files.pl
.LINK
    https://www.powershellgallery.com/packages/OftenOn/1.0.8/Content/Private%5CConvertFrom-CIDR.ps1
.NOTES
    nExoR ::))o-
    version 220609
        last changes
        - 220609 display fix
        - 220608 v1
        - 220601 initialized

    #TO|DO
    - networks and NICs are not all resources with IP - extend
    - although oen may provide a mask it doesn't make sens... azure searches are string-based not comparing IPs
      i don't even think there is any sens in making script more complex.
    - no paging for search - if there will be more than 100 results - they will not be shown
#>
#requires -Modules eNlib,Az.ResourceGraph
param(
    [string]$IP,
    #search for all devices within provided net
    [switch]$all
)
class AzResourceGraphException : Exception {
    [string] $additionalData

    AzResourceGraphException($Message, $additionalData) : base($Message) {
        $this.additionalData = $additionalData
    }
}

function ConvertFrom-CIDR {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param (
        [Parameter(Mandatory)]
            [string] $IPAddress
    )

    #validate IP
    [regex]$rxIP = "^(?<IP>(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?))(?:/(?<mask>[0-9]{1,2}))?$"
    if(-not ($IPAddress -match $rxIP)) {
        write-log "it's not a valid IP address" -type error
        exit
    }
    $ip = [IPAddress] $Matches['IP']
    if($null -eq $Matches['mask']) {
        $suffix = 24 #default if not provided
    } else {
        $suffix = [int] $Matches['mask']
    }
    $mask = ("1" * $suffix) + ("0" * (32 - $suffix))
    $mask = [IPAddress] ([Convert]::ToUInt64($mask, 2))

    @{
        IPAddress  = $ip.ToString()
        CIDR       = $IPAddress
        CIDRSuffix = $suffix
        NetworkID  = ([IPAddress] ($ip.Address -band $mask.Address)).IPAddressToString
        SubnetMask = $mask.IPAddressToString
    }
}
#IP without zero's  -> 10.34.0.0 will return "10.34" part - searches are text based so all reasources with IP starting with '10.34' will be searched
[regex]$rxSubNet = "^(.*?)(?:\.0)*$"
[regex]$rxMask = "^(?:.*?)/([\d]{1,2})$"
$fullIP = ConvertFrom-CIDR $IP
$IP = $fullIP.IPAddress
$netIP = $fullIP.NetworkID
$mask = $fullIP.SubnetMask
write-log "using mask $mask" -type info

if($IP -eq $netIP) {
    write-log "$IP is a network address. searching all objects within this network." -type info
    [string]$IP = $rxSubNet.Match($netIP).Groups[1].value
} 
if($all.IsPresent) {
    [string]$IP = $rxSubNet.Match($netIP).Groups[1].value
    write-log "'all' flag - will search all devices within $netIP"
}

write-log "searching networks containing $netIP..."
try {
    $result = Search-AzGraph -Query "resources 
        | where type =~ 'microsoft.network/virtualNetworks' and properties.addressSpace.addressPrefixes contains '$netIP'
        | join kind=inner (resourceContainers 
            | where type =~ 'microsoft.resources/subscriptions' 
            | project subscriptionId,subscriptionName=name) on subscriptionId
        | project type='NETWORK',subscriptionName,resourceGroup,name,addressSpace = properties.addressSpace.addressPrefixes
    " -ErrorAction SilentlyContinue -ErrorVariable $graphError 
    if ($null -ne $graphError) {
        $errorJSON = $graphError.ErrorDetails.Message | ConvertFrom-Json
        throw [AzResourceGraphException]::new($errorJSON.error.details.code, $errorJSON.error.details.message)
    }
} catch [AzResourceGraphException] {
    Write-Log "An error on KQL query" -type error
    Write-Log $_.Exception.message
    Write-Log $_.Exception.additionalData
} catch {
    Write-Log "An error occurred in the script" -type error
    Write-Log $_.Exception.message
}
if($result.data.count -eq 0) {
    write-log "no network objects found matching $netIP..." -type warning
} else {
    write-log $result.GetEnumerator() -noTimeStamp
    $trueNetMask = $rxMask.Match($result.addressSpace).Groups[1].value
    if($trueNetMask -ne $fullIP.CIDRSuffix) {
        write-log "network mask read from the object ($trueNetMask) differs from provided ($($fullIP.CIDRSuffix))" -type warning
    } 
}

write-log "searching resource with IP $IP..."
try {
    $result = Search-AzGraph -Query "resources 
        | where type =~ 'Microsoft.Network/networkInterfaces' and properties.ipConfigurations[0].properties.privateIPAddress contains '$IP' 
        | extend sName = tostring(properties.ipConfigurations[0].properties.subnet.id) 
        | extend type = iff(isnull(properties.virtualMachine),properties.ipConfigurations[0].name,'virtualMachine')
        | join kind=inner (resourceContainers 
            | where type =~ 'microsoft.resources/subscriptions' 
            | project subscriptionId,subscriptionName=name) on subscriptionId 
        | project type,subscriptionName, resourceGroup,vNet = extract('virtualNetworks/(.+?)/',1,sName),subnetName = extract('subnets/(.+?)$',1,sName),name,privateIp = properties.ipConfigurations[0].properties.privateIPAddress
    " -ErrorAction SilentlyContinue -ErrorVariable $graphError
    if ($null -ne $graphError) {
        $errorJSON = $graphError.ErrorDetails.Message | ConvertFrom-Json
        throw [AzResourceGraphException]::new($errorJSON.error.details.code, $errorJSON.error.details.message)
    }
} catch [AzResourceGraphException] {
    Write-Log "An error on KQL query" -type error
    Write-Log $_.Exception.message
    Write-Log $_.Exception.additionalData
} catch {
    Write-Log "An error occurred in the script" -type error
    Write-Log $_.Exception.message
}
if($result.data.count -eq 0) {
    write-log "no resources found matching $IP" -type warning
} else {
    write-log $result.GetEnumerator() -noTimeStamp
}

write-log 'done' -type ok

