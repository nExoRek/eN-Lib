<#
.SYNOPSIS
    search Azure resource by IP address.
.DESCRIPTION
    searches Azure Resources and Networks using provided IP and Mask in CIDR format.
.EXAMPLE
    .\search-AzureByIP.ps1

.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 220601
        last changes
        - 220601 initialized

    #TO|DO
    - resolve automatically full CIDR IP - change regex
#>

param(
    [string]$IP,
    [validateRange(8,32)]
    [int]$mask=24
)

function convert-CIDRMaskToIP {
    param([int]$CIDRmask)

    [IPAddress] $MASK = 0;
    $MASK.Address = ([UInt32]::MaxValue) -shl (32 - $CIDRmask) -shr (32 - $CIDRmask)
    return $MASK
}
#validate IP
$rxIP = "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
if($IP -notmatch $rxIP) {
    write-log "it's not a valid IP address" -type error
    exit
}

$lookForResource = $true
[ipaddress]$IP = $IP
$netMask = convert-CIDRMaskToIP -CIDRmask $mask
write-log "using mask $($netMask.IPAddressToString)" -type info
$netIP = [ipaddress]($IP.Address -band $netMask.Address)
if($IP -eq $netIP) {
    write-log "$($IP.IPAddressToString) is a network address. checking networks only" -type info
    $lookForResource = $false
}

write-log "searching networks containing $IP..."
Search-AzGraph -Query "resources 
    | where type =~ 'microsoft.network/virtualNetworks' and properties.addressSpace.addressPrefixes contains '$IP'
    | join kind=inner (resourceContainers 
        | where type =~ 'microsoft.resources/subscriptions' 
        | project subscriptionId,subscriptionName=name) on subscriptionId
    | project subscriptionName,resourceGroup,name,addressSpace = properties.addressSpace.addressPrefixes
" | Format-List

if($lookForResource) {
    write-log "searching resources..."
    Search-AzGraph -Query "resources 
        | where type =~ 'Microsoft.Network/networkInterfaces' and properties.ipConfigurations[0].properties.privateIPAddress contains '$IP' 
        | extend sName = tostring(properties.ipConfigurations[0].properties.subnet.id) 
        | extend type = iff(isnull(properties.virtualMachine),properties.ipConfigurations[0].name,'virtualMachine')
        | join kind=inner (resourceContainers 
            | where type =~ 'microsoft.resources/subscriptions' 
            | project subscriptionId,subscriptionName=name) on subscriptionId 
        | project subscriptionName, resourceGroup,vNet = extract('virtualNetworks/(.+?)/',1,sName),subnetName = extract('subnets/(.+?)$',1,sName),name,privateIp = properties.ipConfigurations[0].properties.privateIPAddress,type
    " | Format-List
}
write-log 'done' -type ok

