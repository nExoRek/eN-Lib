<#
.SYNOPSIS
   simple script getting NIC vendor by checking MAC address OUI table. 
   use -webAPI to use remote query. otherwise out.txt file will be downloaded locally.
.LINK
   oui file: http://standards-oui.ieee.org/oui.txt
   remote query API: http://www.macvendorlookup.com/api
.NOTES
   nExoR 2o16
#>
[cmdletbinding(DefaultParameterSetName='mini')]
param(
    [parameter(position=0,mandatory=$true,valueFromPipeline=$true,ParameterSetName='full')]
    [parameter(position=0,mandatory=$true,valueFromPipeline=$true,ParameterSetName='webAPI')]
        [string]$macAddress,
        #MAC address - may be formed with 'AA:BB', 'AA-BB' or concete form 'AABB' 
    [parameter(ParameterSetName='webAPI',position=1)]
        [switch]$webAPI,
        #use remote webAPI and do not download&cache oui file locally - by default oui.txt is downloaded locally
    [parameter(ParameterSetName='full',position=1)]
        [switch]$showCompanyInfo
        #show additional lines from oui file - mostly they contain company address
) 

$macAddressToVerify=$macAddress.Replace(':','').Replace('-','').ToLower()
$macAddressToVerify=$macAddressToVerify.Substring(0,6)
if($macAddressToVerify -notmatch [regex]'[0-9a-f]{6}') {
    throw 'NOT VALID MAC ADDRESS'   
}
Write-Verbose "lookup for $macAddress..."

if($webAPI) {
    Write-Verbose 'getting info from macvendrolookup...'
    $result=Invoke-WebRequest "http://www.macvendorlookup.com/api/v2/$macAddressToVerify"
    $result=ConvertFrom-Json $result.Content   
} else {
    if(-not (Test-Path .\oui.txt)) {
        Write-Verbose 'oui file not found. downloading from IEEE.org...'
        Start-BitsTransfer -Source 'http://standards-oui.ieee.org/oui.txt' -Destination 'oui.txt' -TransferType Download
    }

    if($showCompanyInfo) {
        $result=Select-String -path .\oui.txt -Pattern $macAddressToVerify -Context 0,4
    } else {
        $result=Select-String -path .\oui.txt -Pattern $macAddressToVerify 
        [regex]$rxVendor='\(base 16\)\s+(?<vendor>.*)'
        $result -match $rxVendor|Out-Null
        $result="$macAddress;$($Matches['vendor'])"
    }
}

if($result) {
    $result
} else {
    throw "Vendor not found for $macAddress"
}

