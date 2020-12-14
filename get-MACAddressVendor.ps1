<#
.SYNOPSIS
   simple script getting NIC vendor by checking MAC address OUI table. 
   use -webAPI to use remote query. otherwise out.txt file will be downloaded locally.
.EXAMPLE
    .\get-MACAddressVendor.ps1 00-FF-E2

    lookup for vendor of 00:FF:E2 MAC prefix.
.EXAMPLE
    Get-NetAdapter|select -ExpandProperty macaddress|.\get-MACAddressVendor.ps1

    check vendors for all local network card.
.LINK
    https://w-files.pl
.LINK
    oui file: http://standards-oui.ieee.org/oui.txt
.LINK
    remote query API: http://www.macvendorlookup.com/api

.NOTES
    nExoR 2o16
    ver. 201214
        - 201214 lift and shift. webAPI still unavailable.
#>
[cmdletbinding()]
param(
    #MAC address - may be formed with 'AA:BB', 'AA-BB' or concete form 'AABB' 
    [parameter(position=0,mandatory=$true,valueFromPipeline=$true)]
        [string]$macAddress,
    #use remote webAPI and do not download&cache oui file locally - by default oui.txt is downloaded locally
    [parameter(position=1)]
        [switch]$webAPI,
    #show additional lines from oui file - mostly they contain company address
    [parameter(position=2)]
        [switch]$showCompanyInfo
) 
begin {}
process {
    $macAddressToVerify=$macAddress.Replace(':','').Replace('-','').ToLower()
    $macAddressToVerify=$macAddressToVerify.Substring(0,6)
    if($macAddressToVerify -notmatch [regex]'[0-9a-f]{6}') {
        throw 'NOT VALID MAC ADDRESS'   
    }
    Write-Verbose "lookup for $macAddress..."

    if($webAPI.IsPresent) {
        Write-Verbose 'getting info from macvendrolookup...'
        $result=Invoke-WebRequest "http://www.macvendorlookup.com/api/v2/$macAddressToVerify"
        $result=ConvertFrom-Json $result.Content   
    } else {
        if(-not (Test-Path .\oui.txt)) {
            Write-Verbose 'oui file not found. downloading from IEEE.org...'
            Start-BitsTransfer -Source 'http://standards-oui.ieee.org/oui.txt' -Destination 'oui.txt' -TransferType Download
        } else {
            Write-Verbose "using local copy of oui.txt. delete the file to force download again."
        }

        if($showCompanyInfo.IsPresent) {
            $result=Select-String -path .\oui.txt -Pattern $macAddressToVerify -Context 0,4
        } else {
            $result=Select-String -path .\oui.txt -Pattern $macAddressToVerify 
            [regex]$rxVendor='\(base 16\)\s+(?<vendor>.*)'
            $result -match $rxVendor|Out-Null
            if([string]::isNullOrEmpty($Matches['vendor']) ) {
                $result="$macAddressToVerify;vendor not found"
            } else {
                $result="$macAddressToVerify;$($Matches['vendor'])"
            }
        }
    }

    $result
}
end{}
