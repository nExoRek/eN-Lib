<#
.SYNOPSIS
    simple script using 'whatismyipaddress.com' to query for external IP. 
.DESCRIPTION
    regular rundoes not use any specific API so potentially vulnerable for a page changes.
    pureIP parameter is using API ipv4bot.whatismyipaddress.com and returns IP only which may 
    be usefull when you need to pipeline the values
.EXAMPLE
    .\show-eNLibMyExternalIP.ps1
    connect to whatismyipaddress.com and query for external IP
.EXAMPLE
    .\show-eNLibMyExternalIP.ps1 -extended
    connect to whatismyipaddress.com and query for external IP then query again for whois information 
.NOTES
    2o2o.o9.3o ::))o- 
#>
[cmdletbinding()]
param(
    [parameter(mandatory=$false,position=0,ParameterSetName='extended')]
        [switch]$extended,
        #show extended information 
    [parameter(mandatory=$false,position=0)]
        [switch]$pureIP
        #output only IP number for pipelining
)

if($pureIP) {
    $result=Invoke-WebRequest ipv4bot.whatismyipaddress.com -UserAgent Chrome
    return $result.content
}

try {
    $page=Invoke-WebRequest http://whatismyipaddress.com/ -UserAgent Chrome -ContentType "text/plain; charset=1252"
} catch {
    $_
    exit -1
}

[regex]$rxIP="\d{1,3}[.]\d{1,3}[.]\d{1,3}[.]\d{1,3}"
[regex]$rxISP="ISP:.*"
[regex]$rxCity="City:.*"
[regex]$rxRegion="Region:.*"
[regex]$rxCountry="Country:.*"
[regex]$rxExtIP="IP:</th><td>(?<extIP>[\d.]+)</td>"
[regex]$rxExtDecimal="Decimal:</th><td>(?<extDecimal>[\d]+)</td>"
[regex]$rxExtHostname="Hostname:</th><td>(?<extHostname>[\w\d._-]+)</td>"
[regex]$rxExtASN="ASN:</th><td>(?<extASN>[\d.]+)</td>"
[regex]$rxExtISP="ISP:</th><td>(?<extISP>.*)</td>"
[regex]$rxExtOrganization="Organization:</th><td>(?<extOrganization>.*)</td>"
[regex]$rxExtServices="Services:</th><td>(?<extServices>.*)</td>"
[regex]$rxExtAssignment="Assignment:</th><td><a href=\'/dynamic-static\'>(?<extAssignment>.*)</a></td>"
[regex]$rxExtContinent="Continent:</th><td>(?<extContinent>.*)</td>"
[regex]$rxExtCountry="Country:</th><td>(?<extCountry>.*)<img src"
[regex]$rxExtCity="City:</th><td>(?<extCity>.*)</td>"
[regex]$rxExtPostal="Postal Code:</th><td>(?<extPostal>.*)</td>"

$page.AllElements|? id -eq 'section_left'|%{ 
    $IP=$rxIP.Match($_.outerText)
    $ISP=$rxISP.Match($_.outerText)
    $City=$rxCity.Match($_.outerText)
    $Region=$rxRegion.Match($_.outerText)
    $Country=$rxCountry.Match($_.outerText)
}

if ($extended) { 
    try {
        $pageExtended=Invoke-WebRequest "http://whatismyipaddress.com/ip/$($IP.value)" -UserAgent Chrome -ContentType "text/plain; charset=utf-8"
    } catch {
        $_
        exit -1
    }
    write-host "IP:             $($rxExtIP.Match($pageExtended).groups['extIP'].value)"
    write-host "Decimal:        $($rxExtDecimal.Match($pageExtended).groups['extDecimal'].value)"
    write-host "Hostname:       $($rxExtHostname.Match($pageExtended).groups['extHostname'].value)"
    write-host "ASN:            $($rxExtASN.Match($pageExtended).groups['extASN'].value)"
    write-host "ISP:            $($rxExtISP.Match($pageExtended).groups['extISP'].value)"
    write-host "Organization:   $($rxExtOrganization.Match($pageExtended).groups['extOrganization'].value)"
    write-host "Services:       $($rxExtServices.Match($pageExtended).groups['extServices'].value)"
    write-host "Assginment:     $($rxExtAssignment.Match($pageExtended).groups['extAssignment'].value)"
    write-host "Continent:      $($rxExtContinent.Match($pageExtended).groups['extContinent'].value)"
    write-host "Country:        $($rxExtCountry.Match($pageExtended).groups['extCountry'].value)"
    write-host "City:           $($rxExtCity.Match($pageExtended).groups['extCity'].value)"
    write-host "Postal Code:    $($rxExtPostal.Match($pageExtended).groups['extPostal'].value)"
    write-host -ForegroundColor Green "done."
} else {
    write-host -ForegroundColor RED "External IP: $IP"
    write-host $ISP
    write-host $City
    write-host $Region
    write-host $Country
    write-host -ForegroundColor Green "done."
}
