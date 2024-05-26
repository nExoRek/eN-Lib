<#
.SYNOPSIS
    merge AD and Entra ID reports into single user activity report
.DESCRIPTION
    using outputs from search-eNInactiveADObjects.ps1 and search-eNInactiveEntraUsers.ps1
.EXAMPLE
    .\report-eNInactiveHybridUsers.ps1

    
.INPUTS
    two CSV - from AD and Entra ID
.OUTPUTS
    merged report
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 240520
        last changes
        - 240520 initialized

    #TO|DO
    * edge scenarios - eg. the same UPN on both sides, but account is not hybrid; maybe some other i did not expect?
#>
[CmdletBinding()]
param (
    #output file from AD
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSVAD,
    #output file from Entra ID
    [Parameter(mandatory=$true,position=1)]
        [string]$inputCSVEntraID,
    #key attribute to match the users, default userPrincipalName
    [Parameter(mandatory=$false,position=2)]
        [validateSet('userPrincipalName','mail')]
        [string]$matchBy = 'userPrincipalName'
    
)
$VerbosePreference = 'Continue'
$exportCSVFile = "mergedUsers-{0}.csv" -f (get-date -Format "yyMMdd-hhmm")

Write-Verbose "loading CSV files.."
$ADData = load-CSV $inputCSVAD -header @('samaccountname','userPrincipalName','enabled','givenName','surname','displayName','mail','description','daysInactive') -headerIsCritical
$EntraIDData = load-CSV $inputCSVEntraID -header @('id','displayname','givenname','surname','accountenabled','userprincipalname','mail','userType','Hybrid','givenname','surname','userprincipalname','userType','mail','daysInactive') -headerIsCritical

#make a copy from Entra list - this will be used as a 'metaverse', with Entra ID values already filled in
$aggregatedUserInfo = $EntraIDData.psobject.Copy()
#extend object with AD attributes with 'AD_' prefix
foreach($propertyName in ( ($ADData[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
    if($propertyName -ne $matchBy) {
        $aggregatedUserInfo | Add-Member -MemberType NoteProperty -Name "AD_$propertyName" -Value ''
    }
}
#prepare template - class replacement
[psCustomObject]$recordTemplate =@{}
foreach($propertyName in ( ($aggregatedUserInfo[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
    $recordTemplate | Add-Member -MemberType NoteProperty -Name $propertyName -Value ''
} 

$ADData = $ADData | Select-Object *,'used' #'used' will be a flag to show that there was a match with Entra ID. 

Write-Verbose "matching EntraID with AD objects..."
foreach($entraID in $aggregatedUserInfo) {
    #to have all cloud native object set as forever inactive in AD
    $entraID."AD_daysInactive" = 23000 
    $ADData | ? {$_.$matchBy -eq $entraID.$matchBy} | % {
        $_.used = $true #set as already matched to lated filter based on this attribute
        #rewrite all attribute values from AD object to metaverse object, using 'AD_' prefix for attributes
        #but skip maching attribute since it's the same for both
        foreach($propertyName in ( ($_.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
            if($propertyName -ne $matchBy -and $propertyName -ne 'used') {
                $entraID."AD_$propertyName" = $_.$propertyName
            }
        } 
    }
}

Write-Verbose "list all non-matched values from AD list and add them to the final list..."
foreach($adObj in ($ADData|? used -ne $true) ) {
    $tmp = $recordTemplate.psobject.copy()
    foreach($propertyName in ( ($adObj.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
        if($propertyName -eq 'used') { continue }
        if($propertyName -eq $matchBy) { 
            $tmp.$propertyName = $adObj.$propertyName
        } else {
            $tmp."AD_$propertyName" = $adObj.$propertyName
        }
        #to help filtering in Excel by daysInactive, set as 10000 - never active in Entra ID
        $tmp.daysInactive = 23000
    } 

    $aggregatedUserInfo += $tmp
}
#export all results, extending with Hybrid_daysInactive attribute being lower of the comparison between EID and AD
$aggregatedUserInfo | 
    Select-Object *,@{L='Hybrid_daysInactive';E={($_.daysInactive,$_.AD_daysInactive|Measure-Object -Minimum).minimum}} |
    Export-Csv -Encoding unicode -NoTypeInformation $exportCSVFile
Write-Verbose "merged report saved to '$exportCSVFile'."
