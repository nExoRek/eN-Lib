<#
.SYNOPSIS
    DRAFT - the whole logic is on a very initial state with hardcoded assumptions.

    merge AD, Entra ID and Exchange reports into single user activity report

.DESCRIPTION
    using outputs from get-eNReportADObjects.ps1 (AD), get-eNReportEntraUsers.ps1 (EntraID) and get-eNReportMailboxes.ps1 (EXO)
    to combine them into a single view on the accounts, mailboxes and licenses
.EXAMPLE
    .\join-eNReportHybridUsersInfo.ps1

    
.INPUTS
    CSV report from other scripts
.OUTPUTS
    merged report
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 240718
        last changes
        - 240718 initiated as a bigger project, extended with Exchange checking
        - 240627 add displayname as matching attribute. forceHybrid is for now default and parameter doesn't do anything
        - 240520 initialized

    #TO|DO
    * edge scenarios - eg. the same UPN on both sides, but account is not hybrid; maybe some other i did not expect?
    * change hybrid user detection
    * allow multi-value checks for forced hybrid merge
    * currently matching is ONLY in forced hybrid... which should not be the case
    * change time values to represent the same 'never' value
#>
#requires -module eNLib
[CmdletBinding()]
param (
    #output file from AD
    [Parameter(mandatory=$false,position=0)]
        [string]$inputCSVAD,
    #output file from Entra ID
    [Parameter(mandatory=$false,position=1)]
        [string]$inputCSVEntraID,
    #output file from Exchange Online 
    [Parameter(mandatory=$false,position=3)]
        [string]$inputCSVEXO,
    #force match for non-hybrid users - low accuracy... key attribute to match the users, default userPrincipalName
    [Parameter(mandatory=$false,position=2)]
        [validateSet('userPrincipalName','mail','displayName','all')]
        [string]$matchBy = 'all'
)
#$VerbosePreference = 'Continue'
$exportCSVFile = "HybridUserReport-{0}.csv" -f (get-date -Format "yyMMdd-hhmm")

Write-log "loading CSV files.." -type info
$reports = 0
if($inputCSVAD) {
    $ADData = load-CSV $inputCSVAD -header @('samaccountname','userPrincipalName','enabled','givenName','surname','displayName','mail','description','daysInactive') -headerIsCritical
    $reports++
}
if($inputCSVEntraID) {
    $EntraIDData = load-CSV $inputCSVEntraID -header @('id','displayname','givenname','surname','accountenabled','userprincipalname','mail','userType','Hybrid','givenname','surname','userprincipalname','userType','mail','daysInactive') -headerIsCritical
    $reports++
}
if($inputCSVEXO) {
    $EXOData = load-CSV $inputCSVEXO -header @('RecipientType','RecipientTypeDetails','emails','WhenMailboxCreated','LastInteractionTime','LastUserActionTime','TotalItemSize','ExchangeObjectId') -headerIsCritical
    $reports++
}
if($reports -lt 2) {
    Write-Log "at least two reports are required for merge" -type error
    return
}

#TODO - report should always have all the fields - metafile should be a static schema, and fields populated or not
#make a copy from Entra list - this will be used as a 'metaverse', with Entra ID values already filled in
$aggregatedUserInfo = $EntraIDData.psobject.Copy()
#extend object with AD attributes with 'AD_' prefix
foreach($propertyName in ( ($ADData[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
    if($propertyName -ne $matchBy) {
        $aggregatedUserInfo | Add-Member -MemberType NoteProperty -Name "AD_$propertyName" -Value ''
    }
}
#extend object with EXO attributes with 'EXO_' prefix
foreach($propertyName in ( ($EXOData[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
    if($propertyName -notmatch 'userPrincipalName|Identity|DisplayName|FirstName|LastName|enabled') { #skip dupes
        $aggregatedUserInfo | Add-Member -MemberType NoteProperty -Name "EXO_$propertyName" -Value ''
    }
}
#prepare template - class replacement
[psCustomObject]$recordTemplate =@{}
foreach($propertyName in ( ($aggregatedUserInfo[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
    $recordTemplate | Add-Member -MemberType NoteProperty -Name $propertyName -Value ''
} 

$ADData = $ADData | Select-Object *,'used' #'used' will be a flag for forced hybrid match to show that there was a match with Entra ID

Write-Verbose "matching EntraID with AD objects..."
foreach($entraID in $aggregatedUserInfo) {
    #to have all cloud native object set as forever inactive in AD
    $entraID."AD_daysInactive" = 23000 
    $ADData | ? {$_.userPrincipalName -eq $entraID.userPrincipalName -or $_.displayName -eq $entraID.displayName -or $_.mail -eq $entraID.mail} | % { #TODO - may have duplicates.
        $_.used = $true #set as already matched to later filter based on this attribute
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

#EXO
foreach($recipient in $EXOData) {
    if($recipient.userPrincipalName) { #only mailboxes have UPNs
        $aggregatedUserInfo | ? userPrincipalName -eq $recipient.userPrincipalName | %{ #locate entry by UPN 
            foreach($propertyName in ( ($recipient.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                if($propertyName -notmatch 'userPrincipalName|Identity|DisplayName|FirstName|LastName|enabled') { #skip dupes
                    $_."EXO_$propertyName" = $recipient.$propertyName
                }
            } 
        } 
    }
}

#export all results, extending with Hybrid_daysInactive attribute being lower of the comparison between EID and AD
$aggregatedUserInfo | 
    Select-Object *,@{L='Hybrid_daysInactive';E={($_.daysInactive,$_.AD_daysInactive|Measure-Object -Minimum).minimum}} |
    Export-Csv -Encoding unicode -NoTypeInformation $exportCSVFile
Write-Log "merged report saved to '$exportCSVFile'." -type ok
write-log "converting..."
&(convert-CSV2XLS $exportCSVFile)
write-log "done." -type ok
