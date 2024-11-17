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
    * BUILD SCHEMA - currently script is totally screwed and it assumed that all files are present.
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
        [string]$matchBy = 'userPrincipalName'
)

# Function to update information from different data sources
function Update-MetaverseData {
    param (
        #metaverse object to work on
        [Parameter(Mandatory,Position = 0)]
            [hashtable]$mv,
        #key object ID, 
        [Parameter(Mandatory,Position = 1)]
            [int]$objectID,
        #object with new values
        [Parameter(Mandatory,Position = 2)]
            [PSobject]$dataSource
    )

    if(-not $mv.ContainsKey($objectID)) {
        # If the objectID with a given ID does not exist in the metaverse - thow an error
        throw -1
    }

    # Merge attributes for the specified person
    foreach ($propertyName in ( ($dataSource.psobject.Properties | ? memberType -eq 'NoteProperty')).name) {
        $mv[$objectID][$propertyName] = $dataSource.$propertyName
    }
    Write-Verbose "metaverse object $objectID has been updated"
}

function Add-MetaverseData {
    param (
        #metaverse object to work on
        [Parameter(Mandatory,Position = 0)]
            [hashtable]$mv,
        #object with new values
        [Parameter(Mandatory,Position = 1)]
            [PSObject]$dataSource
    )

    function new-objectID {
        $newID = 0
        if($mv.count -eq 0) { return 0 } #mv is empty - initialize
        foreach($mvOID in $mv.Keys) {
            if($mvOID -gt $newID) { $newID = $mvOID }
        }
        return ($newID + 1)
    }

    $newID = new-objectID
    $mv[$newID] = @{} #initialise a new entry
    #FIX change to externally defined object schema
    $newEntry = @{
        "AD_samaccountname"="";"AD_userPrincipalName"="";"AD_enabled"="";"AD_givenName"="";"AD_surname"="";"AD_displayName"="";"AD_mail"="";"AD_description"="";"AD_lastLogonDate"="";"AD_daysInactive"="";"AD_PasswordLastSet"="";"AD_distinguishedname"="";"AD_parentOU"="";
        "DisplayName"="";"UserType"="";"AccountEnabled"="";"GivenName"="";"Surname"="";"UserPrincipalName"="";"Mail"="";"MFAStatus"="";"Hybrid"="";"LastLogonDate"="";"LastNILogonDate"="";"licenses"="";"Id"="";"daysInactive"="";
        "Identity"="";"EXO_DisplayName"="";"EXO_FirstName"="";"EXO_LastName"="";"EXO_RecipientType"="";"EXO_RecipientTypeDetails"="";"EXO_emails"="";"EXO_WhenMailboxCreated"="";"EXO_userPrincipalName"="";"EXO_enabled"="";"EXO_LastInteractionTime"="";"EXO_LastUserActionTime"="";"EXO_TotalItemSize"="";"EXO_ExchangeObjectId"=""    
    } 
    # prepare new entry rewriting object property values to hashtable 
    foreach ($propertyName in ( ($dataSource.psobject.Properties | ? memberType -eq 'NoteProperty')).name) {
        
        #TODO - add update of chosen attributes only, not the whole object
        $newEntry.$propertyName = $dataSource.$propertyName
    }
    $mv[$newID] = $newEntry
    Write-Verbose "metaverse object ID $newID has been added to MV table"
}

# Function to search the metaverse for a specific key-value match
function Search-Metaverse {
    <#
    .SYNOPSIS
        Search the Metaverse table
    .DESCRIPTION
        here be dragons
    .EXAMPLE
        Search-Metaverse -mv $myMetaVerse -......
    
        
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241106
            last changes
            - 241106 initialized
    
        #TO|DO
        - description
        - different types of varaibles [int/string]
        - lookup for substring and whole words
        - currently using 'first match' - should return an array for numerous matches
    #>
    [CmdletBinding(DefaultParameterSetName = 'byObject')]
    param (
        #metaverse object to search thru
        [parameter(Mandatory,position=0)]
            [validateNotNullOrEmpty()]
            [hashtable]$mv,
        #substring to search for
        [parameter(Mandatory,position=1,ParameterSetName = 'single')]
            [string]$lookupValue,
        #name of the stored object parameter to use in search. 
        [parameter(position=2,ParameterSetName = 'single')]
            [string]$columnName,
        #pass hashtable to be used for search
        [Parameter(Mandatory,position=1, ParameterSetName = 'byObject')]
            [PSObject]$lookupTable
    )

    if($PSCmdlet.ParameterSetName -eq 'single') {
        $lookupTable = @{ 
            $columnName = $lookupValue 
        }
    }

    $foundMatches = @()
    foreach ($mvKey in $mv.Keys) {
        $element = $mv[$mvKey]
        foreach($lookupColumn in $lookupTable.Keys) {
            if(-not $element.ContainsKey($lookupColumn)) { #key exists check
                #TODO ADD SOME ERROR HANDLING
                continue
            } 
            $lookupValue = $lookupTable[$lookupColumn]
            if([string]::isNullOrEmpty($lookupValue)) { #lookup value must not be null
                #maybe some warning info here?
                continue 
            }            
            if ($element[$lookupColumn] -match $lookupvalue) {
                $returnedResult = @{
                    mvID = $mvKey
                    elementProperty = $lookupColumn
                    elementValue = $element[$lookupColumn]
                }
                [array]$foundMatches += $returnedResult
                #FIX - it should just add a mach, but do not allow to make a dupe. for now - first match exist
                return $foundMatches
            } 
        }
    }
<# that supposed to lookup for a match on any column
            if($columnName) {
            } else {
            foreach ($elementKey in $element.Keys) {
                if ($element[$elementKey] -match $lookupvalue) {
                    $returnedResult = @{
                        mvID = $mvKey
                        elementProperty = $elementKey
                        elementValue = $element[$elementKey]
                    }
                    $foundMatches += $returnedResult
                }
            }
        }
    }
#>
    return $foundMatches
}


#$VerbosePreference = 'Continue'
$exportCSVFile = "CombinedReport-{0}.csv" -f (get-date -Format "yyMMdd-hhmm")

#report should always have all the fields - metafile should be a static schema
$metaverseUserInfo = @{}

Write-log "loading CSV files.." -type info
$reports = 0
if($inputCSVEntraID) {
    $EntraIDData = load-CSV $inputCSVEntraID -header @('id','displayname','givenname','surname','accountenabled','userprincipalname','mail','userType','Hybrid','givenname','surname','userprincipalname','userType','mail','daysInactive') -headerIsCritical
    $reports++
}
if($inputCSVAD) {
    $ADData = load-CSV $inputCSVAD -header @('samaccountname','userPrincipalName','enabled','givenName','surname','displayName','mail','description','daysInactive') -headerIsCritical
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

#start from populating EntraID
if($EntraIDData) {
    Write-Verbose "filling EntraID user info..."
    foreach($entraIDEntry in $EntraIDData) {
        Add-MetaverseData -mv $metaverseUserInfo -dataSource $entraIDEntry
    }
}

if($ADData) {
    Write-Verbose "adding AD user info..."
    foreach($ADuser in $ADData) {
        #first change property names, so they are not clashing between systems
        $AD_ADuser = New-Object -TypeName PSObject -Property @{}
        foreach($propertyName in ( ($ADuser.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
            #if([string]::IsNullOrEmpty($ADUser.$propertyName)
            $AD_ADUser | Add-Member -MemberType NoteProperty -Name "AD_$propertyName" -Value $ADUser.$propertyName
        } 

        #check if user already exists from Entra source
        $matchedEID = $false
        if($EntraIDData) {
            #$entraFound = Search-Metaverse -mv $metaverseUserInfo -lookupValue $AD_ADuser."AD_$matchBy" -columnName $matchBy
            [array]$entraFound = Search-Metaverse -mv $metaverseUserInfo -lookupTable @{ 
                #$matchBy = $AD_ADuser."AD_$matchBy" 
                userPrincipalName = $AD_ADuser."AD_userPrincipalName"
                displayName       = $AD_ADuser."AD_displayName"
                mail              = $AD_ADuser."AD_mail"
            }
            #add checks on different attributes 
            # ($_.displayName -eq $ADuser.displayName) -or ($_.mail -eq $ADuser.mail) -and ($_.UserType -eq 'Member')} 
            if($entraFound.count -gt 1) {
                write-verbose "duplicate found"
                Write-Verbose $entraFound
                continue
            } 
            if($entraFound.count -eq 1) {
                foreach($propertyName in ( ($entraFound[0].psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                    if($null -ne $entraFound[0].$propertyName) {
                        $AD_ADUser.$propertyName = $entraFound[0].$propertyName
                    }
                } 
                #$entraFound."AD_daysInactive" = 23000 
                Write-verbose 'matched-adding'
                Update-MetaverseData -mv $metaverseUserInfo -dataSource $AD_ADuser -objectID $entraFound[0].mvID
                $matchedEID = $true
            }

        }
        if(-not $matchedEID) {
            Write-verbose 'non-ad-adding'
            Add-MetaverseData -mv $metaverseUserInfo -dataSource $AD_ADuser
        }
    }
}


#EXO
foreach($recipient in $EXOData) {
    if($recipient.userPrincipalName) { #only mailboxes have UPNs
        $metaverseUserInfo | ? userPrincipalName -eq $recipient.userPrincipalName | %{ #locate entry by UPN 
            foreach($propertyName in ( ($recipient.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                if($propertyName -notmatch 'userPrincipalName|Identity|DisplayName|FirstName|LastName|enabled') { #skip dupes
                    $_."EXO_$propertyName" = $recipient.$propertyName
                }
            } 
        } 
    }
}

#export all results, extending with Hybrid_daysInactive attribute being lower of the comparison between EID and AD
$metaverseUserInfo.Keys | %{ 
    $metaverseUserInfo[$_] |
        Select-Object "AD_samaccountname","AD_userPrincipalName","AD_enabled","AD_givenName","AD_surname","AD_displayName","AD_mail","AD_description","AD_lastLogonDate","AD_daysInactive","AD_PasswordLastSet","AD_distinguishedname","AD_parentOU",
        "DisplayName","UserType","AccountEnabled","GivenName","Surname","UserPrincipalName","Mail","MFAStatus","Hybrid","LastLogonDate","LastNILogonDate","licenses","Id","daysInactive",
        "Identity","EXO_DisplayName","EXO_FirstName","EXO_LastName","EXO_RecipientType","EXO_RecipientTypeDetails","EXO_emails","EXO_WhenMailboxCreated","EXO_userPrincipalName","EXO_enabled","EXO_LastInteractionTime","EXO_LastUserActionTime","EXO_TotalItemSize","EXO_ExchangeObjectId",@{L='Hybrid_daysInactive';E={($_.daysInactive,$_.AD_daysInactive|Measure-Object -Minimum).minimum}}
 } | Export-Csv -Encoding unicode -NoTypeInformation $exportCSVFile

Write-Log "merged report saved to '$exportCSVFile'." -type ok
write-log "converting..."
&(convert-CSV2XLS $exportCSVFile)
write-log "done." -type ok
