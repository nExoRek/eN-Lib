<#
.SYNOPSIS
    wrote for QoD teams migration - after teams are migrated, you want to change names.
.DESCRIPTION
    QoD do not allow to change teams names during migration to target tenant. so you migrate first,
    then you need to ensure that there are no duplicates (same names in source and target), then
    change names with required prefix/sufix using groupID, to avoid renaming incorrect group. 
    so the flow looks like:
    1. connect-MicrosoftTeams to source tenant
    2. run compare-TeamsBetweenTenants.ps1 
    3. connect-MicrosoftTeams to target tenant 
    4. run compare-TeamsBetweenTenant.ps1 -target -fileName <output-from-previous>
    5. check duplicates, choose proper maching - delete rows not matching with source.

    now you have a list of target teams to rename, avoiding dups
.EXAMPLE
    .\compare-TeamsBetweenTenants.ps1
    
    will enumerate Teams to a CSV file
.EXAMPLE
    .\compare-TeamsBetweenTenants.ps1 -target -file <source-CSV-file>
    
    will query target tenant on each row from <source-CSV-file> and fill 'target' columns.
.INPUTS
    None.
.OUTPUTS
    CSV file.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201126
        last changes
        - 201126 initialized
#>
#requires -module MicrosoftTeams
[CmdletBinding(DefaultParameterSetName='source')]
param (
    #indicate that it is initial run in source tenant
    [Parameter(ParameterSetName='source',mandatory=$true,position=0)]
        [switch]$sourceTenant,
    #indicate that it is target tenant 
    [Parameter(ParameterSetName='target',mandatory=$true,position=0)]
        [switch]$tagetTenant,
    #CSV file name - output for source, input for target
    [Parameter(ParameterSetName='source',mandatory=$false,position=1)]
    [Parameter(ParameterSetName='target',mandatory=$true,position=1)]
        [string]$fileName,
    #delimiter for CSV
    [Parameter(mandatory=$false,position=2)]
        [string][validateSet(',',';')]$delimiter=';'
)
function start-Logging {
    param()

    $scriptBaseName = ([System.IO.FileInfo]$PSCommandPath).basename
    $logFolder = "$PSScriptRoot\Logs"

    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder|Out-Null
            write-host "$LogFolder created."
        } catch {
            $_
            exit -2
        }
    }

    $script:logFile="{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
    write-Log "*logging initiated $(get-date)" -silent -skipTimestamp
    write-Log "*script parameters:" -silent -skipTimestamp
    if($script:PSBoundParameters.count -gt 0) {
        write-log $script:PSBoundParameters -silent -skipTimestamp
    } else {
        write-log "<none>" -silent -skipTimestamp
    }
    write-log "***************************************************" -silent -skipTimestamp
}
function write-log {
    param(
        #message to display - can be an object
        [parameter(mandatory=$true,position=0)]
              $message,
        #adds description and colour dependently on message type
        [parameter(mandatory=$false,position=1)]
            [string][validateSet('error','info','warning','ok')]$type,
        #do not output to a screen - logfile only
        [parameter(mandatory=$false,position=2)]
            [switch]$silent,
        # do not show timestamp with the message
        [Parameter(mandatory=$false,position=3)]
            [switch]$skipTimestamp
    )

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {$message=''}
    $message=($message|out-String).trim() 
    
    try {
        if(-not $skipTimestamp) {
            $message = "$(Get-Date -Format "hh:mm:ss>") "+$type.ToUpper()+": "+$message
        }
        Add-Content -Path $logFile -Value $message
        if(-not $silent) {
            switch($type) {
                'error' {
                    write-host -ForegroundColor Red $message
                }
                'info' {
                    Write-Host -ForegroundColor DarkGray $message
                }
                'warning' {
                    Write-Host -ForegroundColor Yellow $message
                }
                'ok' {
                    Write-Host -ForegroundColor Green $message
                }
                default {
                    Write-Host $message 
                }
            }
        }
    } catch {
        Write-Error 'not able to write to log. suggest to cancel the script run.'
        $_
    }    
}
function load-CSV {
    param(
        [parameter(mandatory=$true,position=0)]
            [string]$inputCSV,
        [parameter(mandatory=$true,position=1)]
            [string[]]$header,
        #expected header
        [parameter(mandatory=$false,position=2)]
            [switch]$headerIsCritical,
        #this flag will exit on load if any column is missing. 
        [parameter(mandatory=$false,position=3)]
            [string]$delimiter=','
    )

    try {
        $CSVData=import-csv -path "$inputCSV" -delimiter $delimiter -Encoding UTF8
    } catch {
        Write-log "error opening $inputCSV`: $($_.exception). quitting." -type error 
        exit -1
    }

    $csvHeader=$CSVData|get-Member -MemberType NoteProperty|select-object -ExpandProperty Name
    $hmiss=@()
    foreach($el in $header) {
        if($csvHeader -notcontains $el) {
            Write-log "$el column missing in imported csv" -type warning
            $hmiss+=$el
        }
    }
    if($hmiss) {
        if($headerIsCritical) {
            Write-log "Wrong CSV header. check delimiter used. quitting." -type error
            exit -2
        }
        $ans=Read-Host -Prompt "some columns are missing. type 'add' to add them, 'c' to continue or anything else to cancel"
        switch($ans) {
            'add' {
                foreach($newCol in $hmiss) {
                    $CSVData|add-member  -MemberType NoteProperty -Name $newCol -value ''
                }
                write-log "header extended" -type info
            }
            'c' {
                write-log "continuing without header change" -type info
            }
            default {
                write-log "cancelled. exitting." -type info
                exit -7
            }
        }
    }
    return $CSVData
}
start-Logging

if($sourceTenant.IsPresent) {
    if( [string]::IsNullOrEmpty($fileName) ) {
        $fileName="TeamsListSource-$(get-date -format yyMMddHHmm).csv"
    } 
    write-log "dumping Teams information from source tenant..." -type info
    get-team|
        Select-Object @{N='source_DisplayName';E={$_.DisplayName}},@{N='source_GroupId';E={$_.GroupId}}, `
            @{N='Source_MailNickName';E={$_.MailNickName}},@{N='Source_Visibility';E={$_.Visibility}}, `
            target_DisplayName,target_groupID,target_MailNickName,target_Visibility|
        Export-csv -NoTypeInformation -Delimiter $delimiter -Encoding UTF8 $fileName
    write-log "list saved as $fileName." -type ok
    write-log "now close the session, open a session to target tenant and run with -target and -filename $fileName parameters"
} else {
    $header=@(
        "source_DisplayName","source_GroupId","Source_MailNickName","Source_Visibility",
        "target_DisplayName","target_groupID","target_MailNickName","target_Visibility"
    )
    $sourceTeamsList=load-CSV $fileName -headerIsCritical -header $header -delimiter $delimiter
    $allTeamsResult=@()
    write-log "creating copare Teams list from target tenant of $($sourceTeamsList.count) teams..."
    foreach($team in $sourceTeamsList) {
        write-log "checking $($team.source_displayname)..." -type info
        $getTeam=get-team -DisplayName $team.source_DisplayName
        if( [string]::IsNullOrEmpty($getTeam) ) {
            $team.target_DisplayName = "NOT FOUND IN TARGET"
            $allTeamsResult+=$team
        } else {
            foreach($dupe in $getTeam) {
                $tempTeam=New-Object -TypeName psobject -Property @{
                    'source_DisplayName'=$team.'source_DisplayName'
                    'source_groupID'=$team.'source_groupID'
                    'source_MailNickName'=$team.'source_MailNickName'
                    'source_Visibility'=$team.'source_Visibility'
                }
                $tempTeam|Add-Member -MemberType NoteProperty -Name 'target_Displayname' -Value $dupe.displayName
                $tempTeam|Add-Member -MemberType NoteProperty -Name 'target_groupID' -Value $dupe.GroupId
                $tempTeam|Add-Member -MemberType NoteProperty -Name 'target_MailNickName' -Value $dupe.MailNickName
                $tempTeam|Add-Member -MemberType NoteProperty -Name 'target_Visibility' -Value $dupe.Visibility
                $allTeamsResult+=$tempTeam
            }
        }

    }
    $outFile="teamsComparison-$(get-date -Format yyMMddHHmm).csv"
    $allTeamsResult|export-csv -nti -Delimiter $delimiter -Encoding UTF8 $outFile
    write-log "comparison saved in $outFile"
    write-log "done." -type ok
} 

