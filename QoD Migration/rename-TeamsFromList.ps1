<#
.SYNOPSIS
    single-step script to ranme target teams with given prefix.
.DESCRIPTION
    support script for Quest on Demand which do not support on-the-fly teams name change.
    this script assumes that teams are already migrated and now we need to rename them. 
    QoD export is an input, givin a list of migrated teams. for each team, column 'target team name'
    is used to locate team and rename to $prefix+'target team name'.
.EXAMPLE

.INPUTS
    CSV created by QoD
.OUTPUTS
    none.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201203
        last changes
        - 201203 initialized
#>
#requires -module MicrosoftTeams
[CmdletBinding()]
param (
    #input list file
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSV,
    #prefix to add to each team
    [Parameter(mandatory=$true,position=1)]
        [string]$prefix,
    #delimiter for CSV files
    [Parameter(mandatory=$false,position=2)]
        [string][validateSet(',',';')]$delimiter=';'
)
function start-Logging {
    param()

    $scriptRun                          = $PSCmdlet.MyInvocation.MyCommand #(get-variable MyInvocation -scope 1).Value.MyCommand
    [System.IO.fileInfo]$scriptRunPaths = $scriptRun.Path 
    $scriptBaseName                     = $scriptRunPaths.BaseName
    $scriptFolder                       = $scriptRunPaths.Directory.FullName
    $logFolder                          = "$scriptFolder\Logs"

    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder|Out-Null
            write-host "$LogFolder created."
        } catch {
            $_
            exit -1
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

    if(-not (test-path $inputCSV) ) {
        write-log "$inputCSV not found." -type error
        exit -1
    }

    try {
        $CSVData=import-csv -path "$inputCSV" -delimiter $delimiter -Encoding UTF8
    } catch {
        Write-log "not able to open $inputCSV. quitting." -type error 
        exit -2
    }

    $csvHeader=$CSVData|get-Member -MemberType NoteProperty|select-object -ExpandProperty Name
    $hmiss=@()
    foreach($el in $header) {
        if($csvHeader -notcontains $el) {
            Write-log """$el"" column missing in imported csv" -type warning
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
$header=@('Target Team Name')
$teamsList=load-CSV -inputCSV $inputCSV -delimiter $delimiter -headerIsCritical -header $header
foreach($team in $teamsList) {
    $ttName=$($team.'Target Team Name')
    write-log "gretting ""$ttName""..."
    $teamToRename=get-team -DisplayName $ttName
    if($ttName.count) {
        write-log "uppss.. more then one team with this name found!" -type error
        write-log $teamToRename -type info
        continue
    }
    try {
        $teamToRename|sew-team -DisplayName ($prefix+' '+$ttName)
        write-log """$ttName"" renamed to $($prefix+' '+$ttName)."
    } catch {
        write-log "error renaming ""$ttName"": $($_.Exception)"
        continue
    }
}

write-log "done." -type ok