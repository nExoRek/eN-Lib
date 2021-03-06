<#
.SYNOPSIS
    transform columns in CSV to prepare values for new tenant.
.DESCRIPTION
    <temp support script for migration project>
.EXAMPLE
    .\transform-mailbxoCSVvalues.ps1 -inputCSV sourceMailboxes.csv -delimiter ','
    
.INPUTS
    source mailbox list CSV
.OUTPUTS
    target mailbox list CSV
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201113
        last changes
        - 201113 beta 1
        - 201111 initialized
#>
[CmdletBinding()]
param (
    #input file name
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSV,
    #export CSV file name
    [Parameter(mandatory=$false,position=1)]
        [string]$outputCSV="transformed-$(get-date -format 'yyMMddHHmm').csv",
    #target domain name ofr transformation
    [Parameter(mandatory=$true,position=2)]
        [string]$targetDomain,
    #delimiter for CSV
    [Parameter(mandatory=$false,position=3)]
        [string][validateSet(';',',')]$delimiter=';'
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

    try {
        $CSVData=import-csv -path "$inputCSV" -delimiter $delimiter -Encoding UTF8
    } catch {
        Write-log "not able to open $inputCSV. quitting." -type error 
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
if(-not (test-path $inputCSV)) {
    write-log "can't read message file $inputCSV" -type error
    exit -1
}
$header=@('Type','Name','Source Mailbox')
$entryList=load-CSV -inputCSV $inputCSV -header $header -headerIsCritical -delimiter $delimiter
$prefixToAdd="SDS"

[regex]$emailName='(?i)(?<eName>[\w\d_.\-\+]+)@'
$resultList=@()
foreach($row in $entryList) {
    if($row.type -ne 'SharedMailbox') { 
        write-log "not a shared mailbox. skipping" -type info -silent
        continue 
    }
    $newEntry=[PSObject]$row
    $newEntry|add-member -MemberType NoteProperty -Name 't_name' -Value ($prefixToAdd+' '+$row.name)
    $alias=$prefixToAdd+'.'+ ( $emailName.Match($row."source mailbox").groups['eName'].value )
    $newEntry|Add-Member -MemberType NoteProperty -Name 't_alias' -value $alias
    $newEntry|Add-Member -MemberType NoteProperty -Name 't_targetMail' -value ($alias+'@'+$targetDomain)
    $resultList+=$newEntry
}
$resultList|export-csv -nti -Delimiter $delimiter -Path $outputCSV
write-log "$outputCSV exported."
write-log "done." -type ok
