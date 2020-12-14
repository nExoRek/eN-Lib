<#
.SYNOPSIS
    help in testing for set-mailboxOoOForwarding.ps1 - quickly revert the changes.
.DESCRIPTION
    CSV requires single column - "Source email address". 
.EXAMPLE
    .\disable-mailboxOoOForwarding.ps1 -inputListCSV .\migratedUsers.csv
    
    disables autoforward and OoO for users in CSV file
.INPUTS
    CSV list with valid email addresses.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201103
        last changes
        - 201103 initialized
#>
[CmdletBinding()]
param (
    [Parameter(mandatory=$true,position=0)]
        [string]$inputListCSV,
    [Parameter(mandatory=$false,position=1)]
        [string][validateSet(',',';')]$delimiter=','
)
function start-Logging {
    param(
        #create log in profile folder rather than script run path
        [Parameter(mandatory=$false,position=0)]
            [switch]$userProfilePath
    )

    $scriptBaseName = ([System.IO.FileInfo]$PSCommandPath).basename
    if($userProfilePath.IsPresent) {
        $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
    } else {
        $logFolder = "$PSScriptRoot\Logs"
    }

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
function check-ExchangeConnection {
    param(
        [parameter(mandatory=$false,position=0)]
            [validateSet('OnPrem','EXO')][string]$ExType='EXO',
            #defines if you need to check on-premise Exchange or Exchange Online.
        [parameter(mandatory=$false,position=1)]
            [switch]$isNonCritical
            #defines if connection to Exchange is critical. by default script will exit when not connected.
    )

    $exConnection=$false
    foreach($session in $(get-PSSession)) {
        if($session.ConfigurationName -eq 'Microsoft.Exchange') {
            if($ExType -eq 'EXO' -and $session.ComputerName -eq 'outlook.office365.com') {
                $exConnection=$true
            }
            if($ExType -eq 'OnPrem' -and $session.ComputerName -ne 'outlook.office365.com') {
                $exConnection=$true
            }
        }
    }
    if(-not $exConnection) {
        if($isNonCritical) {
            write-Log "connection to $ExType not established. functionality will be reduced." -type warning
        } else {
            write-log "connection to $ExType not established. you need to connect first. quitting." -type error
            exit -1
        }
    }
    
}
start-Logging
check-ExchangeConnection

$header=@("Source email address")
$mailboxList=load-CSV -inputCSV $inputListCSV -header $header -headerIsCritical -delimiter ';'
foreach($eMail in $mailboxList) {
    write-log "processing $($eMail."Source email address")..." -type info
    set-mailboxAutoReplyConfiguration -identity $eMail."Source email address" -autoReplyState disabled
    set-mailbox  -identity $eMail."Source email address" -ForwardingSmtpAddress $null
}
write-log "all done." -type ok