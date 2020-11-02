<#
.SYNOPSIS
    Setup OoO information on the mailbox and forwarding rule.
.DESCRIPTION
    created for migration support during switchover time. sets forwarding and OoO on the mailbox.
.EXAMPLE
    .\Untitled-1
    Explanation of what the example does
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201102
        last changes
        - 201102 initialized
#>
[CmdletBinding(DefaultParameterSetName="CSV")]
param (
    #by default - use bulk import and set it for list of mailboxes
    [Parameter(ParameterSetName="CSV",mandatory=$true,position=0)]
        [string]$inputList,
    #for single-user - provide source and target emails 
    [Parameter(ParameterSetName="single",mandatory=$true,position=0)]
        [string]$sourceMail,
    [Parameter(ParameterSetName="single",mandatory=$true,position=1)]
        [string]$targetMail,
    #html message file to include in the OoO message
    [Parameter(ParameterSetName="CSV",mandatory=$true,position=1)]
    [Parameter(ParameterSetName="single",mandatory=$true,position=2)]
        [string]$messageFile,
    #delimiter for CSVs
    [Parameter(ParameterSetName="CSV",mandatory=$false,position=2)]
    [Parameter(ParameterSetName="single",mandatory=$false,position=3)]
        [string]$delimiter=';'
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
if(-not (test-path $messageFile)) {
    write-log "can't read message file $messageFile" -type error
    exit -1
}
$message=Get-Content $messageFile
check-ExchangeConnection

$header=@('sourceMail','targetMail')
if($PSCmdlet.ParameterSetName -eq 'CSV') {
    $mailboxList=load-CSV -inputCSV $inputList -header $header -headerIsCritical -delimiter ';'
} else {
    $mailboxList=@(
        [PSObject]@{
            sourceMail=$sourceMail
            targetMail=$targetMail
        }
    )
}
foreach($eMail in $mailboxList) {
    $currentMessage = $message.Replace('[targetMail]',$eMail.targetMail)

    try {
        #https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailboxautoreplyconfiguration?view=exchange-ps
        set-mailboxAutoReplyConfiguration -identity $eMail.sourceMail -autoReplyState Enabled -ExternalAudience All `
            -InternalMessage $currentMessage -ExternalMessage $currentMessage
    } catch {
        write-log "can't set mailbox OoO" -type error
        write-log $_.exception -type error
        continue
    }
    try {
        set-mailbox  -identity $eMail.sourceMail -ForwardingSmtpAddress $eMail.targetMail
    } catch {
        write-log "can't set mailbox forwarding to $($eMail.targetMail)" -type error
        write-log $_.exception -type error
    }
}
write-log "done." -type ok