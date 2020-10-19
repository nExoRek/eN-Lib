<#
.SYNOPSIS
    eN's support functions library.
.DESCRIPTION
    <some day>
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201018
#>
function start-Logging {
    <#
    .SYNOPSIS
        initilizes log file under $logFile variable for write-log function.
    .DESCRIPTION
        Long description
    .EXAMPLE
        start-Logging

        simply initializes the log file.
    .EXAMPLE
        start-Logging -logPath c:\temp\myLogs\somelog.log

        initializes the log file as c:\temp\myLogs\somelog.log .
    .INPUTS
        None.
    .OUTPUTS
        log file under $logFile variable.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 201018
    #>
    param(
        # full name for custom log file. log folder must already exist.
        [Parameter(mandatory=$false,position=0)]
            [string]$fileFullPath
    )

    $scriptRun = (get-variable MyInvocation -scope 2).Value.MyCommand
    <# DEBUG #>
        $PSCmdlet.MyInvocation|select name,commandtype,parameters
        get-variable MyInvocation
        (get-variable MyInvocation -scope 1).Value.MyCommand|select name,commandtype,parameters
        $scriptRun|select name,commandtype,parameters
    <##>
    #check if not run outside script
    if( $scriptRun.commandType -ne 'ExternalScript' ) {
        write-host "don't run this function outside script" -ForegroundColor Red
        remove-module -name eNLib
        return $null
    }
    [System.IO.fileInfo]$scriptRunPaths = $scriptRun.Path 
    $scriptBaseName = $scriptRunPaths.BaseName

    if($fileFullPath) { #custom log file name
        $logFolder=Split-Path $fileFullPath -Parent
        if( [string]::IsNullOrEmpty($logFolder) ) { $logFolder = '.' }
        if(-not (test-path $logFolder) ) { 
            write-error "$logFolder doesn't exist. LOG NOT INITIALIZED"
            return
        }
        $fileName=Split-Path $fileFullPath -leaf
        if( [string]::IsNullOrEmpty($fileName) ) { 
            $script:logFile="{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        } 
        $script:logFile=$fileFullPath
    } else {
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
    }
    write-Log "*logging initiated $(get-date)" -silent -noTimestamp
    write-Log "*script parameters:" -silent -noTimestamp
    if($script:PSBoundParameters.count -gt 0) {
        write-log $script:PSBoundParameters -silent -noTimestamp
    } else {
        write-log "<none>" -silent -noTimestamp
    }
    write-log "***************************************************" -silent -noTimestamp
}
Export-ModuleMember -Function start-Logging

function write-log {
    <#
    .SYNOPSIS
        replacement for write-host, forking information to a log file.
    .DESCRIPTION
        automates forking of output on two different endpoints - on the host, using write-host
        and to the file, appening its content.
        write-log converts everything to a string, so you can use it for virtually any type of 
        variable. additionaly it adds timestamp, message type header and color (on host).

        information is written to a $logFile - you must initialize the value with 'start-Logging' 
        or configure it manually.

    .EXAMPLE
        .\write-log "all is fine"

        shows 'all is fine' on the screen and to the log file.
    .EXAMPLE
        .\write-log -message "trees are green" -type ok

        shows 'trees are green' in Green colour, and send text to a log file.
    .EXAMPLE
        $someObject=get-process
        .\write-log -message $someObject -type info -noTimestamp -silent

        outputs processes object to the log file as -silent disables screen output. it will lack
        timestamp in a message header but will contain '[INFO]' block.
    .INPUTS
        None.
    .OUTPUTS
        text log file under $logFile
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 201018
    #>
    
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
            [switch]$noTimestamp
    )
    #ensure log is initialized
    try {
        Get-Item $script:logFile -errorAction Stop
    } catch {
        Write-Verbose "$script:logFile is not a proper output. try start-Logging to initialize file correctly."
        return
    }

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {$message=''}
    $message=($message|out-String).trim() 
    
    try {
        if(-not $noTimestamp) {
            $message = "$(Get-Date -Format "hh:mm:ss>") "+$type.ToUpper()+": "+$message
        }
        Add-Content -Path $script:logFile -Value $message
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
Export-ModuleMember -Function write-Log

function get-ExchangeConnectionStatus {
    param(
        [parameter(mandatory=$false,position=0)][validateSet('OnPrem','EXO')][string]$ExType='EXO'
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
    return $exConnection
}
Export-ModuleMember -Function get-ExchangeConnectionStatus