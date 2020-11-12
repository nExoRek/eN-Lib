<#
.SYNOPSIS
    enforce MFA on a user.
.EXAMPLE
    .\enable-MFAforMSOLUser.ps1 -UserPrincipalName nexor@w-files.pl 

    single user MFA enablement.
.EXAMPLE
    cat userfile.txt|.\enable-MFAforMSOLUser.ps1 

    bulk user enablement by pipelinging names on the script.
.INPUTS
    User UPN(s).
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201112
        last changes
        - 201112 initialized
#>
#requires -module MSOnline
[CmdletBinding()]
param (
    #User Principal Names pipelined from console
    [Parameter(ValueFromPipeline=$true,mandatory=$true,position=0)]
        [string]$UserPrincipalName
)
begin {
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
    start-Logging
    try {
        $msoldomain=Get-MsolDomain -ErrorAction Stop
    } catch {
        write-log "you're probably not connected - use connect-MSOLService first." -type error
        exit -1
    } 
    $auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $auth.RelyingParty="*"
    $auth.State="Enforced"
}
process {
    write-Log "processing $UserPrincipalName..." -type info -silent
    try {
        Set-MsolUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationRequirements @($auth) 
        write-log "MFA enabled for $UserPrincipalName" -type info
    } catch {
        write-log "error enabling MFA for $UserPrincipalName" -type error
        continue
    }
}
end {
    write-log "done." -type ok
}