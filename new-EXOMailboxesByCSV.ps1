<#
.SYNOPSIS
    create new cloud-native mailboxes based on information from CSV.
.DESCRIPTION
    created for particualar migration project....
.EXAMPLE
    .\new-SharedMailboxesFromCSV.ps1 -inputCSV targetSharedMailboxes.csv
    
.INPUTS
    shared mailbox list
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201214
        last changes
        - 201214 start-logging update
        - 201113 beta 1
        - 201109 initialized
#>
[CmdletBinding()]
param (
    [Parameter(mandatory=$true,position=0)]
        [string]$inputCSV,
    #delimiter for CSV files
    [Parameter(mandatory=$false,position=1)]
        [string][validateSet(';',',')]$delimiter=';'
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
function new-RandomPassword {
    param( 
        [int]$length=8,
        [int][validateSet(1,2,3,4)]$uniqueSets=4,
        [int][validateSet(1,2,3)]$specialCharacterRange=1
            
    )
    function generate-Set {
        param(
            # set up password length
            [int]$length,
            # number of 'sets of sets' defining complexity range
            [int]$setSize
        )
        $safe=0
        while ($safe++ -lt 100) {
            $array=@()
            1..$length|%{
                $array+=(Get-Random -Maximum ($setSize) -Minimum 0)
            }
            if(($array|Sort-Object -Unique|Measure-Object).count -ge $setSize) {
                return $array
            } else {
                Write-Verbose "[generate-Set]bad array: $($array -join ',')"
            }
        }
        return $null
    }
    #prepare char-sets 
    $smallLetters=$null
    97..122|%{$smallLetters+=,[char][byte]$_}
    $capitalLetters=$null
    65..90|%{$capitalLetters+=,[char][byte]$_}
    $numbers=$null
    48..57|%{$numbers+=,[char][byte]$_}
    $specialCharacterL1=$null
    @(33;35..38;43;45..46;95)|%{$specialCharacterL1+=,[char][byte]$_} # !"#$%&
    $specialCharacterL2=$null
    58..64|%{$specialCharacterL2+=,[char][byte]$_} # :;<=>?@
    $specialCharacterL3=$null
    @(34;39..42;44;47;91..94;96;123..125)|%{$specialCharacterL3+=,[char][byte]$_} # [\]^`  
      
    $ascii=@()
    $ascii+=,$smallLetters
    $ascii+=,$capitalLetters
    $ascii+=,$numbers
    if($specialCharacterRange -ge 2) { $specialCharacterL1+=,$specialCharacterL2 }
    if($specialCharacterRange -ge 3) { $specialCharacterL1+=,$specialCharacterL3 }
    $ascii+=,$specialCharacterL1
    #prepare set of character-sets ensuring that there will be at least one character from at least 3 different sets
    $passwordSet=generate-Set -length $length -setSize $uniqueSets 

    $password=$NULL
    0..($length-1)|% {
        $password+=($ascii[$passwordSet[$_]] | Get-Random)
    }
    return $password
}

start-Logging
if(-not (test-path $inputCSV)) {
    write-log "can't read message file $inputCSV" -type error
    exit -1
}
$header=@('t_targetMail','t_name','t_alias')
$mailboxList=load-CSV -inputCSV $inputCSV -header $header -headerIsCritical -delimiter $delimiter
check-ExchangeConnection

foreach($mailbox in $mailboxList) {
    if($mailbox.type -eq 'SharedMailbox') {
        write-log "creating shared mailbox $($mailbox.t_name) ..." -type info
        $UPN=$mailbox.t_targetMail
        $displayName = $mailbox.t_name
        $alias=$mailbox.t_alias
        $accountPassword=new-RandomPassword
        try {
            new-mailbox -shared `
                -name $displayName `
                -displayName $displayName `
                -alias $alias `
                -PrimarySMTPAddress $UPN `
                -Password (ConvertTo-SecureString -String $accountPassword -AsPlainText -Force) 
            write-log "shared mailbox $UPN created." -type info
        } catch {
            write-log "not able to create mailbox with error: $($_.exception)" -type error
        }
        #for shared mailbox - no need to export passwords, as accounts are disabled.
    } else {
            #-ResetPasswordOnNextLogon $true `
            #below doesn't work for shared mailbox.
            #-MicrosoftOnlineServicesID $UPN `
            write-log "read $($mailbox.name) but it's not type shared. skipping" -type info -silent
    }
}
write-log "all done." -type ok