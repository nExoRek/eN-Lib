﻿<#
.SYNOPSIS
    script granting permissions to shared mailboxes. created for migration project but useful in everyday
    EXO administration. 
    grants 'full access' and 'send as'. 
.DESCRIPTION
    script has been created for customers implementing hybrid Exchange and to support bulk permissioning
    operations for shared mailboxes. it may work in bulk mode using CSV or for single user. 
    CSV file need to have two columns: 
        'emailAddress' : any email alias of shared mailbox 
        'grantAccessTo': semi-colon separated list of email aliases of users that need to have access 
                         to shared mailbox.
    this is 'compact' version - no additional libraries are necessary (functions included in-line)
    in future - it is planned to have a support module with all functions in there. 
.EXAMPLE
    .\grant-eNLibSharedMailboxAccess.ps1 -inputCSV c:\temp\listOfMailboxes.csv -delimiter ';'
    bulk permission of mailboxes based on CSV import
.EXAMPLE
    .\grant-eNLibSharedMailboxAccess.ps1 -sharedMbxName shared@w-files.pl -grantTo nexor@w-files.pl
    grants Full Access and Send As to a signgle user 

.NOTES
    nExoR ::))o-
    ver.20200930
    - 20200930 githubbed
    - 20200914 enforce EXO cmdlets
    - 20200821 initiate-logging
    - 20200727 better error handling
    - 20200724 writelog standardisation
    - 20200721 minor information change
    - 20200720 loadCSV bugfix
    - 20200629 trim 
    - 20200623 v1
 
#>
[cmdletbinding(DefaultParameterSetName="CSV")]
param(
    #CSV input file with user information, for which IDs will be generated. header: "emailAddress","grantAccessTo"
    [parameter(ParameterSetName="CSV",mandatory=$true,position=0)]
        [string]$inputCSV,
    #CSV delimiter character    
    [parameter(ParameterSetName="CSV",mandatory=$false,position=1)]
        [string][validateSet(',',';')]$delimiter=',',

    [parameter(ParameterSetName="single",mandatory=$true,position=0)]
        [string]$sharedMbxName,
    #you can provide semicolon delimited list of mailboxes (don't forget to quote the string)
    [parameter(ParameterSetName="single",mandatory=$true,position=1)]
        [string]$grantTo
)

function write-log {
    param(
        [parameter(mandatory=$true,position=0)]
            $message,
        [parameter(mandatory=$false,position=1)]
            [string][validateSet('error','info','warning','ok')]$type,
        #do not output to a screen - logfile only
        [parameter(mandatory=$false,position=2)]
            [switch]$silent
    )

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    $message=($message|out-String).trim() 
    
    try {
        switch($type) {
            'error' {
                if(-not $silent) { write-host -ForegroundColor Red $message }
                Add-Content -Path $logFile -Value "$(Get-Date -Format "hh:mm:ss>") ERROR: $message"
            }
            'info' {
                if(-not $silent) { Write-Host -ForegroundColor DarkGray "INFO: $message" }
                Add-Content -Path $logFile -Value "$(Get-Date -Format "hh:mm:ss>") INFO: $message"
            }
            'warning' {
                if(-not $silent) { Write-Host -ForegroundColor Yellow "WARNING: $message" }
                Add-Content -Path $logFile -Value "$(Get-Date -Format "hh:mm:ss>") WARNING: $message"
            }
            'ok' {
                if(-not $silent) { Write-Host -ForegroundColor Green "$message" }
                Add-Content -Path $logFile -Value "$(Get-Date -Format "hh:mm:ss>") OK: $message"
            }
            default {
                if(-not $silent) { Write-Host $message }
                Add-Content -Path $logFile -Value "$(Get-Date -Format "hh:mm:ss>") $message"
            }
        }
    } catch {
        Write-Error 'not able to write to log. suggest to cancel the script run.'
        $_
    }    
}
function initiate-Logging {
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
    write-Log "*logging initiated $(get-date)" -type info -silent
    write-Log "*script parameters:" -type info -silent
    foreach($param in $scriptRun.parameters) {
        write-log (Get-Variable -Name $Param.Values.Name -ErrorAction SilentlyContinue ) -silent
    }
    write-log "***************************************************" -type info -silent
}
function grant-SharedMailboxPermissions {
    param(
        [string]$shared,
        [string]$accessTo
    )

    #error handling is incorrect - both functions do not cast errors 
    $retValue=$null
    #here comes mailbox permission setting https://docs.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps
    try {
        add-mailboxpermission -identity $shared -accessrights FullAccess -user $accessTo -errorAction stop
    } catch {
        $retValue="mbxperm: $($_.Exception)"
    }
    #https://docs.microsoft.com/en-us/powershell/module/exchange/add-recipientpermission?view=exchange-ps
    try {
        add-recipientpermission -identity $shared -accessrights SendAs -trustee $accessTo -confirm:$false
    } catch {
        $retValue+="recperm: $($_.Exception)"
    }
   return $retValue
}
function check-ExchangeConnection {
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
function load-CSV {
    param(
        [string]$inputCSV,
        [string[]]$expectedHeader,
        [string]$delimiter
   )

    $CSVData=@()
    try {
        $CSVData=import-csv -delimiter $delimiter -path $inputCSV -Encoding UTF8
    } catch {
        Write-log -type error -message "not able to open $inputCSV"
        exit -1
    }

    $csvHeader=$CSVData|get-Member -MemberType NoteProperty|select-object -ExpandProperty Name
    $hmiss=@()
    foreach($el in $expectedHeader) {
        if($csvHeader -notcontains $el) {
            Write-log -type error "$el column missing in imported csv"
            $hmiss+=$el
        }
    }
    if($hmiss) {
        write-log -type error "some columns are missing. check delimiter. quitting."
        exit -2
    }
    Write-log -type info "loaded $($CSVData.count) records from CSV file"
    return $CSVData
}

<##################################
#
#          SCRIPT  BODY
#
###################################>
initiate-Logging

if(-not (check-ExchangeConnection)) {
    write-log -message "you need Exchange Online connection. quitting." -type error
    exit -13
}

if($PSCmdlet.ParameterSetName -eq 'CSV') {
    #information required in CSV
    $header=@("emailAddress","grantAccessTo")

    #import CSV file
    $userList=load-CSV -inputCSV $inputCSV -expectedHeader $header -delimiter $delimiter
} else {
    $userList=@(@{emailAddress=$sharedMbxName;grantAccessTo=$grantTo})
}

foreach($user in $userList) {
    
    write-log -message "PROCESSING <$($user.emailAddress)> -> <$($user.grantAccessTo)>" 
    if(-not $user.grantAccessTo) {
        write-log -message "grant-to info empty. skipping" -type warning
        continue
    }
    if(-not $user.emailAddress) {
        write-log -message "no Shared Mbx email - not able to process. skipping." -type warning
        continue
    }

    #check for mailbox existence. it must be EXO mailbox -> get-mailbox
    $shared=$user.emailAddress.trim()
    if(-not (get-mailbox $shared -errorAction SilentlyContinue)) {
        Write-log -message "mailbox $shared not found. skipping" -type error
        continue
    }

    foreach($grantPerm in $user.grantAccessTo.split(';') ) {

        #check for recipient existence - it may be any existent mail-enabled object -> get-recipient
        $accessTo=$grantPerm.trim()        
        if(-not (get-recipient $accessTo -errorAction SilentlyContinue)) {
            Write-log -message "recipient $accessTo not found. skipping" -type error
            continue
        }
        $gperm=grant-SharedMailboxPermissions -shared $shared -accessTo $accessTo
        if( (-not $gperm) ) {
            write-log -message "[shared $shared] error adding permissions to $accessTo`:" -type error
            #write-log -message $gperm
        } else { 
            write-log -message "[shared $shared] access for $accessTo granted" -type ok
        }
    }
}

write-log -message "done. check $logFile" -type ok

