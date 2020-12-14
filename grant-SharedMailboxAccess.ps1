<#
.SYNOPSIS
    script granting permissions to shared mailboxes. created for migration project but useful in everyday
    EXO administration. 
    grants 'full access' and 'send as' or 'send on behalf' for a mailbox.
.DESCRIPTION
    script has been created for customers implementing hybrid Exchange and to support bulk permissioning
    operations for shared mailboxes. it may work in bulk mode using CSV or for single user. 
    CSV file need to have three columns: 
        'emailAddress' : any email alias of shared mailbox 
        'grantAccessTo': semi-colon separated list of email aliases of users that need to have access 
                         to shared mailbox.
        'accessType'   : optional, default is 'sendAs+Full'. can be sendAs or sendOnBehalf or FullAccess, or combination
    
    'accessType' may be empty - then default will be 'sendAs+FullAccess'
    
    this is 'compact' version - no additional libraries are necessary (functions included in-line)
    
.EXAMPLE
    .\grant-eNLibSharedMailboxAccess.ps1 -inputCSV c:\temp\listOfMailboxes.csv -delimiter ';'

    bulk permission of mailboxes based on CSV import
.EXAMPLE
    .\grant-eNLibSharedMailboxAccess.ps1 -sharedMbxName shared@w-files.pl -grantTo nexor@w-files.pl

    grants Full Access and SendAs (no accessType defaults to) to a single mailbox.
.EXAMPLE
    .\grant-eNLibSharedMailboxAccess.ps1 -sharedMbxName shared@w-files.pl -grantTo nexor@w-files.pl -accessType FullAccess,SendOnBehalf

    grants Full Access and SendOnBehalf to a single mailbox
.NOTES
    nExoR ::))o-
    ver.201214
    - 201214 start-logging update
    - 201116 minor standardization fixes
    - 201115 sendOfBehalf/sendAs handling thru accessType parameter
    - 200930 githubbed
    - 200914 enforce EXO cmdlets
    - 200821 initiate-logging
    - 200727 better error handling
    - 200724 writelog standardisation
    - 200721 minor information change
    - 200720 loadCSV bugfix
    - 200629 trim 
    - 200623 v1
 
#>
#requires -module ExchangeOnlineManagement
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
        [string]$grantTo,
    #permission type
    [parameter(ParameterSetName="single",mandatory=$false,position=2)]
        [string[]][validateSet('SendAs','SendOnBehalf','FullAccess')]$accessType=@('sendAs','FullAccess')
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
function grant-SharedMailboxPermissions {
    param(
        [string]$shared,
        [string]$accessTo,
        [string][validateSet('SendAs','SendOnBehalf','FullAccess')]$accessType
    )

    switch($access) {
        'FullAccess' {
            #here comes mailbox permission setting https://docs.microsoft.com/en-us/powershell/module/exchange/add-mailboxpermission?view=exchange-ps
            try {
                add-mailboxpermission -identity $shared -accessrights FullAccess -user $accessTo -errorAction stop
            } catch {
                write-log "error adding FullAccess permissions to $accessTo : $($_.Exception)"
                return $null
            }
        }
        "SendAs" {
            try {
                add-recipientpermission -identity $shared -accessrights SendAs -trustee $accessTo -confirm:$false -ErrorAction stop
            } catch {
                write-log "error adding SendAs permissions to $accessTo : $($_.Exception)"
                return $null
            }
        }
        #https://docs.microsoft.com/en-us/powershell/module/exchange/add-recipientpermission?view=exchange-ps
        #for sendonbehalf: https://docs.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-permissions-for-recipients
        "SendOnBehalf" {
            try {
                set-mailbox -Identity $shared -GrantSendOnBehalfTo $accessTo -Confirm:$false -ErrorAction stop
            } catch {
                write-log "error adding SendOnBehalf permissions to $accessTo : $($_.Exception)"
                return $null
            }
        }
    }
    return $null
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
            Write-log -type error """$el"" column missing in imported csv"
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
start-Logging

if(-not (check-ExchangeConnection)) {
    write-log -message "you need Exchange Online connection. quitting." -type error
    exit -13
}

if($PSCmdlet.ParameterSetName -eq 'CSV') {
    #information required in CSV
    $header=@("emailAddress","grantAccessTo","AccessType")

    #import CSV file
    $userList=load-CSV -inputCSV $inputCSV -expectedHeader $header -delimiter $delimiter
} else {
    $userList=@( @{ 
        emailAddress = $sharedMbxName
        grantAccessTo = $grantTo
        accessType = $accessType
    } )

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
    if($PSCmdlet.ParameterSetName -eq 'CSV' -and [string]::IsNullOrEmpty($user.accessType) ) {
        $accessType = @('SendAs','FullAccess')
    } else {
        $accessType = ($user.accessType).split(',;')
    }

    #check for mailbox existence. it must be EXO mailbox -> get-mailbox
    $shared=$user.emailAddress.trim()
    if(-not (get-mailbox $shared -errorAction SilentlyContinue)) {
        Write-log -message "mailbox $shared not found. skipping" -type error
        continue
    }

    foreach($grantPerm in $user.grantAccessTo.split(';,') ) {
        #check for recipient existence - it may be any existent mail-enabled object -> get-recipient
        $accessTo=$grantPerm.trim()        
        if(-not (get-recipient $accessTo -errorAction SilentlyContinue)) {
            Write-log -message "recipient $accessTo not found. skipping" -type error
            continue
        }
        foreach($access in $accessType) {
            $gperm=grant-SharedMailboxPermissions -shared $shared -accessTo $accessTo -accessType $access
            if( (-not $gperm) ) {
                #write-log -message "[shared $shared] error adding $access to $accessTo" -type error
            } else { 
                write-log -message "[shared $shared] $access for $accessTo granted" -type ok
            }
        }
    }
}

write-log -message "done. check $logFile" -type ok

