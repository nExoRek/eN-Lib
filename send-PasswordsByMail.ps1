<#
.SYNOPSIS
    support script for migration project. dispatches passwords to users via email on EXO or SendGrid, using CSV as input.
.DESCRIPTION
    support script for bulk user operations created for the sake of migration. it addresses the need to send out
    passwords for newly created accounts (eg. in a new tenant) using users' email addresses from current environment.
    you can set up subject and add attachment if needed.
    special considerations:
    * in order to send password via Exchange Online with MFA-enabled account, you need to provide Application Password
      instead of regular user password. 
    * if you're using sendgrid, username is 'apikey', so you need to define 'From' address. locate $sendGridFromAddress 
      variable and set it for value of your choice.

    email should be *encrypted*, you can can e.g. create Transport Rule in Exchange Online to encrypt the messages 
    from service account that is used for email dispatch, based on specific subject or any other attribute specific for 
    the project.

    email body may contain variables, which will be replaced by values from imported CSV file:
    [GN] - will be replaced by value from "First Name" column
    [NTLOGIN] - will be replaced by value from "NT Login" column
    [PASSWORD] - will be replaced by value from "password" column
    [EMAIL] - will be replaced by value from "Target email address" column

    CSV header required: "First Name";"NT Login";"Source email address";"Target email address";"password"
    they were created for particular project reuqirement.

.EXAMPLE
    .\send-PasswordsByMail.ps1 -inputListCSV migratedUsers.csv -delimiter ';'

    dispatches emails to all users from CSV saved with Polish regional settings, with default subject.
.EXAMPLE
    .\send-PasswordsByMail.ps1 -inputListCSV migratedUsers.csv -subject 'Automated Migration Information' -attachment .\welcome.docx,.\configGuide.pdf

    dispatches emails to all users from CSV, adding subject and some guides and welcome documents as attachments.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201118
        last changes
        - 201118 test-path moed to load-csv
        - 201116 use SendGrid
        - 201112 multiple attachments fix
        - 201104 target address added as variable, saveCreds
        - 201102 initialized
#>
[CmdletBinding()]
param (
    #CSV file containing emails and passwords
    [Parameter(mandatory=$true,position=0)]
        [string]$inputListCSV,
    # email subject 
    [Parameter(mandatory=$false,position=1)]
        [string]$subject="automated message",
    #full path to a files to be attached, comma delimited
    [Parameter(mandatory=$false,position=2)]
        [string[]]$attachment,
    #once credentials are saved, script will use them instead of querying
    [Parameter(mandatory=$false,position=3)]
        [switch]$saveCredentials,
    #Use SendGrid instead of EXO
    [Parameter(mandatory=$false,position=4)]
        [switch]$useSendGrid,
    #CSV delimiter
    [Parameter(mandatory=$false,position=5)]
        [string][validateSet(',',';')]$delimiter=','
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

if( $NULL -ne $attachment ) {
    if (-not (test-path $attachment -ErrorAction SilentlyContinue)) {
        write-log -type error "attachment $attachment not found."
        exit -2
    }
}

$sendGridFromAddress = "automation@w-files.pl" #SET IT UP IF YOU'RE USING SENDGRID
$header=@('First Name','NT Login','Source email address','password','Target email address')
$recipientList=load-CSV -delimiter $delimiter -headerIsCritical -header $header -inputCSV $inputListCSV
$sendMailParam=@{
    Subject=$subject
    Body=""
    BodyAsHtml=$true
    SmtpServer = ''
    UseSSL = $true
    Port = 587
    From = "" 
    Credential =  ''
    To = ""
}

#get credential - GUI or saved
if($useSendGrid.IsPresent) {
    $credsFile=$env:USERNAME+'_sendSG.crds'
} else {
    $credsFile=$env:USERNAME+'_sendEXO.crds'
}
if(test-path $credsFile) {
    $myCreds = Import-CliXml -Path $credsFile
    write-log "used saved credentials in $credsFile" -type info
} else {
    $myCreds=Get-Credential
    if($NULL -eq $myCreds) {
        write-log 'Cancelled.' -type error
        exit -3
    }
    if($saveCredentials) {
        $myCreds | Export-Clixml -Path $credsFile
        write-log "credentials saved as $credsFile" -type info
    }
}
#set up creds in parameters
if($useSendGrid.IsPresent) {
    $sendMailParam.SmtpServer = 'smtp.sendgrid.net'
    $sendMailParam.From = $sendGridFromAddress #configure any valid user from your domain
    $sendMailParam.Credential =  $myCreds #sendgird auth is using 'apikey' user and API key as password
} else {
    $sendMailParam.SmtpServer = 'smtp.office365.com'
    $sendMailParam.From = $myCreds.UserName
    $sendMailParam.Credential =  $myCreds
}

$messageBody="<b>Hello, [GN]</b><p />
your new login is: [NTLOGIN]<br />
your new email is: [EMAIL]<br />
one time password: [PASSWORD]<br />
<i>have a great new experience!</i><p />
"
if($attachment) {
    $sendMailParam.Add('attachment',$attachment)
}

foreach($recipient in $recipientList) {
    $mailTo=$recipient.'Source email address'
    write-log "sending email to $mailTo ..." -type info
    $sendMailParam['To'] = $mailTo
        $body=$messageBody.Replace('[PASSWORD]',$recipient.password)
        $body=$body.Replace('[NTLOGIN]',$recipient.'NT Login')
        $body=$body.Replace('[GN]',$recipient.'First Name')
        $body=$body.Replace('[EMAIL]',$recipient.'')
    $sendMailParam['Body']=$body
    try {
        Send-MailMessage @sendMailParam
        write-log "sent." -type info
    } catch {
        Write-Log "Finished with error: $($_.Exception)" -type error
        continue
    }
}
write-log 'All done.' -type ok
