﻿<#
.SYNOPSIS
    change user primary SMTP addresses. may be used in bulk mode or for single user.
.DESCRIPTION
    script is automating change of Primary SMTP address, written mainly having hybid Exchange and 
    migration projects on mind. allows to be used in bulk mode using CSV file. CSV must have two 
    columns - 'SAMAccountName' of AD user and 'newPrimarySMTP' which should self-explaining. 
    script is veryfing if chosen email is on Accepted Domains list but beside that there is not much 
    error checking or verifiation 
    
                    USE WITH CARE

    this is 'compact' version - no additional libraries are necessary (functions included in-line)
.EXAMPLE 
    .\change-PrimarySMTP.ps1 -inputCSV c:\temp\userList.csv -delimiter ';'
    bulk mode usefull during migration - you create account and mailboxes, migrate content and then
    swithover to a new environment. before you do switchover you don't want to have a source domain 
    in your environment. then you can change all emails in bulk during the switchover. 

.EXAMPLE 
    .\change-PrimarySMTP.ps1 -samaccountname myADUser -newPrimarySMTP my.AD.user@new.domain
    will change primary SMTP for a single user in AD. if 'new.domain' is not on Accepted Domains 
    list, email will not be changed

.EXAMPLE 
    .\change-PrimarySMTP.ps1 -samaccountname myADUser -newPrimarySMTP my.AD.user@new.domain -disableDomainVerification
    will change primary SMTP for a single user in AD skipping domain check. not sure why would you like 
    to do that but... 

.NOTES
    nExoR ::))o-
    ver.201214
    last changes
    - 201214 start-logging and write-log update
    - 202015 verificaiton of domain, contains is case sensitive? O_o
    - 200930 check if on accepted domain list
    - 200916 beta, signle mode, standardized functions.


    #bugs and TODOs:
     merge with template and group aliases

 #>
#requires -modules ActiveDirectory
[cmdletbinding(DefaultParameterSetName='CSV')]
param( 
    [parameter(mandatory=$true,position=0,ParameterSetName='CSV')]
        [string]$inputCSV,
    [parameter(mandatory=$false,position=1,ParameterSetName='CSV')]
        [string]$delimiter=',',

    [parameter(mandatory=$true,position=0,ParameterSetName='single')]
        [string]$samAccountName,
    [parameter(mandatory=$true,position=1,ParameterSetName='single')]
        [string]$NewPrimarySMTP,

    [parameter(mandatory=$false,position=2,ParameterSetName='CSV')]
    [parameter(mandatory=$false,position=2,ParameterSetName='single')]
        [switch]$disableDomainVerification

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

start-Logging

[regex]$rxEmail='^[\w\d_.\-\+]+@(?<domain>[\w\d_.\-]+)$'

try {
    $AcceptedDomainList = Get-AcceptedDomain|Select-Object @{N='domain';E={($_.domainname.address).toLower()}}
} catch {
    write-log 'not able to get accepted domain list. check Exchange connection.' -type error
    exit -1
}

if($PSCmdlet.ParameterSetName -eq 'CSV') {
    $header=@('samaccountname','newPrimarySMTP')
    $userList=load-CSV -header $header -delimiter $delimiter -inputCSV $inputCSV -headerIsCritical
} else {
    $userList=@()
    $userList+=@{
        samaccountname=$samAccountName
        newPrimarySMTP=$NewPrimarySMTP
    }
}


foreach($user in $userList) {

    write-log "processing $($user.samaccountname)" -type info

    if(-not $disableDomainVerification ) {
        $emailDomain=$rxEmail.Match($user.newPrimarySMTP).groups['domain'].value
        write-log "email domain: $emailDomain" -type info -silent
        if(-not $AcceptedDomainList.domain.contains($emailDomain.toLower()) ) {
            write-log "$NewPrimarySMTP is not an accepted domain. skipping" -type error
            continue
        }
    }

    $ADu=Get-ADUser $user.samaccountname -Properties proxyaddresses
    if([string]::IsNullOrEmpty($ADu) ) {
        write-host "$($user.samaccountname) not found in AD. skipping"
        continue
    }
    
    $newProxyAddresses=@()
    foreach($a in $ADu.proxyAddresses) {
        if($a -cmatch '^SMTP:') {
            Write-verbose "FOUND $a"
            $newProxyAddresses+=$a.toLower()
        } else {
            $newProxyAddresses+=$a
        }
    }
    $newProxyAddresses+="SMTP:$($user.newPrimarySMTP)"

    Write-Log "new 'would-be' aliases:" -type info
    $newProxyAddresses
    write-log "setting parameters" -type info
    try {
        Set-ADUser $ADu -Replace @{proxyAddresses=$newProxyAddresses} -EmailAddress $($user.newPrimarySMTP).replace("SMTP:",'')
        write-log "ok." -type ok
    } catch {
        write-log $_.Exception -type error
    }
}
Write-log "check $logFile" -type ok
write-log "done." -type ok
