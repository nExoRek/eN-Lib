<#
.SYNOPSIS
    simple script semi-automating basic system checks. requires manual review of the generated logs
.DESCRIPTION
    this is a 3 step process and each step might be maually disabled by using appropriate switches.
.EXAMPLE
    .\get-ServerReview.ps1
    
.INPUTS
    None.
.OUTPUTS
    seperate log files for each step
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250131
        last changes
        - 250131 changed logging to unattended mode, sfc checks and disk checks changed, app and system logs splited
        - 250108 initialized

    #TO|DO
    - error handling. currently none. 
#>

[CmdletBinding()]
param (
    #disable disk scans
    [Parameter(mandatory=$false,position=0)]
        [switch]$noDiskScan,
    #disable sfc scan
    [Parameter(mandatory=$false,position=1)]
        [switch]$noSFC,
    #disable eventlogs
    [Parameter(mandatory=$false,position=2)]
        [switch]$noEventLogCheck
    
)
$VerbosePreference = 'Continue'
$logList = @()
$runDate = [string](Get-Date -Format yyMMddHHmm)

if(-not $noDiskScan) {
    write-verbose "STEP 1 - DISK CHECKS"

    $volumes = Get-Volume|? {$_.drivetype -eq 'Fixed' -and $_.driveletter}
    $outFile = "c:\temp\{0}-{1}-diskChecks.log" -f $runDate,$Env:COMPUTERNAME
    $logList += $outFile
    foreach($volume in $volumes.DriveLetter){
        "checking disk $volume ..." | Tee-Object -FilePath $outFile -Append
        Repair-Volume -DriveLetter $volume -Scan | out-file $outFile -Append
    }
} 

if(-not $noEventLogCheck) {
    write-verbose "STEP 2 - EventLog dump"
    $days = 60
    # Calculate the start date (60 days ago)
    $startDate = (Get-Date).AddDays(-$days)
    $eventLogs = @()

    # Get the event logs from the Application log
    # Filter for events with levels: Warning (2), Error (1), and Critical (0)
    write-verbose "getting application logs for the last $days days"
    $outFile = "C:\temp\{0}-{1}-EventLogs-Application.csv" -f $runDate,$Env:COMPUTERNAME
    $logList += $outFile
    $eventLogs += Get-WinEvent -FilterHashtable @{ 
        LogName = "Application"
        StartTime =  $startDate
        Level = @(1, 2, 3)  # 1 = Error, 2 = Warning, 3 = Critical
    } | Select-Object LogName,TimeCreated, LevelDisplayName, Id, Message
    $eventLogs | Export-Csv -Path $outFile -NoTypeInformation -Force

    $outFile = "C:\temp\{0}-{1}-EventLogs-System.csv" -f $runDate,$Env:COMPUTERNAME
    $logList += $outFile
    write-verbose "getting system logs for the last $days days"
    $eventLogs += Get-WinEvent -FilterHashtable @{ 
        LogName = "System"
        StartTime =  $startDate
        Level = @(1, 2, 3)  # 1 = Error, 2 = Warning, 3 = Critical
    } | Select-Object LogName,TimeCreated, LevelDisplayName, Id, Message

    # Export the filtered logs to a CSV file
    $eventLogs | Export-Csv -Path $outFile -NoTypeInformation -Force
}
if(-not $noSFC) {
    write-verbose "STEP 3 - OS consistency check"
    $outFile = "c:\temp\{0}-{1}-consistency.log" -f $runDate,$Env:COMPUTERNAME
    $logList += $outFile
    $scanResult = &"c:\windows\system32\sfc.exe" /scannow
    $scanResult | out-file $outFile

    #although it's not 1oo% the same as SFC it's very close and can be used as a workaround and for better reporting and automation:
    #Repair-WindowsImage -Online -ScanHealth | out-file $outFile
    #may also be automated on error with:
    #Repair-WindowsImage -Online -RestoreHealth
}

write-verbose "logs exported:"
$logList
write-verbose "done"
