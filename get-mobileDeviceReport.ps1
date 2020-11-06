<#
.SYNOPSIS
    Prepare report on mobile devices for EXO users.
.DESCRIPTION
.EXAMPLE
    .\get-mobileDeviceReport.ps1
    
    simply get the report to CSV file with default name.
.EXAMPLE
    .\get-mobileDeviceReport.ps1 -extendedStats -delimiter ';'

    get device report with some additional stats, report file will use semicolon as separator.
.INPUTS
    None.
.OUTPUTS
    report CSV file
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201106
        last changes
        - 201106 initialized
#>
[CmdletBinding()]
param (
    # add additional device stats (slower)
    [Parameter(mandatory=$false,position=0)]
        [switch]$extendedStats,
    #output CSV name
    [Parameter(mandatory=$false,position=1)]
        [string]$reportFile,
    # delimiter user for CSV output
    [Parameter(mandatory=$false,position=2)]
        [string][validateSet(',',';')]$delimiter=';'
)
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

if(-not (get-ExchangeConnectionStatus)) {
    write-host "must connect to EXO before running the script." -ForegroundColor Red
    exit -1
}

if( [string]::IsNullOrEmpty($reportFile) ) {
    $reportFile = "$PSScriptRoot\MobileReport-$( (Get-Date).ToString('yyMMdd') ).csv" 
} 

write-host "getting recipient list..."
$recipients = Get-Recipient -RecipientTypeDetails "UserMailbox" -ResultSize unlimited|Select-Object -ExpandProperty PrimarySmtpAddress
write-host "found $($recipients.count) user mailboxes"

$finalReport = @()
foreach ($smtpaddr in $recipients) {
    write-host "getting devices of $smtpaddr..."
    $mobileDevices = Get-MobileDevice -Mailbox $smtpaddr
    if($mobileDevices) {
        write-verbose "found $($mobileDevices.count) devices."
    } else {
        Write-Verbose "no devices."
    }

    foreach ($device in $mobileDevices) {

        $finalReport += New-Object psobject -Property @{ 
            mailboxSMTP = $smtpaddr

            friendlyName = $device.FriendlyName
            deviceID = $device.DeviceID
            deviceAccessState = $device.deviceAccessState
            deviceType = $device.deviceType
            deviceModel = $device.DeviceModel
            deviceOS = $device.DeviceOS
            isDisabled = $device.isDisabled
            whenCreated = $device.whenCreated
            whenChanged = $device.WhenChanged
            GUID = $device.GUID 
        }

        if($extendedStats) {
            write-host "getting extended device stats..."
            $mobileStats = Get-MobileDeviceStatistics -Identity $device.Identity

            $finalReport.IsRemoteWipeSupported.add('IsRemoteWipeSupported' ,$mobileStats.IsRemoteWipeSupported )
            $finalReport.LastSyncAttemptTime.add('LastSyncAttemptTime' ,$mobileStats.LastSyncAttemptTime )
            $finalReport.LastSuccessSync.add('LastSuccessSync' ,$mobileStats.LastSuccessSync )
            $finalReport.LastDeviceWipeRequest.add('LastDeviceWipeRequest' ,$mobileStats.lastdevicewiperequestor )
        }
    }
}

$finalReport | Export-Csv -NoTypeInformation -Delimiter $delimiter -Path $reportFile
write-host -ForegroundColor Green "report generated to $reportFile ."


