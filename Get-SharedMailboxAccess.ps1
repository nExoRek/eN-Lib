<#
.SYNOPSIS
  list mailbox permissions - fullAccess, SendOnBehalf and SendAs. 
.DESCRIPTION
  migration support script to backup source mailboxes accesses. after emails are translated to target
  values - they can be later used by 'grant-SharedMailboxAccess.ps1' to write them back to target EXO.
.EXAMPLE
  get-mailbox -resultsize unlimited| .\Get-SharedMailboxAccess.ps1 -exportCSV

  reads all mailboxes and relevant permissions, and back them up in CSV file
.EXAMPLE
  get-mailbox -resultsize unlimited| .\Get-SharedMailboxAccess.ps1

  reads all mailboxes and relevant permissions, and display them to screen
.INPUTS
  Mailbox objects
.OUTPUTS
  list of permissions - to CSV or on screen.
.LINK
  https://w-files.pl
.NOTES
  nExoR ::))o-
  version 201115
    last changes
    - 201115 initialized
#>
#requires -module ExchangeOnlineManagement
[cmdletbinding(DefaultParameterSetName='pipe')]
param( 
  #mailbox object for pipeline
  [Parameter(ParameterSetName='pipe',ValueFromPipeline=$true,mandatory=$true,position=0)]
      [object]$EXOmailbox,
  #automatically output to CSV
  [Parameter(mandatory=$false,position=1)]
      [switch]$exportToCSV,
  #CSV delimiter
  [Parameter(mandatory=$false,position=2)]
      [string][validateSet(';',',')]$delimiter=';'
) 
 
begin { 
  #Getting Mailbox permission 
  function Get-AccessPermissions { 
    param(
      #passed exomailbox object
      [Parameter(mandatory=$true,position=0)]
          [object]$mailbox
    )

    #Getting delegated Fullaccess permission for mailbox 
    $FullAccessPermissions = (Get-MailboxPermission -Identity $mailbox.UserPrincipalName | 
      Where-Object { 
        ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") 
      }
    ).User 
    if ([string]$FullAccessPermissions -ne "") { 
      $retObject += new-object -TypeName psobject -Property @{
        'displayName' = $mailbox.Displayname
        'UserPrinciPalName' = $mailbox.UserPrincipalName
        'emailAddress' = $mailbox.PrimarySMTPAddress
        'accessType' = "FullAccess" 
        'grantAccessTo' = ($FullAccessPermissions -join ',')
        'aliases' = ( ( ( $mailbox.EmailAddresses | Select-String -CaseSensitive "smtp:" ) -join ',' ) -replace 'smtp:','' )
        'MBXType' = $mailbox.RecipientTypeDetails  
      }
      write-host "found full access for $($mailbox.Displayname)"
    }
    return $retObject
  }
  function get-SendAsPermissions {
    param(
      #passed exomailbox object
      [Parameter(mandatory=$true,position=0)]
          [object]$mailbox
    )
  
    #Getting delegated SendAs permission for mailbox 
    $SendAsPermissions = (Get-RecipientPermission -Identity $mailbox.UserPrincipalName | Where-Object { 
      -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21")) 
    }).Trustee 
    if ([string]$SendAsPermissions -ne "") { 
      $retObject += new-object -TypeName psobject -Property @{
        'displayName' = $mailbox.Displayname
        'UserPrinciPalName' = $mailbox.UserPrincipalName
        'emailAddress' = $mailbox.PrimarySMTPAddress
        'accessType' = "sendAs"
        'grantAccessTo' = ($SendAsPermissions -join ',')
        'aliases' = ( ( ( $mailbox.EmailAddresses | Select-String -CaseSensitive "smtp:" ) -join ',' ) -replace 'smtp:','' )
        'MBXType' = $mailbox.RecipientTypeDetails  
      } 
      write-host "found sendAs for $($mailbox.Displayname)"
    } 
    return $retObject
  }
  function get-SendOnBehalfPermissions {
    param(
      #passed exomailbox object
      [Parameter(mandatory=$true,position=0)]
          [object]$mailbox
    )

    #Getting delegated SendOnBehalf permission for mailbox 
    $SendOnBehalfPermissions = $mailbox.GrantSendOnBehalfTo 
    if ( -not [string]::IsNullOrEmpty($SendOnBehalfPermissions) ) { 
      $UserWithAccess = @()
      foreach ($SendOnBehalfName in $SendOnBehalfPermissions) { 
        $filter = "name -eq ""$SendOnBehalfName"""
        $onBehalfUserName = (get-recipient -filter $filter).PrimarySmtpAddress
        $UserWithAccess += $onBehalfUserName
      } 
      
      $retObject += new-object -TypeName psobject -Property @{
        'displayName' = $mailbox.Displayname
        'UserPrinciPalName' = $mailbox.UserPrincipalName
        'emailAddress' = $mailbox.PrimarySMTPAddress
        'accessType' = "SendOnBehalf"
        'grantAccessTo' = [string]($userwithAccess -join ',')
        'aliases' = ( ( $mailbox.EmailAddresses | Select-String -CaseSensitive "smtp:" ) -join ',' ) -replace 'smtp:',''
        'MBXType' = $mailbox.RecipientTypeDetails  
      } 
      write-host "found sendOnBehalf $($mailbox.Displayname)"
    } 
    return $retObject
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
  ###############################  MAIN  ##################################

  if(-not (check-ExchangeConnection)) {
    write-log -message "you need Exchange Online connection. quitting." -type error
    exit -13
  }

  #Set output file 
  $ExportCSV = ".\SharedMBPermissionReport_$(Get-Date -format yyMMddhhmm).csv" 
  $AllResults = @() 

}

process {
  write-host -ForegroundColor DarkGray "processing $($EXOmailbox.displayName)..."
  $AllResults += Get-AccessPermissions $EXOmailbox
  $AllResults += get-SendAsPermissions $EXOmailbox
  $AllResults += get-SendOnBehalfPermissions $EXOmailbox
}

end {
  #Open output file after execution  
  Write-Host "`nScript executed successfully "
  if($exportToCSV.IsPresent) {
    $AllResults | Select-Object MBXType,emailAddress,grantAccessTo,AccessType,DisplayName,Aliases | Export-Csv -Path $ExportCSV -NoTypeInformation -encoding UTF8 -Delimiter $delimiter
    Write-Host "Detailed report available in: $ExportCSV"  -ForegroundColor Green
  } else {
    $AllResults
  }
}