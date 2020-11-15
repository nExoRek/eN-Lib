<#
.SYNOPSIS
  list mailbox permissions.
.DESCRIPTION
  later
.EXAMPLE
  .\Get-EXOMailboxPermissions.ps1
  Explanation of what the example does
.INPUTS
  None.
.OUTPUTS
  None.
.LINK
  https://w-files.pl
.NOTES
  nExoR ::))o-
  version 201115
    last changes
    - 201115 initialized
#>
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
      [string][validateSet(';',',')]$delimiter=';',
  [switch]$FullAccess, 
  [switch]$SendAs, 
  [switch]$SendOnBehalf
) 
 
begin { 
  #Getting Mailbox permission 
  function Get_MBPermission { 
    param(
      #passed exomailbox object
      [Parameter(mandatory=$true,position=0)]
          [object]$mailbox
    )

    $thisResults=@()
    $upn = $mailbox.UserPrincipalName
    $DisplayName = $mailbox.Displayname 
    $MBType = $mailbox.RecipientTypeDetails 
    $PrimarySMTPAddress = $mailbox.PrimarySMTPAddress
    $EmailAddresses = $mailbox.EmailAddresses
    $EmailAlias = ""

    foreach ($EmailAddress in $EmailAddresses) {
      if ($EmailAddress -clike "smtp:*") {
        if ($EmailAlias -ne "") {
          $EmailAlias = $EmailAlias + ","
        }
        $EmailAlias = $EmailAlias + ($EmailAddress -Split ":" | Select-Object -Last 1 )
      }
    }

    #Getting delegated Fullaccess permission for mailbox 
    if (($FilterPresent -ne $true) -or ($FullAccess.IsPresent)) { 
      $FullAccessPermissions = (Get-MailboxPermission -Identity $upn | 
        Where-Object { 
          ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY" -or $_.User -match "S-1-5-21") 
        }).User 
      if ([string]$FullAccessPermissions -ne "") { 
        $AccessType = "FullAccess" 
        $userwithAccess= $FullAccessPermissions -join ','
        <#
        foreach ($FullAccessPermission in $FullAccessPermissions) { 
          if ($UserWithAccess -ne "") {
            $UserWithAccess = $UserWithAccess + ","
          }
          $UserWithAccess = $UserWithAccess + $FullAccessPermission 
        } 
        #>
        $thisResults += new-object -TypeName psobject -Property @{
          'displayName' = $Displayname
          'UserPrinciPalName' = $upn
          'emailAddress' = $PrimarySMTPAddress
          'accessType' = $AccessType
          'grantAccessTo' = $userwithAccess
          'aliases' = $EmailAlias
          'MBXType' = $MBType   
        } 
        write-host "found full access for $displayName"
      }
    } 
  
    #Getting delegated SendAs permission for mailbox 
    if (($FilterPresent -ne $true) -or ($SendAs.IsPresent)) { 
      $SendAsPermissions = (Get-RecipientPermission -Identity $upn | Where-Object { 
        -not (($_.Trustee -match "NT AUTHORITY") -or ($_.Trustee -match "S-1-5-21")) 
      }).Trustee 
      if ([string]$SendAsPermissions -ne "") { 
        $AccessType = "SendAs" 
        $userwithAccess= $SendAsPermissions -join ','
        <#
        foreach ($SendAsPermission in $SendAsPermissions) { 
          if ($UserWithAccess -ne "") {
            $UserWithAccess = $UserWithAccess + ","
          }
          $UserWithAccess = $UserWithAccess + $SendAsPermission 
        } 
        #>
        $thisResults += new-object -TypeName psobject -Property @{
          'displayName' = $Displayname
          'UserPrinciPalName' = $upn
          'emailAddress' = $PrimarySMTPAddress
          'accessType' = $AccessType
          'grantAccessTo' = $userwithAccess
          'aliases' = $EmailAlias
          'MBXType' = $MBType   
        } 
        write-host "found sendAs for $displayName"
      } 
    } 
  
    #Getting delegated SendOnBehalf permission for mailbox 
    if (($FilterPresent -ne $true) -or ($SendOnBehalf.IsPresent)) { 
      $SendOnBehalfPermissions = $_.GrantSendOnBehalfTo 
      if ([string]$SendOnBehalfPermissions -ne "") { 
        $UserWithAccess = "" 
        $AccessType = "SendOnBehalf" 
        foreach ($SendOnBehalfPermissionDN in $SendOnBehalfPermissions) { 
          if ($UserWithAccess -ne "") {
            $UserWithAccess = $UserWithAccess + ","
          }
          $filter = "name -eq ""$SendOnBehalfPermissionDN"""
          $onBehalfUserName = (get-recipient -filter $filter).PrimarySmtpAddress
          $UserWithAccess = $UserWithAccess + $onBehalfUserName
        } 
        $thisResults += new-object -TypeName psobject -Property @{
          'displayName' = $Displayname
          'UserPrinciPalName' = $upn
          'emailAddress' = $PrimarySMTPAddress
          'accessType' = $AccessType
          'grantAccessTo' = $userwithAccess
          'aliases' = $EmailAlias
          'MBXType' = $MBType   
        } 
        write-host "found sendOnBehalf $displayName"
      } 
    } 
    return $thisResults
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
  $Results = @() 

  #Check for AccessType filter 
  if (($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent)) {
    $FilterPresent = $true
  } 

}

process {
  write-host -ForegroundColor DarkGray "processing $($EXOmailbox.displayName)..."
  $readPermissions = Get_MBPermission $EXOmailbox
  if($NULL -ne $readPermissions) {
    $Results += $readPermissions
  }
}

end {
  #Open output file after execution  
  Write-Host "`nScript executed successfully "
  if($exportToCSV.IsPresent) {
    $Results | Select-Object MBXType,emailAddress,grantAccessTo,AccessType,DisplayName,Aliases | Export-Csv -Path $ExportCSV -NoTypeInformation -encoding UTF8 -Delimiter $delimiter
    Write-Host "Detailed report available in: $ExportCSV"  -ForegroundColor Green
  } else {
    $Results
  }
}