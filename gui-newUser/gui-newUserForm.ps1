<#
.SYNOPSIS
    *THIS IS DEMO ONLY for WGUiSW snack* (http://wguisw.org)
    Script simplifies creation of a new user account in Hybrid environment.
.DESCRIPTION
    with hybrid identity it is hard to manage user Exchange parameters as some are managed from on-premise.
    this script is not fully functional nither is universalized... my trash basicaly for a presentation (;
.EXAMPLE
    .\gui-newUserForm.ps1
    
    runs the script showing GUI wizard to follow.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201130
        last changes:
        - 201130 WGUiSW version
        - 201106 multidomain  fixes, licenses
        - 201020 fixes
        - 201016 initialization

#>
#requires -module ActiveDirectory, ExchangeOnlineManagement
[cmdletbinding()]
param()

#STATIC DEFINITIONS
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
[regex]$rxEmail='^[\w\d_.\-\+]+@(?<domain>[\w\d_.\-]+\.[\w\d]+)$'
[regex]$rxValidMailCharacters='[\w\d.@-]'
[regex]$rxValidNameCharacters='[\w\d. ''-]'
[regex]$rxValidPhoneCharacters='[\d-+ ]'

$USEONLINE=$false

$validateTenantDomainName="emprirebm.onmicrosoft.com"
$CreationTargetOUs=@{
    CH="OU=CH,OU=companyUsers,DC=w-files,DC=lab"
    FR="OU=FR,OU=companyUsers,DC=w-files,DC=lab"    
    PL="OU=PL,OU=companyUsers,DC=w-files,DC=lab"
}

#region FUNCTIONS

# Create Icon Extractor Assembly
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
 
namespace System
{
    public class IconExtractor
    {
 
     public static Icon Extract(string file, int number, bool largeIcon)
     {
      IntPtr large;
      IntPtr small;
      ExtractIconEx(file, number, out large, out small, 1);
      try
      {
       return Icon.FromHandle(largeIcon ? large : small);
      }
      catch
      {
       return null;
      }
 
     }
     [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
     private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
 
    }
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

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
            [switch]$skipTimestamp,
        # show also in text form
        [Parameter(mandatory=$false,position=4)]
            [system.Windows.Forms.TextBox]$txtBox
    )

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {$message=''}
    $message=($message|out-String).trim() 

    if($txtBox) {
        $txtBox.AppendText($message)
        $txtBox.AppendText([System.Environment]::NewLine)
    }
    
    try {
        if(-not $skipTimestamp) {
            $message = "$(Get-Date -Format "HH:mm:ss>") "+$type.ToUpper()+": "+$message
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
function get-AzureADConnectionStatus {
    param(
        #defines if connection is critical. by default script will exit when not connected.
        [parameter(mandatory=$false,position=0)]
            [switch]$isNonCritical
    )

    $testAAD=$null
    try {
        $testAAD=Get-AzureADDomain|out-null
    } catch {
        if($isNonCritical) {
            write-Log "connection to AAD not established. funcionality will be reduced." -type warning
        } else {
            write-Log "connection to AAD not established. please use connect-AzureAD first. quitting" -type error
            exit -1            
        }
    }
    return $testAAD
}
function Remove-Diacritics {
    param ([String]$src = [String]::Empty)

    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % { 
        if([int][char]$_ -eq 322) { #ł is not handled correctly by .NET function
            [void]$sb.Append('l')
        } else {
            if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($_)
            }
        }
    }
    return $sb.ToString()
}
function check-dupe {
    param(
        #email to check dupe for
        [parameter(mandatory=$true,position=0)]
            [string]$email,
        #by default checks against Ex onprem. with this switch - against EXO
        [parameter(mandatory=$false,position=1)]
            [switch]$useEXO
    )
    
    if($email -notmatch "^[\w\.\d-]+@[\w\.\d-]+\.[\w\d]+$") {
        write-log "this is not valid email [$email]" -type error
        return -1
    }

    if($useEXO) {
        #-identity is actually looking for any possible match on these parameters, so if the email exisit
        #on any of below, it will return result.
        #$filter="EmailAddresses -like ""*$email*"" -or PrimarySmtpAddress -eq ""$email"" -or WindowsEmailAddress -eq ""$email"""
        $dupeMail=get-recipient -identity $email -ErrorAction SilentlyContinue
        if($dupeMail) {
            write-log "dupe found:"
            write-log $dupeMail
            return "[$($dupeMail.RecipientTypeDetails)][$($dupeMail.alias)]"
        } else {
            return $null
        }
    } else {
        $filter="mail -eq ""$email"" -or proxyaddresses -like ""*smtp:$email*"" -or userprincipalname -eq ""$email"""
        $dupeMail=get-adobject -filter $filter -Properties samAccountName
        if($dupeMail) {
            if($dupeMail.objectClass -eq 'contact') {
                return "[CONTACT][$($dupeMail.distinguishedName)]"
            }
            return "[$($dupeMail.samAccountName)][$($dupeMail.distinguishedName)]"
        } else { 
            return $null
        }
    }
}
function get-validNameString {
    param(
        # string to verify
        [Parameter(mandatory=$true,position=0)]
            [string]$str
    )

    #if([string]::IsNullOrEmpty($str) ) { return ' ' }
    $outStr=''
    foreach($c in $str.ToCharArray() ) {
        if( $c -match $rxValidNameCharacters ) {
            $outStr+=$c
        }
    }
    return (Get-Culture).textinfo.totitlecase($outStr.tolower())
}
function get-validMailString {
    param(
        # string to verify
        [Parameter(mandatory=$true,position=0)]
            [string]$str
    )

    $outStr=''
    foreach($c in $str.ToCharArray() ) {
        if( $c -match $rxValidMailCharacters ) {
            $outStr+=$c
        }
    }
    return $outStr
}
function new-displayName {
    $objDN=@()
    if(-not [string]::IsNullOrEmpty($txtGivenName.text.trim()) ) { $objDN+=$txtGivenName.text.trim() }
    if(-not [string]::IsNullOrEmpty($txtMiddleName.text.trim()) ) { $objDN+=$txtMiddleName.text.trim() }
    if(-not [string]::IsNullOrEmpty($txtSurname.text.trim()) ) { $objDN+=$txtSurname.text.trim() }
    if(-not [string]::IsNullOrEmpty($txtSurnameExt.text.trim()) ) { $objDN+=$txtSurnameExt.text.trim() }
    if($chbIncludeExtDN.checked -eq $true) { $objDN+='external' }
    #if($isExternal) { $objDN+='external' }
    $outStr=$objDN -join ' '
    return $outStr
}
function get-validMobileNumber {
    param(
        # string to check against
        [Parameter(mandatory=$true,position=0)]
            [string]$str
    )

    $outStr=''
    foreach($c in $str.ToCharArray() ) {
        if( $c -match $rxValidPhoneCharacters ) {
            $outStr+=$c
        }
    }
    return $outStr

}
function new-eMailFromTemplate {
    param( )
    #%gn.%sn@w-files.pl
    #%gn.%mn.%sn.%se@w-files.pl
    #%gn.%sn.%se@w-files.pl
    #%gn.%mn.%sn@w-files.pl
    $mailTemplate=$cbMailTemplate.text
    $mailData=[ordered]@{}
    if($mailTemplate.Contains("%gn") -and -not [string]::IsNullOrEmpty($txtGivenName.Text) ) { 
        $gn4mail=Remove-Diacritics $txtGivenName.Text.replace(' ','')
        $mailData.add('givenName',$gn4mail) 
    }
    if($mailTemplate.Contains("%mn") -and -not [string]::IsNullOrEmpty($txtMiddleName.Text) ) { 
        $mn4mail=Remove-Diacritics $txtMiddleName.Text.replace(' ','')
        $mailData.add('middleName',$mn4mail) 
    }
    if($mailTemplate.Contains("%sn") -and -not [string]::IsNullOrEmpty($txtSurname.Text) ) { 
        $sn4mail=Remove-Diacritics $txtSurname.Text.replace(' ','')
        $mailData.add('surname',$sn4mail) 
    }
    if($mailTemplate.Contains("%se") -and -not [string]::IsNullOrEmpty($txtSurnameExt.Text) ) { 
        $se4mail=Remove-Diacritics $txtSurnameExt.Text.replace(' ','')
        $mailData.add('extSurname',$se4mail) 
    }
    if($cbUserType.text -eq 'external') {
        $strDomain="@external.w-files.pl"
    } else {
        $strDomain="@w-files.pl"
    }
    $outMail='' 
    foreach($c in (($mailData.values -join '.')+$strDomain).ToCharArray() ) {
        if($c -match $rxValidMailCharacters) {
            $outMail+=$c
        }
    }

    return $outMail.tolower()
}
function set-MailTemplateValues {
    $cbMailTemplate.Items.Clear()
    if($cbUserType.text -eq 'external') {
        $strDomain="@external.w-files.pl"
    } else {
        $strDomain="@w-files.pl"
    }
    [void]$cbMailTemplate.Items.Add("%gn.%sn$strDomain")
    [void]$cbMailTemplate.Items.Add("%gn.%mn.%sn.%se$strDomain")
    [void]$cbMailTemplate.Items.Add("%gn.%mn.%sn$strDomain")
    [void]$cbMailTemplate.Items.Add("%gn.%sn.%se$strDomain")
    $cbMailTemplate.SelectedIndex = 0
    $cbMailTemplate.refresh()
}
function new-SAMname {
    [cmdletbinding(defaultParameterSetName='default')]
    param(
        #user type to know number to generate
        [parameter(mandatory=$true,position=0)]
            [string]$uType,
        #query Azure AD instead of AD
        [parameter(ParameterSetName="AAD",mandatory=$false,position=1)]
            [switch]$useAAD,
        #query EXO 
        [parameter(ParameterSetName="EXO",mandatory=$false,position=1)]
            [switch]$useEXO,
        # number of retries for quering agains AD in case dupe has been found
        [parameter(mandatory=$false,position=2)]
            [int]$maxRetries=10
    )

    for($try=0;$try -lt $maxRetries;$try++) {
        $rndNumber="{0:D5}" -f ( Get-Random -Minimum 0 -Maximum 99999 )
        switch($uType) {
            {$_ -match "internal|archive"} {
                #convetion assumes values 0-5 but leading zero causes issues in many automations (convertion to number)
                $userDigit=Get-Random -Minimum 1 -Maximum 5
                #this was used in first version of bulk script.
                #$sam="5$rndNumber" 
                $sam="$userDigit$rndNumber"
            }
            {$_ -match "^external"} {
                $sam= "8$rndNumber"
            }
            {$_ -match "^shared"} {
                $sam= "6$rndNumber"
            }
            {$_ -match "^service"} {
                $sam= "6$rndNumber"
            }

            default {
                return -3
            }
        }

        #TESTS for existence of SAM in environment   
        write-log "[SAMGEN] ensuring user $sam does not exist in AD" -silent
        if($useAAD) {
            $test=Get-AzureADUser -SearchString "$sam@w-files.pl"
            if($test) {
                write-log -type warning -message "[SAMGEN] dupe found in AAD. duplicating object: $($test.displayName) $($test.usageLocation)" -silent
                continue
            }
        } elseif($useEXO) {
            $test=get-recipient "$sam@w-files.pl" -ErrorAction SilentlyContinue
            if($test) {
                write-log -type warning -message "[SAMGEN] dupe found in EXO. duplicating object type: $($test.RecipientTypeDetails)" 
                continue
            }

        } else {
            try {
                $test=get-aduser -Identity $sam -ErrorAction SilentlyContinue
                write-log -type warning -message "[SAMGEN] dupe found in AD. duplicating object: $($test.samaccountname) $($test.distinguishedname)" -silent
                continue
            } catch {
                #that is expected - not found
            }
        }

        write-log "[SAMGEN] ensuring user $sam does not exist in Reservations file or this run" -silent
        if($script:SAMReservations -contains $sam) {
            write-log -type error -message "[SAMGEN] dupe found in reservation list" -silent
            continue
        }

        #no dupes - add current number to reservation list to ensure it will not be generated twice and return SAM
        return $sam
    }
    return -7 #still dupe - should never enter this line
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
function new-CloudUser {
    write-log "creating $SAMAccountName..." -silent -txtBox $txtLog
    #$txtLog.text+="creating $SAMAccountName...`n"

    #other attributes pushed by 'replace'
    $OtherAttributes=@{
        msExchUsageLocation         = $msExchUsageLocation
        proxyAddresses              = @("smtp:$SAMAccountName@w-files.pl")
        mailNickname                = $SAMAccountName
        msExchRemoteRecipientType   = 1 #provision mailbox
        msExchRecipientDisplayType  = -2147483642 #remoteMailboxUser
    }
    foreach($emailAlias in $userAliases) {
        $OtherAttributes.proxyAddresses+="smtp:($emailAlias.toString())"
    }

        #add non-empty values to new-ADUser 
    $ADUser = @{
        SAMAccountName      = $SAMAccountName
        UserPrincipalName   = $UserPrincipalName
        GivenName           = $GivenName
        Surname             = $Surname
        Displayname         = $Displayname
        Name                = $Name
        EmailAddress        = $EmailAddress
        Path                = $Path
        OtherAttributes     = $OtherAttributes
    }
    switch($userType) {
        {$_ -match 'shared|archive'} {
            $ADUser.add('enabled',$false)
            $OtherAttributes.add('msExchRecipientTypeDetails','34359738368')
        }
        default {
            #user mailbox
            $ADUser.add('enabled',$true)
            $OtherAttributes.add('msExchRecipientTypeDetails','2147483648')
        }
    }     

    if($mobilePhone) { $ADUser.add('mobilePhone',$mobilePhone) }
    if($title) { $ADUser.add('title',$title) }
    if($department) { $ADUser.add('department',$department) }
    if($description) { $ADUser.add('description',$description) }
    if($manager) { $ADUser.add('manager',$manager) }

    $ADUser.add('accountPassword',(ConvertTo-SecureString -String $accountPassword -AsPlainText -Force) )
    write-log "basic AD attributes: " -silent -txtBox $txtLog
    #$txtLog.text+="basic AD attributes:`n"
    write-log $ADUser -silent -skipTimestamp -txtBox $txtLog
    #$txtLog.text+="$ADUser`n"
    write-log "password=""$accountPassword""" -silent -skipTimestamp -txtBox $txtLog
    #$txtLog.text+="password=""$accountPassword""`n"
    write-log "extended attributes:" -silent -txtBox $txtLog
    #$txtLog.text+="extended attributes:`n"
    write-log $OtherAttributes -silent -skipTimestamp -txtBox $txtLog
    #$txtLog.text+="$OtherAttributes`n"

 
    write-log "Creating user..." -txtBox $txtLog
    #$txtLog.text+="Creating user...`n"
    try {
        new-aduser @ADUser
        write-log -message 'user created' -type ok -txtBox $txtLog
        #$txtLog.text+="CREATED`n"
    } catch {
        write-log -message 'not able to create user' -type error -txtBox $txtLog
        write-log $_.Exception -type error -txtBox $txtLog
        #$txtLog.text+="ERROR CREATING USER`n"
        #$txtLog.text+="$($_.Exception)`n"
    }

    if($licenseGroup -ne 'none') {
        write-log "adding to a $licenseGroup group..." -txtBox $txtLog
        try {
            Add-ADGroupMember -Identity $licenseGroup -Members $SAMAccountName
        } catch {
            #write-log -message "error adding to $licenseGroup" -type error
            write-log $_.exception -type error -txtBox $txtLog
        }
    }

}
function search-forManager {
    param (
        [Parameter(mandatory=$false,position=0)]
            [string]$managerName
    )

    if( [string]::IsNullOrEmpty($managerName) ) { return '' }
    write-log -type info -message "trying to find manager object..."
    #try finding manager, treating input as: SAM/DN,UPN,displayname, mail
    $managerObject=$null
        #first try directly - as SAM/DN
    try {
        $managerObject=Get-ADUser $managerName -Properties * -ErrorAction SilentlyContinue 
    } catch { <#not found#> }
        #if not found - try using as UPN
    if(-not $managerObject) {
        $f="UserPrincipalName -eq `"$managerName`""
        $managerObject=Get-ADUser -filter $f -Properties * -ErrorAction SilentlyContinue
    }
        #if still not found - check displayname
    if(-not $managerObject) {
        $f="displayname -eq `"$managerName`""
        $managerObject=Get-ADUser -filter $f -Properties * -ErrorAction SilentlyContinue
    }
        #if still not found - check mail
    if(-not $managerObject) {
        $f="mail -eq `"$managerName`""
        $managerObject=Get-ADUser -filter $f -Properties * -ErrorAction SilentlyContinue
    }
        #still not found? now cast error
    if(-not $managerObject) {
        write-log "$managerName not found in AD" -silent
        return $null
    } else {
        return $managerObject
    }
}
function check-domainIsAccepted {
    param (
        [Parameter(mandatory=$false,position=0)]
            [string]$email
    )

    $emailDomain=$rxEmail.Match($email).groups['domain'].value
    if($NULL -ne $AcceptedDomainList) {
        if(-not $AcceptedDomainList.contains($emailDomain.toLower()) ) {
            write-log "$($emailDomain.toLower()) is not an accepted domain." -type error -silent
            return $false
        }
    }
    return $emailDomain
}
#endregion FUNCTIONS

#region MAINFORM
$NewUserMain = New-Object system.Windows.Forms.Form
$NewUserMain.ClientSize = New-Object System.Drawing.Point(630, 500)
$NewUserMain.text = "New User Form"
$NewUserMain.TopMost = $false
$NewUserMain.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
#$NewUserMain.KeyPreview = $true
#region Add_Application_icon
$appIconBase64="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"
$appIconBytes = [Convert]::FromBase64String($appIconBase64)
$imageStream = New-Object IO.MemoryStream($appIconBytes, 0, $appIconBytes.Length)
$imageStream.Write($appIconBytes, 0, $appIconBytes.Length);
$NewUserMain.Icon  = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $imageStream).GetHIcon())
#endregion Add_Application_icon

#region UPPER_MENU_DROPDOWNS
#menu
$menuMain = New-Object System.Windows.Forms.MenuStrip

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFile.Text = "File"
[void]$menuMain.Items.Add($menuFile)

$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit.Image = [System.IconExtractor]::Extract('shell32.dll', 10, $true)
$menuExit.ShortcutKeys = "Control, X"
$menuExit.Text = "Exit"
$menuExit.Add_Click( { $NewUserMain.Close() })
[void]$menuFile.DropDownItems.Add($menuExit)

#dropdowns
$cbCountry = New-Object system.Windows.Forms.ComboBox
$cbCountry.text = "country"
$cbCountry.width = 80
$cbCountry.height = 30
$cbCountry.location = New-Object System.Drawing.Point(10, 30)
foreach($ou in $CreationTargetOUs.keys) {
    [void]$cbCountry.Items.Add($ou)
}

$cbUserType = New-Object system.Windows.Forms.ComboBox
$cbUserType.text = "Internal"
$cbUserType.width = 100
$cbUserType.height = 30
$cbUserType.location = New-Object System.Drawing.Point(120, 30)
[void]$cbUserType.Items.Add("Internal")
[void]$cbUserType.Items.Add("External")
[void]$cbUserType.Items.Add("Shared")

$lblUPN = New-Object system.Windows.Forms.Label
$lblUPN.text = "Principal Name"
$lblUPN.AutoSize = $true
$lblUPN.width = 100
$lblUPN.height = 20
$lblUPN.location = New-Object System.Drawing.Point(330, 30)
$lblUPN.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$txtUPN = New-Object system.Windows.Forms.TextBox
$txtUPN.multiline = $false
$txtUPN.ReadOnly = $true
$txtUPN.width = 180
$txtUPN.height = 20
$txtUPN.location = New-Object System.Drawing.Point(440, 30)
$txtUPN.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$cbLicense = New-Object system.Windows.Forms.ComboBox
$cbLicense.text = "choose license"
$cbLicense.width = 180
$cbLicense.height = 30
$cbLicense.location = New-Object System.Drawing.Point(440, 55)
    [void]$cbLicense.Items.Add("none")
    [void]$cbLicense.Items.Add("o365-E1")
    [void]$cbLicense.Items.Add("o365-E1-audioconferencing")
    [void]$cbLicense.Items.Add("o365-E1-trial")
    [void]$cbLicense.Items.Add("Teams-E1-trial-audioconferencing")
    [void]$cbLicense.Items.Add("o365-E3")
    [void]$cbLicense.Items.Add("o365-E3-audioconferencing")
    [void]$cbLicense.Items.Add("o365-E5")
    [void]$cbLicense.Items.Add("o365-management")
    [void]$cbLicense.Items.Add("o365-Kiosk")
#endregion UPPER_MENU_DROPDOWNS

#region PERSONAL_INFORMATION
$gbNames = New-Object system.Windows.Forms.Groupbox
$gbNames.width = 290
$gbNames.height = 190
$gbNames.text = "Personal Information"
$gbNames.location = New-Object System.Drawing.Point(15, 60)

$txtGivenName = New-Object system.Windows.Forms.TextBox
$txtGivenName.multiline = $false
$txtGivenName.width = 160
$txtGivenName.height = 20
$txtGivenName.location = New-Object System.Drawing.Point(10, 30)

$lblGivenName = New-Object system.Windows.Forms.Label
$lblGivenName.text = "Given Name"
$lblGivenName.AutoSize = $true
$lblGivenName.width = 50
$lblGivenName.height = 20
$lblGivenName.location = New-Object System.Drawing.Point(190, 30)

$txtMiddleName = New-Object system.Windows.Forms.TextBox
$txtMiddleName.multiline = $false
$txtMiddleName.width = 160
$txtMiddleName.height = 20
$txtMiddleName.location = New-Object System.Drawing.Point(10, 60)

$lblSecondName = New-Object system.Windows.Forms.Label
$lblSecondName.text = "Middle Name"
$lblSecondName.AutoSize = $true
$lblSecondName.width = 50
$lblSecondName.height = 20
$lblSecondName.location = New-Object System.Drawing.Point(190, 60)

$txtSurname = New-Object system.Windows.Forms.TextBox
$txtSurname.multiline = $false
$txtSurname.width = 160
$txtSurname.height = 20
$txtSurname.location = New-Object System.Drawing.Point(10, 90)

$lblSurName = New-Object system.Windows.Forms.Label
$lblSurName.text = "Surname"
$lblSurName.AutoSize = $true
$lblSurName.width = 50
$lblSurName.height = 20
$lblSurName.location = New-Object System.Drawing.Point(190, 90)

$txtSurnameExt = New-Object system.Windows.Forms.TextBox
$txtSurnameExt.multiline = $false
$txtSurnameExt.width = 160
$txtSurnameExt.height = 20
$txtSurnameExt.location = New-Object System.Drawing.Point(10, 120)

$lblSurnameExt = New-Object system.Windows.Forms.Label
$lblSurnameExt.text = "Surname ext."
$lblSurnameExt.AutoSize = $true
$lblSurnameExt.width = 50
$lblSurnameExt.height = 20
$lblSurnameExt.location = New-Object System.Drawing.Point(190, 120)

$chbIncludeExtDN = New-Object System.Windows.Forms.Checkbox 
$chbIncludeExtDN.Location = New-Object System.Drawing.Size(240,150) 
$chbIncludeExtDN.Size = New-Object System.Drawing.Size(50,20)
$chbIncludeExtDN.Text = "ext."

$lblDisplayName = New-Object system.Windows.Forms.Label
$lblDisplayName.text = "Display Name be there"
$lblDisplayName.AutoSize = $true
$lblDisplayName.width = 260
$lblDisplayName.height = 20
$lblDisplayName.location = New-Object System.Drawing.Point(15, 150)
$lblDisplayName.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)

$gbNames.controls.AddRange(@(
    $txtGivenName,$lblGivenName,$txtMiddleName,$lblSecondName,
    $txtSurname,$lblSurName,$txtSurnameExt,$lblSurnameExt, 
    $lblDisplayName
))
#endregion PERSONAL_INFORMATION

#region ORG_INFO
$gbInfo = New-Object system.Windows.Forms.Groupbox
$gbInfo.width = 295
$gbInfo.height = 160
$gbInfo.text = "Oranization Information"
$gbInfo.location = New-Object System.Drawing.Point(320, 90)

$txtDepartment = New-Object system.Windows.Forms.TextBox
$txtDepartment.multiline = $false
$txtDepartment.width = 170
$txtDepartment.height = 20
$txtDepartment.location = New-Object System.Drawing.Point(10, 30)

$lblDepartment = New-Object system.Windows.Forms.Label
$lblDepartment.text = "Department"
$lblDepartment.AutoSize = $true
$lblDepartment.width = 50
$lblDepartment.height = 20
$lblDepartment.location = New-Object System.Drawing.Point(200, 30)

$txtTitle = New-Object system.Windows.Forms.TextBox
$txtTitle.multiline = $false
$txtTitle.width = 170
$txtTitle.height = 20
$txtTitle.location = New-Object System.Drawing.Point(10, 60)

$lblTitle = New-Object system.Windows.Forms.Label
$lblTitle.text = "Title"
$lblTitle.AutoSize = $true
$lblTitle.width = 50
$lblTitle.height = 20
$lblTitle.location = New-Object System.Drawing.Point(200, 60)

$txtManager = New-Object system.Windows.Forms.TextBox
$txtManager.multiline = $false
$txtManager.width = 170
$txtManager.height = 20
$txtManager.location = New-Object System.Drawing.Point(10, 90)

$lblManager = New-Object system.Windows.Forms.Label
$lblManager.text = "Manager"
$lblManager.AutoSize = $true
$lblManager.width = 50
$lblManager.height = 20
$lblManager.location = New-Object System.Drawing.Point(200, 90)

$btCheckManager = New-Object system.Windows.Forms.Button
$btCheckManager.text = '>'
$btCheckManager.width = 20
$btCheckManager.height = 20
$btCheckManager.location = New-Object System.Drawing.Point(270, 90)

$txtDescription = New-Object system.Windows.Forms.TextBox
$txtDescription.multiline = $false
$txtDescription.width = 170
$txtDescription.height = 20
$txtDescription.location = New-Object System.Drawing.Point(10, 120)

$lblDescription = New-Object system.Windows.Forms.Label
$lblDescription.text = "Description"
$lblDescription.AutoSize = $true
$lblDescription.width = 50
$lblDescription.height = 20
$lblDescription.location = New-Object System.Drawing.Point(200, 120)

$gbInfo.controls.AddRange(@(
    $txtDepartment, $txtTitle, $txtManager, $btCheckManager, $lblDepartment, 
    $lblTitle, $lblManager, $txtDescription, $lblDescription
))
#endregion ORG_INFO

#region CONTACT_INFO
$gbContact = New-Object system.Windows.Forms.Groupbox
$gbContact.width = 600
$gbContact.height = 190
$gbContact.text = "Contact Information"
$gbContact.location = New-Object System.Drawing.Point(15, 260)

$lblMailTemplate = New-Object system.Windows.Forms.Label
$lblMailTemplate.text = "Primary SMTP"
$lblMailTemplate.AutoSize = $true
$lblMailTemplate.width = 50
$lblMailTemplate.height = 20
$lblMailTemplate.location = New-Object System.Drawing.Point(10, 30)
$lblMailTemplate.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$lblMailTemplateInfo = New-Object system.Windows.Forms.Label
$lblMailTemplateInfo.text = "Choose email template:"
$lblMailTemplateInfo.AutoSize = $true
$lblMailTemplateInfo.width = 50
$lblMailTemplateInfo.height = 20
$lblMailTemplateInfo.location = New-Object System.Drawing.Point(120, 10)
$lblMailTemplateInfo.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$cbMailTemplate = New-Object system.Windows.Forms.ComboBox
$cbMailTemplate.width = 170
$cbMailTemplate.height = 20
$cbMailTemplate.location = New-Object System.Drawing.Point(120, 30)
set-MailTemplateValues

$txtMail = New-Object system.Windows.Forms.textBox
$txtMail.multiline = $false
$txtMail.text = "@w-files.pl"
$txtMail.width = 280
$txtMail.height = 20
$txtMail.location = New-Object System.Drawing.Point(10, 60)
$txtMail.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$lblMobile = New-Object system.Windows.Forms.Label
$lblMobile.text = "Mobile Phone"
$lblMobile.AutoSize = $true
$lblMobile.width = 50
$lblMobile.height = 20
$lblMobile.location = New-Object System.Drawing.Point(10, 155)
$lblMobile.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$txtMobile = New-Object system.Windows.Forms.TextBox
$txtMobile.multiline = $false
$txtMobile.width = 130
$txtMobile.height = 20
$txtMobile.location = New-Object System.Drawing.Point(120, 155)
$txtMobile.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$lblAlias = New-Object system.Windows.Forms.Label
$lblAlias.text = "Mail Aliases"
$lblAlias.AutoSize = $true
$lblAlias.width = 50
$lblAlias.height = 20
$lblAlias.location = New-Object System.Drawing.Point(310, 10)
$lblAlias.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$txtAlias = New-Object system.Windows.Forms.TextBox
$txtAlias.multiline = $false
$txtAlias.width = 220
$txtAlias.height = 20
$txtAlias.location = New-Object System.Drawing.Point(310, 30)
$txtAlias.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$btAddAlias = New-Object system.Windows.Forms.Button
$btAddAlias.text = "ADD"
$btAddAlias.width = 50
$btAddAlias.height = 20
$btAddAlias.location = New-Object System.Drawing.Point(540, 30)
$btAddAlias.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$lbAlias = New-Object system.Windows.Forms.ListBox
$lbAlias.Multicolumn = $true
$lbAlias.width = 220
$lbAlias.height = 130
$lbAlias.location = New-Object System.Drawing.Point(310, 60)

$btRemoveAlias = New-Object system.Windows.Forms.Button
$btRemoveAlias.text = "DEL"
$btRemoveAlias.width = 50
$btRemoveAlias.height = 20
$btRemoveAlias.location = New-Object System.Drawing.Point(540, 60)
$btRemoveAlias.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$gbContact.controls.AddRange(@($lblMailTemplate, $lblMailTemplateInfo, $cbMailTemplate, $txtMail, $txtMobile, $lblMobile, $lblAlias, $txtAlias, $btAddAlias, $lbAlias, $btRemoveAlias))
#endregion CONTACT_INFO

$btCheckReview = New-Object system.Windows.Forms.Button
$btCheckReview.text = 'CHECK&&REVIEW'
$btCheckReview.width = 130
$btCheckReview.height = 30
$btCheckReview.location = New-Object System.Drawing.Point(480, 460)

$NewUserMain.controls.AddRange(@($menuMain, $cbCountry, $cbUserType, $lblUPN, $txtUPN, $cbLicense, $gbNames, $gbInfo, $gbContact, $btCheckReview))
#endregion MAINFORM

#region REVIEW_FORM
$reviewForm = New-Object system.Windows.Forms.Form
$reviewForm.ClientSize = New-Object System.Drawing.Point(630, 495)
$reviewForm.text = "REVIEW USER VALUES"
$reviewForm.TopMost = $false
$reviewForm.icon = [System.Drawing.SystemIcons]::Question
$reviewForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$gbReviewInfo = New-Object system.Windows.Forms.Groupbox
$gbReviewInfo.width = 600
$gbReviewInfo.height = 475
$gbReviewInfo.text = "User attribute values"
$gbReviewInfo.location = New-Object System.Drawing.Point(15, 10)

#region OBJECT_INFO
$gbReviewObjectInfo = New-Object system.Windows.Forms.Groupbox
$gbReviewObjectInfo.width = 570
$gbReviewObjectInfo.height = 125
$gbReviewObjectInfo.text = "Object Information"
$gbReviewObjectInfo.location = New-Object System.Drawing.Point(15, 15)

$lblReviewuserType = New-Object system.Windows.Forms.Label
$lblReviewuserType.text = "user type:"
$lblReviewuserType.width = 80
$lblReviewuserType.height = 20
$lblReviewuserType.location = New-Object System.Drawing.Point(15, 15)

$lblReviewuserTypeValue = New-Object system.Windows.Forms.Label
$lblReviewuserTypeValue.AutoSize = $true
$lblReviewuserTypeValue.width = 70
$lblReviewuserTypeValue.height = 20
$lblReviewuserTypeValue.location = New-Object System.Drawing.Point(120, 15)
$lblReviewuserTypeValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewSAM = New-Object system.Windows.Forms.Label
$lblReviewSAM.text = "SAM:"
$lblReviewSAM.width = 40
$lblReviewSAM.height = 20
$lblReviewSAM.location = New-Object System.Drawing.Point(190, 15)

$lblReviewSAMValue = New-Object system.Windows.Forms.Label
$lblReviewSAMValue.AutoSize = $true
$lblReviewSAMValue.width = 70
$lblReviewSAMValue.height = 20
$lblReviewSAMValue.location = New-Object System.Drawing.Point(230, 15)
$lblReviewSAMValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewUPN = New-Object system.Windows.Forms.Label
$lblReviewUPN.text = "UPN:"
$lblReviewUPN.width = 30
$lblReviewUPN.height = 20
$lblReviewUPN.location = New-Object System.Drawing.Point(320, 15)

$lblReviewUPNValue = New-Object system.Windows.Forms.Label
$lblReviewUPNValue.AutoSize = $true
$lblReviewUPNValue.width = 200
$lblReviewUPNValue.height = 20
$lblReviewUPNValue.location = New-Object System.Drawing.Point(380, 15)
$lblReviewUPNValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)

$lblReviewOU = New-Object system.Windows.Forms.Label
$lblReviewOU.text = "Organizational Unit:"
$lblReviewOU.width = 90
$lblReviewOU.height = 25
$lblReviewOU.location = New-Object System.Drawing.Point(15, 35)

$lblReviewOUValue = New-Object system.Windows.Forms.Label
$lblReviewOUValue.AutoSize = $true
$lblReviewOUValue.width = 370
$lblReviewOUValue.height = 25
$lblReviewOUValue.location = New-Object System.Drawing.Point(120, 35)

$lblReviewPassword = New-Object system.Windows.Forms.Label
$lblReviewPassword.text = "One Time Password:"
$lblReviewPassword.width = 90
$lblReviewPassword.height = 25
$lblReviewPassword.location = New-Object System.Drawing.Point(15, 65)
#$lblReviewPassword.visible = $false

$txtReviewPasswordValue = New-Object system.Windows.Forms.TextBox
$txtReviewPasswordValue.ReadOnly = $true
$txtReviewPasswordValue.width = 150
$txtReviewPasswordValue.height = 20
$txtReviewPasswordValue.location = New-Object System.Drawing.Point(120, 65)

$lblReviewLicense = New-Object system.Windows.Forms.Label
$lblReviewLicense.text = "License:"
$lblReviewLicense.width = 90
$lblReviewLicense.height = 25
$lblReviewLicense.location = New-Object System.Drawing.Point(15, 90)

$lblReviewLicenseValue = New-Object system.Windows.Forms.Label
$lblReviewLicenseValue.text = ""
$lblReviewLicenseValue.width = 150
$lblReviewLicenseValue.height = 25
$lblReviewLicenseValue.location = New-Object System.Drawing.Point(120, 90)


$gbReviewObjectInfo.Controls.AddRange(@(
    $lblReviewuserType, $lblReviewuserTypeValue, $lblReviewSAM, $lblReviewSAMValue, 
    $lblReviewUPN, $lblReviewUPNValue, $lblReviewOU, $lblReviewOUValue, $lblReviewPassword, 
    $txtReviewPasswordValue, $lblReviewLicense, $lblReviewLicenseValue
))
#endregion OBJECT_INFO

#region PERSONAL_INFO
$gbReviewPersonalInfo = New-Object system.Windows.Forms.Groupbox
$gbReviewPersonalInfo.width = 280
$gbReviewPersonalInfo.height = 100
$gbReviewPersonalInfo.text = "Personal Information"
$gbReviewPersonalInfo.location = New-Object System.Drawing.Point(15, 145)

$lblReviewGivenName = New-Object system.Windows.Forms.Label
$lblReviewGivenName.text = "Given Name:"
$lblReviewGivenName.width = 80
$lblReviewGivenName.height = 20
$lblReviewGivenName.location = New-Object System.Drawing.Point(15, 15)

$lblReviewGivenNameValue = New-Object System.Windows.Forms.Label
$lblReviewGivenNameValue.AutoSize = $true
$lblReviewGivenNameValue.width = 185
$lblReviewGivenNameValue.height = 20
$lblReviewGivenNameValue.location = New-Object System.Drawing.Point(120, 15)
$lblReviewGivenNameValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)

$lblReviewMiddleName = New-Object system.Windows.Forms.Label
$lblReviewMiddleName.text = "Middle Name:"
$lblReviewMiddleName.width = 80
$lblReviewMiddleName.height = 20
$lblReviewMiddleName.location = New-Object System.Drawing.Point(300, 15)

$lblReviewMiddleNameValue = New-Object System.Windows.Forms.Label
$lblReviewMiddleNameValue.AutoSize = $true
$lblReviewMiddleNameValue.width = 185
$lblReviewMiddleNameValue.height = 20
$lblReviewMiddleNameValue.location = New-Object System.Drawing.Point(390, 15)

$lblReviewSurname = New-Object system.Windows.Forms.Label
$lblReviewSurname.text = "Surname:"
$lblReviewSurname.width = 80
$lblReviewSurname.height = 20
$lblReviewSurname.location = New-Object System.Drawing.Point(15, 35)

$lblReviewSurnameValue = New-Object System.Windows.Forms.Label
$lblReviewSurnameValue.AutoSize = $true
$lblReviewSurnameValue.width = 200
$lblReviewSurnameValue.height = 20
$lblReviewSurnameValue.location = New-Object System.Drawing.Point(120, 35)
$lblReviewSurnameValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)

$lblReviewSurnameExt = New-Object system.Windows.Forms.Label
$lblReviewSurnameExt.text = "Sn extention:"
$lblReviewSurnameExt.width = 80
$lblReviewSurnameExt.height = 20
$lblReviewSurnameExt.location = New-Object System.Drawing.Point(300, 35)

$lblReviewSurnameExtValue = New-Object System.Windows.Forms.Label
$lblReviewSurnameExtValue.AutoSize = $true
$lblReviewSurnameExtValue.width = 200
$lblReviewSurnameExtValue.height = 20
$lblReviewSurnameExtValue.location = New-Object System.Drawing.Point(390, 35)

$lblReviewDisplayName = New-Object system.Windows.Forms.Label
$lblReviewDisplayName.text = "Display Name:"
$lblReviewDisplayName.width = 80
$lblReviewDisplayName.height = 20
$lblReviewDisplayName.location = New-Object System.Drawing.Point(15, 55)

$lblReviewDisplayNameValue = New-Object System.Windows.Forms.Label
$lblReviewDisplayNameValue.AutoSize = $true
$lblReviewDisplayNameValue.width = 200
$lblReviewDisplayNameValue.height = 20
$lblReviewDisplayNameValue.location = New-Object System.Drawing.Point(120, 55)
$lblReviewDisplayNameValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)

$gbReviewPersonalInfo.Controls.AddRange( @(
    $lblReviewGivenName, $lblReviewGivenNameValue, $lblReviewSurname, $lblReviewSurnameValue,
    #$lblReviewMiddleName, $lblReviewMiddleNameValue, $lblReviewSurnameExt, $lblReviewSurnameExtValue,
    $lblReviewDisplayName, $lblReviewDisplayNameValue
))
#endregion PERSONAL_INFO

#region ORG_INFO
$gbReviewOrgInfo = New-Object system.Windows.Forms.Groupbox
$gbReviewOrgInfo.width = 285
$gbReviewOrgInfo.height = 100
$gbReviewOrgInfo.text = "Organizational Information"
$gbReviewOrgInfo.location = New-Object System.Drawing.Point(300, 145)

$lblReviewDepartment = New-Object system.Windows.Forms.Label
$lblReviewDepartment.text = "Department:"
$lblReviewDepartment.AutoSize = $true
$lblReviewDepartment.width = 100
$lblReviewDepartment.height = 20
$lblReviewDepartment.location = New-Object System.Drawing.Point(15, 15)

$lblReviewDepartmentValue = New-Object System.Windows.Forms.Label
$lblReviewDepartmentValue.width = 140
$lblReviewDepartmentValue.height = 20
$lblReviewDepartmentValue.location = New-Object System.Drawing.Point(110, 15)
$lblReviewDepartmentValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewTitle = New-Object system.Windows.Forms.Label
$lblReviewTitle.text = "Title:"
$lblReviewTitle.AutoSize = $true
$lblReviewTitle.width = 100
$lblReviewTitle.height = 20
$lblReviewTitle.location = New-Object System.Drawing.Point(15, 35)

$lblReviewTitleValue = New-Object System.Windows.Forms.Label
$lblReviewTitleValue.width = 140
$lblReviewTitleValue.height = 20
$lblReviewTitleValue.location = New-Object System.Drawing.Point(110, 35)
$lblReviewTitleValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewDescription = New-Object system.Windows.Forms.Label
$lblReviewDescription.text = "Description:"
$lblReviewDescription.AutoSize = $true
$lblReviewDescription.width = 100
$lblReviewDescription.height = 20
$lblReviewDescription.location = New-Object System.Drawing.Point(15, 55)

$lblReviewDescriptionValue = New-Object System.Windows.Forms.Label
$lblReviewDescriptionValue.width = 140
$lblReviewDescriptionValue.height = 20
$lblReviewDescriptionValue.location = New-Object System.Drawing.Point(110, 55)
$lblReviewDescriptionValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewCountry = New-Object system.Windows.Forms.Label
$lblReviewCountry.text = "Country:"
$lblReviewCountry.AutoSize = $true
$lblReviewCountry.width = 100
$lblReviewCountry.height = 20
$lblReviewCountry.location = New-Object System.Drawing.Point(15, 75)

$lblReviewCountryValue = New-Object System.Windows.Forms.Label
$lblReviewCountryValue.width = 140
$lblReviewCountryValue.height = 20
$lblReviewCountryValue.location = New-Object System.Drawing.Point(110, 75)
$lblReviewCountryValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$gbReviewOrgInfo.Controls.AddRange(@(
    $lblReviewDepartment, $lblReviewDepartmentValue, $lblReviewDescription, $lblReviewDescriptionValue
    $lblReviewTitle, $lblReviewTitleValue, 
    $lblReviewCountry, $lblReviewCountryValue
))
#endregion ORG_INFO

#region CONTACT_INFO
$gbReviewContactInfo = New-Object system.Windows.Forms.Groupbox
$gbReviewContactInfo.width = 570
$gbReviewContactInfo.height = 180
$gbReviewContactInfo.text = "User Contact Information"
$gbReviewContactInfo.location = New-Object System.Drawing.Point(15, 250)

$btNowaNazwak = New-Object System.Windows.Forms.Button
$btNowaNazwak.Location = New-Object System.Drawing.Size(30,30) 
$btNowaNazwak.Size = New-Object System.Drawing.Size(100,20)
$btNowaNazwak.Text = "TEST"

$lblReviewManager = New-Object system.Windows.Forms.Label
$lblReviewManager.text = "Manager:"
$lblReviewManager.AutoSize = $true
$lblReviewNowa Nazwa .width = 100
$lblReviewNowa Nazwa .height = 20
$lblReviewNowa Nazwa .location = New-Object System.Drawing.Point(15, 15)

$lblReviewManagerValue = New-Object System.Windows.Forms.Label
$lblReviewManagerValue.width = 400
$lblReviewManagerValue.height = 20
$lblReviewManagerValue.location = New-Object System.Drawing.Point(110, 15)
$lblReviewManagerValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 7, [System.Drawing.FontStyle]::Bold)

$lblReviewMobile = New-Object system.Windows.Forms.Label
$lblReviewMobile.text = "Mobile Phone:"
$lblReviewMobile.AutoSize = $true
$lblReviewMobile.width = 100
$lblReviewMobile.height = 20
$lblReviewMobile.location = New-Object System.Drawing.Point(15, 35)

$lblReviewMobileValue = New-Object System.Windows.Forms.Label
$lblReviewMobileValue.width = 200
$lblReviewMobileValue.height = 20
$lblReviewMobileValue.location = New-Object System.Drawing.Point(110, 35)
$lblReviewMobileValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewEmail = New-Object system.Windows.Forms.Label
$lblReviewEmail.text = "Primary eMail:"
$lblReviewEmail.AutoSize = $true
$lblReviewEmail.width = 100
$lblReviewEmail.height = 20
$lblReviewEmail.location = New-Object System.Drawing.Point(15, 55)

$lblReviewEmailValue = New-Object System.Windows.Forms.Label
$lblReviewEmailValue.width = 400
$lblReviewEmailValue.height = 20
$lblReviewEmailValue.location = New-Object System.Drawing.Point(110, 55)
$lblReviewEmailValue.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)

$lblReviewAliases = New-Object system.Windows.Forms.Label
$lblReviewAliases.text = "Aliases:"
$lblReviewAliases.AutoSize = $true
$lblReviewAliases.width = 70
$lblReviewAliases.height = 20
$lblReviewAliases.location = New-Object System.Drawing.Point(15, 75)

$lbReviewAliasesValue = New-Object system.Windows.Forms.ListBox
$lbReviewAliasesValue.Multicolumn = $true
$lbReviewAliasesValue.width = 175
$lbReviewAliasesValue.height = 100
$lbReviewAliasesValue.location = New-Object System.Drawing.Point(110, 75)

$gbReviewContactInfo.Controls.AddRange(@(
    $lblReviewEmail, $lblReviewEmailValue, $lblReviewMobile, $lblReviewMobileValue, 
    $lblReviewAliases, $lbReviewAliasesValue, $lblReviewManager, $lblReviewManagerValue
))
#endregion CONTACT_INFO

#region BOTTOM_BUTTONS
$btReviewContinue = New-Object system.Windows.Forms.Button
$btReviewContinue.text = 'CREATE >>>'
$btReviewContinue.width = 130
$btReviewContinue.height = 30
$btReviewContinue.location = New-Object System.Drawing.Point(460, 435)
$btReviewContinue.font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$btReviewBack = New-Object system.Windows.Forms.Button
$btReviewBack.text = '<<< BACK'
$btReviewBack.width = 130
$btReviewBack.height = 30
$btReviewBack.location = New-Object System.Drawing.Point(10, 435)
$btReviewBack.font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
#endregion BOTTOM_BUTTONS
$gbReviewInfo.controls.AddRange( @(
    $gbReviewObjectInfo, $gbReviewPersonalInfo, $gbReviewContactInfo, $gbReviewOrgInfo
    $btReviewBack, $btReviewContinue
) )

$reviewForm.Controls.AddRange( @($gbReviewInfo) )
#endregion REVIEW_FORM

#region CREATE_LOG_FORM
$LogForm = New-Object system.Windows.Forms.Form
$LogForm.ClientSize = New-Object System.Drawing.Point(630, 470)
$LogForm.text = "CREATING USER OBJECT..."
$LogForm.TopMost = $false
$LogForm.icon = [System.Drawing.SystemIcons]::Exclamation
$LogForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$gbLog = New-Object system.Windows.Forms.Groupbox
$gbLog.width = 600
$gbLog.height = 450
$gbLog.text = "LOG:"
$gbLog.location = New-Object System.Drawing.Point(15, 10)

$txtLog = New-Object system.Windows.Forms.textBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.WordWrap = $true
$txtLog.width = 570
$txtLog.height = 380
$txtLog.location = New-Object System.Drawing.Point(15, 15)

$btFinish = New-Object system.Windows.Forms.Button
$btFinish.text = 'FINISHED'
$btFinish.width = 130
$btFinish.height = 30
$btFinish.location = New-Object System.Drawing.Point(460, 410)
$btFinish.font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

$gbLog.Controls.AddRange(@(
    $txtLog, $btFinish
))

$LogForm.Controls.AddRange(@(
    $gbLog
))
#endregion CREATE_LOG_FORM

#region INTERFACE_MAIN_FUNCTIONS
$NewUserMain.add_Load({
    $NewUserMain.BringToFront()
    set-MailTemplateValues
})
$NewUserMain.add_Closing({
    [System.Windows.Forms.Application]::Exit()
})
$cbUserType.add_SelectedIndexChanged({
    if($USEONLINE) {
        $txtUPN.Text=(new-SAMname -uType ($cbUserType.Text) -useEXO)+'@w-files.pl'
    } else {
        $txtUPN.Text=(new-SAMname -uType ($cbUserType.Text) )+'@w-files.pl'
    }
    if($cbUserType.text -eq 'external') {
        $gbNames.Controls.Add($chbIncludeExtDN)
    } else {
        if( $gbNames.Controls.Contains($chbIncludeExtDN) ) {
            $gbNames.Controls.Remove($chbIncludeExtDN)
        }
    }
    set-MailTemplateValues
})
$cbMailTemplate.add_SelectedIndexChanged({
    $txtMail.Text=new-eMailFromTemplate
})
$chbIncludeExtDN.add_CheckStateChanged({
    if($chbIncludeExtDN.checked) {
         $lblDisplayName.text+=' external'
    } else {
        $lblDisplayName.text=$lblDisplayName.text.replace(' external','')
    }
})

$btCheckManager.add_Click({
    $manago=search-forManager -managerName $txtManager.Text
    if($NUll -eq $manago) {
        [System.Windows.Forms.MessageBox]::show($this,"$($txtManager.Text) not found.",'NOT FOUND','OK',[System.Windows.Forms.MessageBoxIcon]::Error )
    } else {
        [System.Windows.Forms.MessageBox]::show($this,"$($txtManager.Text) found:`nSAM: $($manago.samaccountname)`ndisplayName: $($manago.name)",'NOT FOUND','OK',[System.Windows.Forms.MessageBoxIcon]::Information )
        $txtManager.Text=$manago.samaccountname
    }
})
$btCheckReview.add_Click({
    if([string]::IsNullOrEmpty($txtGivenName.Text) -and [string]::IsNullOrEmpty($txtSurname.Text) -and $cbUserType.Text -eq 'shared' ) {
        [System.Windows.Forms.MessageBox]::show($this,"Givenname and Surname can't be both empty.",'SYNTAX','OK',[System.Windows.Forms.MessageBoxIcon]::Error )
        return
    }
    if( ([string]::IsNullOrEmpty($txtGivenName.Text) -or [string]::IsNullOrEmpty($txtSurname.Text)) -and ($cbUserType.Text -eq 'internal' -or $cbUserType.Text -eq 'external') ) {
        [System.Windows.Forms.MessageBox]::show($this,"You need to provide Givenname AND Surname for the user.",'SYNTAX','OK',[System.Windows.Forms.MessageBoxIcon]::Error )
        return
    }
    if($cbCountry.Text -eq 'country') {
        [System.Windows.Forms.MessageBox]::show($this,"You must choose user country",'SYNTAX','OK',[System.Windows.Forms.MessageBoxIcon]::Error ) 
        return
    }
    if($cbLicense.text -eq 'choose license') {
        [System.Windows.Forms.MessageBox]::show($this,"You must choose a license",'SYNTAX','OK',[System.Windows.Forms.MessageBoxIcon]::Error ) 
        return
    }
    #write-host 'review goes here'
    $NewUserMain.Cursor=[System.Windows.Forms.Cursor]::WaitCursor
    $reviewForm.ShowDialog()
})

$btAddAlias.add_Click({
    if($txtAlias.Text -notmatch $rxEmail) {
        [System.Windows.Forms.MessageBox]::show($this,"$($txtAlias.Text) doesn't seem a valid email.",'wrong email value','OK',[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    $emailDomain = check-domainIsAccepted -email $txtAlias.Text
    if(-not $emailDomain) {
        [System.Windows.Forms.MessageBox]::show($this,"$($txtAlias.Text) can't be added as `nit is not an Accepted Domain.",'alias from unknown domain','OK',[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $txtAlias.Text = Remove-Diacritics -src $txtAlias.Text
    $lbAlias.Items.Add($txtAlias.Text)
    $txtAlias.Text=''
})
$btRemoveAlias.add_Click({
    $lbAlias.Items.Remove($lbAlias.SelectedItem)
})
$txtAlias.add_KeyPress({
    if( $txtAlias.text) {
        $cursor=$txtAlias.SelectionStart
        $txtAlias.text=get-validMailString $txtAlias.text
        $txtAlias.Select($cursor, 0);
    }
})
$txtMobile.add_KeyPress({
    if($txtMobile.text) {
        $cursor=$txtMobile.SelectionStart
        $txtMobile.text=get-validMobileNumber $txtMobile.text
        $txtMobile.Select($cursor, 0);
    }
})
$txtGivenName.add_KeyPress({
    if($txtGivenName.text) {
        $cursor=$txtGivenName.SelectionStart
        $txtGivenName.text=get-validNameString $txtGivenName.text
        $txtGivenName.Select($cursor, 0);
    }
})
$txtGivenName.add_KeyUp({
    $txtMail.Text=new-eMailFromTemplate
    $lblDisplayName.text = new-displayName
})
$txtMiddleName.add_KeyPress({
    if($txtMiddleName.text) {
        $cursor=$txtMiddleName.SelectionStart
        $txtMiddleName.text=get-validNameString $txtMiddleName.text
        $txtMiddleName.Select($cursor, 0);
    }
})
$txtMiddleName.add_KeyUp({
    $txtMail.Text=new-eMailFromTemplate
    $lblDisplayName.text = new-displayName
})
$txtSurname.add_KeyPress({
    if($txtSurname.text) {
        $cursor=$txtSurname.SelectionStart
        $txtSurname.text=get-validNameString $txtSurname.text
        $txtSurname.Select($cursor, 0);
    }
})
$txtSurname.add_KeyUp({
    $txtMail.Text=new-eMailFromTemplate
    $lblDisplayName.text = new-displayName
})
$txtSurnameExt.add_KeyPress({
    if($txtSurnameExt.text) {
        $cursor=$txtSurnameExt.SelectionStart
        $txtSurnameExt.text=get-validNameString $txtSurnameExt.text
        $txtSurnameExt.Select($cursor, 0);
    }
})
$txtSurnameExt.add_KeyUp({
    $txtMail.Text=new-eMailFromTemplate
    $lblDisplayName.text = new-displayName
})
#endregion INTERFACE_MAIN_FUNCTIONS

#region INTERFACE_REVEW_FUNCTIONS
$reviewForm.add_Load({
    #$reviewForm.Show()
    #object
    $lblReviewuserTypeValue.Text = $cbUserType.Text
    $lblReviewUPNValue.text = $txtUPN.text
    $lblReviewUPNValue.Text -match "(?<sam>[\w\d_-]+)@w-files.pl"
    $lblReviewSAMValue.text = $Matches['sam']
    $lblReviewOUValue.Text = $CreationTargetOUs[$cbCountry.Text]
    $txtReviewPasswordValue.text = new-RandomPassword #-length 20
    if($cbLicense.text -eq 'none') {
        $lblReviewLicenseValue.text = "none"
    } else {
        $lblReviewLicenseValue.text = "Licenses-MS-"+$cbCountry.Text+'-'+$cbLicense.text
    }
    #personal
    $lblReviewGivenNameValue.text = $txtGivenName.text.trim()
    $lblReviewMiddleNameValue.text = $txtMiddleName.text.trim()
    $lblReviewSurnameValue.text = $txtSurname.text.trim()
    $lblReviewSurnameExtValue.text = $txtSurnameExt.text.trim()
    $lblReviewDisplayNameValue.text = $lblDisplayName.Text
    #contact
    if($EXOConnection) {
        $testMail=check-domainIsAccepted $txtMail.text
        if(-not $testMail) {
            $lblReviewEmailValue.ForeColor = 'red'
            $lblReviewEmailValue.Text='<not accepted domain>'+$txtMail.text
            $gbReviewInfo.controls.remove($btReviewContinue)
        } else {
            if($USEONLINE) {
                $testMail=check-dupe -email $txtMail.text -useEXO
            } else {
                $testMail=check-dupe -email $txtMail.text
            }
            if([string]::IsNullOrEmpty($testMail) ) {
                $lblReviewEmailValue.ForeColor = 'black' 
                $lblReviewEmailValue.Text = $txtMail.text
                $gbReviewInfo.controls.Add($btReviewContinue)
            } elseif($testMail -eq -1) {
                    $lblReviewEmailValue.ForeColor = 'red' 
                    $lblReviewEmailValue.Text = "<invalid>$($txtMail.text)"
                    $gbReviewInfo.controls.remove($btReviewContinue)
            } else {
                    $lblReviewEmailValue.ForeColor = 'red' 
                    $lblReviewEmailValue.Text = "<dupped>$($txtMail.text)"
                    $gbReviewInfo.controls.remove($btReviewContinue)
            }
        }
    } else {
        $lblReviewEmailValue.ForeColor = 'Yellow'
        $lblReviewEmailValue.Text = $txtMail.text
        $gbReviewInfo.controls.Add($btReviewContinue)
    }
    $lblReviewMobileValue.Text = $txtMobile.Text
    $lbReviewAliasesValue.Items.Clear()
    foreach($alias in $lbAlias.Items) {
        if($EXOConnection) {
            if($USEONLINE) {
                $testMail=check-dupe -email $alias.toString() -useEXO
            } else {
                $testMail=check-dupe -email $alias.toString() 
            }
            if([string]::IsNullOrEmpty($testMail) ) {
                $lbReviewAliasesValue.Items.add($alias)
            } elseif($testMail -eq -1) {
                $lbReviewAliasesValue.Items.add("<invalid>$($alias)")
                $gbReviewInfo.controls.remove($btReviewContinue)
            } else {
                $lbReviewAliasesValue.Items.add("<dupped>$($alias)")
                $gbReviewInfo.controls.remove($btReviewContinue)
            }
        } else {
            $lbReviewAliasesValue.Items.Add($alias)
        }
    
    }
    #org
    $lblReviewDepartmentValue.text = $txtDepartment.text.trim()
    $lblReviewDescriptionValue.text = $txtDescription.Text.trim()
    $lblReviewTitleValue.text = $txtTitle.Text.trim()
    $lblReviewCountryValue.text = $cbCountry.Text
    $manago=search-forManager -managerName $txtManager.text
    if($NULL -eq $manago) {
        $lblReviewManagerValue.text='<not found>'
        $lblReviewManagerValue.ForeColor = 'Red'
    } else {
        $lblReviewManagerValue.text = $manago
        $lblReviewManagerValue.ForeColor = 'Black'
    }

    $txtReviewPasswordValue.SelectionStart = 0;  
    $txtReviewPasswordValue.SelectionLength = $txtReviewPasswordValue.Text.Length; 
    $txtReviewPasswordValue.focus()
    $reviewForm.Cursor=[System.Windows.Forms.Cursor]::Default
})
$btReviewContinue.add_Click({
    $LogForm.ShowDialog()
})
$btReviewBack.add_Click({
    $reviewForm.Close()
})

#endregion INTERFACE_REVEW_FUNCTIONS

#region INTERFACE_LOG_FUNCTIONS
$LogForm.add_Load({
    $logForm.Show()
    $logForm.Activate()
    #pass variables to form...
    $userType = $lblReviewuserTypeValue.text
    $msExchUsageLocation = $lblReviewCountryValue.text
    $userAliases = $lbReviewAliasesValue.Items
    $SAMAccountName = $lblReviewSAMValue.Text
    $UserPrincipalName = $lblReviewUPNValue.Text
    $GivenName = $lblReviewGivenNameValue.Text
    $Surname = $lblReviewSurnameValue.Text
    $Displayname = $lblReviewDisplayNameValue.Text
    $Name = $lblReviewDisplayNameValue.Text
    $EmailAddress = $lblReviewEmailValue.Text
    $Path = $lblReviewOUValue.Text
    $mobilePhone = $lblReviewMobileValue.Text
    $title = $lblReviewTitleValue.text
    $department = $lblReviewDepartmentValue.Text
    $description = $lblReviewDescriptionValue.Text
    $licenseGroup = $lblReviewLicenseValue.Text
    if($lblReviewManagerValue.Text -ne '<not found>') {
        $manager = $lblReviewManagerValue.Text
    }
    $accountPassword = $txtReviewPasswordValue.Text
    #...and create user
    new-CloudUser

})
$logForm.add_Closing({
    $LogForm.Close()
    $reviewForm.Close()
    $NewUserMain.Close()
})
$btFinish.add_Click({
    #$LogForm.Close()
    #$reviewForm.Close()
    #$NewUserMain.Close()
    $NewUserMain.Dispose()
    $reviewForm.Dispose()
    $LogForm.Dispose()
})
#endregion INTERFACE_LOG_FUNCTIONS

####################################################
#                       BODY                       #
####################################################
start-Logging
$EXOConnection=$true
if($USEONLINE) { 
    if(-not (get-ExchangeConnectionStatus)) {
        write-Log "EXO connection required to check mail duplicates. `
        continuing without connection may lead to synchronization errors" -type error
        $EXOConnection=$false
    }
} else {
    if(-not (get-ExchangeConnectionStatus -ExType OnPrem)) {
        write-Log "No Exchange connection which is required. `
        continuing without connection may lead to synchronization errors" -type error
        $EXOConnection=$false
    }
}
if($USEONLINE) {
    #check if you are connected to proper tenant!
    $acceptedDomain=Get-AcceptedDomain
    $AcceptedDomainList = $acceptedDomain|Select-Object -ExpandProperty name
    $currentTenantDomain=($acceptedDomain|? initialdomain -eq $true).DomainName
    if( $currentTenantDomain -ne $validateTenantDomainName ) {
        write-log "you are connected to $currentTenantDomain and i was expecting $validateTenantDomainName." -type error
        exit -9
    }
    $UPN=new-SAMname -uType internal -useEXO #-useAAD
} else {
    $UPN=new-SAMname -uType internal
}
$txtUPN.Text=$UPN+'@w-files.pl'
[void]$NewUserMain.ShowDialog()

#i have no idea why... but on Terminal environment application do not run again.
#adding below lines, although seems idiotic and sensless - fixes the problem /:
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)

[System.GC]::Collect() 

write-log 'done.' -type ok
