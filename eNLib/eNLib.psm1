<#
.SYNOPSIS
    eN's support functions library.
.DESCRIPTION
    most commonly required functions such as better logging support for scripts, forking
    information to screen and file, enchanced GUI controls such as input box, most common
    connection accelerators.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 210205
    changes
        - 210205 write-log fixes
        - 210202 tuned write-log and start-logging, fixes and logical separation. v0.9
        - 201018 initialize 
#>

#################################################### GENERAL
$logFile=''

function start-Logging {
    <#
    .SYNOPSIS
        initilizes log file under $logFile variable for write-log function.
    .DESCRIPTION
        all scripts require logging mechinism. write-log function forking each output to screen and to logfile
        is a most common function i use in my scripts. in order to simplify $logFile variable - this function
        initilized environment for write-log function. 
    .EXAMPLE
        start-Logging

        simply initializes the log file with generic name, saved in 'Logs' subfolder under script run path. 
    .EXAMPLE
        start-Logging -logFileName c:\temp\myLogs\somelog.log

        initializes the log file as c:\temp\myLogs\somelog.log .
    .EXAMPLE
        start-Logging -userProfilePath

        initializes the log file in Logs subfolder uder user profile path
    .INPUTS
        None.
    .OUTPUTS
        log file under $logFile variable.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210205
        changes:
            - 210205 fixes to logfilename initialization
            - 210203 added 'logFolder' and proper log initilization when called indirectly
            - 210127 v1
            - 201018 initialize
    #>
    [CmdletBinding(DefaultParameterSetName='FilePath')]
    param(
        # full name for custom log file. 
        [Parameter(ParameterSetName='FilePath',mandatory=$false,position=0)]
            [string]$logFileName,
        #create log in profile folder rather than script run path
        [Parameter(ParameterSetName='userProfile',mandatory=$false,position=0)]
            [alias('useProfile')]
            [switch]$userProfilePath,
        #similar to logFileName, but takes folder only and log file name is generic.
        [Parameter(ParameterSetName='Folder',mandatory=$false,position=0)]
            [string]$logFolder
    )

    #check if not run outside script
    #if( $scriptRun.commandType -ne 'ExternalScript' ) {
    #    write-host "don't run this function outside script" -ForegroundColor Red
    #    remove-module -name eNLib
    #    return $null
    #}
    $scriptBaseName = ([System.IO.FileInfo]$($MyInvocation.PSCommandPath)).basename
    if([string]::isNullOrEmpty($scriptBaseName) ) {
        $scriptBaseName = 'console'
    }
    switch($PSCmdlet.ParameterSetName) {
        'userProfile' {
            $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
            $script:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        'Folder' {
            $script:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        default {
            #by default 'filepath' is used and empty 
            if ( [string]::IsNullOrEmpty($logFileName) ) {
                $logFolder="$($MyInvocation.PSScriptRoot)\Logs"
                $script:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)   
            } else {
                $logFolder = Split-Path $logFileName -Parent
                if([string]::isNullOrEmpty($logFolder) ) {
                    $logFolder = $MyInvocation.PSScriptRoot
                }
                $logFile = Split-Path $logFileName -Leaf
                if( test-path $logFile -PathType Container ) {
                    write-host "$logFileName seems to be an existing folder. use 'logFolder' parameter or change log name. quitting." -ForegroundColor Red
                    exit -1
                }
                $script:logFile = "$logFolder\$logFile"
            }
        }
    }

    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder|Out-Null
            write-host "$LogFolder created."
        } catch {
            write-error $_.exception
            exit -2
        }
    }
    write-Log "*logging initiated $(get-date) in $($script:logFile)" -skipTimestamp #-silent
    write-Log "*script parameters:" -silent -skipTimestamp
    if($script:PSBoundParameters.count -gt 0) {
        write-log $script:PSBoundParameters -silent -skipTimestamp
    } else {
        write-log "<none>" -silent -skipTimestamp
    }
    write-log "***************************************************" -silent -skipTimestamp
}
#Export-ModuleMember -Function start-Logging

function write-log {
    <#
    .SYNOPSIS
        replacement for write-host, forking information to a log file and screen.
    .DESCRIPTION
        automates forking of output on two different endpoints - on the host, using write-host
        and to the file, appening its content.
        write-log converts everything to a string, so you can use it for virtually any type of 
        variable. additionaly it adds timestamp, message type header and color (on host).

        information is written to a $logFile - you must initialize the value with 'start-Logging' 
        or configure it manually.

    .EXAMPLE
        .\write-log "all is fine"

        output 'all is fine' on the screen and to the log file.
    .EXAMPLE
        .\write-log -message "trees are green" -type ok

        shows 'trees are green' in Green colour, and send text to a log file.
    .EXAMPLE
        $someObject=get-process
        .\write-log -message $someObject -type info -noTimestamp -silent

        outputs processes object to the log file as -silent disables screen output. it will lack
        timestamp in a message header but will contain '[INFO]' block.
    .INPUTS
        None.
    .OUTPUTS
        text log file under $logFile
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210203
        changes:
            - 210205 fix when run directly from console, init fixes
            - 210203 properly initiating log with new start-logging, when called indirectly
            - 210127 v1
            - 201018 initialize
    #>
    
    param(
        #message to display - can be an object
        [parameter(mandatory=$false,position=0)]
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

    #if function is called without pre-initialize with start-logging, run it to initialize log.
    if( [string]::isNullOrEmpty($script:logFile) ){
        #these need to be calculated here, as $myinvocation context changes giving library name instead of script
        if( [string]::isNullOrEmpty($MyInvocation.PSCommandPath) ) { #it's run directly from console.
            $scriptBaseName = 'console'
        } else {
            $scriptBaseName = ([System.IO.FileInfo]$($MyInvocation.PSCommandPath)).basename 
        }
        if([string]::isNullOrEmpty($MyInvocation.PSScriptRoot) ) { #it's run directly from console
            $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
        } else {
            $logFolder = "$($MyInvocation.PSScriptRoot)\Logs"
        }
        $logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        start-Logging -logFileName $logFile
    }
    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {$message=''}
    $message=($message|out-String).trim() 

    try {
        $finalMessageString=@()
        if(-not $skipTimestamp) {
            $finalMessageString += "$(Get-Date -Format "hh:mm:ss>") "
        }
        if(-not [string]::IsNullOrEmpty( $type) ) { 
            $finalMessageString += $type.ToUpper()+": " 
        }
        $finalMessageString += $message
        $message=$finalMessageString -join ''
        Add-Content -Path $script:logFile -Value $message
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
        $_.exception
    }      
}
#Export-ModuleMember -Function write-Log

function new-RandomPassword {
    <#
    .SYNOPSIS
        generate random password with given ranges
    .DESCRIPTION
        #TODO
    .EXAMPLE
        #TODO
    .INPUTS
        None.
    .OUTPUTS
        string of random characters
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210127
            last changes
            - 210127 initialized
    #>
    
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
#Export-ModuleMember -Function new-randomPassword

#################################################### PowerShell GUI
function get-AnswerBox {
    <#
    .SYNOPSIS
        win32 forms message box to get YES/NO input from user
    .DESCRIPTION
        replacement for simple messageBox giving option to customize buttons, giving option
        to add some additional information
    .EXAMPLE
        $response =  get-answerBox -OKButtonText 'YES' -CancelButtonText 'NO' -info 'choose your answer' -detailedInfo 'do you find this function useful?'
        if($response) {
            write-host 'thank you!'
        }
    .INPUTS
        None.
    .OUTPUTS
        true/false
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210127
            last changes
            - 210127 module
            - 210110 initialized
        
        TO|DO
         - icon
         - docked layouts
    #>
    
    param(
        [string]$OKButtonText = "OK",
        [string]$CancelButtonText = "Cancel",
        [string]$info = "Which option?",
        [string]$detailedInfo = "What is your choice:"
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $info
    $form.Size = New-Object System.Drawing.Size(300,120)
    $form.StartPosition = 'CenterScreen'
    $form.Icon = [System.Drawing.SystemIcons]::Question
    $form.Topmost = $true
   
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(65,50)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = $OKButtonText
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(160,50)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = $CancelButtonText
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,30)
    $label.Text = $detailedInfo
    $form.Controls.Add($label)
   
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $true
    } 
    return $false
}
#Export-ModuleMember -Function get-AnswerBox

function get-valueFromInputBox {
    <#
    .SYNOPSIS
        simple input message box function for PS GUI scripts
    .EXAMPLE
        $response = get-valueFromInputBox -title 'WARNING!' -text "type 'YES' to continue" -type Warning
        if($null -eq $response) {
            'cancelled'
            exit 0
        }
        if($response -ne 'YES') {
            'not correct. quitting'
            exit -1
        } else {
            "you agreed, let's continue"
        }
        write-host 'code to execute here'
    .INPUTS
        None.
    .OUTPUTS
        User Input or $null for cancel
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210113
            last changes
            - 210113 initialized
        
        TO|DO
        - docked layout
    #>
    
    param(
        [parameter(mandatory=$false,position=0)]
            [string]$title='input',
        [parameter(mandatory=$false,position=1)]
            [string]$text='put your input',
        [parameter(mandatory=$false,position=2)]
            [validateSet('Asterisk','Error','Exclamation','Hand','Information','None','Question','Stop','Warning')]
            [string]$type='Question'
    )
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $promptWindowForm = New-Object system.Windows.Forms.Form
    $promptWindowForm.ClientSize = '250,100'
    $promptWindowForm.text = $title
    $promptWindowForm.BackColor = "#ffffff"
    $promptWindowForm.AutoSize = $true
    $promptWindowForm.StartPosition = 'CenterScreen'
    $promptWindowForm.Icon = [System.Drawing.SystemIcons]::$type
    $promptWindowForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

    $lblPromptInfo = New-Object System.Windows.Forms.Label 
    $lblPromptInfo.Location = New-Object System.Drawing.Size(10,5) 
    $lblPromptInfo.Size = New-Object System.Drawing.Size(230,40)
    $lblPromptInfo.Text = $text

    $txtUserInput = New-Object system.Windows.Forms.TextBox
    $txtUserInput.multiline = $false
    $txtUserInput.ReadOnly = $false
    $txtUserInput.width = 230
    $txtUserInput.height = 25
    $txtUserInput.location = New-Object System.Drawing.Point(10, 50)

    $btOK = New-Object System.Windows.Forms.Button
    $btOK.Location = New-Object System.Drawing.Size(30,80) 
    $btOK.Size = New-Object System.Drawing.Size(70,20)
    $btOK.ForeColor = "Green"
    $btOK.Text = "Continue"
    $btOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btCancel = New-Object System.Windows.Forms.Button
    $btCancel.Location = New-Object System.Drawing.Size(150,80) 
    $btCancel.Size = New-Object System.Drawing.Size(70,20)
    $btCancel.ForeColor = "Red"
    $btCancel.Text = "Cancel"
    $btCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $promptWindowForm.AcceptButton=$btOK
    $promptWindowForm.CancelButton=$btCancel
    $promptWindowForm.Controls.AddRange(@($lblPromptInfo, $txtUserInput,$btOK,$btCancel))
    $promptWindowForm.Topmost = $true
    $promptWindowForm.Add_Shown( { $promptWindowForm.Activate();$txtUserInput.Select() })
    $result=$promptWindowForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $response = $txtUserInput.Text
        return $response
    }
    else {
        return $null
    }   
}
#Export-ModuleMember -Function get-valueFromInputBox

#################################################### OFFICE 365
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
#Export-ModuleMember -Function get-ExchangeConnectionStatus

function connect-Azure {
    try {
        $AzSourceContext=Get-AzContext
    } catch {
        write-log $_.exception -type error
        exit -1
    }
    if([string]::IsNullOrEmpty( $AzSourceContext ) ) {
            write-log "you need to be connected before running this script. use connect-AzAccount first." -type error
            exit -1
    }
    write-log "connected to $($AzSourceContext.Subscription.name) as $($AzSourceContext.account.id)" -silent -type info
    write-host "Your Azure connection:"
    write-host "  subscription: " -noNewLine
    write-host -foreground Yellow "$($AzSourceContext.Subscription.name)"
    write-host "  connected as: " -noNewLine 
    write-host -foreground Yellow "$($AzSourceContext.account.id)"
    Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"
}
#Export-ModuleMember -Function connect-Azure

Export-ModuleMember -Function * -Variable 'logFile'

