<#
.SYNOPSIS
    eN's support functions library.
.DESCRIPTION
    most commonly required functions such as better logging support for scripts, forking
    information to screen and file, enchanced GUI controls such as input box, most common
    connection accelerators and CSV manipulation.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 241114
    changes
        - 241114 major changes and fixes to the load-CSV, connect-Azure - environments [1.4.0]
        - 241029 [1.3.34]
        - 241007 CSV2XLS fixes and -open switch
        - 220523 silent mode for CSV/XLS, default message for get-valueFromInputBox
        - 220423 updates to select-directory
        - 220418 convert-CSV2XLS PS7 fix [1.3.33]
        - 220412 setting 'persistent' as option was a mistake - this must be a default behaviour [1.3.32]
        - 220411 convert-CSV2XLS out folder change
        - 220407 rare error to persistent flag (requires rethinking) 
        - 220403 fix for autosave during XLS2CSV load [1.3.31]
        - 220328 get-CSVDelimiter universalized and TAB delim added for detection [1.3.30]
                 major changes in write-log/start-logging
                 fix to extract icon - incompatibilities between PS5 & PS7
                 quickEdit mode typedef tuning
        - 220321 loading in PS 7.x of icon extractor didn't work
        - 220301 write-log error handling fix [1.3.22]
        - 220203 fixed retuned values from mutlichoice [1.3.21]
        - 220202 multichoice for select-ADObject [1.3.2]
        - 210810 select-OrganizationalUnit replaced with select-ADObject and proxy function for backward compatibility [1.3.1]
        - 210609 set-QuickEditMode function [1.3.0]
        - 210524 fix to select-Directory
        - 210520 fixes to select-OU, new select-Directory,select-File [1.2.0]
        - 210507 write-log null detection fix [1.1.8], get-valueFromInputBox
        - 210430 covert-CSV2XLS #typedef
        - 210422 again fixes to exit
        - 210421 write-log $message fix
        - 210408 fixes to import-* function exit
        - 210402 write-log 3rd output init, get-valueFromInput fix
        - 210329 write-log init fix
        - 210321 select-OU extention
        - 210317 upgrade to CSVtoXLS, get-ValueFromInput, delimiter detection, select-OU
        - 210315 many changes to CSV functions, experimental import-XLS
        - 210309 proper pipelining for CSV convertion, get-AzureADConnectionStatus
        - 210308 select-OU, convert-XLS2CSV, convert-CSV2XLS
        - 210302 write-log fix, check-exoconnection ext
        - 210301 connect-azure fix
        - 210219 import-structuredCSV function added, with alias load-csv, fix to connect-azure
        - 210212 wl fix
        - 210210 write-log and start-logging init fix
        - 210209 get-answerBox changes, get-valueFromInputBox, wl fix
        - 210206 write-log accepts all unnamed parameters as messages
        - 210205 write-log fixes
        - 210202 tuned write-log and start-logging, fixes and logical separation. v0.9
        - 201018 initialize 
#>

#################################################### GENERAL
function start-Logging {
    <#
    .SYNOPSIS
        initilizes log file under $logFile variable for write-log function.
    .DESCRIPTION
        all scripts require logging mechinism. write-log function forking each output to the screen and to the 
        logfile is a very convinient way keeping all logs consistent with very nice host output. for those who
        use scripts with 'always verbose' paradigm meaning that while running a script you want to have information
        what is going on (on screen) and have it logged in case anything goes wrong.
        this function initilizes $logFiles variable creation, and initiates the log file itself. in order to 
        ease the creation there are several ways of initilizing $logFile:
        - using write-log directly
            - from console host
            - from script
        - using start-logging function directly 
            - no parameters - defaults to $ScriptRoot/Logs folder or user documents/Logs if run from console
            - using 'useProfile' - to store logs in User Documents/Logs directory
            - using 'logFolder' parameter to define particular folder for logs
            - using 'logFileName' - (exclusive to logFolder) presenting full path for the log or logfile name 
        logFile name changes automatically when script name changes. you can use 'persistent' switch to make given 
        logFile persistent for all scripts run later - this is especially important if you use externall calls 
        (invoke-command or &) so they are still logging to the same single file. 

        script initializes $logFiles array variable to keep tract on log file names run from different context.
        write-log creates simple [string]$LogFile variable so you can easly reference the log file name in your 
        scripts.
    .EXAMPLE
        write-log 

        using write-log will inderectly call start-logging function and initialize the log file with generic name,
        saved in 'Logs' subfolder under script run path or 'documents' folder if run directly from console. 
    .EXAMPLE
        start-Logging -logFileName c:\temp\myLogs\somelog.log
        write-log 'test message'

        initializes the log file as c:\temp\myLogs\somelog.log .
    .EXAMPLE
        start-Logging -logFileName c:\temp\myLogs\somelog.log -persistent:$false
        write-log 'test message'

        initializes the log file as c:\temp\myLogs\somelog.log and makes it non-persistent - next script run from this 
        script will generate new logfile name. this is usefull for 'launchers' - scripts that launch series of other 
        scripts that supposed to run with its own logfiles
    .EXAMPLE
        start-Logging -logFileName somelog.log
        write-log 'test message'

        initializes the log file as .\somelog.log .
    .EXAMPLE
        start-Logging -logFolder c:\temp\myLogs
        write-log 'test message'

        initializes the log file under c:\temp\myLogs\ folder with generic name containing script name and date.
    .EXAMPLE
        start-Logging -userProfilePath
        write-log 'test message'

        initializes the log file in Logs subfolder uder user profile path with generic name containing script name and date.
    .INPUTS
        None.
    .OUTPUTS
        log file under $logFile variable.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 220412
        changes:
            - 220412 persistence set as default - multilevel calls are not expected by default
            - 220328 rewritten with many fixes, and mostly - supports multi-level calls. when calling script-from-script.
                persistent switch added which related to multi-level calls. 
            - 210408 breaks
            - 210210 removing recurrency to write-log (loop elimination)
            - 210205 fixes to logfilename initialization
            - 210203 added 'logFolder' and proper log initilization when called indirectly
            - 210127 v1
            - 201018 initialize
    #>
    [CmdletBinding(DefaultParameterSetName='FilePath')]
    param(
        #provide custom name or full file path for log file. 
        [Parameter(ParameterSetName='FilePath',mandatory=$false,position=0)]
            [string]$logFileName,
        #create log file with automatic name in the user profile folder rather than script run path
        [Parameter(ParameterSetName='userProfile',mandatory=$false,position=0)]
            [alias('useProfile')]
            [switch]$userProfilePath,
        #similar to logFileName, but takes folder only and log file name is generic.
        [Parameter(ParameterSetName='Folder',mandatory=$false,position=0)]
            [string]$logFolder,
        #make this logFile persisent till the end of the PS Session or re-running start-Logging (write-log will not generate new name)
        [Parameter(mandatory=$false,position=1)]
            [switch]$persistent = $true
    )
    #prepare baseName of the logFile 
    #main global object used to keep record
    if(-not $global:logFiles) {
        $global:logFiles = @()
        for($lvl = 0; $lvl -lt 10; $lvl++ ) {       #yeah.. hardcoding such limits is a risk, but I assume script will not be nested/stacked more 1o times
            $global:logFiles += [PSCustomObject]@{   #need to be array as $script context is broken and to handle nested invocation need to keep seperate values
                logName = ''                        #actual logFile name declared
                persistent = $false                 #enforce all scripts to use this name until directly changed with start-logging
                lastScriptUsed = ''                 #name of the script that created the logFile
            } 
        }
    } 
    $scriptCallStack = (Get-PSCallStack | Where-Object {$_.command -ne 'write-log' -and $_.command -ne 'start-logging' -and $_.ScriptName -notmatch "\.psm1$"} )
    $runLevel = $scriptCallStack.count - 1
    if( -not ($scriptCallStack | ? ScriptName) ) { #if run from console - set the logfile name as 'console'
        $scriptBaseName = 'console'
    } else {
        $scriptBaseName = ([System.IO.FileInfo]$scriptCallStack[0].ScriptName).basename #after removing write-log and start-logging from callStack, next is a script that called
    }
    $logFiles[$runLevel].lastScriptUsed = $scriptBaseName
    
    #dependently on the parameters prepare actual logFile name and folder
    switch($PSCmdlet.ParameterSetName) {
        'userProfile' {
            $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
            $logFiles[$runLevel].logName = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        'Folder' {
            $logFiles[$runLevel].logName = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        'filePath' {
            if ( [string]::IsNullOrEmpty($logFileName) ) { #start-logging used without any parameters (default)
                if($scriptBaseName -eq 'console') {
                    $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
                    $logFiles[$runLevel].logName = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)   
                } else {
                    $logFolder = "$(split-path $scriptCallStack[0].scriptname -Parent)\Logs"
                    $logFiles[$runLevel].logName = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)   
                }
            } else { #$logFileName provided . it can be: 1. file, 2. folder, 3. fullpath
                if( [string]::isNullOrEmpty( (split-path $logFileName) ) ) { #no path - only file name
                    if( $scriptBaseName -eq 'console' ) { #run directly from console
                        $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
                    } else {
                        $logFolder = $MyInvocation.PSScriptRoot
                    }
                } else {    
                    if( test-path $logFileName -PathType Container ) { #folder ...ups!
                        write-host "$logFileName seems to be an existing folder. use 'logFolder' parameter or change log name. quitting." -ForegroundColor Red
                        return
                    }
                    $logFolder = split-path $logFileName -Parent
                    $logFileName = split-path $logFileName -Leaf
                }
                $logFiles[$runLevel].logName = "$logFolder\$logFileName"
            }
        }
        default {
            write-host -ForegroundColor Magenta 'very strange error'
            return
        }
    }

    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder -ErrorAction Stop|Out-Null
            write-host "$LogFolder created."
        } catch {
            write-error $_.exception
            break
        }
    }
    if($persistent.IsPresent) {
        $logFiles[$runLevel].Persistent = $true
    } else {
        $logFiles[$runLevel].Persistent = $false
    }
    "*logging initiated $(get-date) in $($logFiles[$runLevel].logName)"|Out-File $logFiles[$runLevel].logName -Append
    write-host "*logging initiated $(get-date) in $($logFiles[$runLevel].logName)"
    "*script parameters:"|Out-File $logFiles[$runLevel].logName -Append
    if($script:PSBoundParameters.count -gt 0) {
        $script:PSBoundParameters|Out-File $logFiles[$runLevel].logName -Append
    } else {
        "<none>"|Out-File $logFiles[$runLevel].logName -Append
    }
    "***************************************************"|Out-File $logFiles[$runLevel].logName -Append
}
function write-log {
    <#
    .SYNOPSIS
        replacement for tee-object and write-host, forking information to a log file and screen
        (and possibly to third object), with flexible log initialization and colouring.
    .DESCRIPTION
        function has been developed with a paradigm 'always verbose' - if you need to see script
        run, and same time have it in a log file. it automates forking of output on two different 
        endpoints - on the host, using write-host and to the file, appening its content, similarly 
        to tee-object, but adds more options like message tag (error, info, warning) with cloured 
        output and timestamps.
        ...actually there is option to fork on third source, described later...

        write-log converts everything to a string, so you can use it for virtually any type of 
        variable - including objects. 

        in order to use write-log, $logFiles variable requires to be set up. you can initialize the 
        value directly with 'start-Logging', configure $logFile manually or simply run write-log to 
        have it initialized automatically. by default logs are stored in $PSScriptRoot/Logs directory 
        with generic file name. if you need special location refer to start-logging help how to
        initialize variable. 

        function may also be used directly from command line - in this scenario log file will be created 
        in Logs directory under User Documents folder. file with be named 'console-<date>.log'.

        THIRD OBJECT

          third object was introduced to help developing GUI apps and ability to show log on Forms
        elements, but may be used for regular strings as well and may be easily extended on other types.
        referenced objects must be provided as '([ref]$OBJECTNAME)' - with parenthesis and [ref]. 
        as example, lets assume you have a Forms Label to show progress:

              $LabelStatus = New-Object System.Windows.Forms.Label

        you can then use write-log to fork on the screen, the file and show it on the label:
        write-log "something happens" -type info -thirdOutput $labelStatus

        GLOBAL VARIABLES
        start-logging initializes $logFiles variable which is an array, allowing to store different log
        file names for scripts run on different stack levels (when you do dot sourcing, &, invokes etc).
        this variable persists gloabally so the next script has 'a memory' to reuse the name.

        write-log creates simple [string]$logFile globally so you can easily reference the name in your 
        srcipts e.g.:

        write-log "script run finished. check $logFile for details" -type ok

    .EXAMPLE
        write-log "all is fine"

        output '<timestamp> all is fine' on the screen and to the log file.
    .EXAMPLE
        write-log all is fine -skip

        outputs 
            all is fine
        on the screen and to the log file - all unnamed parameters are displayed
    .EXAMPLE
        write-log -message "trees`nare`ngreen" -type ok

        shows:
        '<timestamp> OK: trees
        are
        green'
        in Green colour, and send text to a log file.
    .EXAMPLE
        start-Logging -logFileName ("{0}{1}{2}{3}{4}{5}" -f $PSScriptRoot,'\Logs\_',$env:USERNAME,'myScript-',$(Get-Date -Format yyMMddHHmm),'.log')
        write-log -message $someObject -type info 

        if you need to create a logfile with some custom name or location - just use 'start-Logging' function which initializes 
        logfile with required values. 
    .EXAMPLE
        $someObject=get-process
        write-log -message $someObject -type info -skipTimeStamp -silent

        outputs processes object to the log file as -silent disables screen output. it will lack
        timestamp in a message header but will contain 'INFO:' block.
    .INPUTS
        None.
    .OUTPUTS
        text log file under $logFile
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 220407
        changes:
            - 220407 tiny fix to persistant logging
            - 220328 v3 rewritten with many fixes, and mostly - supports multi-level calls. when calling script-from-script.
                     skipTimeStamp changed to noTimeStamp.
            - 220301 error handling for add-content - issue found when trying to write to network drives and timeout occurs. 
            - 210526 ...saga with catching $null continues
            - 210507 rare issue with message type check
            - 210421 interpreting $message elements fix
            - 210402 3rd output init
            - 210329 write-log init fix
            - 210302 do not convert to 'out-string' when it's already a string
            - 210212 imporper name when calling from console thru module
            - 210210 v2. init finally works!
            - 210209 when initialized on console, wl was not creating script log and using console file. 
            - 210209 init from console fix
            - 210206 valueFromRemainingArguments 
            - 210205 fix when run directly from console, init fixes
            - 210203 properly initiating log with new start-logging, when called indirectly
            - 210127 v1
            - 201018 initialize

        #TO|DO
        - colouring codes for text - change screen text colour on ** <y></y> <r></r> <g></g> 
    #>
    #DO NOT ADD 'parameter' data as it will break ValueFromRemainingArgs taking over numbered parameters    
    param(
        #message to display - can be an object
        [parameter(ValueFromRemainingArguments=$true,mandatory=$false,position=0)]
            $message,
        #adds description and colour dependently on message type
            [string][validateSet('error','info','warning','ok')]$type,
        #do not output to a screen - logfile only
            [switch]$silent,
        # do not show timestamp with the message
            [Alias('skipTimestamp')]
            [switch]$noTimeStamp,
        #experimantal - 3rd output object so you can add virtually anything that accepts text to be set to. 
            [ref][alias('thirdOutput')]$externallyReferencedObject
    )

#region INIT_LOG_FILE_NAME
    #0. no logFile - new
    #1. logfile & persistent - keep the same
    #       logfile and not persisent:
    #   2. different script name - new
    #   3. same script and the same level (invocations) - keep the same
    #   4. same script but different level - new
    $scriptCallStack = (Get-PSCallStack | Where-Object {$_.command -ne 'write-log' -and $_.command -ne 'start-logging' -and $_.ScriptName -notmatch "\.psm1$"} )
    $runLevel = $scriptCallStack.count - 1
    #$scriptCallStack|fl|Out-Host
    $scriptBaseName = ([System.IO.FileInfo]$scriptCallStack[0].ScriptName).basename 
    if( $logFiles ) { #$logFiles already initialized
        if( -not ($logFiles | Where-Object persistent) ) { #logFiles set but not Persistent 
            if([string]::isNullOrEmpty($scriptBaseName)) { #run directly from console
                $scriptBaseName = 'console'
            } 
            if($logFiles[$runLevel].lastScriptUsed -ne $scriptBaseName) { #logFiles exists and for the same script - check if the same level
                start-Logging
            }
            $LogFile = $logFiles[$runLevel].logName
        } else { #if persisent - then don't generate new
            $LogFile = ($logFiles | Where-Object persistent)[0].logName
        }
    } else {
        #no $logFiles - create new
        start-Logging 
        $LogFile = $logFiles[$runLevel].logName
    }
#endregion INIT_LOG_FILE_NAME

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {
        $message=''
    } else {
        #ValueFromRemainingArguments changes how variable is presented. running "write-log 'output text'" will pass 'output text' as a "List`1" object
        #as that's how 'valueFromRemainingArguments' is bulding variable and ruturns an array. but if passed directly with "write-log -message 'output text'" 
        #retruns true type of the passed variable - in this case [string]. 
        try {
            switch( $message.GetType().name )  {
                'List`1' { 
                    for($count=0;$count -lt $message.count;$count++) {
                        if($message[$count].GetType().name -ne 'string') {
                            $message[$count]=($message[$count]|out-String).trim()+"`n" 
                        } else {
                            $message[$count]+=' '
                        }
                    }
                    $message=$message -join ''
                }
                'String' {
                    #do nothing - string is fine.
                }
                'Default' {
                    ($message|out-String).trim()+"`n" 
                }
            }
        } catch {
            #there simply is no way to catch all nulls - any type of comparison generates exceptions. 
            $message=''
        }
    }
#region 3rdOUTPUT
    if($externallyReferencedObject) {
        if($externallyReferencedObject.Value.GetType().FullName -match 'System.Windows.Forms') {
            if($externallyReferencedObject.Value.multiline) {
                $externallyReferencedObject.Value.Text += "$message`n`r"
            } else {
                $externallyReferencedObject.Value.Text = $message
            }
            $externallyReferencedObject.Value.refresh()
        } else {
            $externallyReferencedObject.Value = $message
        }
    }
#endregion 3rdOUTPUT

#region FILE_OUTPUT
    try {
        $finalMessageString=@()
        if(-not $noTimestamp) {
            $finalMessageString += "$(Get-Date -Format "hh:mm:ss>") "
        }
        if(-not [string]::IsNullOrEmpty( $type) ) { 
            $finalMessageString += $type.ToUpper()+": " 
        }
        $finalMessageString += $message
        $message=$finalMessageString -join ''
        try {
            Add-Content -Path $LogFile -Value $message -ErrorAction Stop
            $global:logFile = $logFile
        } catch {
            "ERROR WRITING TO LOG FILE: $($_.exception)" | out-host 
        }
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
#endregion FILE_OUTPUT

}
function get-CSVDelimiter {
    <#
    .SYNOPSIS
        support function for CSV import. primitive function trying to detect delimiter used in CSV file.
    .DESCRIPTION
        different languages have different separators. CSV actually means 'Country-specyfic Separator Value' when you use Excel.
        to avoid necessity to transporm semicolon-separated to/from comma-separated, this function counts characters to guess 
        the separator. 
        it requires at least two data lines as the trick is to compare characters in the first and the second line to chose
        which occures in consistent number.

        currently function checks free most common separators: ';', ',' and TAB. otherwise returns default ',' as English 
        regionals are the most common.

    .EXAMPLE
        get-CSVDelimiter c:\temp\mydata.csv
        
    .INPUTS
        csv table
    .OUTPUTS
        [char]
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210219
            last changes
            - 210219 initialized
    
        #TO|DO
    #>
    
    param(
        [string]$inputCSV
    )
    #header check - at least 2 lines required
    $FirstLines = Get-Content $inputCSV -Head 3
    if($FirstLines[0] -match '^#') { #comments and #TYPE defs from export - skip
        $FirstLines=$FirstLines[1..2]
    }
    if($FirstLines.count -lt 2) {
        write-log "$inputCSV is not proper stuctured CSV - at least 2 lines expected (header and data)." -type error
        return $null
    }

    #this is very simple delimiter check based on number of columns in two first lines
    $FirstLines=$FirstLines -replace '''.*?''|".*?"','ANTI-DELIMITER' #change all quoted strings to simple string to avoid quoted delimiter characters
    $delims=@(",",";","`t")
    $current = ','
    $maxCount = 0
    foreach($delimiter in $delims) {
        $fl = $FirstLines[0].split($delimiter).Length - 1
        $sl = $FirstLines[1].split($delimiter).Length - 1
        if( ($sl - $fl) -eq 0 ) {
            if( $maxCount -lt $fl ) { 
                $maxCount = $fl
                $current = $delimiter
            }
        }
    }
    write-log "'$current' detected as delimiter." -type info
    return $current
}
function import-structuredCSV {
    <#
    .SYNOPSIS
        loads CSV file with header check and auto delimiter detection. 
    .DESCRIPTION
        support function to gather data from CSV file with ability to ensure it is correct CSV file by
        enumerating header. 
        if you operate on data you need to ensure that it is CORRECT file, and not some random CSV. 
        extremally usuful in the projects when you use xls/csv as data providers and need to ensure
        that you stick to the standard column names. You have an ability to force adding columns when only 
        several are obligatory and others are not (default) or define string checking making entire header 
        structure critical.

        additionally you can manipulate parameter names during the import by adding prefix and suffix to 
        parameter names - e.g. you import CSV with columns 'username' and 'activity' but want to have
        'AD_username' and 'AD_actitivity' for easier recognition. 

    .EXAMPLE
        $data = load-csv c:\temp\ADUserActivity.csv

        imports CSV, automatically detecting delimiter 

    .EXAMPLE
        $inputCSV = "c:\temp\ADUserActivity.csv"
        $header=@('username','activity')
        $data = load-CSV -header $header -headerIsCritical -delimiter ';' -inputCSV $inputCSV

        above code will load CSV expecting minimum of 'username' and 'activity' columns to be present 
        (there might be more). 
        since 'headerIsCritical' flag is added, script will terminate if any of 
        these columns is missing.
        delimiter enforces semicolon as CSV delimiter.
        
    .EXAMPLE
        $inputCSV = "c:\temp\ADUserActivity.csv"
        $header=@('username','activity')
        $data = load-CSV -header $header -inputCSV $inputCSV -prefix 'AD_'

        above code will import CSV while ensuring that columns 'username' and 'activity' exist. If any 
        of the column is not found in the CSV, script will ask what to do - add them, terminate or 
        simply continue.
        imported data columns/attribute names will be prefixed with 'AD_' - here 'AD_username' and 
        'AD_activity'.
        delimiter is detected automatically.
        
    .EXAMPLE
        $inputCSV = "c:\temp\ADUserActivity.csv"
        $header=@('AD_username','AD_activity')
        $data = load-CSV -header $header -inputCSV $inputCSV -transformation @{'username' = 'AD_UserName';'password' = 'AD_Password'}

        above code will import CSV while ensuring that columns 'AD_username' and 'AD_activity' are present
        but not in CSV but AFTER TRANSFORMATION. transformation is processed PRELOADING - beafore header is checked
        
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 221114
            last changes
            - 221114 transformation table for column names
            - 221113 fixes to PS7 - isPresent doesn't work on non-switch parametrs, encoding 
            - 221112 attribute prefix/suffix while loading
            - 210523 silent mode
            - 210421 exit/return/break tuning
            - 210317 delimiter detection as function
            - 210315 finbished auto, non-terminating from console, header not mandatory
            - 210311 auto delimiter detection, min 2lines
            - 210219 initialized
    
        #TO|DO
    #>
    param(
        #path to CSV file containing data
        [parameter(mandatory=$true,position=0)]
            [string]$inputCSV,
        #expected header to check if this is the CSV you're actually expecting
        [parameter(mandatory=$false,position=1)]
            [string[]]$header,
        #this flag causes exit on load if any column is missing. 
        [parameter(mandatory=$false,position=2)]
            [switch]$headerIsCritical,
        #CSV delimiter if different then regional settings. auto - tries to detect between comma and semicolon. uses comma if failed.
        [parameter(mandatory=$false,position=3)]
            [string]$delimiter='auto',
        #CSV encoding - deafult vlaue of ansi is breaking diactritics. here UTF8 is chosen, but for Azure outputs it's recommended to use 
        [parameter(mandatory=$false,position=4)]
            [validateSet('ansi','ascii','bigendianunicode','bigendianutf32','oem','unicode','utf7','utf8','utf8BOM','utf8NoBOM','utf32')]
            [string]$encoding='utf8',
        #add prefix to all column names *after checking the CSV header*
        [Parameter(mandatory=$false,position=5)]
            [string]$prefix,
        #add suffix to all column names *after checking the CSV header*
        [Parameter(mandatory=$false,position=6)]
            [string]$suffix,
        #column name transformation table. transformation is proccessed *before checking the header*
        [Parameter(mandatory=$false,position=7)]
            [hashtable]$transformationTable,
        #silent - no output on screen. my script are in always-verbose logic, so this is opposite to regular PS, 'silent' allows to disable output  
        [Parameter(mandatory=$false,position=8)]
            [switch]$silent
    )

    if($silent.IsPresent) {
        $PSDefaultParameterValues=@{"write-log:silent"=$true}
    }

    if(-not (test-path $inputCSV) ) {
        write-log "$inputCSV not found." -type error
        return
    }

    if($delimiter -eq 'auto') {
        $delimiter = get-CSVDelimiter -inputCSV $inputCSV
        if($null -eq $delimiter) {
            return
        }
    }

    try {
        $CSVData = import-csv -path "$inputCSV" -delimiter $delimiter -Encoding $encoding
    } catch {
        Write-log "not able to import $inputCSV. $($_.exception)" -type error 
        return
    }

#region tranformation
    if($transformationTable) {
        $CSVData | ForEach-Object {
            foreach($propertyName in ( ($_.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                if($transformationTable.ContainsKey($propertyName) ) {
                    $_ | Add-Member -MemberType NoteProperty -Name $transformationTable[$propertyName] -Value $_.$propertyName
                    $_.PSObject.Properties.Remove($propertyName)
                }
            } 
        }
    }    
#endregion transformation

#region header check
    if($null -ne $header) {
        $csvHeader = $CSVData | get-Member -MemberType NoteProperty | select-object -ExpandProperty Name
        $hmiss = @()
        foreach($el in $header) {
            if($csvHeader -notcontains $el) {
                Write-log """$el"" column missing in imported csv" -type warning
                $hmiss += $el
            }
        }
        if($hmiss) {
            if($headerIsCritical) {
                Write-log "Wrong CSV header. check delimiter used. quitting." -type error
                return
            }
            $ans = Read-Host -Prompt "some columns are missing. type 'add' to add them, 'c' to continue or anything else to cancel"
            switch($ans) {
                'add' {
                    foreach($newCol in $hmiss) {
                        write-host "adding $newCol"
                        $CSVData | add-member  -MemberType NoteProperty -Name $newCol -value ''
                    }
                    write-log "header extended" -type info
                }
                'c' {
                    write-log "continuing without header change" -type info
                }
                default {
                    write-log "cancelled. exitting." -type info
                    return
                }
            }
        }
    }
#endregion header check

#region addPrefix
    if(-not [string]::isNullOrEmpty($prefix)) {
        $CSVData | ForEach-Object {
            foreach($propertyName in ( ($_.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                $_ | Add-Member -MemberType NoteProperty -Name "$prefix$propertyName" -Value $_.$propertyName
                $_.PSObject.Properties.Remove($propertyName)
            } 
        }
    }
#endregion addPrefix

#region addSuffix
    if(-not [string]::isNullOrEmpty($suffix)) {
        $CSVData | ForEach-Object {
            foreach($propertyName in ( ($_.psobject.Properties | ? memberType -eq 'NoteProperty')).name )  {
                $_ | Add-Member -MemberType NoteProperty -Name "$propertyName$suffix" -Value $_.$propertyName
                $_.PSObject.Properties.Remove($propertyName)
            } 
        }
    }
#endregion addSuffix

    return $CSVData
}
set-alias -Name load-CSV -Value import-structuredCSV

function convert-XLStoCSV {
    <#
    .SYNOPSIS
        export all tables in XLSX files to CSV files. enumerates all sheets, and each table goes to another file.
        if sheet does not contain table - whole sheet is saved as csv
    .DESCRIPTION
        if file contains information outside of table objects - they will not be exported.
        files will be named after the sheet name + table/worksheet name and placed in seperate directory.

        separate script with ability to drag'n'drop may be downloaded from
        https://github.com/nExoRek/eN-Lib/blob/master/convert-XLSX2CSV.ps1
    .EXAMPLE
        convert-XLS2CSV -fileName .\myFile.xlsx

        extracts tables/worksheets to CSV files under folder named after file
    .EXAMPLE
        ls *.xlsx | convertTo-CSVFromXLS

        converts all xlsx file in current directory to series of CSVs. 
    .INPUTS
        XLSX file.
    .OUTPUTS
        Series of CSV files representing tables and/or worksheets (if lack of tables).
    .LINK
        https://w-files.pl
    .LINK
        https://github.com/nExoRek/eN-Lib/blob/master/convert-XLSX2CSV.ps1
        drag'n'drop version - separate file.
    .NOTES
        nExoR ::))o-
        version 231016
            last changes
            - 231016 return/exit, cleanup
            - 220523 silent mode - for import-xls
            - 220403 autosave error when not on OD
            - 220401 stupid autosave behaviour, file open error handling
            - 210422 ...again fixes to exit/break/return
            - 210408 proper 'run from console' detection and exit
            - 210317 firstWorksheet, suppress directory creation info
            - 210315 error detection during creation
            - 210309 proper pipeline
            - 210308 module function
            - 201121 output folder changed, descirption, do not export hidden by default, saveAs CSVUTF8
            - 201101 initialized
        TO|DO 
        - explore silent mode
    #>
    [cmdletbinding()]
    param(
        # XLSX file name to be converted to CSV files
        [Parameter(ParameterSetName='byName',mandatory=$true,position=0,ValueFromPipeline)]
            [string]$XLSfileName,
        # XLSX file object to be converted to CSV files
        [Parameter(ParameterSetName='byObject',mandatory=$true,position=0,ValueFromPipeline)]
            [System.IO.FileInfo]$XLSFile,
        #export only first worksheet, not all
        [Parameter(mandatory=$false,position=1)]
            [switch]$firstWorksheetOnly,
        #include hidden worksheets? 
        [Parameter(mandatory=$false,position=2)]
            [switch]$includeHiddenWorksheets,
        #silent - no output on screen. my scripts are built with always-verbose logic, opposite to regular PS 
        [Parameter(mandatory=$false,position=3)]
            [switch]$silent
    )

    begin {
        if($silent.IsPresent) {
            $PSDefaultParameterValues=@{"write-log:silent"=$true}
        }
        try{
            $Excel = New-Object -ComObject Excel.Application
        } catch {
            write-log "not able to initialize Excel lib. requires Excel to run.`n$($_.exception)" -type error
            break
        }
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
    }

    process {
        if($PSCmdlet.ParameterSetName -eq 'byName') {
            if(-not (test-path $XLSFileName)) {
                write-log "$XLSfileName not found. exitting" -type error
                return
            }
            $XLSFile=get-Item $XLSfileName
        }
        if($XLSFile.Extension -notmatch '\.xls.$') {
            write-log "$($XLSFile.Name) doesn't look like excel file. exitting" -type error
            return
        }
        $outputFolder=$XLSFile.DirectoryName+'\'+$XLSFile.BaseName+'.exported'
        if( -not (test-path($outputFolder)) ) {
            new-Item -ItemType Directory $outputFolder|Out-Null
        }
        try {
            $workBookFile = $Excel.Workbooks.Open($XLSFile)
        } catch {
            write-log "can't open $XLSfileName. $($_.Exception)" -type error
            return
        }
        #if file is opened from OneDrive and workbook is set to autosave - additional sheet will autoamtically be saved, although script is deleting it /:
        try {
            $workBookFile.autoSaveOn = $false
        } catch { 
            #silence out error when already disabled 
        }

        #excel file save statics
        $fileType=62 #CSVUTF8 https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
        $localLanguage=$true
        write-log "converting $($XLSFile.Name) tables to CSV files..." -type info
        $CSVFileList=@()

        foreach($worksheet in $workBookFile.Worksheets) {
            if($worksheet.Visible -eq $false -and -not $includeHiddenWorksheets.IsPresent) {
                write-log "worksheet $($worksheet.name) found but it is hidden. use -includeHiddenWorksheets to export" -type info
                continue
            }
            Write-log "found worksheet: $($worksheet.name)" -type info
            $tableList=$worksheet.listObjects|Where-Object SourceType -eq 1
            if($tableList) {
                foreach($table in $tableList ) {
                    Write-log "found table $($table.name) on $($worksheet.name)" -type info
                    $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^\w\d\-_\.]', '') + '_' + ($table.name -replace '[^\w\d]', '') + '.csv'
                    $tempWS=$workBookFile.Worksheets.add()
                    $table.range.copy()|out-null
                    $tempWS.paste($tempWS.range("A1"))
                    $tempWS.SaveAs($exportFileName, $fileType,$null,$null,$null,$null,$sddToMRU,$null,$null,$localLanguage)
                    write-log "$($table.name) saved as $exportFileName"
                    $tempWS.delete()
                    Remove-Variable -Name tempWS
                    $CSVFileList += get-Item $exportFileName
                }
            } else {
                Write-log "$($worksheet.name) does not contain tables. exporting whole sheet..." -type info
                $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^a-zA-Z0-9\-_]', '') + '_sheet.csv'
                $worksheet.SaveAs($exportFileName, $fileType,$null,$null,$null,$null,$sddToMRU,$null,$null,$localLanguage)
                write-log "worksheet $($worksheet.name) saved as $exportFileName"
                $CSVFileList += get-Item $exportFileName
            }
            if($firstWorksheetOnly) {
                break
            }
        }
        $Excel.Workbooks.Close()
    }

    end {
        $Excel.Quit()
        #any method of closing Excel file is not working 1oo%. there are scenarios where excel process stays in memory.
        #Remove-Variable -name workBookFile
        #Remove-Variable -Name excel
        #[gc]::collect()
        #[gc]::WaitForPendingFinalizers()
        #https://social.technet.microsoft.com/Forums/lync/en-US/81dcbbd7-f6cc-47ec-8537-db23e5ae5e2f/excel-releasecomobject-doesnt-work?forum=ITCG
        while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) ){}
        Write-log "done and cleared." -type ok
        return $CSVFileList
    }
}
Set-Alias -Name convert-XLS2CSV -Value convert-XLStoCSV
function convert-CSVtoXLS {
    <#
    .SYNOPSIS
        Converts CSV file into XLS with table.
    .DESCRIPTION
        creates XLXS out of CSV file and formats data as a table of preferable style.
    .EXAMPLE
        convert-CSV2XLSX c:\temp\test.csv -delimiter ','
        
        Converts test.csv to test.xlsx enforcing comma as delimiter in CSV interpretation
    .EXAMPLE
        ls *.csv | convert-CSV2XLS -outputFileName myfile.xlsx

        converts all csv files in current directory into sinlge xls file with multiple worksheets.
    .EXAMPLE
        start (convert-CSVtoXLS myfile.csv)

        convrts file and opens it in Excel
    .INPUTS
        CSV file or file name
    .OUTPUTS
        XLSX file.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241007
            last changes
            - 241007 CSV2XLS fixes and -open switch
            - 220523 silent mode
            - 220418 further fixes for PS7
            - 220411 destination folder changed to CSV location - to mach convert-XLS2CSV behaviour
            - 210430 #typedef skip
            - 210422 ...again fixes to exit/break/return
            - 210408 breaks
            - 210402 proper 'run from console' detection and exit
            - 210317 processing multiple CSV will create single XLS, delimiter autodetection, output file name
            - 210309 proper pipelining
            - 210308 module function
            - 201123 initialized
        
        TO|DO
    #>
    [CmdletBinding()]
    param (
        #CSV file name to convert
        [Parameter(ParameterSetName='byName',mandatory=$true,position=0,ValueFromPipeline)]
            [string]$CSVfileName,
        #CSV file object to convert
        [Parameter(ParameterSetName='byObject',mandatory=$true,position=0,ValueFromPipeline)]
            [System.IO.FileInfo]$CSVfile,
        #output XLSX file name
        [Parameter(mandatory=$false,position=1)]
            [alias('outputFileName')]
            [string]$XLSfileName=$null,
        #style intensity
        [Parameter(mandatory=$false,position=2)]
            [alias('intensity')]
            [string][validateSet('Light','Medium','Dark')]$tableStyleIntensity='Medium',
        #style number
        [Parameter(mandatory=$false,position=3)]
            [alias('nr')]
            [int]$tableStyleNumber=21,
        #open excel file automatically after conversion
        [Parameter(mandatory=$false,position=4)]
            [alias('run')]
            [switch]$openOnConversion,
        #CSV delimiter character
        [Parameter(mandatory=$false,position=5)]
            [string]$delimiter='auto',
        #silent - no output on screen. my script are in always-verbose logic, so this is opposite to regular PS 
        [Parameter(mandatory=$false,position=6)]
            [switch]$silent
    )
    

    begin {
        if($silent.IsPresent) {
            $PSDefaultParameterValues=@{"write-log:silent"=$true}
        }
        try{
            $Excel = New-Object -ComObject Excel.Application
        } catch {
            write-log "not able to initialize Excel lib. requires Excel to run.`n$($_.Exception)" -type error
            break
        }
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add(1)
        $XLSfileList = @()
        $counter=0
        if($delimiter -eq 'auto') {
            $autoDelimiter=$true
        }
        if(![string]::isNullOrEmpty($XLSfileName) ){
            $XLSFolder = Split-Path $XLSfileName -Parent
            if([string]::isNullOrEmpty($XLSFolder)) {
                $XLSfileName = ($pwd).path +'\'+$XLSfileName
            } else {
                $XLSFileName = (Resolve-Path $XLSFolder).Path + '\' + (Split-Path $XLSfileName -Leaf)
            }
            if($XLSFileName -notmatch "\.xls[x]?") {
                $XLSfileName+=".xlsx"
            }
            write-log "creating $XLSfileName excel file..." -type info
        }
    }

    process {
        #$ErrorActionPreference="SilentlyContinue"
        #read CSV
        if($PSCmdlet.ParameterSetName -eq 'byName') {
            if(-not (test-path $CSVfileName) ) {
                write-host -ForegroundColor Red "file $CSVfileName is not accessible"
                return
            }
            #typedef skip from PSobjects
            $TYPEline = $false
            if((Get-Content $CSVfileName -Head 1) -match "^#") {
                $TYPEline = $true
                Get-Content $CSVfileName|select-object -Skip 1|Out-File "$CSVfileName-tmp" -Encoding utf8
                $CSVFile = get-Item "$CSVfileName-tmp"
            } else {
                $CSVFile = get-Item $CSVfileName
            }
        } 

        #convert output file name to full path
        if([string]::isNullOrEmpty($XLSfileName)) {
            $XLSfileName = ($CSVfile.DirectoryName) + '\' + $CSVFile.BaseName + '.xlsx'
            write-log "creating $XLSfileName excel file..." -type info
        } 

        try {
            write-log "adding $($CSVfile.Name) data as worksheet..." -type info
            if($autoDelimiter) {
                $delimiter = get-CSVDelimiter -inputCSV $CSVfile.FullName
                if([string]::isNullOrEmpty($delimiter) ) {  
                    $delimiter=','
                }
            }
            if($counter++ -gt 0) {
                $worksheet = $workbook.worksheets.add([System.Reflection.Missing]::Value,$workbook.Worksheets.Item($workbook.Worksheets.count))
            }
            $worksheet = $workbook.worksheets.Item($workbook.Worksheets.count)
            if($CSVfile.BaseName.Length -gt 20) {
                $wksName = $CSVfile.BaseName.Substring(0,19)
            } else {
                $wksName = $CSVfile.BaseName
            }
            $worksheet.name = $wksName
            ### Build the QueryTables.Add command and reformat the data
            $TxtConnector = ("TEXT;" + $CSVFile.FullName)
            $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
            $query = $worksheet.QueryTables.item($Connector.name)
            $query.TextFileOtherDelimiter = $delimiter
            $query.TextFileParseType  = 1
            $query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
            $query.AdjustColumnWidth = 1
            $query.TextFilePlatform = 65001
            ### Execute & delete the import query
            $query.Refresh() | out-null
            $range=$query.ResultRange
            $query.Delete()

            #can't load assembly on PS7 and don't have access to enums. so far didn't found a method to load types on PS7 correctly.
            #https://github.com/PowerShell/PowerShell/issues/12052
            $Table = $worksheet.ListObjects.Add(
                #[Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,  
                1,
                $Range, 
                "importedCSV",
                1 #[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
                )

            <#
            TableStyle:
            - Light
            - Medium
            - Dark
            tableStyleNumber:
            - 1,8,15 black
            - 2,9,16 navy blue
            - 3,1o,17 orange
            - 4,11,18 gray
            - 5,12,19 yellow
            - 6,13,2o blue
            - 7,14,21 green
            #>
            $tableStyle=[string]"$tableStyleIntensity$tableStyleNumber"
            $Table.TableStyle = "TableStyle$tableStyle" #green with gray shadowing

        } catch {
            write-log "error converting CSV to XLS: $($_.exception)" -type error
            return -2         
        }
        if($TYPEline) {
            remove-item "$CSVfileName-tmp"
        }
    }

    end {
        $errorSaving=$false
        try {
            $worksheet.SaveAs($XLSfileName, 51,$null,$null,$null,$null,$null,$null,$null,'True') #|out-null
        } catch {
            write-log "error saving $XLSfileName. $($_.exception)" -type error
            $errorSaving=$true
        }
        $workbook = $null
        $Excel.Workbooks.Close()
        $Excel.Quit()
        while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) ){}
        if(!$errorSaving) {
            write-log "convertion done, saved as $XLSfileName"
            $XLSfileList += (Get-Item $XLSfileName)
            Write-log "done and cleared." -type ok
            if($openOnConversion) {
                & $XLSfileName
            }
            return $XLSfileList
        } else {
            return $null
        }

    }
}
Set-Alias -Name convert-CSV2XLS -Value convert-CSVtoXLS
function import-XLS {
    <#
    .SYNOPSIS
        EXPERIMENTAL - importing XLS as table object, using convert+load
    .DESCRIPTION
    .INPUTS
        XLS file.
    .OUTPUTS
        table
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 220523
            last changes
            - 220523 silent import
            - 210317 use of firstWorksheet
            - 210315 initialized
    
        #TO|DO
         - add cleanup of files after xls output
    #>
    
    param(
        # XLSX file name to be converted to CSV files
        [Parameter(ParameterSetName='byName',mandatory=$true,position=0)]
            [string]$XLSfileName
    )

    $tempCSV = convert-XLStoCSV -XLSfileName $XLSfileName -firstWorksheetOnly -silent
    if( $tempCSV ) {
        return (import-structuredCSV -inputCSV $tempCSV[0].FullName)
    }
}
function new-RandomPassword {
    <#
    .SYNOPSIS
        generate random password with given char ranges (complexity) and lenght
    .DESCRIPTION
        by default it genrates 8-long string with 
    .EXAMPLE
        $pass = new-RandomPassword

        generated 8-char long semi-complex password
    .EXAMPLE
        $pass = new-RandomPassword -specialCharacterRange 3

        generated 8-char long password with full complexity
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
    #TO|DO
    - character sets: 
      - upper case letters
      - lower case letters 
      - digits
      - spec0 = '.-_ '
      - spec1
      - spec2
      - spec3
    - at least one char per set - redo
    - avoid similarities (Il, 0O)
    #>
    
    param( 
        #password length
        [Parameter(mandatory=$false,position=0)]
            [int]$length=8,
        #password complexity based on a range of special characters.
        [Parameter(mandatory=$false,position=1)]
            [int][validateSet(1,2,3)]$specialCharacterRange=1,
        #uniquness - related to complexity, recommended to leave. this guarantee that password will have characters from given number of char sets.
        [Parameter(mandatory=$false,position=2)]
            [int][validateSet(1,2,3,4)]$uniqueSets=4
            
    )
    function new-CharSet {
        param(
            # set up password length
            [int]$length,
            # number of 'sets of sets' defining complexity range
            [int]$setSize
        )
        $safe=0
        while ($safe++ -lt 100) {
            $array=@()
            1..$length|ForEach-Object{
                $array+=(Get-Random -Maximum ($setSize) -Minimum 0)
            }
            if(($array|Sort-Object -Unique|Measure-Object).count -ge $setSize) {
                return $array
            } else {
                Write-Verbose "[new-CharSet]bad array: $($array -join ',')"
            }
        }
        return $null
    }
    #prepare char-sets 
    $smallLetters=$null
    97..122|ForEach-Object{$smallLetters+=,[char][byte]$_}
    $capitalLetters=$null
    65..90|ForEach-Object{$capitalLetters+=,[char][byte]$_}
    $numbers=$null
    48..57|ForEach-Object{$numbers+=,[char][byte]$_}
    $specialCharacterL1=$null
    @(33;35..38;43;45..46;95)|ForEach-Object{$specialCharacterL1+=,[char][byte]$_} # !"#$%&
    $specialCharacterL2=$null
    58..64|ForEach-Object{$specialCharacterL2+=,[char][byte]$_} # :;<=>?@
    $specialCharacterL3=$null
    @(34;39..42;44;47;91..94;96;123..125)|ForEach-Object{$specialCharacterL3+=,[char][byte]$_} # [\]^`  
      
    $ascii=@()
    $ascii+=,$smallLetters
    $ascii+=,$capitalLetters
    $ascii+=,$numbers
    if($specialCharacterRange -ge 2) { $specialCharacterL1+=,$specialCharacterL2 }
    if($specialCharacterRange -ge 3) { $specialCharacterL1+=,$specialCharacterL3 }
    $ascii+=,$specialCharacterL1
    #prepare set of character-sets ensuring that there will be at least one character from at least $uniqueSets different sets
    $passwordSet = new-CharSet -length $length -setSize $uniqueSets 

    $password=$NULL
    0..($length-1)|ForEach-Object {
        $password+=($ascii[$passwordSet[$_]] | Get-Random)
    }
    return $password
}

function Set-QuickEditMode {
    <#
    .SYNOPSIS
        function allowing to disable/enable Quick Edit Mode for current PS host session.
    .DESCRIPTION
        accidental mouse-press on PS screen will lead to script pause. this is real problem - especially if
        you're providing scripts to unaware users. this simple function taken from CodeOverflow allows
        to control Quick Edit Mode setting for current PS host. this will allow to disable this 
        feature before running the script.
    .EXAMPLE
        PS C:\> set-QuickEditMode -DisableQuickEdit
        disables Quick Edit mode for current PS Session
    .EXAMPLE
        PS C:\> set-QuickEditMode 
        enables Quick Edit mode for current PS Session
    .LINK
        source code taken from:
        https://stackoverflow.com/questions/30872345/script-commands-to-disable-quick-edit-mode/42792718
    .NOTES
        nExoR ::))o-
        version 210609
            last changes
            - 210609 initialized
    #>
    param(
        [Parameter(Mandatory=$false)]
            [switch]$DisableQuickEdit
    )

    add-type -TypeDefinition @" 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

public static class DisableConsoleQuickEdit {
    const uint ENABLE_QUICK_EDIT = 0x0040;
    // STD_INPUT_HANDLE (DWORD): -10 is the standard input device.
    const int STD_INPUT_HANDLE = -10;
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")]
    static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")]
    static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    public static bool SetQuickEdit(bool SetEnabled) {
        IntPtr consoleHandle = GetStdHandle(STD_INPUT_HANDLE);
        // get current console mode
        uint consoleMode;
        if (!GetConsoleMode(consoleHandle, out consoleMode)) {
            // ERROR: Unable to get console mode.
            return false;
        }
        // Clear the quick edit bit in the mode flags
        if (SetEnabled) {
            consoleMode &= ~ENABLE_QUICK_EDIT;
        } else {
            consoleMode |= ENABLE_QUICK_EDIT;
        }
        // set the new mode
        if (!SetConsoleMode(consoleHandle, consoleMode)) {
            // ERROR: Unable to set console mode
            return false;
        }
        return true;
    }
}
"@ -Language CSharp
    
    if( [DisableConsoleQuickEdit]::SetQuickEdit($DisableQuickEdit) ) {
        Write-Log "QuickEdit settings has been updated." -type info 
    } else {
        Write-Log "Something went wrong." -type info
    }
}

#################################################### PowerShell GUI
function get-answerBox {
    <#
    .SYNOPSIS
        win32 forms message box to get YES/NO input from user
    .DESCRIPTION
        replacement for simple messageBox giving option to customize buttons, giving option
        to add some additional information
    .EXAMPLE
        $response =  get-answerBox -OKButtonText 'YES' -CancelButtonText 'NO' -info 'choose your answer' -message 'do you find this function useful?'
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
        version 210209
            last changes
            - 210209 detailedInfo -> message (alias left for compatibility), autosize, anchors
            - 210208 icon, tune, info -> title
            - 210127 module
            - 210110 initialized
        
        #TO|DO
    #>
    
    param(
        #OK button text
        [Parameter(mandatory=$false,position=0)]
            [string]$OKButtonText = "OK",
        #Canel button text
        [Parameter(mandatory=$false,position=1)]
            [string]$CancelButtonText = "Cancel",
        #title bar text
        [Parameter(mandatory=$false,position=2)]
            [string]$title = "Which option?",
        #message text
        [Parameter(mandatory=$false,position=3)]
            [alias('detailedInfo')]
            [string]$message = "What is your choice:",
        #messagebox icon
        [Parameter(mandatory=$false,position=4)]
            [validateSet('Asterisk','Error','Exclamation','Hand','Information','None','Question','Stop','Warning')]
            [string]$icon='Question'
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    $messageBoxForm = New-Object System.Windows.Forms.Form
    $messageBoxForm.Text = $title
    $messageBoxForm.Size = New-Object System.Drawing.Size(300,120)
    $messageBoxForm.AutoSize = $true
    $messageBoxForm.StartPosition = 'CenterScreen'
    $messageBoxForm.FormBorderStyle = 'Fixed3D'
    $messageBoxForm.Icon = [System.Drawing.SystemIcons]::$icon
    $messageBoxForm.Topmost = $true
    $messageBoxForm.MaximizeBox = $false
   
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(50,50)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Anchor = 'left,bottom'
    $okButton.Text = $OKButtonText
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $messageBoxForm.AcceptButton = $okButton
    $messageBoxForm.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(160,50)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Anchor = 'right,bottom'
    $cancelButton.Text = $CancelButtonText
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $messageBoxForm.CancelButton = $cancelButton
    $messageBoxForm.Controls.Add($cancelButton)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    
    $label.AutoSize = $true
    $label.Anchor = 'left,top'
    $label.Text = $message
    $messageBoxForm.Controls.Add($label)
   
    $result = $messageBoxForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $true
    } 
    return $false
}

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
            exit
        } else {
            "you agreed, let's continue"
        }
        write-host 'code to execute here'
    .EXAMPLE
        $computerName = get-valueFromInbox -title 'Provide computer name' -maxChars 15 -allowedCharacters '[a-zA-Z0-9_-]'

        limit input to 15 characters and allow only letters,digits, underscore and minus.
    .INPUTS
        None.
    .OUTPUTS
        User Input or $null for cancel
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210523
        last changes
            - 210523 default message to be displayed while loading
            - 210507 maxChars fix
            - 210402 allowcharacter check worked for last character only.
            - 210317 allowCharacter 
            - 210209 anchored layout
            - 210113 initialized
        
        #TO|DO
    #>
    
    param(
        [parameter(mandatory=$false,position=0)]
            [string]$title='input',
        [parameter(mandatory=$false,position=1)]
            [alias('text')]
            [string]$message='put your input',
        [parameter(mandatory=$false,position=2)]
            [validateSet('Asterisk','Error','Exclamation','Hand','Information','None','Question','Stop','Warning')]
            [string]$type='Question',
        #maximum number of characters allowed
        [Parameter(mandatory=$false,position=3)]
            [alias('maxChars')]
            [int]$maxCharacters = 30,
        #regular expression limiting characters -eg 'only digits'
        [Parameter(mandatory=$false,position=4)]
            [alias('regex')]
            [regex]$allowedCharacters,
        #default value to be shown on screen
        [Parameter(mandatory=$false,position=5)]
            [string]$defaultMessage
    )
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $promptWindowForm = New-Object system.Windows.Forms.Form
    $promptWindowForm.Size = New-Object System.Drawing.Size(250,140)
    $promptWindowForm.text = $title
    $promptWindowForm.BackColor = "#ffffff"
    $promptWindowForm.AutoSize = $true
    $promptWindowForm.StartPosition = 'CenterScreen'
    $promptWindowForm.FormBorderStyle = 'Fixed3D'
    $promptWindowForm.MaximizeBox = $false
    $promptWindowForm.Icon = [System.Drawing.SystemIcons]::$type
    $promptWindowForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

    $lblPromptInfo = New-Object System.Windows.Forms.Label 
    $lblPromptInfo.Location = New-Object System.Drawing.Size(10,5) 
    #$lblPromptInfo.Size = New-Object System.Drawing.Size(230,40)
    $lblPromptInfo.AutoSize = $true
    $lblPromptInfo.MinimumSize = New-Object System.Drawing.Size(235,10)
    $lblPromptInfo.Anchor = 'left,top'
    $lblPromptInfo.Text = $message

    $txtUserInput = New-Object system.Windows.Forms.TextBox
    $txtUserInput.multiline = $false
    $txtUserInput.ReadOnly = $false
    $txtUserInput.MinimumSize = New-Object System.Drawing.Size(230,25)
    $txtUserInput.autosize = $true
    $txtUserInput.Anchor = "none" #effectively - center
    $txtUserInput.MaxLength = $maxCharacters
    $txtUserInput.location = New-Object System.Drawing.Point(0, 35)

    $btOK = New-Object System.Windows.Forms.Button
    $btOK.Location = New-Object System.Drawing.Size(30,70) 
    $btOK.Size = New-Object System.Drawing.Size(70,20)
    $btOK.ForeColor = "Green"
    $btOK.Anchor = "left,bottom"
    $btOK.Text = "Continue"
    $btOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btCancel = New-Object System.Windows.Forms.Button
    $btCancel.Location = New-Object System.Drawing.Size(150,70) 
    $btCancel.Size = New-Object System.Drawing.Size(70,20)
    $btCancel.ForeColor = "Red"
    $btCancel.Anchor = "right,bottom"
    $btCancel.Text = "Cancel"
    $btCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $promptWindowForm.AcceptButton=$btOK
    $promptWindowForm.CancelButton=$btCancel
    $promptWindowForm.Controls.AddRange(@($lblPromptInfo, $txtUserInput,$btOK,$btCancel))
    $promptWindowForm.Topmost = $true
    $promptWindowForm.Add_Shown( { $promptWindowForm.Activate();$txtUserInput.Select() })

    $txtUserInput.add_KeyUp({
        if($allowedCharacters -and $txtUserInput.text) {
            $cursor=$txtUserInput.SelectionStart
            if($txtUserInput.text[$cursor-1] -notmatch $allowedCharacters) {
                $tempText=''
                for($len=0;$len -lt $txtUserInput.text.Length;$len++) {
                    if($len -ne $cursor-1) { $tempText+=$txtUserInput.text[$len] }
                }
                $txtUserInput.Text=$tempText
                $txtUserInput.Select($cursor-1, 0);
            }
        }
    })

    $promptWindowForm.add_Shown({
        if($defaultMessage) {
            $txtUserInput.Text = $defaultMessage
            $txtUserInput.SelectionStart = 0
            $txtUserInput.SelectionLength = $defaultMessage.Length
        }
    })
    $result = $promptWindowForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $response = $txtUserInput.Text
        return $response
    }
    else {
        return $null
    }   
}
function get-Icon {
    param( 
        [int]$iconNumber,
        [string]$fileContaining = 'Shell32.dll'
    ) 
    #icon extractor 
    if($PSVersionTable.PSVersion.Major -le 5) {
        $ref = @('System.Drawing')
    } else {
        $ref = @('System.Drawing.Common','System.Runtime.InteropServices')
    }
Add-Type -TypeDefinition @"
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
    try {
        return Icon.FromHandle(largeIcon ? large : small);
    } catch {
        return null;
    }
    }
    [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
    }
}
"@ -ReferencedAssemblies $ref

    return [System.IconExtractor]::Extract($fileContaining, $iconNumber, $true)
}
function select-Directory {
    <#
    .SYNOPSIS
        accelerator function allowing to select system directory with GUI.
    .DESCRIPTION
        function is using winforms treeView to display directory structure. returns folder path on select
        or folder object if 'object' parameter used.
        WARNING! text search looks up only for loaded branches. use 'loadAll' to be able to seach entire branch.
        best is to combine with 'startingDirectory' for performance reasons - loading whole disk structure may take
        a long, long time.         
    .EXAMPLE
        cd (select-Directory)
        
        displays forms treeview enabling to choose directory using GUI and changes location
    .INPUTS
        None.
    .OUTPUTS
        directory name/object
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 220423
            last changes
            - 220423 added disk dorpdown and filter, fixes to search, 'hidden' changed to 'force'
            - 210524 fix to shortcut folders
            - 210521 optimization, tuning
            - 210519 initialized
    
        #TO|DO
        - multichoice
        - default folder should select it but let app open
    #>
    [cmdletbinding()]
    param(
        #starting folder (tree root)
        [Parameter(mandatory=$false,position=0)]
            [string]$startingDirectory='\',
        #search mask for files
        [Parameter(mandatory=$false,position=1)]
            [string]$filter,
        #include files
        [Parameter(mandatory=$false,position=2)]
            [switch]$files,
        #return directory path as string instead of 'folder object'
        [Parameter(mandatory=$false,position=3)]
            [switch]$object,
        #enable text search box - will load entire tree which might take time, but will allow to search thru entire tree. best used with -startingDirectory for subfolders
        [Parameter(mandatory=$false,position=4)]
            [switch]$loadAll,
        #show hidden folders
        [Parameter(mandatory=$false,position=5)]
            [alias('hidden')]
            [switch]$force
    )

    Function add-Nodes {
        param(
            $node,
            [int]$localDepth=0
        )
        write-verbose "check $($node.tag.name)"
        if($node.tag.unfolded -eq $false -and $node.tag.type -eq 'DirectoryInfo') {
            write-verbose "addingNode $($node.tag.name)"
            #directories first, later files [if chosen]
            $listParams = @{
                ErrorAction = 'SilentlyContinue'
                Path = $node.tag.FullName
                Directory = $true
            } 
            if($force.IsPresent) {
                $listParams.Add("Force",$true)
            }
            $SubDirList = Get-ChildItem @listParams
            $lblLoading.Text = "Loading $($node.tag.name) ($($SubDirList.count) subs)..."
            $loading.refresh()
            $node.tag.unfolded = $true
            foreach ( $dir in $SubDirList ) {
                $NodeSub = $Node.Nodes.Add($dir.Name)
                $NodeSub.tag = [psobject]@{
                    fullName = $dir.FullName
                    unfolded = $false
                    name = $dir.name
                    type = $dir.gettype().name
                }
                $script:NodeList += $NodeSub.tag
            }
            #load files [if -files switch present]
            if($files.IsPresent) {
                $listParams = @{
                    ErrorAction = 'SilentlyContinue'
                    Path = $node.tag.FullName
                    file = $true
                }
                if($filter) {
                    $listParams.Add("Filter",$filter)
                }
                if($force.IsPresent) {
                    $listParams.Add("Force",$true)
                }
                $fileList = Get-ChildItem @listParams
                $node.tag.unfolded = $true
                foreach ( $file in $fileList ) {
                    $NodeSub = $Node.Nodes.Add($file.Name)
                    $NodeSub.tag = [psobject]@{
                        fullName = $file.FullName
                        unfolded = $true #files are not to be unfolded...
                        name = $file.name
                        type = $file.gettype().name
                    }
                    $script:NodeList += $NodeSub.tag
                }
            }
            if($localDepth -gt 0) { 
                foreach($SubNode in $node.Nodes) {
                    add-Nodes -node $SubNode -localDepth ($localDepth - 1)
                }
            }
        } else {
        }
    }
    Function select-NodeByPath {
        param (
            [Parameter(Mandatory=$true)]
            [System.Windows.Forms.TreeView]$TreeView,
    
            [Parameter(Mandatory=$true)]
            [string]$Path
        )
    
        # Split the path into its components
        $pathComponents = $Path.Split([IO.Path]::DirectorySeparatorChar)
    
        # Start with the root node
        $currentNodes = $TreeView.Nodes
    
        foreach ($component in $pathComponents) {
            $foundNode = $null
            
            write-host $component
            foreach ($node in $currentNodes) {
                write-host "test: $($node.text)"
                if ($node.Text -eq $component) {
                    $foundNode = $node
                    #break
                }
            }
    
            # If the component wasn't found, exit
            if (-not $foundNode) {
                #return $null
            }
    
            # Otherwise, move to the child nodes for the next iteration
            $currentNodes = $foundNode.Nodes
        }
    
        # Select the found node
        $TreeView.SelectedNode = $foundNode
    }

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    #region LOADINGFORM
    $loading = New-Object system.Windows.Forms.Form
    $loading.ClientSize = New-Object System.Drawing.Point(250, 30)
    $loading.StartPosition = 'CenterScreen'
    $loading.TopMost = $true
    $loading.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10)
    $loading.ControlBox = $false
    $loading.FormBorderStyle = 'FixedSingle'
    $loading.padding = 0

    $lblLoading = New-Object System.Windows.Forms.Label 
    $lblLoading.Location = New-Object System.Drawing.Size(10,10) 
    $lblLoading.Size = New-Object System.Drawing.Size(250,30)
    $loading.Controls.Add($lblLoading)

    $loading.add_Shown({
        [System.windows.Forms.Application]::UseWaitCursor = $true
        [System.windows.Forms.Application]::DoEvents()        
    })
    $loading.add_Deactivate({
        [System.windows.Forms.Application]::UseWaitCursor = $false
        [System.windows.Forms.Application]::DoEvents()        
    })
    #endregion LOADINGFORM

    #region FORM
    $formFolders = New-Object System.Windows.Forms.Form
    $formFolders.Text = "select Directory under $startingDirectory"
    $formFolders.MinimumSize = New-Object System.Drawing.Size(300,500)
    $formFolders.AutoSize = $true
    $formFolders.StartPosition = 'CenterScreen'
    $formFolders.Icon = [System.Drawing.SystemIcons]::Question
    $formFolders.Topmost = $true
    $formFolders.MaximizeBox = $false
    $formFolders.dock = "fill"

    #region shortcut_buttons
    $lpUpperMenu = new-object system.windows.forms.TableLayoutPanel
    $lpUpperMenu.Anchor = 'left,right'
    $lpUpperMenu.Padding = 0
    $lpUpperMenu.ColumnCount = 2
    #$lpUpperMenu.RowCount = 1
    #$lpUpperMenu.Dock = [System.Windows.Forms.DockStyle]::Fill

    $cbDrives = New-Object System.Windows.Forms.Combobox
    $cbDrives.Size = New-Object System.Drawing.Size(40,20)
    $cbDrives.Text = "C:"
    $cbDrives.Anchor = 'none'
    foreach($vol in (get-volume | Where-Object DriveLetter)){
        [void]$cbDrives.Items.Add("$($vol.DriveLetter):\")
    }
    $cbDrives.add_SelectedIndexChanged({
        param($sender,$e)

        try {
            $initialDirectory = Get-Item $sender.text -force -ErrorAction Stop
        } catch {
            (new-object System.Windows.Forms.ToolTip -Property @{
                ToolTipIcon = 3
                isBalloon = $true
            }).show('access error',$cbDrives,3000)
        
            return
        }
        $treeView.Nodes.Clear()
        $rootNode = $treeView.Nodes.Add($initialDirectory.FullName)
        $rootNode.Tag=[psobject]@{
            fullName = $initialDirectory.FullName
            unfolded = $false
            name = $initialDirectory.Name
            type = $initialDirectory.gettype().name
        }
        
        $DEPTH = 0
        if($loadAll) {
            $DEPTH = 1000
            write-log "LOADING FULL TREE (will take time)..." -type warning
        }
        add-Nodes $rootNode -localDepth $DEPTH
        $treeView.refresh()
        $formFolders.refresh()
        $treeView.nodes[0].Expand()
    })
    $lpUpperMenu.Controls.Add($cbDrives,0,0)

    $gbShortcuts = new-object system.windows.forms.groupBox
    $gbShortcuts.Text = 'Shortcuts'
    $gbShortcuts.Anchor = 'left,right'
    $gbShortcuts.Height = 55

    $lpIcons = new-object system.windows.forms.TableLayoutPanel
    $lpIcons.ColumnCount = 6
    $lpIcons.Dock = [System.Windows.Forms.DockStyle]::Fill
    #debug
    #$lpIcons.BackColor = 'Gray'
    #$lpIcons.padding = 1
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$lpIcons.ColumnStyles.Add( (new-object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
                  
    $pbxDocument = New-Object System.windows.forms.pictureBox
    $pbxDocument.Size = New-Object System.Drawing.Size(25,25)
    $pbxDocument.SizeMode = 'StretchImage'
    $pbxDocument.Image = get-Icon -iconNumber 1
    $pbxDocument.add_Click({
        select-NodeByPath -Path ([Environment]::GetFolderPath("MyDocuments")) -treeView $treeView
        #$result.value = ([Environment]::GetFolderPath("MyDocuments")) 
        #$formFolders.DialogResult = [System.Windows.Forms.DialogResult]::OK
        #$formFolders.Close()
    })
    $lpIcons.controls.add($pbxDocument,1,0)

    $pbxDownloads = New-Object System.Windows.Forms.pictureBox
    $pbxDownloads.Size = New-Object System.Drawing.Size(25,25)
    $pbxDownloads.SizeMode = 'StretchImage'
    $pbxDownloads.Image = get-Icon -iconNumber 122
    $pbxDownloads.add_Click({
        $result.value = ((New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path)
        $formFolders.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $formFolders.Close()
    })
    $lpIcons.controls.add($pbxDownloads,2,0)
    
    $pbxDesktop = New-Object System.Windows.Forms.pictureBox
    $pbxDesktop.Size = New-Object System.Drawing.Size(25,25)
    $pbxDesktop.SizeMode = 'StretchImage'
    $pbxDesktop.Image = get-Icon -iconNumber 34
    $pbxDesktop.add_Click({
        $result.value = ([Environment]::GetFolderPath("Desktop"))
        $formFolders.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $formFolders.Close()
    })
    $lpIcons.controls.add($pbxDesktop,3,0)

    $pbxTemp = New-Object System.Windows.Forms.pictureBox
    $pbxTemp.Size = New-Object System.Drawing.Size(25,25)
    $pbxTemp.SizeMode = 'StretchImage'
    $pbxTemp.Image = get-Icon -iconNumber 35
    $pbxTemp.add_Click({
        $result.value = ($env:temp) 
        $formFolders.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $formFolders.Close()
    })
    $lpIcons.controls.add($pbxTemp,4,0)
   
    $gbShortcuts.Controls.Add($lpIcons)
    $lpUpperMenu.Controls.Add($gbShortcuts,1,0)
    #endregion shortcut_buttons
    
    $txtSearch = New-Object system.Windows.Forms.TextBox
    $txtSearch.multiline = $false
    $txtSearch.ReadOnly = $false
    $txtSearch.MinimumSize = new-object System.Drawing.Size(300,20)
    $txtSearch.AutoSize = $true
    $txtSearch.Height = 20
    $txtSearch.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
    $txtSearch.Location = new-object System.Drawing.Point(3,3)
    $txtSearch.TabIndex = 2
    $txtSearch.add_gotFocus({
        $okButton.Enabled = $false
    })

    $txtSearch.add_KeyUp({
        #param($sender,$e)
        $searchTimer.start()
    })

    #regular Tree View component - after loading
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Dock = 'Fill'
    $treeView.CheckBoxes = $false
    $treeView.Name = 'treeView'
    $treeView.TabIndex = 1
    $treeview.add_beforeExpand({
        param($sender, $e)
        write-verbose "beforeExand: $($e.node.tag.name)"
        $treeView.SelectedNode = $e.node
    })
    $treeview.add_afterSelect({
        param($sender, $e)
        write-verbose "afterSelect: $($e.node.tag.name)"
        $okButton.Enabled = $true
        #$e.node.Expand()
        $loading.Show()
        foreach($subNode in $e.node.Nodes) {
            add-Nodes -node $subNode
        }
        $loading.Hide()
        [System.windows.Forms.Application]::UseWaitCursor = $false
        [System.windows.Forms.Application]::DoEvents()        
        $formFolders.Focus()
    })    

    #'shadow' Tree View component used during text search 
    $SearchTreeView = New-Object System.Windows.Forms.TreeView
    $SearchTreeView.Dock = 'Fill'
    $SearchTreeView.CheckBoxes = $false
    $SearchTreeView.name = 'SearchTreeView'
     
    #region OKCANCEL
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Anchor = 'left'
    $okButton.Text = "OK"
    $okButton.Enabled = $false
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.TabIndex = 3
    $formFolders.AcceptButton = $okButton

    $okButton.add_Click({
        $currentView=($mainTable.controls|Where-Object name -match 'treeView')
        if($currentView.name -eq 'treeView') {
                $result.value = $currentView.SelectedNode.tag.FullName
        } else {
                $result.value = $currentView.SelectedNode.text
        }
        $formFolders.close()
    })
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.anchor = 'right'
    $cancelButton.Text = "Cancel"
    $cancelButton.TabIndex = 4
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $formFolders.CancelButton = $cancelButton
    #endregion OKCANCEL

    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.AutoSize = $true
    $mainTable.ColumnCount = 2
    $mainTable.RowCount = 4
    $mainTable.Dock = "fill"
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,65)) )
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    
    $mainTable.Controls.add($lpUpperMenu,0,0)
    $mainTable.SetColumnSpan($lpUpperMenu,2)
    $mainTable.controls.Add($txtSearch,1,0)
    $mainTable.SetColumnSpan($txtSearch,2)
    $mainTable.Controls.Add($treeView,2,0)
    $mainTable.SetColumnSpan($treeView,2)
    $mainTable.Controls.add($okButton,3,0)
    $mainTable.Controls.add($cancelButton,3,1)
    
    $formFolders.Controls.Add($mainTable)

    $searchTimer = new-object System.Windows.Forms.Timer
    $searchTimer.Interval = 1000
    #endregion FORM

    #region FORM_FUNCTIONS
    $formFolders.add_Load({
        $toolTip = new-object System.Windows.Forms.ToolTip
        $toolTip.SetToolTip($pbxDocument,'My Documents')
        $toolTip.SetToolTip($pbxDownloads,'Downloads')
        $toolTip.SetToolTip($pbxDesktop,'Desktop')
        $toolTip.SetToolTip($pbxTemp,'Temp')   
        if(!$loadAll.IsPresent) {
            $toolTip.SetToolTip($txtSearch,'directories are not fully loaded - results will be limited. check "loadAll" flag usage.')
        }
        $formFolders.refresh()
    })
    $formFolders.add_Shown({
        $initialDirectory = Get-Item $startingDirectory -force
        $rootNode = $treeView.Nodes.Add($initialDirectory.FullName)
        $rootNode.Tag=[psobject]@{
            fullName = $initialDirectory.FullName
            unfolded = $false
            name = $initialDirectory.Name
            type = $initialDirectory.gettype().name
        }
        
        $DEPTH = 0
        if($loadAll) {
            $DEPTH = 1000
            write-log "LOADING FULL TREE (will take time)..." -type warning
        }
        add-Nodes $rootNode -localDepth $DEPTH
        $treeView.refresh()
        $formFolders.refresh()
        $treeView.nodes[0].Expand()
    })
    $formFolders.add_Closing({
        $loading.dispose()
    })

    $searchTimer.add_Tick({
        if($txtSearch.Text.Length -gt 1 -and ($mainTable.Controls|Where-Object name -eq 'treeView')) {
            $mainTable.Controls.Remove($treeView)
            $mainTable.controls.add($searchTreeView,2,0)
            $mainTable.SetColumnSpan($searchTreeView,2)
            $formFolders.Refresh()
        } 
        if($txtSearch.Text.Length -le 1) {
            $mainTable.Controls.Remove($searchTreeView)
            $mainTable.Controls.Add($treeView,2,0)
            $formFolders.Refresh()
        }
        if($txtSearch.Text.Length -gt 1) {
            $searchTreeView.Nodes.Clear()
            foreach($n in $NodeList) {
                try {
                    if($txtSearch.Text.indexof('*') -ge 0) {
                        if($n.name -like $txtSearch.Text) {
                            $searchTreeView.Nodes.Add($n.FullName)
                        }
                    } else {
                        if($n.name -match $txtSearch.Text) {
                            $searchTreeView.Nodes.Add($n.FullName)
                        }
                    }
                } catch { }#eliminate regexp errors
            }
            
        }
        $searchTimer.stop()
    })
    $SearchTreeView.add_afterSelect({
        $okButton.Enabled = $true
    })
    #endregion FORM_FUNCTIONS

    $script:NodeList=@()
    $result = @{ Value='' }
   
    $ret = $formFolders.ShowDialog() 
    if($ret -eq [System.Windows.Forms.DialogResult]::OK) {
        if($object.IsPresent) {
            return (Get-Item $result.value -force)
        } else {
            return $result.value
        }
    }
}
function select-File {
    <#
    .SYNOPSIS
        wrapper function for select-directory with 'files' flag
    .NOTES
        nExoR ::))o-
        version 220423
            last changes
            - 220423 filter
            - 210520 initialized
    
        #TO|DO
    #>
    
    param(
        #starting folder (tree root)
        [Parameter(mandatory=$false,position=0)]
        [string]$startingDirectory='\',
        #search mask for files and folders
        [Parameter(mandatory=$false,position=1)]
            [string]$filter,
        #return directory path as string instead of 'folder object'
        [Parameter(mandatory=$false,position=2)]
            [switch]$object,
        #enable text search box - will load entire tree which might take time, but will allow to search thru entire tree. best used with -startingDirectory for subfolders
        [Parameter(mandatory=$false,position=3)]
            [switch]$loadAll,
        #show hidden folders
        [Parameter(mandatory=$false,position=4)]
            [alias('hidden')]
            [switch]$force
    )

    select-Directory -files @PSBoundParameters 
}
function select-ADObject {
    <#
    .SYNOPSIS
        accelerator function allowing to select OU with GUI.
    .DESCRIPTION
        function is using winforms treeView to display OU structure. returns DistinguishedName on select
        or OU object if 'object' parameter used.
        WARNING! text search looks up only for loaded branches. use 'loadAll' to be able to seach entire branch.
        best is to combine with 'startingOU' for performance reasons - loading whole AD structure may take a long
        time. 
    .EXAMPLE
        $ou = select-OU
        
        displays forms treeview enabling to choose OU from the tree.
    .EXAMPLE
        new-ADUser -name 'some user' -path (select-OrganizationalUnit -start OU=LU,DC=w-files,DC=pl -loadAll)

        allows to select OU starting from OU=LU and preloading entire tree
    .EXAMPLE 
        $ou = select-OU -object
        $ou.ObjectGUID

        returns full OU object instead of distinguishedName only. 
    .INPUTS
        None.
    .OUTPUTS
        DistinguishedName
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 220203
            last changes
            - 220203 fixed how multichoice is returned
            - 220202 improvements
            - 220201 multichoice initial, load fixes
            - 210520 icons, load improvements, behaviour fixes
            - 210511 return object
            - 210321 loadAll
            - 210317 rootNode, disableRoot
            - 210308 initialized
    
        #TO|DO
        - [optimization] loading and viewing is very slow, using a lot of recursion and re-reading. especialy painful on slow DC connections.
          should use some caching mechanism... 
        - [optimization] multichoice should use additional table to store chosen values - to quickly return values
    #>
    
    param(
        #starting OU (tree root)
        [Parameter(mandatory=$false,position=0)]
            [string]$startingOU=(get-ADRootDSE).defaultNamingContext,
        #root node can't be selected
        [Parameter(mandatory=$false,position=1)]
            [switch]$disableRoot,
        #return OU object instead of string name
        [Parameter(mandatory=$false,position=2)]
            [switch]$object,
        #enable text search box - will load entire tree which might take time, but will allow to search thru entire tree
        [Parameter(mandatory=$false,position=3)]
            [switch]$loadAll,
        # do not load leaf objects - only OU structure
        [Parameter(mandatory=$false,position=4)]
            [validateSet('computer','user','group','organizationalUnit')]
            [string]$filterObject,
        #enable multichoice
        [Parameter(mandatory=$false,position=5)]
            [switch]$multichoice,
        #if critical - will exit instead of returning false
        [Parameter(mandatory=$false,position=6)]
            [switch]$isCritical
    )

    Function add-Nodes {
        param(
            [Parameter(Mandatory=$true,position=0)]
                $node,
            [Parameter(Mandatory=$false,position=1)]
                [int]$localDepth=0,
            [Parameter(Mandatory=$false,position=2)]
                [ValidateSet('computer','group','user','organizationalUnit')]
                [string[]]$filterObject
        )

        write-verbose ("(adding node) {0} | type: {1} | deph: {2}" -f $node.text, $node.tag.type, $localDepth)
        if($node.tag.type -notmatch 'organizationalUnit|container|root') { return } 
        if([string]::IsNullOrEmpty($filterObject)) {
            $Filter = "*"
        } elseif($filterObject -eq 'organizationalUnit') {
            $Filter="objectClass -eq 'organizationalUnit' -or objectClass -eq 'container'"
        } else {
            $Filter="objectClass -eq '$filterObject' -or objectClass -eq 'organizationalUnit' -or objectClass -eq 'container'"
        } 
        if($node.tag.unfolded -eq $false ) {
            try {
                $OUobjects = get-ADObject -Filter $Filter -SearchBase $node.tag.distinguishedName -SearchScope OneLevel|Sort-Object @{E={$_.ObjectClass};Ascending=$false},name
            } catch {
                write-log "error getting objects using provided values. $($_.exception)" -type error
                break
            }
            $node.tag.unfolded = $true
            foreach($obj in $OUobjects) {
                if([string]::isNullOrEmpty($obj.name)) { continue }
                try {
                    $NodeSub = $Node.Nodes.Add($obj.Name)
                } CATCH {
                    write-host 'err' -ForegroundColor red
                    $obj
                }
                $NodeSub.tag = [psobject]@{
                    distinguishedName = $obj.DistinguishedName
                    unfolded = $false
                    name = $rxADObjName.Match($obj.DistinguishedName).groups[1].value
                    type = $obj.objectClass
                }

                switch($Obj.objectClass) {
                    'computer' { $NodeSub.ImageIndex = $nodeSub.SelectedImageIndex = 2 }
                    'group' { $NodeSub.ImageIndex = $nodeSub.SelectedImageIndex = 3 }
                    'user' { $NodeSub.ImageIndex = $nodeSub.SelectedImageIndex = 4 }
                    'contact' { $NodeSub.ImageIndex = $nodeSub.SelectedImageIndex = 5 }
                    {($_ -eq 'organizationalUnit') -or ($_ -eq 'container')} { 
                        $NodeSub.ImageIndex = 0
                        $nodeSub.SelectedImageIndex = 1 
                    }
                    default { 
                        #write-log "unknown AD object type: $($obj.objectClass)"
                        $NodeSub.ImageIndex = $nodeSub.SelectedImageIndex = 6
                    }
                }
                $script:NodeList += $NodeSub.tag
                if($localDepth -gt 0) { 
                    $addNodes=@{
                        node = $NodeSub
                        localDepth = $localDepth - 1
                    }
                    if($filterObject) {
                        $addNodes.add('filterObject',$filterObject)
                    }
                    add-Nodes @addNodes
                }
            }
        } else {
            foreach($SubNode in $node.Nodes) {
                if($localDepth -gt 0) { 
                    $addNodes=@{
                        node = $SubNode
                        localDepth = $localDepth -1
                    }
                    if($filterObject) {
                        $addNodes.add('filterObject',$filterObject)
                    }
                    add-Nodes @addNodes
                }
            }
        }
    }

    Function list-Nodes {
        param(
            #reference to an object
            [parameter(mandatory=$true,position=0)]
                $nodeSet,
            #flat - for searchlist and treeview for regular treeview object
            [parameter(mandatory=$false,position=1)]
                [validateSet('flat','treeView')]
                $type='treeView'
        ) 

        write-verbose ("(list nodes) {0} : {1}" -f $nodeSet.checked,$nodeSet.Tag.distinguishedName)
        if($type -eq 'treeView') {
            foreach($node in $nodeSet.nodes) {
                list-Nodes $node
            }
            if($nodeSet.checked) {
                if([string]::isNullOrEmpty($filterObject) -or ($filterObject -eq $nodeSet.tag.type) ) {
                    $nodeSet.tag.distinguishedName
                }
            }
        } else {
            foreach($node in $nodeSet.nodes) {
                if($node.checked) {
                    if([string]::isNullOrEmpty($filterObject) -or ($filterObject -eq $node.tag.type) ) {
                        $node.text
                    }
                }
            }
        }
    }

    Function check-Nodes {
        param(
            [parameter(Mandatory=$true,Position=0)]
                $nodeSet,
            [parameter(Mandatory=$true,Position=1)]
                [bool]$set
        )

        foreach($node in $nodeSet.Nodes) {
            $node.checked = $set
            check-Nodes -nodeSet $node -set $set
        }
    }

    [regex]$rxADObjName="^(?:OU|CN)=(.*?),"
    $script:NodeList=@()

    #region FORM
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $Form.Text = "Select AD object under $startingOU"
    $form.MinimumSize = New-Object System.Drawing.Size(300,500)
    $Form.AutoSize = $true
    $Form.StartPosition = 'CenterScreen'
    $Form.Icon = [System.Drawing.SystemIcons]::Question
    $Form.Topmost = $true
    $Form.MaximizeBox = $false
    $form.dock = "fill"
   
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Dock = 'Fill'
    if($multichoice.IsPresent) {
        $treeView.CheckBoxes = $true
    } else {
        $treeView.CheckBoxes = $false
    }
    $treeView.Name = 'treeView'

        $treeViewImageList = new-object System.Windows.Forms.ImageList
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 4 -fileContaining 'imageres.dll') ) #0 folder
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 6 -fileContaining 'imageres.dll') ) #1 opened folder
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 104 -fileContaining 'imageres.dll') ) #2 computer
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 74 -fileContaining 'imageres.dll') ) #3 group
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 208 -fileContaining 'imageres.dll') ) #4 user
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 124 -fileContaining 'imageres.dll') ) #5 contact
        $treeViewImageList.Images.Add( (get-Icon -iconNumber 63 -fileContaining 'imageres.dll') ) #6 other
    $treeView.ImageList = $treeViewImageList
    $treeView.ImageIndex = 0
    $treeView.SelectedImageIndex = 1

    $treeview.add_afterSelect({
        param($sender,$e)
        [System.Windows.Forms.Cursor]::Current = 'WaitCursor'
        [System.Windows.Forms.Application]::UseWaitCursor=$true
        if(!$script:COLLAPSING) {
            $okButton.Enabled = $true
            $e.node.Expand()
            $addNodes=@{
                node = $treeView.SelectedNode
                localDepth = 1
            }
            if($filterObject) {
                $addNodes.add('filterObject',$filterObject)
            }
            add-Nodes @addNodes
            if($treeView.SelectedNode.text -eq $startingOU -and $disableRoot) {
                $okButton.Enabled = $false
            } else {
                $okButton.Enabled = $true
            }
        } else {
            $script:COLLAPSING = $false
        }
        [System.Windows.Forms.Application]::UseWaitCursor=$false
    })
    $treeView.add_BeforeCollapse({
        $script:COLLAPSING = $true
    })
    $treeview.add_beforeExpand({
        param($sender, $e)
        $treeView.SelectedNode = $e.node
    })
    $treeView.add_afterCheck({
        param($sender, $e)
        [System.Windows.Forms.Cursor]::Current = 'WaitCursor'
        [System.Windows.Forms.Application]::UseWaitCursor=$true
        $addNodes=@{
            node = $e.node
            localDepth = 1000
        }
        if($filterObject) {
            $addNodes.add('filterObject',$filterObject)
        }
        add-Nodes @addNodes
        check-Nodes -nodeSet $e.node -set $e.node.checked
        [System.Windows.Forms.Application]::UseWaitCursor=$false
        [System.Windows.Forms.Cursor]::Current = 'Default'
    })
     
    $rootNode = $treeView.Nodes.Add($startingOU)
    $rootNode.Tag=[psobject]@{
        distinguishedName = $startingOU
        unfolded = $false
        type = "root"
    }
    
    $SearchTreeView = New-Object System.Windows.Forms.TreeView
    $SearchTreeView.Dock = 'Fill'
    if($multichoice.IsPresent) {
        $SearchTreeView.CheckBoxes = $true
    } else {
        $SearchTreeView.CheckBoxes = $false
    }
    $SearchTreeView.name = 'SearchTreeView'
     
    $SearchTreeView.add_afterSelect({
        $okButton.Enabled = $true
    })    
    $SearchTreeView.add_afterCheck({
        $okButton.Enabled = $true
    })    
    
    $txtSearch = New-Object system.Windows.Forms.TextBox
    $txtSearch.multiline = $false
    $txtSearch.ReadOnly = $false
    $txtSearch.MinimumSize = new-object System.Drawing.Size(300,20)
    $txtSearch.Height = 20
    $txtSearch.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
    $txtSearch.Location = new-object System.Drawing.Point(3,3)
    $txtSearch.add_KeyUp({
        #param($sender,$e)
        if($txtSearch.Text.Length -gt 1 -and ($mainTable.Controls|Where-Object name -eq 'treeView')) {
            $mainTable.Controls.Remove($treeView)
            $mainTable.controls.add($searchTreeView,1,0)
            $mainTable.SetColumnSpan($searchTreeView,2)
            $form.Refresh()
        } 
        if($txtSearch.Text.Length -le 1) {
            $mainTable.Controls.Remove($searchTreeView)
            $mainTable.Controls.Add($treeView,1,0)
            $form.Refresh()
        }
        if($txtSearch.Text.Length -gt 1) {
            $searchTreeView.Nodes.Clear()
            foreach($n in $NodeList) {
                if($n.name -match $txtSearch.Text) {
                    $searchTreeView.Nodes.Add($n.distinguishedName)
                }
            }
        }
    })
    $txtSearch.add_gotFocus({
        $okButton.Enabled = $false
    }) 

#region MAINFORM-END
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Anchor = 'left'
    $okButton.Text = "OK"
    $okButton.Enabled = $false
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.anchor = 'right'
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.CancelButton = $cancelButton

    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.AutoSize = $true
    $mainTable.ColumnCount = 2
    $mainTable.RowCount = 3
    $mainTable.Dock = "fill"
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    [void]$mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )
    
    $mainTable.controls.Add($txtSearch,0,0)
    $mainTable.SetColumnSpan($txtSearch,2)
    $mainTable.Controls.Add($treeView,1,0)
    $mainTable.SetColumnSpan($treeView,2)
    $mainTable.Controls.add($okButton,2,0)
    $mainTable.Controls.add($cancelButton,2,1)
    
    $form.Controls.Add($mainTable)
#endregion MAINFORM-END

#region SCRIPT_BODY
    $COLLAPSING = $false
    $DEPTH = 1
    if($loadAll) {
        $DEPTH = 1000
        write-log "LOADING FULL TREE..." -type warning
    } else {
        write-log "loading..." -type info
    }
    $addNodes=@{
        node = $rootNode
        localDepth = $DEPTH
    }
    if($filterObject) {
        $addNodes.add('filterObject',$filterObject)
    }
    #[System.Windows.Forms.Application]::UseWaitCursor=$true
    $form.UseWaitCursor=$true
    add-Nodes @addNodes
    $treeView.nodes[0].Expand()
    #[System.Windows.Forms.Application]::UseWaitCursor=$false
    $form.UseWaitCursor=$false

    $result = $Form.ShowDialog()
#endregion SCRIPT_BODY

#region results
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $currentView=($mainTable.controls|? name -match 'treeView')
        if($currentView.name -eq 'treeView') { #return from regular treeview
            if($object.IsPresent) { #return as object
                if($multichoice.IsPresent) {
                    return (list-Nodes $currentView|%{Get-ADObject $_ -properties *})
                } else {
                    return (Get-ADObject $currentView.SelectedNode.tag.distinguishedName -properties *)
                }
            } else { #return as string - default
                if($multichoice.IsPresent) {
                    return (list-Nodes $currentView)
                } else {
                    return $currentView.SelectedNode.tag.distinguishedName
                }
            }
        } else { #return from 'search' space - which is not a treeview anymore
            if($object.IsPresent) {
                if($multichoice.IsPresent) {
                    return (list-Nodes $currentView -type flat|%{Get-ADObject $_ -properties *})
                } else {
                    return (Get-ADObject $currentView.SelectedNode.text -properties *)
                }
            } else {
                if($multichoice.IsPresent) {
                    return (list-Nodes $currentView -type flat)
                } else {
                    return $currentView.SelectedNode.text
                }
            }
        }

    } 
    if($isCritical.IsPresent) {
        write-log "cancelled."
        break
    } 
    return $false
#endregion results
 
}

function select-OrganizationalUnit {
    <#
    .SYNOPSIS
        proxy function for backward compatibility - replaced by select-ADObject
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210810
            last changes
            - 210810 initialized
    
        #TO|DO
    #>
    param(
        #starting OU (tree root)
        [Parameter(mandatory=$false,position=0)]
            [string]$startingOU=(get-ADRootDSE).defaultNamingContext,
        #root node can't be selected
        [Parameter(mandatory=$false,position=1)]
            [switch]$disableRoot,
        #return OU object instead of string name
        [Parameter(mandatory=$false,position=2)]
            [switch]$object,
        #enable text search box - will load entire tree which might take time, but will allow to search thru entire tree
        [Parameter(mandatory=$false,position=3)]
            [switch]$loadAll,
        #if critical - will exit instead of returning false
        [Parameter(mandatory=$false,position=4)]
            [switch]$isCritical
    ) 
    $runParam = @{
        startingOU = $startingOU
        filterObject = 'organizationalUnit'
    }
    if($disableRoot) { $runParam.Add('disableRoot',$true) }
    if($object) { $runParam.Add('object',$true) }
    if($loadAll) { $runParam.Add('loadAll',$true) }
    if($isCritical) { $runParam.Add('isCritical',$true) }
    select-ADObject @runParam
}
set-alias -Name select-OU -Value select-OrganizationalUnit

#################################################### connection checkers
function get-ExchangeConnectionStatus {
    <#
    .SYNOPSIS
        check Ex/EXO connection status.
    .DESCRIPTION
        Exchange is using Remoting commands. this function verifies if session connection exists.
    .EXAMPLE
        get-ExchangeConnectionStatus -isCritical

        checks connection and exits if not present.
    .EXAMPLE
        if(-not get-ExchangeConnectionStatus) {
            write-log "you should connect to Exchange first, scrpt will run with limited options"
        }

        checks connection and warns about lack of Exchange connectivity.
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 validation domain added
            - 210208 initialized
    
        #TO|DO
        - isCritical flag
        - verify domain name

    #>
    
    param(
        #define if connection to EXO or Ex Onprem
        [parameter(mandatory=$false,position=0)]
            [validateSet('OnPrem','EXO')]
            [string]$ExType='EXO',
        #if connection is not established exit with error instead of returning $false.
        [parameter(mandatory=$false,position=1)]
            [switch]$isCritical,
        #you can enforce check against particular tenant so script does not run in improper tenant
        [parameter(mandatory=$false,position=2)]
            [string]$validateDomainName
            
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
    if($isCritical.IsPresent -and !$exConnection) {
        write-log "connection to $ExType not established. quitting." -type error
        exit
    }
    if($validateDomainName) {
        $connectionDomainName = (Get-AcceptedDomain | Where-Object default).Name
        if($connectionDomainName -ne $validateDomainName) {
            write-log "conection established to $connectionDomainName but session expected to $validateDomainName. " -type error
            if($isCritical.IsPresent) {
                exit
            } else {
                $exConnection = $false
            }
        }
    }
    return $exConnection
}
function get-AzureADConnectionStatus {
    param(
        #defines if connection is critical (will exit). by default script will return $null or defalt azure domain name.
        [parameter(mandatory=$false,position=0)]
            [switch]$isCritical        
    )

    $testAAD=$null
    try {
        $testAAD=( Get-AzureADDomain |Where-Object isDefault ).name
    } catch {
        if($isCritical) {
            write-Log "connection to AAD not established. please use connect-AzureAD first. quitting" -type error
            exit         
        } else {
            write-Log "connection to AAD not established." -type warning
        }
    }
    return $testAAD
}
function connect-Azure {
    <#
    .SYNOPSIS
        quick Azure connection check by verifying AzContext.
    .DESCRIPTION
        there is no life session to Azure. Az commandlets are using saved AzContext and token. when 
        token expires, context is returned, but connection attemt will return error. to clean it up
        - best is to clear context and exforce re-authentication.

        function is checking azcontext and test connection by calling get-AzTenant. clears context if 
        connection is broken.
    .EXAMPLE
        connect-Azure

        checks AzContext and connection health
    .EXAMPLE
        connect-Azure

        checks AzContext and connection health
    .INPUTS
        None.
    .OUTPUTS
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 241110
            last changes
            - 241110 Environments
            - 210302 fix to expired token - PSMessageDetail is not populated on many OSes. why? 
            - 210301 proper detection of expired tokens
            - 210220 proper handiling of consonle call - return instead of exit
            - 210219 extended handling of context expiration
            - 210208 initialized
    
        #TO|DO
    #>
    [CmdletBinding()]
    param (
        #Provide cloud type 
        [Parameter(mandatory=$false,position=0)]
        [validateSet('AzureCloud','AzureChinaCloud','AzureUSGovernment')]
            [string]$Environment = 'AzureCloud',
        #confirm the connection
        [Parameter(mandatory=$false,position=1)]
            [switch]$Confirm
        
    )
    Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"

    try {
        $AzSourceContext = Get-AzContext
    } catch {
        write-log $_.exception -type error
        write-log "trying to fix" -type info
        Clear-AzContext -Force
    }
    if([string]::IsNullOrEmpty( $AzSourceContext ) ) {
        write-log "you need to be connected before running this script. use connect-AzAccount first." -type warning
        $AzSourceContext = Connect-AzAccount -Environment $Environment -ErrorAction SilentlyContinue
        if([string]::isNullOrEmpty($AzSourceContext) ) {
            write-log "cancelled"
            if( (Get-PSCallStack).count -gt 2 ) { #run from script
                exit
            } else { #run from console  
                return $null
            }          
        }
        $AzSourceContext = Get-AzContext
    } else { 
        #1. check if context is from a proper Cloud Environment
        if($AzSourceContext.Environment.Name -ne $Environment) {
            write-Log -message "different Cloud environment connected: $($AzSourceContext.Environment.Name); requested: $Environment." -type error
            Clear-AzContext -Force
            write-log "re-run the script."
            if( (Get-PSCallStack).count -gt 2 ) { #run from script
                exit
            } else { #run from console  
                return $null
            }
        }
        #2. token exist, check if it is still working
        try{
            #if access token has been revoked, Az commands return warning "Unable to acquire token for tenant"
            Get-AzSubscription -WarningAction stop|Out-Null
        } catch {
            if($_.Exception -match 'Unable to acquire token for tenant') {
                write-log "token expired, clearing cache" -type info
                Clear-AzContext -Force
                write-log "re-run the script."
                if( (Get-PSCallStack).count -gt 2 ) { #run from script
                    exit
                } else { #run from console  
                    return $null
                }
            } else {
               write-log $_.exception
               return -3
            }
        }
    }
    write-log "connected to $($AzSourceContext.Subscription.name) as $($AzSourceContext.account.id)" -silent -type info
    write-host "Your Azure connection:"
    write-host "  subscription: " -noNewLine
    write-host -foreground Yellow "$($AzSourceContext.Subscription.name)"
    write-host "  connected as: " -noNewLine 
    write-host -foreground Yellow "$($AzSourceContext.account.id)"
    if($Confirm.IsPresent) {
        write-log "Is that a correct account? (press 'y' to continue)" -type warning -skipTimestamp
        $keyPress = [console]::ReadKey($true)
        if ($keyPress.KeyChar -ne 'y') {
            Write-log "wrong account. disconnecting." -type warning
            exit -1
        }        
    }
}

Export-ModuleMember -Function * -Alias 'load-CSV','select-OU','convert-XLS2CSV','convert-CSV2XLS'

