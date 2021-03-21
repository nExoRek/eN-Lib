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
    version 210321
    changes
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
$logFile=''

function start-Logging {
    <#
    .SYNOPSIS
        initilizes log file under $logFile variable for write-log function.
    .DESCRIPTION
        all scripts require logging mechinism. write-log function forking each output to screen and to logfile
        is a very convinient way keeping all logs consistent with very nice host output. 
        this function initilizeds $logFile variable creation, and initiates the log file itself. in order to 
        ease the creation there are several ways of initilizing $logFile:
        - using write-log directly
            - from console host
            - from script
        - using this function directly 
            - no parameters - defaults to $ScriptRoot/Logs folder
            - using 'useProfile' - to store logs in User Documents/Logs directory
            - using 'logFolder' parameter to define particular folder for logs
            - using 'logFileName' - (exclusive to logFolder) presenting full path for the log or logfile name 
    .EXAMPLE
        write-log 

        using write-log will inderctly call start-logging function and initializes the log file with generic name,
        saved in 'Logs' subfolder under script run path. 
    .EXAMPLE
        start-Logging -logFileName c:\temp\myLogs\somelog.log
        write-log 'test message'

        initializes the log file as c:\temp\myLogs\somelog.log .
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

        initializes the log file in Logs subfolder uder user profile path
    .INPUTS
        None.
    .OUTPUTS
        log file under $logFile variable.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210210
        changes:
            - 210210 removing recurrency to write-log (loop elimination)
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

    #write-host -ForegroundColor red ">>$($MyInvocation.PSCommandPath)<<"
    if( [string]::isNullOrEmpty($MyInvocation.PSCommandPath) ) { #if run from console - set the logfile name as 'console'
        $scriptBaseName = 'console'
        $script:lastScriptUsed = 'console'
    } elseif( $MyInvocation.PSCommandPath -match '\.psm1$' ) { #if run from inside module...
        if([string]::isNullOrEmpty($script:lastScriptUsed) ) {
            $scriptBaseName = 'console'
        } else {
            $scriptBaseName = $script:lastScriptUsed
        }
    } else { #regular run from the script - take a script basename as logfile name
        $scriptBaseName = ([System.IO.FileInfo]$($MyInvocation.PSCommandPath)).basename
        $script:lastScriptUsed = $scriptBaseName
    }
    switch($PSCmdlet.ParameterSetName) {
        'userProfile' {
            $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
            $global:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        'Folder' {
            $global:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        }
        'filePath' {
            if($scriptBaseName -eq 'console') {
                $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
                $global:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)   
            }elseif ( [string]::IsNullOrEmpty($logFileName) ) {          #by default 'filepath' is used and empty 
                $logFolder="$($MyInvocation.PSScriptRoot)\Logs"
                $global:logFile = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)   
            } else {
                #logfile can be: 1. file, 2. folder, 3. fullpath
                $logFolder = Split-Path $logFileName -Parent
                if([string]::isNullOrEmpty($logFolder) ) { #logfile name without full path, name only
                    if( [string]::isNullOrEmpty($MyInvocation.PSScriptRoot) ) { #run directly from console
                        $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
                    } else {
                        $logFolder = $MyInvocation.PSScriptRoot
                    }
                }
                $logFile = Split-Path $logFileName -Leaf
                if( test-path $logFile -PathType Container ) {
                    write-host "$logFileName seems to be an existing folder. use 'logFolder' parameter or change log name. quitting." -ForegroundColor Red
                    exit -1
                }
                $global:logFile = "$logFolder\$logFile"
            }
        }
        default {
            write-host -ForegroundColor Magenta 'very strange error'
            exit -666
        }
    }

    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder -ErrorAction Stop|Out-Null
            write-host "$LogFolder created."
        } catch {
            write-error $_.exception
            exit -2
        }
    }
    "*logging initiated $(get-date) in $($global:logFile)"|Out-File $global:logFile -Append
    write-host "*logging initiated $(get-date) in $($global:logFile)"
    "*script parameters:"|Out-File $global:logFile -Append
    if($script:PSBoundParameters.count -gt 0) {
        $script:PSBoundParameters|Out-File $global:logFile -Append
    } else {
        "<none>"|Out-File $global:logFile -Append
    }
    "***************************************************"|Out-File $global:logFile -Append
}
function write-log {
    <#
    .SYNOPSIS
        replacement for write-host, forking information to a log file and screen.
    .DESCRIPTION
        automates forking of output on two different endpoints - on the host, using write-host
        and to the file, appening its content.
        write-log converts everything to a string, so you can use it for virtually any type of 
        variable. additionaly it adds timestamp, message type header and color (on host).

        in order to use write-log, $logFile variable requires to be set up. you can initialize the 
        value directly with 'start-Logging', configure $logFile manually or simply run write-log to 
        have it initialized automatically. by default logs are stored in $PSScriptRoot/Logs directory 
        with generic file name. if you need special location refer to start-logging help how to
        initialize variable. 

        function may also be used from command line - in this scenario log file will be created in 
        Logs directory under User Documents folder. file with be named 'console-<date>.log'.

    .EXAMPLE
        write-log "all is fine"

        output 'all is fine' on the screen and to the log file.
    .EXAMPLE
        write-log all is fine

        outputs 
            all 
            is 
            fine
        on the screen and to the log file - all unnamed parameters are displayed
    .EXAMPLE
        write-log -message "trees are green" -type ok

        shows 'trees are green' in Green colour, and send text to a log file.
    .EXAMPLE
        $someObject=get-process
        write-log -message $someObject -type info -noTimestamp -silent

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
        version 210302
        changes:
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
    
    param(
        #message to display - can be an object
        [parameter(ValueFromRemainingArguments=$true,mandatory=$false,position=0)]
            $message,
        #adds description and colour dependently on message type
            [string][validateSet('error','info','warning','ok')]$type,
        #do not output to a screen - logfile only
            [switch]$silent,
        # do not show timestamp with the message
            [switch]$skipTimestamp
    )

    #$callStack = (Get-PSCallStack |Where-Object ScriptName)[-1]
    $callStack = (Get-PSCallStack)[-2]
    if( [string]::isNullOrEmpty($MyInvocation.PSCommandPath) ) { #it's run directly from console.
        $scriptBaseName = 'console'
        $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
    } elseif( -not [string]::isNullOrEmpty( $callStack ) ){ 
        if( $callStack.ScriptName -match '\.psm1$' ) { #run from inside module
            $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
            $scriptBaseName = 'console'
        } else {
            $logFolder = "$(split-path $callStack.scriptname -Parent)\Logs"
            $scriptBaseName = ([System.IO.FileInfo]$callStack.scriptname).basename
        }
    } else {
        $scriptBaseName = ([System.IO.FileInfo]$($MyInvocation.PSCommandPath)).basename 
        $logFolder = "$($MyInvocation.PSScriptRoot)\Logs"
    }
    if( [string]::isNullOrEmpty($global:logFile) -or ( $script:lastScriptUsed -ne $scriptbasename) ) {   
        $script:lastScriptUsed = $scriptBaseName
        $logFileName = "{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
        start-Logging -logFileName $logFileName
    }

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if($null -eq $message) {$message=''}
    if($message.GetType().name -ne 'string') {
        $message=($message|out-String).trim() 
    }

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
        Add-Content -Path $global:logFile -Value $message
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
function get-CSVDelimiter {
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

    #this is very simple delimiter check based on number of columns detected with , and ;
    $FirstLines=$FirstLines -replace '''.*?''|".*?"','ANTI-DELIMITER' #change all quoted strings to simple string to avoid quoted delimiter characters
    $delimiter=',' #set default
    $semi0=$FirstLines[0].split(';').Length -1 
    $semi1=$FirstLines[1].split(';').Length -1
    $colon0=$FirstLines[0].split(',').Length -1
    $colon1=$FirstLines[1].split(',').Length -1
    if( (($semi0 -eq $semi1) -and ($colon0 -ne $colon1)) -or `
        (($semi0 -eq $semi1) -and ($semi0 -gt $colon0)) -or `
        ($colon0 -eq 0 -and $semi0 -gt 0)
        ){
        $delimiter=';'
    } 
    write-log "$delimiter detected as delimiter." -type info
    return $delimiter
}
function import-structuredCSV {
    <#
    .SYNOPSIS
        loads CSV file with header check and auto delimiter detection 
    .DESCRIPTION
        support function to gather data from CSV file with ability to ensure it is correct CSV file by
        enumerating header. if you operate on data you need to ensure that it is CORRECT file, and not some
        random CSV. extremally usuful in the projects when you use xls/csv as data providers and need to ensure
        that you stick to the standard column names.
        with non-critical header function allows to add missing columns.
    .EXAMPLE
        $inputCSV = "c:\temp\mydata.csv"
        $header=@('column1','column2')
        $data = load-CSV -header $header -headerIsCritical -delimiter ';' -inputCSV $inputCSV
        
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210317
            last changes
            - 210317 delimiter detection as function
            - 210315 finbished auto, non-terminating from console, header not mandatory
            - 210311 auto delimiter detection, min 2lines
            - 210219 initialized
    
        #TO|DO
        - add nonCritical header + crit header handling
        - silent mode
    #>
    param(
        #path to CSV file containing data
        [parameter(mandatory=$true,position=0)]
            [string]$inputCSV,
        #expected header
        [parameter(mandatory=$false,position=1)]
            [string[]]$header,
        #this flag causes exit on load if any column is missing. 
        [parameter(mandatory=$false,position=2)]
            [switch]$headerIsCritical,
        #CSV delimiter if different then regional settings. auto - tries to detect between comma and semicolon. uses comma if failed.
        [parameter(mandatory=$false,position=3)]
            [string]$delimiter='auto'
    )

    $exit=$true #should process exit or return? return when run from console, exit when from script.
    if( (Get-PSCallStack).count -le 2 ) { #run from console
        $exit=$false
    }          

    if(-not (test-path $inputCSV) ) {
        write-log "$inputCSV not found." -type error
        if($exit) {
            exit -1
        } else {
            return -1
        }
    }

    if($delimiter='auto') {
        $delimiter=get-CSVDelimiter -inputCSV $inputCSV
        if($null -eq $delimiter) {
            if($exit) {
                exit -1
            } else {
                return -1
            }
        }
    }

    try {
        $CSVData=import-csv -path "$inputCSV" -delimiter $delimiter -Encoding UTF8
    } catch {
        Write-log "not able to import $inputCSV. $($_.exception)" -type error 
        if($exit) {
            exit -2
        } else {
            return -2
        }
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
            if($exit) {
                exit -3
            } else {
                return -3
            }
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
                if($exit) {
                    exit 0
                } else {
                    return 0
                }
            }
        }
    }
    return $CSVData
}
set-alias -Name load-CSV -Value import-structuredCSV

function convert-XLStoCSV {
    <#
    .SYNOPSIS
        export all tables in XLSX files to CSV files. enumerates all sheets, and each table goes to another file.
        if sheet does not contain table - whole sheet is saved as csv
    .DESCRIPTION
        if file contain information out of table objects - they will be exported as a whole worksheet.
        files will be named after the sheet name + table/worksheet name and placed in seperate directory.

        separate script with ability to drang'n'drop may be downloaded from
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
        version 210317
            last changes
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
            [switch]$includeHiddenWorksheets
    )

    begin {
        $exit=$true #should process exit or return? return when run from console, exit when from script.
        if( (Get-PSCallStack).count -le 2 ) { #run from console
            $exit=$false
        }          
        try{
            $Excel = New-Object -ComObject Excel.Application
        } catch {
            write-log "not able to initialize Excel lib. requires Excel to run" -type error
            if($exit) {
                exit -1
            } else {
                return -1
            }
        }
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
    }

    process {
        if($PSCmdlet.ParameterSetName -eq 'byName') {
            if(-not (test-path $XLSFileName)) {
                write-log "$XLSfileName not found. exitting" -type error
                if($exit) {
                    exit -2
                } else {
                    return -2
                }
            }
            $XLSFile=get-Item $XLSfileName
        }
        if($XLSFile.Extension -notmatch '\.xls') {
            write-log "$($XLSFile.Name) doesn't look like excel file. exitting" -type error
            if($exit) {
                exit -3
            } else {
                return -3
            }
        }
        $outputFolder=$XLSFile.DirectoryName+'\'+$XLSFile.BaseName+'.exported'
        if( -not (test-path($outputFolder)) ) {
            new-Item -ItemType Directory $outputFolder|Out-Null
        }
        $workBookFile = $Excel.Workbooks.Open($XLSFile)

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
        #any method of closing Excel file is not working 1oo% there are scenarios where excel process stays in memory.
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
        creates XLXS out of CSV file and formats data as a table.
    .EXAMPLE
        convert-CSV2XLSX c:\temp\test.csv -delimiter ','
        
        Converts test.csv to test.xlsx 
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
        version 210317
            last changes
            - 210317 processing multiple CSV will create single XLS, delimiter autodetection, output file name
            - 210309 proper pipelining
            - 210308 module function
            - 201123 initialized
        
        TO|DO
        - add silent mode 
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
        #CSV delimiter character
        [Parameter(mandatory=$false,position=4)]
            [string]$delimiter='auto'
    )

    begin {
        $exit=$true #should process exit or return? return when run from console, exit when from script.
        if( (Get-PSCallStack).count -le 2 ) { #run from console
            $exit=$false
        }          
    
        try{
            $Excel = New-Object -ComObject Excel.Application
        } catch {
            write-host -ForegroundColor Red "not able to initialize Excel lib. requires Excel to run"
            write-host -ForegroundColor red $_.Exeption
            if($exit) {
                exit -1
            } else {
                return -1
            }
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
            $parent = Split-Path $XLSfileName -Parent
            if([string]::isNullOrEmpty($parent)) {
                $XLSfileName = ($pwd).path +'\'+$XLSfileName
            } else {
                $XLSFileName = (Resolve-Path $parent).Path + '\' + (Split-Path $XLSfileName -Leaf)
            }
            if($XLSFileName -notmatch "\.xls[x]?") {
                $XLSfileName+=".xlsx"
            }
            write-log "creating $XLSfileName excel file..." -type info
        }
    }

    process {
        #$ErrorActionPreference="SilentlyContinue"
        if($PSCmdlet.ParameterSetName -eq 'byName') {
            if(-not (test-path $CSVfileName) ) {
                write-host -ForegroundColor Red "file $CSVfileName is not accessible"
                if($exit) {
                    exit -2
                } else {
                    return -2
                }
            }
            $CSVFile=get-childItem $CSVfileName
        } 
        #convert output file name to full path
        if([string]::isNullOrEmpty($XLSfileName)) {
            $XLSfileName= ($pwd).path + '\' + $CSVFile.BaseName + '.xlsx'
            write-log "creating $XLSfileName excel file..." -type info
        } 

        try {
            write-log "adding $($CSVfile.Name) data as worksheet..." -type info
            if($autoDelimiter) {
                $delimiter = get-CSVDelimiter -inputCSV $CSVfile.FullName
                if([string]::isNullOrEmpty($delimiter) ) {  
                    $delimiter=','
                    #return -1
                }
            }
            if($counter++ -gt 0) {
                $worksheet = $workbook.worksheets.add([System.Reflection.Missing]::Value,$workbook.Worksheets.Item($workbook.Worksheets.count))
            }
            $worksheet = $workbook.worksheets.Item($workbook.Worksheets.count)
            $worksheet.name = $CSVfile.BaseName
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

            $Table = $worksheet.ListObjects.Add(
                [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
                $Range, 
                "importedCSV",
                [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
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
            $Table.TableStyle = "TableStyle$tableStyle" #green with with gray shadowing

        } catch {
            write-log "error converting CSV to XLS: $($_.exception)" -type error
            return -2         
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
        version 210317
            last changes
            - 210317 use of firstWorksheet
            - 210315 initialized
    
        #TO|DO
    #>
    
    param(
        # XLSX file name to be converted to CSV files
        [Parameter(ParameterSetName='byName',mandatory=$true,position=0)]
            [string]$XLSfileName
    )

    write-log "EXERIMAENTAL FUNCTION" -type warning
    $tempCSV = convert-XLStoCSV -XLSfileName $XLSfileName -firstWorksheetOnly
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
            exit -1
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
        version 210317
        last changes
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
            [regex]$allowedCharacters
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
    $txtUserInput.MaxLength = $maxChars
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
            if($txtUserInput.text[-1] -notmatch $allowedCharacters) {
                $cursor=$txtUserInput.SelectionStart
                $txtUserInput.text=$txtUserInput.text -replace ".$"
                $txtUserInput.Select($cursor-1, 0);
            }
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

function select-OrganizationalUnit {
    <#
    .SYNOPSIS
        accelerator function allowing to select OU with GUI.
    .DESCRIPTION
        function is using winforms treeView to display OU structure. returns DistinguishedName on select.
    .EXAMPLE
        $ou = select-OU
        
        displays forms treeview enabling to choose OU from the tree.
    .EXAMPLE
        new-ADComputer -name 'mycomputer' -path (select-OrganizationalUnit -start OU=LUcomps,DC=w-files,DC=pl -enableSearch)

        launches search windows with OU structure, starting from 'LUcomps', with search box.
    .INPUTS
        None.
    .OUTPUTS
        DistinguishedName
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210321
            last changes
            - 210321 enableSearch
            - 210317 rootNode, disableRoot
            - 210308 initialized
    
        #TO|DO
        - icons
        - multichoice
    #>
    
    param(
        #starting OU (tree root)
        [Parameter(mandatory=$false,position=0)]
            [string]$startingOU=(get-ADRootDSE).defaultNamingContext,
        #root node can't be selected
        [Parameter(mandatory=$false,position=1)]
            [switch]$disableRoot,
        #enable text search box - will load entire tree - MAY TAKE LONG TIME - use for subOUs rather then full tree 
        [Parameter(mandatory=$false,position=2)]
            [switch]$enableSearch,
        #if critical - will exit instead of returning false
        [Parameter(mandatory=$false,position=3)]
            [switch]$isCritical
    )

    Function add-Nodes ( $Node) {
        #write-host $costam
        $nodeList=@()
        $SubOU = Get-ADOrganizationalUnit -SearchBase $node.tag.distinguishedName -SearchScope OneLevel -filter *
        if($node.tag.unfolded -eq $false) {
            $node.tag.unfolded = $true
            foreach ( $ou in $SubOU ) {
                $NodeSub = $Node.Nodes.Add($ou.Name)
                $NodeSub.tag = [psobject]@{
                    distinguishedName = $ou.DistinguishedName
                    unfolded = $false
                    name = $rxOUName.Match($ou.DistinguishedName).groups[1].value
                }
                $NodeList += $NodeSub.tag
                if($enableSearch) { 
                    add-Nodes $NodeSub 
                }
            }
        }
        return $nodeList
    }

    function get-Nodes( [System.Windows.Forms.TreeNodeCollection] $nodes) {
        foreach ($n in $nodes) {
            $n
            Get-Nodes($n.Nodes)
        }
    }

    [regex]$rxOUName="^OU=(.*?),"

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = New-Object System.Windows.Forms.Form
    $Form.Text = "Select OU under $startingOU"
    #$Form.Size = New-Object System.Drawing.Size(300,500)
    $form.MinimumSize = New-Object System.Drawing.Size(300,500)
    $Form.AutoSize = $true
    $Form.StartPosition = 'CenterScreen'
    #$Form.FormBorderStyle = 'Fixed3D'
    $Form.Icon = [System.Drawing.SystemIcons]::Question
    $Form.Topmost = $true
    $Form.MaximizeBox = $false
    $form.dock = "fill"
   
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Dock = 'Fill'
    $treeView.CheckBoxes = $false
    $treeView.row
    $treeView.Name = 'treeView'
     
    $rootNode = $treeView.Nodes.Add($startingOU)
    $rootNode.Tag=[psobject]@{
        distinguishedName = $startingOU
        unfolded = $false
    }
    $nodeList=add-Nodes $rootNode

    $SearchTreeView = New-Object System.Windows.Forms.TreeView
    $SearchTreeView.Dock = 'Fill'
    $SearchTreeView.CheckBoxes = $false
    $SearchTreeView.row
    $SearchTreeView.name = 'SearchTreeView'
     
    $SearchRootNode = $SearchTreeView.Nodes.Add($startingOU)
    $SearchRootNode.Tag=$startingOU
        
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Anchor = 'left'
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.anchor = 'right'
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.CancelButton = $cancelButton
    
    $txtSearch = New-Object system.Windows.Forms.TextBox
    $txtSearch.multiline = $false
    $txtSearch.ReadOnly = $false
    $txtSearch.MinimumSize = new-object System.Drawing.Size(300,20)
    $txtSearch.Height = 20
    $txtSearch.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
    $txtSearch.Location = new-object System.Drawing.Point(3,3)

    $nodes = get-Nodes($treeView.Nodes)
    $txtSearch.add_KeyUp({
        #param($sender,$e)
        if($txtSearch.Text.Length -gt 1 -and ($mainTable.Controls|? name -eq 'treeView')) {
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

    $mainTable = New-Object System.Windows.Forms.TableLayoutPanel
    $mainTable.AutoSize = $true
    $mainTable.ColumnCount = 2
    $mainTable.RowCount = 3
    $mainTable.Dock = "fill"
    $mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )|out-null
    $mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)) )|out-null
    $mainTable.RowStyles.Add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,30)) )|out-null
    
    $mainTable.controls.Add($txtSearch,0,0)
    $mainTable.SetColumnSpan($txtSearch,2)
    $mainTable.Controls.Add($treeView,1,0)
    $mainTable.SetColumnSpan($treeView,2)
    $mainTable.Controls.add($okButton,2,0)
    $mainTable.Controls.add($cancelButton,2,1)

    
    $form.Controls.Add($mainTable)

    $treeview.add_afterSelect({
        add-Nodes $treeView.SelectedNode 
        if($treeView.SelectedNode.text -eq $startingOU -and $disableRoot) {
            $okButton.Enabled = $false
        } else {
            $okButton.Enabled = $true
        }
    })
    
    $result = $Form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $currentView=($mainTable.controls|? name -match 'treeView')
        if($currentView.name -eq 'treeView') {
            return $currentView.SelectedNode.tag.distinguishedName
        } else {
            return $currentView.SelectedNode.text
        }

    } 
    if($isCritical.IsPresent) {
        write-log "cancelled."
        exit 0
    } 
    return $false
    
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
        exit -1
    }
    if($validateDomainName) {
        $connectionDomainName = (Get-AcceptedDomain | Where-Object default).Name
        if($connectionDomainName -ne $validateDomainName) {
            write-log "conection established to $connectionDomainName but session expected to $validateDomainName. " -type error
            if($isCritical.IsPresent) {
                exit -1
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
            exit -1            
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
        there is no session to Azure and Az commandlets are using saved AzContext and token. when 
        token expires, context is returned, but connection attemt will return error. to clean it up
        - best is to clear context and exforce re-authentication.

        function is checking azcontext and test connection by calling get-AzTenant. clears context if 
        connection is broken.
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
        version 210302
            last changes
            - 210302 fix to expired token - PSMessageDetail is not populated on many OSes. why? 
            - 210301 proper detection of expired tokens
            - 210220 proper handiling of consonle call - return instead of exit
            - 210219 extended handling of context expiration
            - 210208 initialized
    
        #TO|DO
    #>
    
    try {
        $AzSourceContext=Get-AzContext
    } catch {
        write-log $_.exception -type error
        write-log "trying to fix" -type info
        Clear-AzContext -Force
    }
    if([string]::IsNullOrEmpty( $AzSourceContext ) ) {
        write-log "you need to be connected before running this script. use connect-AzAccount first." -type warning
        $AzSourceContext = Connect-AzAccount -ErrorAction SilentlyContinue
        if([string]::isNullOrEmpty($AzSourceContext) ) {
            if( (Get-PSCallStack).count -gt 2 ) { #run from script
                write-log "cancelled"
                exit 0
            } else { #run from console  
                write-log "cancelled"
                return $null
            }          
        }
        $AzSourceContext = Get-AzContext
    } else { #token exist, check if it is still working
        try{
            #if access token has been revoked, Az commands return warning "Unable to acquire token for tenant"
            Get-AzSubscription -WarningAction stop|Out-Null
        } catch {
            if($_.Exception -match 'Unable to acquire token for tenant') {
                write-log "token expired, clearing cache" -type info
                Clear-AzContext -Force
                write-log "re-run the script."
                if( (Get-PSCallStack).count -gt 2 ) { #run from script
                    exit -1
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
    Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"
}

Export-ModuleMember -Function * -Variable 'logFile' -Alias 'load-CSV','select-OU','convert-XLS2CSV','convert-CSV2XLS'
