<#
.SYNOPSIS
    export all tables in XLSX files to CSV files. enumerates all sheets, and each table goes to another file.
    if sheet does not contain table - whole sheet is saved as csv
.DESCRIPTION
    if file contain information out of table objects - they will be exported as a whole worksheet.
    files will be named after the sheet name + table/worksheet name and placed in seperate directory.

    using such exports a lot? why not create a desktop shortcut to just drag'n'drop xlsx files into?
    create a shortcut and type: 
    C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noprofile -file "<path to script>\convert-XLSX2CSV.ps1"
    enjoy quick xlsx->convert with one mouse move (=
.EXAMPLE
    .\convert-XLSX2CSV.ps1 -fileName .\myFile.xlsx

    saves tables/worksheets to CSV files.
.INPUTS
    XLSX file.
.OUTPUTS
    Series of CSV files representing tables and/or worksheets (if lack of tables).
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201121
        last changes
        - 201121 output folder changed, descirption, do not export hidden by default, saveAs CSVUTF8
        - 201101 initialized
#>
[cmdletbinding()]
param(
    # XLSX file to be converted to CSV files
    [Parameter(mandatory=$true,position=0)]
        [string]$fileName,
    #include hidden worksheets? 
    [Parameter(mandatory=$false,position=1)]
        [switch]$includeHiddenWorksheets
)

begin {
    #region initial_checks
    if(-not (test-path $fileName)) {
        write-host "$fileName not found. exitting"
        exit -1
    }
    $file=get-childItem $fileName
    if($file.Extension -ne '.xlsx') {
        write-host "$fileName doen't look like excel file. exitting"
        exit -2
    }
    try{
        $Excel = New-Object -ComObject Excel.Application
    } catch {
        write-host -ForegroundColor Red "not able to initialize Excel lib. requires Excel to run"
        write-host -ForegroundColor red $_.Exeption
        exit -3
    }
    #endregion initial_checks
    #$outputFolder=$PSScriptRoot + '\' + $file.BaseName #decided that output is better to have in original file location rather then script root
    $outputFolder=$file.DirectoryName+'\'+$file.BaseName+'.exported'
    if( -not (test-path($outputFolder)) ) {
        new-Item -ItemType Directory $outputFolder
    }
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $workBookFile = $Excel.Workbooks.Open($fileName)

}
process {
    write-host verbose "converting $fileName tables to CSV files..."

    foreach($worksheet in $workBookFile.Worksheets) {
        if($worksheet.Visible -eq $false -and -not $includeHiddenWorksheets.IsPresent) {
            write-verbose "worksheet $($worksheet.name) found but it is hidden. use -includeHiddenWorksheets to export"
            continue
        }
        Write-Verbose "worksheet $($worksheet.name)"
        $tableList=$worksheet.listObjects|Where-Object SourceType -eq 1
        if($tableList) {
            foreach($table in $tableList ) {
                Write-Verbose "found table $($table.name) on $($worksheet.name)"
                $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^\w\d\-_\.]', '') + '_' + ($table.name -replace '[^\w\d]', '') + '.csv'
                $tempWS=$workBookFile.Worksheets.add()
                $table.range.copy()|out-null
                $tempWS.paste($tempWS.range("A1"))
                $tempWS.SaveAs($exportFileName, 6,$null,$null,$null,$null,$null,$null,$null,'True')
                write-host "$($table.name) saved as $exportFileName"
                $tempWS.delete()
                Remove-Variable -Name tempWS
            }
        } else {
            Write-Verbose "$($worksheet.name) does not contain tables. exporting whole sheet..."
            $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^a-zA-Z0-9\-_]', '') + '_sheet.csv'
            $fileType=62 #CSVUTF8 https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
            $addToMRU=$false #https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.saveas
            $worksheet.SaveAs($exportFileName, $fileType,$null,$null,$null,$null,$sddToMRU,$null,$null,'True')
            write-host "worksheet $($worksheet.name) saved as $exportFileName"
        }
    }
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
    Write-Host -ForegroundColor Green "done and cleared."
}