<#
.SYNOPSIS
    export all tables in XLSX files to CSV files. enumerates all sheets, and each table goes to another file.
    if sheet does not contain table - whole sheet is saved as csv
.DESCRIPTION
    if file contain information out of table objects - they will be exported as a whole worksheet.
    files will be named after the sheet name + table/worksheet name and placed in seperate directory.
.EXAMPLE
    .\convert-XLSX2CSV.ps1 -fileName .\myFile.xlsx

    Explanation of what the example does
.INPUTS
    XLSX file.
.OUTPUTS
    Series of CSV files representing tables and/or worksheets (if lack of tables).
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201101
        last changes
        - 201101 initialized
#>
[cmdletbinding()]
param(
    # XLSX file to be converted to CSV files
    [Parameter(mandatory=$true,position=0)]
        [string]$fileName
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
    $outputFolder=$PSScriptRoot + '\' + $file.BaseName
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
        Write-Verbose "worksheet $($worksheet.name)"
        $tableList=$worksheet.listObjects|Where-Object SourceType -eq 1
        if($tableList) {
            foreach($table in $tableList ) {
                Write-Verbose "found table $($table.name) on $($worksheet.name)"
                $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^a-zA-Z0-9]', '') + '_' + ($table.name -replace '[^a-zA-Z0-9]', '') + '.csv'
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
            $exportFileName=$outputFolder +'\'+($worksheet.name -replace '[^a-zA-Z0-9]', '') + '_whole.csv'
            $worksheet.SaveAs($exportFileName, 6,$null,$null,$null,$null,$null,$null,$null,'True')
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