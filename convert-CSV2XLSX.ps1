<#
.SYNOPSIS
    Converts CSV file into XLS with table.
.DESCRIPTION
    creates XLXS out of CSV file and formats data as a table.
.EXAMPLE
    .\convert-CSV2XLSX.ps1 c:\temp\test.csv -delimiter ','
    
    Converts test.csv to test.xlsx 
.INPUTS
    CSV file.
.OUTPUTS
    XLSX file.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 201123
        last changes
        - 201123 initialized
#>
[CmdletBinding()]
param (
    #CSV file name to convert
    [Parameter(mandatory=$true,position=0)]
        [string]$CSVfileName,
    #style intensity
    [Parameter(mandatory=$false,position=1)]
        [alias('intensity')]
        [string][validateSet('Light','Medium','Dark')]$tableStyleIntensity='Medium',
    #style number
    [Parameter(mandatory=$false,position=2)]
        [alias('nr')]
        [int]$tableStyleNumber=21,
    #Excel output file name 
    [Parameter(mandatory=$false,position=3)]
        [string]$XLSfileName,
    #CSV delimiter character
    [Parameter(mandatory=$false,position=4)]
        [string][validateSet(',',';')]$delimiter=';'
)

if(-not (test-path $CSVfileName) ) {
    write-host -ForegroundColor Red "file $CSVfileName is not accessible"
    exit -1
}

try{
    $Excel = New-Object -ComObject Excel.Application
} catch {
    write-host -ForegroundColor Red "not able to initialize Excel lib. requires Excel to run"
    write-host -ForegroundColor red $_.Exeption
    exit -3
}

$file=get-childItem $CSVfileName
if( [string]::IsNullOrEmpty($XLSfileName) ) {
    $XLSfileName=$file.DirectoryName+'\'+$file.BaseName+'.xlsx'
}
#$table=import-csv -Delimiter $delimiter ';' -Encoding UTF8

write-host 'creating excel file...'
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)
### Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $CSVfileName)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.TextFilePlatform = 65001
### Execute & delete the import query
$query.Refresh() |out-null

#$range=$worksheet.QueryTables[1].ResultRange
$range=$query.ResultRange
$query.Delete()

#$Table = $worksheet.ListObjects.Add($range, "importedCSV")
$Table = $worksheet.ListObjects.Add(
    [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
    $Range, 
    "importedCSV",
    [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
    )
<#
TableStyle
  Light
  Medium
  Dark
    1,8,15 black
    2,9,16 navy blue
    3,1o,17 orange
    4,11,18 gray
    5,12,19 yellow
    6,13,2o blue
    7,14,21 green
    
#>
$tableStyle=[string]"$tableStyleIntensity$tableStyleNumber"
$Table.TableStyle = "TableStyle$tableStyle" #green with with gray shadowing

#$wb = $Excel.Workbooks.Open($CSVfileName)
$worksheet.SaveAs($XLSfileName, 51,$null,$null,$null,$null,$null,$null,$null,'True') #|out-null
$Excel.Quit()
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) ){}

write-host -ForegroundColor Green "convertion done, saved as $XLSfileName"
Write-Host -ForegroundColor Green "done and cleared."
