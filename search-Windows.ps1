<#
.SYNOPSIS
  locate function to quickly find files using Windows Search Index
.NOTES
  author: nExoR 2o15
  version: 25.1o.2o15
  TO|DO 
    add search parameters:
     - search for file contents
     - only in specific files [doc/pdf]
.EXAMPLE
  .\search-Windows *hidden*.ps1

  queries Windows Search for all ps1 files with 'hidden' string
.EXAMPLE
  .\search-Windows doc

  searches for all files and folder named 'doc'. if you are looking for all word files use *.doc* instead.
.LINK
  http://w-files.pl
.LINK
  http://blogs.technet.com/b/heyscriptingguy/archive/2010/05/30/hey-scripting-guy-weekend-scripter-using-the-windows-search-index-to-find-specific-files.aspx
.LINK
  https://msdn.microsoft.com/en-us/library/bb231256%28v=VS.85%29.aspx
#>
#requires -version 5

param(
  [parameter(mandatory=$true)][string]$fileName,
  [int]$pageSize=20000
)

class returnedFile {
  [string]$name
  [string]$path
  [string]$type
  [DateTime]$dateCreated
  [DateTime]$dateModified
  [int]$fileAttributes
  [string]$fileOwner
  [int]$Size
}

$fileName=$fileName.Replace('*','%')
$query = "SELECT `
    System.ItemName, system.ItemPathDisplay, System.ItemTypeText,System.DateCreated,System.DateModified,`
    System.FileAttributes,System.FileOwner,System.Size `
        FROM SystemIndex where system.fileName LIKE '$fileName'"

#https://social.technet.microsoft.com/Forums/scriptcenter/en-US/c262d5de-da6b-4e9e-9aab-2965f3109b8a/vbs-looking-for-script-to-change-windows-search-property?forum=ITCG        
$ADOCommand = New-Object -ComObject adodb.command
$ADOConnection = New-Object -ComObject adodb.connection
$RecordSet = New-Object -ComObject adodb.recordset
$ADOConnection.open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
$RecordSet.pageSize=$pageSize
$RecordSet.open($query, $ADOConnection)
Try { $RecordSet.MoveFirst() }
Catch [system.exception] { "no records returned" }
while(-not($RecordSet.EOF)) 
{
  $locatedFile=[returnedFile]::new()
  $locatedFile.Name=($RecordSet.Fields.Item("System.ItemName")).value
  $locatedFile.Path=($RecordSet.Fields.Item("System.ItemPathDisplay")).value
  $locatedFile.Type=($RecordSet.Fields.Item("System.ITemTypeText")).value
  $locatedFile.DateCreated=($RecordSet.Fields.Item("System.DateCreated")).value 
  $locatedFile.DateModified=($RecordSet.Fields.Item("System.DateModified")).value 
  $locatedFile.FileAttributes=($RecordSet.Fields.Item("System.FileAttributes")).value 
  $locatedFile.FileOwner=($RecordSet.Fields.Item("System.FileOwner")).value 
  $locatedFile.Size=($RecordSet.Fields.Item("System.Size")).value
  $locatedFile
 $RecordSet.MoveNext()
} 

$RecordSet.Close()
$ADOConnection.Close()
$RecordSet = $null
$ADOConnection = $null
[gc]::collect()
