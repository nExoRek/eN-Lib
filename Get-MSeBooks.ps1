<#
.SYNOPSIS
 Free Microsoft eBook Giveaway script
 ...rewritten by nExoR ::))o-
.DESCRIPTION
    script initially found on 
 https://blogs.msdn.microsoft.com/mssmallbiz/2016/07/10/free-thats-right-im-giving-away-millions-of-free-microsoft-ebooks-again-including-windows-10-office-365-office-2016-power-bi-azure-windows-8-1-office-2013-sharepoint-2016-sha/
  but it sucked. so i wrote my own using the idea. 
   
.LINK
  Link to download list of eBooks: http://ligman.me/29zpthb
.LINK
  author:  http://w-files.pl 
.NOTES
    nExoR ::))o- 

#>
[cmdletbinding()]
param(
    [parameter(mandatory=$false,position=1)]
        [switch]$refreshCache,
        #by default script is only listing available books and then cache in local file. this switch enforced qeb query for a book list
    [parameter(mandatory=$false,position=0)]
        [string]$downloadFolder = "$($env:USERPROFILE)\Documents\Downloads\ebooks\"
        #directory where to store books - used for 'download'
)

function get-BookList {
    $bookList = @()
    if (Test-Path $cacheFile) {
        $bookList = Import-Csv -Delimiter ';' $cacheFile
        write-host "list read from cache file. use ""-refreshCache"" to re-read from web source"
    }
    else {
        Write-host "book list cache unavailable - downloading list. this may take a while..."
        # Download the source list of books
        $downLoadList = (Invoke-WebRequest "http://ligman.me/29zpthb").content.split("`n")
        # Remove the first line - it's not a book
        $downLoadList = $downLoadList[1..$downLoadList.count]

        #get info on all books
        Write-host "found $($downLoadList.count) books. getting books titles..."
        $nr=1
        foreach ($bookLink in $downLoadList) {
            try {
                $bookLink=$bookLink.trim()
                write-host "getting title of $nr/$($downLoadList.count)..."
                $nr++
                $header = Invoke-WebRequest $bookLink -Method Head 
                $title = $header.BaseResponse.ResponseUri.Segments[-1]
                $title=$title.replace("%20"," ")
                $bookList += New-Object -type PSobject -Property @{
                    title = $title
                    link  = $bookLink
                    size  = [string]( [math]::round(($header.Headers.'Content-Length')/1MB,1) )+ "MB"
                }
                write-host $title
                #if($nr -gt 10) { break } #for debugging
            }
            catch {
                Write-host -ForegroundColor Yellow "`[4o4`] $bookLink is unavailable."
                $bookList += New-Object -type PSobject -Property @{
                    title = "<unavailable>"
                    link  = $bookLink
                    size  = "<null>"
                }
            }
        }
        $bookList|export-csv -Delimiter ';' -Path $cacheFile -NoTypeInformation
        Write-Verbose "list saved to a cache file $cacheFile"       
    }
    return $bookList
}

function get-BookDownload {
    param(
        $book
    )
    try {
        $saveAs = $downloadFolder + $book.title
        Write-Host "downloading $bookLink -> $saveAs ...."
        Invoke-WebRequest $book.link -OutFile $saveAs
    } catch {
        Write-host -ForegroundColor red "[4o4] $($book.title) is unavailable"
    }
}

$cacheFile = "$PSScriptRoot\msbooks.cache.csv"
$bookList  = @()

#remove cache file if requested
if($refreshCache) { Remove-Item $cacheFile }

#check if download location exists
if(-not (Test-Path $downloadFolder)) {
    New-Item -ItemType Directory -Path $downloadFolder -Force
    Write-Verbose "$downloadFolder created. use ""-downloadFolder"" to define target location."
}

#get the list
$bookList=get-BookList
$downloadList=$bookList|out-gridView -title "Choose books to download" -PassThru
foreach ($book in $downloadList) {
    get-BookDownload -book $book
}
write-host -ForegroundColor green "done."
