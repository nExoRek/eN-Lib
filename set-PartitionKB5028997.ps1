<#
.SYNOPSIS
    semi-automated soluton for KB503441
    - downloads WinRE which is required to enable it back
    - creates input file for diskpart
    ...the rest is in yur hands - check if file is properly generated

    USE ON YOUR OWN RISK
    
.EXAMPLE
    .\set-PartitionForKB.ps1
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.LINK
    https://support.microsoft.com/en-us/topic/kb5028997-instructions-to-manually-resize-your-partition-to-install-the-winre-update-400faa27-9343-461c-ada9-24c8229763bf
.NOTES
    nExoR ::))o-
    version 240410
        last changes
        - 240410 initialized
    #TO|DO
#>
#requires -runAsAdministrator
[cmdletbinding()]
#download winRE - which is not being saved by disabling with reagentc
$ProductName = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\').productName
switch($productName) {
  {$_ -match 'windows 10'}  { $winRElink = 'http://download.w-files.pl/winRM/winRE_win10.wim'}
  {$_ -match 'windows 11'}  { $winRElink = 'http://download.w-files.pl/winRM/winRE_win11.wim' }
  {$_ -match 'Server 2016'} { $winRElink = 'http://download.w-files.pl/winRM/winRE_srv16.wim' }
  {$_ -match 'Server 2019'} { $winRElink = 'http://download.w-files.pl/winRM/winRE_srv19.wim' }
  {$_ -match 'Server 2022'} { $winRElink = 'http://download.w-files.pl/winRM/winRE_srv22.wim' }
  default { 
    "$productName not found"
    $productName = $false
  }
}
if($productName) {
    Invoke-WebRequest -Uri $winRElink -OutFile "$($env:windir)\system32\recovery\winRE.wim"
}
$partitionStyle = (get-disk|? issystem).PartitionStyle
if($partitionStyle -ne 'GPT') {
    return "currently only GPT disks supported. this: $partitionStyle"
}
#GETTING RECOVERY PARTITION
[regex]$rxpart = "\\harddisk([0-9])\\partition([0-9])\\"
$REINFO = reagentc /info
if( ($REINFO|Select-String "Windows RE status") -notmatch "enabled") {
    $REINFO
    return "RE partition not enabled"
}
$locationString = $REINFO|Select-String "Windows RE location"
$m = $rxpart.Matches( $locationString )
if($m.groups.count -ne 3) {
    return "error detecting partition from '$locationString'"
}
if($m.groups[1].value -ne 0) {
    return "detected disk is not 0 which is unexpected. '$locationString'"
}
$REpart = $m.groups[2].value
write-host "Recovery partition detected: $REpart. check with: '$locationString'"
<# old version. more geeky but there were too many issues.
#regexp for getting partition GUID
#$rxGUID=[regex]"(\{[0-9a-f]{8}\-[0-9a-f]{4}\-[0-9a-f]{4}\-[0-9a-f]{4}\-[0-9a-f]{12}\})"
$REvol = get-volume|? FileSystemLabel -match 're tools|winre'
if(!$REvol) {
    Write-Host -ForegroundColor Yellow "can't locate Recovery partition by name."
    return "run get-volume to see which volume is for RE"
}
$guid = $rxGUID.Matches($REvol.UniqueId).groups[1].value
if(!$guid) {
    return "not able to extract GUID"
}
write-host "recovery partition guid: $guid"
$REpart = Get-Partition|? guid -eq $guid
if(!$REpart) {
    return "can't locate proper recovery partition"
}
write-host -ForegroundColor Green "Recovery partition:"
$REvol|select-object FileSystemLabel,uniqueid,size | out-host
#>
$OSPartition = get-partition|? DriveLetter -eq ($env:SystemDrive)[0]
if(!$OSPartition) {
    return "can't locate OS partition"
}
write-host -ForegroundColor Green "OS partition:"
$OSPartition|Select-Object PartitionNumber,DriveLetter,@{l="SizeGB";e={[math]::round($_.size/1GB)}} | out-host
#prepare diskpart input based on gathered information
write-host -ForegroundColor green "preparing diskpart script..."
$diskpartInput = "c:\temp\winre.txt"
#get-list disk
#list part
"select disk $($OSPartition.DiskNumber)"|out-file $diskpartInput -Encoding utf8
"select partition $($OSPartition.PartitionNumber)"|out-file $diskpartInput -Append -Encoding UTF8
"shrink desired=250 minimum=250"|out-file $diskpartInput -Append -Encoding UTF8
"select partition $($REpart)"|out-file $diskpartInput -Append -Encoding UTF8
"delete partition override"|out-file $diskpartInput -Append -Encoding UTF8
#assuming GPT
"create partition primary id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"|out-file $diskpartInput -Append -Encoding UTF8
"gpt attributes =0x8000000000000001"|out-file $diskpartInput -Append -Encoding UTF8
"format quick fs=ntfs label=""Windows RE tools"""|out-file $diskpartInput -Append -Encoding UTF8
write-host -ForegroundColor green "diskpart script:"
Get-Content $diskpartInput|Out-Host
if(!(test-path "$($env:windir)\system32\Recovery\Winre.wim")) {
    write-host -ForegroundColor Red "`nWinRE image is not present! before continuing deploy it first`n"
}
write-host -ForegroundColor green "now you need to:`n * check if disks and partitions were correctly detected"
#disable reagentc and check if winRE file will be properly copied 
write-host "*** commands to fix the partitions:"
write-host "reagentc /disable"
write-host "diskpart /s $diskpartInput"
write-host "reagentc /enable"
write-host "reagentc /info"