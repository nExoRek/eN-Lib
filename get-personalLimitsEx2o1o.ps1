Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
$personalLimits=Get-Mailbox |?{$_.UseDatabaseQuotaDefaults -eq $false}

foreach($pl in $personalLimits){
    $mbx=get-mailboxstatistics $pl 
    add-member -inputObject $pl -MemberType NoteProperty -Name totalItemSize -Value $mbx.totalitemsize.value.toMB()
    add-member -inputObject $pl -MemberType NoteProperty -Name deleteItemSize -Value $mbx.totaldeleteditemsize.value.toMB()
    add-member -inputObject $pl -MemberType NoteProperty -Name databasename -Value $mbx.databasename
    add-member -inputObject $pl -MemberType NoteProperty -Name lastlogontime -Value $mbx.lastlogontime
}
$personalLimits|select name,`
    @{n='sendQuota';e={$_.ProhibitSendQuota.value.toMB()}},`
    @{n='Receive Quota';e={$_.ProhibitSendReceiveQuota.value.toMB()}},`
    @{n='warning';e={$_.IssueWarningQuota.value.toMB()}},`
    totalitemsize,deleteItemSize,databasename,lastlogontime|`
        Export-csv -NoTypeInformation -Delimiter ';' C:\temp\personalLimits.csv -encoding UTF8
    
