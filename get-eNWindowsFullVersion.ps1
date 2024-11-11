$v = [System.Environment]::OSVersion.Version
$fullVer = "{0}.{1}.{2}.{3}" -f $v.Major,$v.Minor,$v.build,(Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\' -Name UBR)
return $fullVer