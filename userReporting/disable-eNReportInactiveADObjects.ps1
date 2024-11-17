[CmdletBinding()]
param (
    #input file
    [Parameter(mandatory=$true,position=0)]
        [string]$inputFile,
    #destination OU for disabled object to move - empty will not move objects.
    [Parameter(mandatory=$false,position=1)]
        [string]$unusedObjectsOU,
    #do not confirm each disable operation - deafult will ask
    [Parameter(mandatory=$false,position=2)]
        [switch]$doNotConfirmDisable,
    #do not confirm each move operation - deafult will ask
    [Parameter(mandatory=$false,position=3)]
        [switch]$doNotConfirmMove,
    #period of inactivity time to disable objects
    [Parameter(mandatory=$false,position=4)]
        [int]$monthsOld = -6
)

try {
    $objects = import-csv -Path $inputFile
} catch {
    return $_
}
if($monthsOld -gt 0) { $monthsOld = -1 * $monthsOld }
$toDisable = @()
foreach($obj in $objects) {  
    Write-Verbose "processing $($obj.samAccountName);$($obj.LastLogonDate);$($obj.enabled)"
    if($obj.LastLogonDate) {
        try {
            $lld = get-date $obj.LastLogonDate
        } catch {
            write-host "not able to detect date $($obj.samAccountName) - '$($obj.LastLogonDate)'"
            continue
        }
    } else {
        $lld = get-date "1/1/1971"
    }
    if( $lld -lt (get-date).AddMonths($monthsOld) -and $obj.enabled -eq 'True') {
        $toDisable += $obj
    }
}
if($toDisable.count -lt 1) {
    return "no unused accounts detected"
}
Write-Warning "list of accounts inactive for $(-1 * $monthsOld) months:"
$toDisable | Select-Object SamAccountName,LastLogonDate,Enabled,distinguishedName | out-host

<#
Write-Host "do you disable listed obejcts and move them to '$unusedObjectsOU'?"
$readHostValue = Read-Host -Prompt "Enter 'Yes' to continue"
if(!($readHostValue -eq 'yes')) {
    return 'cancelling' 
}
#>
write-warning "you can exit in any moment by pressing ctrl-c"
foreach($obj in $toDisable) {
    $willMove = $false
    write-host  "processing $($obj.SamAccountName):$($obj.LastLogonDate):'$($obj.distinguishedName)'..."
    if(!$doNotConfirmDisable) {
        Write-Warning "do you want to disable account?"
        $readHostValue = Read-Host -Prompt "Enter Y to accept"
        if($readHostValue -ne 'y') {
            #$willDisable = $false
            #$willMove = $false
            write-host "skipping account"
            continue
        }
    }
    if($unusedObjectsOU) {
        if(!$doNotConfirmMove) {
            Write-Warning "do you want to move '$($obj.distinguishedName)' to '$unusedObjectsOU'?"
            $readHostValue = Read-Host -Prompt "Enter Y to accept"
            if($readHostValue -eq 'y') {
                $willMove = $true
            }
        } else {
            $willMove = $true
        }
    }
    write-host "disabling $($obj.samAccountName)..."
    disable-ADAccount $obj.distinguishedName
    set-ADObject -Identity $obj.distinguishedName -Description $("object disabled {0} [{1}]" -f $(get-date -Format "dd-MM-yyyy")),$obj.description
    Write-Verbose "ok."
    if($willMove) {
        write-host "moving '$($obj.distinguishedName)' to '$unusedObjectsOU'..."
        Move-ADObject -Identity $obj.DistinguishedName -TargetPath $unusedObjectsOU
        Write-Verbose "ok."
    }
}
write-host 'done.'