<#
.SYNOPSIS
    scan all folders and subfolders for non-inherited NTFS permissions. useful for auditing purposes.
.DESCRIPTION
    here be dragons
.EXAMPLE
    .\get-nonInheritedNFTSpermissions.ps1 d:\share\sharedFolder

    checks the pemission structure in search of any broken inheritance.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 250129
        last changes
        - 250129 initialized

    #TO|DO
    * add information on folders that have 'inheritance' flag disable to distinguish between added and modified permissions
    * currently only showing folders and names with no actual permissions...
#>
[CmdletBinding()]
param(
    [Parameter(mandatory,position=0)]
    [string]$Path
)
#$VerbosePreference = "Continue"
$finalPermissions = @()
$date = [string](Get-Date -Format "yyyyMMdd-HHmmss")
$outFile = "nonInheritedPermissions-{0}.csv" -f $date
$errFile = "errors-{0}.log" -f $date

function Get-nonInheritedPermissions {
    param(
        [string]$FolderPath
    )
    Write-Verbose "Checking $FolderPath"
    try {
        $acl = Get-Acl -Path $FolderPath
    } catch {
        $FolderPath  | Out-File -FilePath $errFile -Append
        $_.Exception | Out-File -FilePath $errFile -Append
    }

    
    foreach ($access in $acl.Access) {
        if (-not $access.IsInherited) {
            $p = [PSCustomObject]@{
                Folder = $FolderPath
                User   = $access.IdentityReference
            }
            write-host $p
            $p
        }
    }
}

function Scan-Directory {
    param(
        [parameter(mandatory)]
        [string]$RootPath
    )

    if (Test-Path $RootPath) {
        $script:finalPermissions += Get-nonInheritedPermissions -FolderPath $RootPath
        
        Get-ChildItem -Path $RootPath -Directory -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $script:finalPermissions += Get-nonInheritedPermissions -FolderPath $_.FullName
        }
    } else {
        Write-error "Path not found: $RootPath"
    }
}

Scan-Directory -RootPath $Path
$finalPermissions | Export-Csv -Path $outFile -NoTypeInformation
write-host -ForegroundColor Green "done."
