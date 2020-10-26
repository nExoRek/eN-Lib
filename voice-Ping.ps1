<#
.SYNOPSIS
  fun script presenting usage of SAPI.spVoice
.EXAMPLE
  .\voicePing.ps1
  Explanation of what the example does
.INPUTS
  None.
.OUTPUTS
  None.
.LINK
  https://w-files.pl
.NOTES
  nExoR ::))o-
  version 201026
    last changes
    - 201026 githubbed old funscript
#>
[CmdletBinding()]
param (
    [Parameter(mandatory=$true,position=0,ValueFromPipeline=$true)]
        [string]$computerName
)

BEGIN {
  function Ping-andAnnounce {
    param(
      # computer IP or name
      [Parameter(mandatory=$true,position=0)]
          [string]$computer
    )
 
    $results = Get-WmiObject -query "SELECT * FROM Win32_PingStatus WHERE Address = '$computer'"
    
    if ($results.StatusCode -eq 0) { 
        $Voice.Speak( "$computer is working fine", 1 )|Out-Null
        Write-Host "$computer is Pingable"
    } else {
        $Voice.Speak( "Alert! Alert! Alert! $computer is down", 1 )|Out-Null
        Write-Host "$computer is not Pingable" -BackgroundColor red 
    }
  }
  $Voice = new-object -com SAPI.SpVoice
}
process {
  Ping-andAnnounce $computerName
}
END {}
