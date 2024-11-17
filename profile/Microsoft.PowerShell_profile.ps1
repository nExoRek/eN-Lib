function prompt { 
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $p=(get-location).path
    if($p.length -gt 20) {
      $p=$p.substring(0,3)+"..."+$p.substring($p.length-12)
    }
    $p="$p :))o-"
    if($isAdmin) { write-host 'A' -NoNewline -ForegroundColor Red } #admin prompt
    write-host $p -NoNewline -ForegroundColor darkgreen
    $host.ui.rawui.WindowTitle=$(get-location).path
    return " "
}
function searchFor {
    param(
      [string]$pattern,
      [int]$context = 0
    )
    Get-ChildItem *.ps1 -Recurse | select-string $pattern -Context $context,$context
}

New-Alias ss Select-String

cd 'C:\_ScriptZ\'
  