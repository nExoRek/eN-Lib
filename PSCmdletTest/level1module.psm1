function f1 {
    [cmdletbinding()]
    param()

    return $PSCommandPath

}

function f2 {
    [cmdletbinding()]
    param()

    f1

}
function level1-callstack {
    Get-PSCallStack
}
Export-ModuleMember -Function *
