function level2-f1 {
    [cmdletbinding()]
    param()

    f1 
}
function level2-callstack {
    [cmdletbinding()]
    param()

    level1-callstack
}
Export-ModuleMember -Function *