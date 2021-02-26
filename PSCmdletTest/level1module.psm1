function f1 {
    [cmdletbinding()]
    param()

    <#
        $myinvocation.MyCommand.Path is empty for functions - functions do not have paths
        $myincocation.mycommand will return function name.
    #>
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
