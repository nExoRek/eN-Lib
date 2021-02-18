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
Export-ModuleMember -Function *
