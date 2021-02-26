#requires -module level1module
[cmdletbinding()]
param( )

write-host -ForegroundColor Magenta "PSCallStack in different contexts`n"

write-host -ForegroundColor yellow  "directly from here"
Get-PSCallStack|out-host

write-host -ForegroundColor yellow "from the function loaded by module"
level1-callstack|out-host
write-host -ForegroundColor yellow  "from the function in module L2 calling function in module L1"
level2-callstack|out-host

write-host -ForegroundColor Magenta "MtInvocation test from different contexts`n"

write-host -ForegroundColor yellow  "first in-script PScommandPath (from myinvocation)"
"{0} -> {1}" -f 'local myinvocation',($myinvocation.MyCommand.Path)
write-host -ForegroundColor yellow  "now same myinovation but run from the function, loaded in a module L1"
"{0} -> {1}" -f 'f1 PScommandPath',(f1)
write-host -ForegroundColor yellow  "similar, but function inside module L1 calls function in the same module"
"{0} -> {1}" -f 'f1 via f2 PScommandPath',(f2)
write-host -ForegroundColor yellow  "more complex - call function in module L2, which is calling function in module L1"
"{0} -> {1}" -f 'f1 via diferent module',(level2-f1)
