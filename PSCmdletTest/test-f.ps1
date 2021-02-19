#requires -module level1module
[cmdletbinding()]
param( )

"{0} -> {1}" -f 'f1 via diferent module',(level2-f1)

"{0} -> {1}" -f 'local myinvocation',($PSCommandPath)
"{0} -> {1}" -f 'f1 myinvocation',(f1)
"{0} -> {1}" -f 'f1 via f2 myinvocation',(f2)

level2-f2