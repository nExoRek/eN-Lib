<#
.SYNOPSIS
     providing psedo-GUI extention to speed up Az-based deployment for 1st line support.
.DESCRIPTION
    module is using out-GridView and some basic win32 forms (from eNLib) to present element of simple
    GUI. 
.EXAMPLE
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 210430
        last changes
        - 210430 naming normalization
        - 210310 vnet fix
        - 210302 functions retun $null on cancel, autoselect, fixes and help 
        - 210301 storageAccount choice fixed, select-networksecuritygroup added, multiselect for some functions
        - 210215 seelct-vnet and select-subnet fixes, select-recoveryContainer
        - 210208 select-recoveryVault, descriptions, fixes...
        - 210202 initialized alpha
    #TO|DO
    - -autoSelectSingleValue to all functions
    - option to repeat-until for resource choices (-repeatUntilCancel)
#>
Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"

#################################################### Az pseudo-GUI Extension
function select-Subscription {
    <#
    .SYNOPSIS
        pseudo-GUI function to allow Azure context change 
    .DESCRIPTION
        #TODO
    .EXAMPLE
        #TODO
    .INPUTS
        None.
    .OUTPUTS
        AzContext
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoselectsingle
            - 210220 fixes to exit logic
            - 210127 initialized
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Subscription*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Select Subscription',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption
    )
    try {
        $subscriptionList = Get-AzSubscription
        if($null -eq $subscriptionList) {
            write-log "No subscriptions found. try changing context." -type warning
            if($isCritical.IsPresent) {
                exit -1
            }
            return -1
        }
        if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($subscriptionList.count) ) {
            write-log "single subscription: $($subscriptionList.name) found. selecting." -type info
            $sourceSubscription = $subscriptionList
        } else {
            write-log $message
            $sourceSubscription = $subscriptionList | out-GridView -title $title -OutputMode Single
            if($null -eq $sourceSubscription) {
                write-log "operation cancelled by the user."
                if($isCritical.IsPresent) {
                    exit 0
                }        
                return $null
            }
            write-log "chosen subscription: $($sourceSubscription.name)" -type info
        }
    } catch {
        write-log "error getting Subscription list $($_.exception)" -type error
        if($isCritical.IsPresent) {
            exit -3
        }        
        return -3
    }
    
    try {
        $AzSourceContext = set-azContext -SubscriptionObject $sourceSubscription -force
        write-log "source context: $($AzSourceContext.Subscription.name)" -silent
    } catch {
        write-log "error changing context. $($_.Exception)" -type error
        if($isCritical.IsPresent) {
            exit -4
        }        
        return -4
    }
    return $AzSourceContext
}
function select-ResourceGroup { 
    <#
    .SYNOPSIS
        pseudo-GUI for ResourceGroup selection
    .DESCRIPTION
        #todo
    .EXAMPLE
        #todo
    .INPUTS
        None.
    .OUTPUTS
        Resource Group object
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoselectsingle
            - 210220 multichoice
            - 210127 initialized
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "select *Resource Group*..." ,
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "select Resource Group",
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #allow multiple choice
        [Parameter(mandatory=$false,position=4)]
            [switch]$multiChoice
    )
    
    write-log $message
    $ogvParam = @{
        title = $title
        OutputMode = 'Single'
    }
    if($multiChoice.IsPresent) {
        $ogvParam.OutputMode = 'Multiple'    
    } 
    $RGList = Get-AzResourceGroup | select-object ResourceGroupName, location, Tags
    if($null -eq $RGList) {
        write-log "No resource groups in this subscription." -type warning
        if($isCritical.IsPresent) {
            exit -1
        }
        return -1
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($RGList.count) ) { #when single object available, 'count' does not exist
        write-log "single RG available: $($RGList.ResourceGroupName) found. selecting." -type info
        return (Get-AzResourceGroup -Name $RGList.ResourceGroupName)
    } else {
        $ResourceGroup = $RGList | Out-GridView @ogvParam
        if ([string]::isNullOrEmpty($ResourceGroup)) {
            Write-log 'Cancelled by user'
            if($isCritical.IsPresent) {
                exit 0
            }
            return $null
        }
        if($multiChoice.IsPresent) {
            foreach($RG in $ResourceGroup) {
                write-log "ResourceGroup $($RG.name) chosen." -type info
                Get-AzResourceGroup -Name $RG.ResourceGroupName
            }

        } else {
            Write-log "RG $($ResourceGroup.ResourceGroupName) chosen. " -type info
            return (Get-AzResourceGroup -Name $ResourceGroup.ResourceGroupName)
        }
    }
}
function select-NetworkSecurityGroup { 
    <#
    .SYNOPSIS
        pseudo-GUI for Network Security Group selection
    .DESCRIPTION
        #todo
    .EXAMPLE
        #todo
    .INPUTS
        None.
    .OUTPUTS
        Network Security Group object
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210301
            last changes
            - 210301 initialized
    #>
    [CmdletBinding(DefaultParameterSetName='byObject')]
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "select *Network Security Group*..." ,
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "select Network Security Group",
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #allow multiple choice
        [Parameter(mandatory=$false,position=4)]
            [switch]$multiChoice,
        #Resource Group object
        [Parameter(parameterSetName='byObject',mandatory=$false,position=5,ValueFromPipeline=$true)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,
        #Parameter help description
        [Parameter(parameterSetName='byName',mandatory=$false,position=5)]
            [string]$ResourceGroupName
    )
    
    write-log $message
    $ogvParam = @{
        title = $title
        OutputMode = 'Single'
    }
    if($multiChoice.IsPresent) {
        $ogvParam.OutputMode = 'Multiple'    
    } 
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $ResourceGroupName = $ResourceGroup.ResourceGroupName
    }
    $NSGParam=@{}
    if(![string]::isNullOrEmpty($ResourceGroupName) ) {
        $NSGParam.add('ResourceGroupName',$ResourceGroupName)
    } 
    $NSGList = Get-AzNetworkSecurityGroup @NSGParam | select-object name, ResourceGroupName, location, Tags
    if($null -eq $NSGList) {
        write-log "No Network Security groups in this subscription." -type warning
        if($isCritical.IsPresent) {
            exit -1
        }
        return -1
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($NSGList.count) ) { #when single object available, 'count' does not exist
        write-log "single NSG available: $($NSGList.Name) found. selecting." -type info
        return (Get-AzNetworkSecurityGroup -Name $NSGList.Name -ResourceGroupName $nsgList.resourceGroupName)
    } else {
        $NSG = $NSGList | Out-GridView @ogvParam
        if ([string]::isNullOrEmpty($NSG)) {
            Write-log 'Cancelled by user'
            if($isCritical.IsPresent) {
                exit 0
            }
            return 0
        }
        if($multiChoice.IsPresent) {
            foreach($nsgitem in $NSG) {
                write-log "NSG $($nsgitem.name) chosen." -type info
                Get-AzNetworkSecurityGroup -Name $nsgitem.name -ResourceGroupName $nsgitem.resourceGroupName
            }

        } else {
            Write-log "NSG $($NSG.ResourceGroupName) chosen. " -type info
            return (Get-AzNetworkSecurityGroup -Name $nsg.name -ResourceGroupName $nsg.resourceGroupName)
        }
    }
}
Set-Alias -Name 'select-NSG' -Value 'select-networkSecurityGroup'
function select-StorageAccount {
    <#
    .SYNOPSIS
        select Storage Account resource visually/
    .DESCRIPTION
        #todo
    .EXAMPLE
        $SA=select-StorageAccount
        
        choose Storage Account using out-gridview and assign under SA variable.
    .INPUTS
        None.
    .OUTPUTS
        Storage Account(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle
            - 210220 multichoice, descritpion
    #>
    [cmdletbinding(DefaultParameterSetName='byObject')]
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'select *Storage Account*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'select Storage Account',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #allow multiple choice
        [Parameter(mandatory=$false,position=4)]
            [switch]$multiChoice,
        #Resource Group object
        [Parameter(parameterSetName='byObject',mandatory=$false,position=5,ValueFromPipeline=$true)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,
        #Parameter help description
        [Parameter(parameterSetName='byName',mandatory=$false,position=5)]
            [string]$ResourceGroupName
    )
    
    $ogvParam = @{
        title = $title
        OutputMode = 'Single'
    }
    if($multiChoice.IsPresent) {
        $ogvParam.OutputMode = 'Multiple'    
    }     
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $ResourceGroupName = $ResourceGroup.ResourceGroupName
    }
    #get-azStorageAccount do not allow to set RG as $null, so must be done with IF
    if([string]::isNullOrEmpty($ResourceGroupName) ) {
        $saList = Get-AzStorageAccount |
            #? {$_.Sku.Name -match "^standard"} |
            select-object StorageAccountName,ResourceGroupName,PrimaryLocation,@{N='SkuName';E={$_.sku.name}}
    } else {
        $saList = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName|
            #? {$_.Sku.Name -match "^standard"} |
            select-object StorageAccountName,ResourceGroupName,PrimaryLocation,@{N='SkuName';E={$_.sku.name}}
    }
    write-log $message
    if([string]::isNullOrEmpty($saList) ) {
        write-host "there are no storage accounts in this Resource Group."
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }        
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($saList.count) ) { #when single object available, 'count' does not exist
        write-log "single SA available: $($saList.StorageAccountName) found. selecting." -type info
        return (Get-AzStorageAccount -Name $saList.StorageAccountName -ResourceGroupName $saList.resourceGroupName)
    } else {
        $storageAccount = $saList | Out-GridView @ogvParam
        if([string]::isNullOrEmpty($storageAccount) ) {
            write-host "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }

        if($multiChoice.IsPresent) {
            foreach($SA in $storageAccount) {
                write-log "SA $($SA.StorageAccountName) chosen." -type info
                Get-AzStorageAccount -ResourceGroupName $SA.ResourceGroupName -Name $SA.StorageAccountName
            }

        } else {
            write-log "SA $($storageAccount.StorageAccountName) chosen." -type info
            return (Get-AzStorageAccount -ResourceGroupName $storageAccount.resourceGroupName -Name $storageAccount.StorageAccountName)
        }    
    }
}
function select-VirtualMachine {
    <#
    .SYNOPSIS
        select VM on subscription or RG level.
    .DESCRIPTION
        uses out-gridView to show list of virtual Machines. by default it shows all VM in current context 
        you can pass Resource Group object to narrow down the list to VMs in current RG.
        returns VM object or 0/-1 in case of cancell/error.
    .EXAMPLE
        $VM = select-ResourceGroup | select-VM
        
        allows to choose Azure VM from given RG.
    .INPUTS
        None.
    .OUTPUTS
        regular VM - Microsoft.Azure.Commands.Compute.Models.PSVirtualMachineInstanceView 
        or PSCustomObject (with 'status' flag)
        or 0  - Cancel
        or -1 - Error
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle, backup status fix
            - 210220 multichoice, backup info
            - 210209 initialized
    #TO|DO

    #>
    
    [CmdletBinding(DefaultParameterSetName='byObject')]
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "choose *VM* ...",
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "choose VM",
        #changes return to terminal - define if cancel/empty is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #allow multiple choice
        [Parameter(mandatory=$false,position=4)]
            [switch]$multiChoice,
        #resource group object containing VMs to list
        [Parameter(ValueFromPipeline=$true,parameterSetName='byObject',mandatory=$false,position=5)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,  
        #resource group name containing VMs to list
        [Parameter(ValueFromPipeline=$true,parameterSetName='byName',mandatory=$false,position=5)]
            [string]$ResourceGroupName,  
        #include backup information on the list
        [Parameter(mandatory=$false,position=6)]
            [switch]$status,
        #include backup information on the list
        [Parameter(mandatory=$false,position=7)]
            [switch]$includeBackupInfo
    )

    write-log $message
    $ogvParam = @{
        title = $title
        OutputMode = 'Single'
    }
    if($multiChoice.IsPresent) {
        $ogvParam.OutputMode = 'Multiple'    
    } 
    $getAzVMParams = @{
        status = $true
    }
    if( -not [string]::isNullOrEmpty($ResourceGroup) ){
        $getAzVMParams.add('ResourceGroup',$ResourceGroup.ResourceGroupName)
    } elseif( -not [string]::isNullOrEmpty($ResourceGroupName) ) {
        $getAzVMParams.add('ResourceGroup',$ResourceGroupName)
    }
    try {
        $VMList = get-AzVM @getAzVMParams | Select-Object name, powerstate, @{N='OsType';E={$_.StorageProfile.OsDisk.OsType}}, ` 
            @{N='VmSize';E={$_.HardwareProfile.VmSize}},Location,ResourceGroupName
    } catch {
        write-log "error getting VM list $($_.exception)" -type error
        if($isCritical.IsPresent) {
            exit -2
        } else {
            return -2
        }
    }
    if($includeBackupInfo.IsPresent) {
        $VMList = $VMList | select-object *,backupVault
        foreach($VM in $VMList) {
            $backupStatus = Get-AzRecoveryServicesBackupStatus -ResourceGroupName $VM.ResourceGroupName -Name $VM.name -Type AzureVM
            if( $backupStatus.BackedUp ) {
                $VM.backupVault = ($backupStatus.VaultId -split '/')[-1]
            } else {
                $VM.backupVault = 'no backup'
            }
        }
    }
    if([string]::isNullOrEmpty($VMList) ) {
        write-log "no VMs found" -type warning
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($vmList.count) ) { #when single object available, 'count' does not exist
        write-log "single VM available: $($vmList.Name) found. selecting." -type info
        if($status.IsPresent) {
            $powerstate = $VMList.powerstate
            return (get-AzVM -Name $VMlist.name -ResourceGroupName $VMlist.ResourceGroupName | select-object *,@{N='powerstate';E={$powerstate}})
        } else {
            return (get-AzVM -Name $VMlist.name -ResourceGroupName $VMlist.ResourceGroupName)
        }
    } else {
        #ACTUAL VM Select
        $selectVM = $VMList | Out-GridView @ogvParam
        if($null -eq $selectVM) {
            write-log "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }

        foreach($VM in $selectVM) {
            if($includeBackupInfo.IsPresent -and $selectVM.backup -ne 'no backup') { 
                write-log "$($VM.name) has backup enabled." -type warning
                if($isCritical.IsPresent) {
                    exit -2
                } else {
                    return -2
                }
            }
            if($status.IsPresent) {
                $powerstate = $VM.powerstate
                $retVM = get-AzVM -Name $VM.name -ResourceGroupName $VM.ResourceGroupName | select-object *,@{N='powerstate';E={$powerstate}}
            } else {
                $retVM = get-AzVM -Name $VM.name -ResourceGroupName $VM.ResourceGroupName 
            }
            write-log "$($retVM.name) chosen type of $($retVM.StorageProfile.OsDisk.OsType)." -type info
            $retVM
        }
    }
}
Set-Alias -Name select-VirtualMachine -Value select-VM
function select-VirtualNetwork {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose AzVNet resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $vNet = select-vNet
        
        choose vNet resource via out-gridview
    .INPUTS
        None.
    .OUTPUTS
        vNet resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210310
            last changes
            - 210310 rgname when no net
            - 210302 autoSelectSingle
            - 210220 mulichoise, help
    #>
    
    [CmdletBinding(DefaultParameterSetName='byObject')]
    param(
        #onscreen message displayed before popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'select *vNet*...',
        #gridview title 
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'select vNet',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #allow multi-choice and return array
        [Parameter(mandatory=$false,position=4)]
            [switch]$multiChoice,
        #resource group object containing vNet
        [Parameter(parameterSetName='byObject',mandatory=$false,position=5,ValueFromPipeline=$true)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,
        #resource group name containing vNet
        [Parameter(parameterSetName='byName',mandatory=$false,position=5)]
            [string]$ResourceGroupName
    )
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $ResourceGroupName=$ResourceGroup.ResourceGroupName
    }
    $ogvParam = @{
        title = $title
        OutputMode = 'Single'
    }
    if($multiChoice.IsPresent) {
        $ogvParam.OutputMode = 'Multiple'    
    } 
    $azvNet=@{}
    if(![string]::isNullOrEmpty($ResourceGroupName) ) {
        $azvNet.add('ResourceGroupName',$ResourceGroupName)
    } 
    $vNetList = Get-AzVirtualNetwork @azvNet | select-object name,location, ResourceGroupName
    if($null -eq $vNetList) {
        write-log "there is no VNets in $ResourceGroupName"
        if($isCritical.IsPresent) {
            exit -1
        }
        return -1
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($vNetList.count) ) { #when single object available, 'count' does not exist
        write-log "single vNet available: $($vNetList.Name) found. selecting." -type info
        return (Get-AzVirtualNetwork -Name $vNetList.Name -ResourceGroupName $vNetList.resourceGroupName)
    } else {
        $vnetChoice = $vNetList | Out-GridView @ogvParam
        if($null -eq $vNetChoice) {
            write-log "cancelled by the user."
            if($isCritical.IsPresent) {
                exit 0
            }
            return $null
        }
        if($multiChoice.IsPresent) {
            foreach($vNet in $vnetChoice) {
                write-log "vNet $($vNet.name) chosen." -type info
                Get-AzVirtualNetwork -ResourceGroupName $vNet.ResourceGroupName -Name $vNet.Name
            }

        } else {
            write-log "vNet $($vNetChoice.name) chosen." -type info
            return (Get-AzVirtualNetwork -ResourceGroupName $vNetChoice.ResourceGroupName -Name $vNetChoice.Name)
        }
    }
}
Set-Alias -Name select-vNet -Value select-VirtualNetwork
function select-Subnet {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose vNet Subnet resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $vSubnet = select-vNet|select-vSubnet
        
        choose Subnet resource via out-gridview
    .INPUTS
        None.
    .OUTPUTS
        Subnet resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle,comment-based-help
    #>
    [cmdletbinding(DefaultParameterSetName='byObject')]
    param(
         #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Subnet*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Select Subnet',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #vNet Object containing subnets to choose
        [Parameter(parameterSetName='byObject',mandatory=$true,position=4,ValueFromPipeline)]
            [Microsoft.Azure.Commands.Network.Models.PSVirtualNetwork]$vNet,
        #vNet name
        [Parameter(parameterSetName='byName',mandatory=$true,position=4,ValueFromPipeline)]
            [string]$vNetName
    )

    if($PSCmdlet.ParameterSetName -eq 'byName') {
        try {  
            $vNet = Get-AzVirtualNetwork -Name $vNetName
        } catch {
            write-log "error getting vNet $vNetName. $($_.exception)" -type error
            if(isCritical.IsPresent) {
                exit -2
            } else {
                return -2
            }
        }
    }
    write-log $message
    $vSubnetList = Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vNet | Select-Object name,AddressPrefix    
    if([string]::isNullOrEmpty($vSubnetList) ) {
        write-host "there are no subnets in this vNet."
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($vSubnetList.count) ) { #when single object available, 'count' does not exist
        write-log "single vSubnet available: $($vSubnetList.Name) found. selecting." -type info
        return (Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vNet -Name $vSubnetList.Name)
    } else {
        $vSubnet = $vSubnetList | Out-GridView -Title $title -OutputMode Single
        if([string]::isNullOrEmpty($vSubnet) ) {
            write-host "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }
        write-log "vSubnet $($vSubnet.name) chosen." -type info
        return (Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vNet -Name $vSubnet.name)
    }    
}
Set-Alias -Name select-Subnet -Value select-vSubnet
function select-KeyVault {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose Key Vault resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $kv = select-KeyVault -autoSelectSingle
        
        choose Key Vault resource via out-gridview and if only one available - return unattended.
    .INPUTS
        None.
    .OUTPUTS
        KeyVault resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle,comment-based-help
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "choose *Key Vault*...",
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "choose Key Vault",
        #changes return to terminal - define if cancel/empty is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,    
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption
    )

    write-log $message
    $KVList = Get-AzKeyVault | Select-Object VaultName,ResourceGroupName,Location
    if([string]::isNullOrEmpty($KVList) ) {
        write-log "no key vaults found." -type warning
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }    
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($KVList.count) ) { #when single object available, 'count' does not exist
        write-log "single KV available: $($KVList.VaultName) found. selecting." -type info
        return (Get-AzKeyVault -ResourceGroupName $kvList.ResourceGroupName -VaultName $kvList.VaultName )
    } else {
        $kv = $KVList | Out-GridView -Title $title -OutputMode Single
        if($null -eq $kv) {
            write-log "cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        } else {
            write-log "KeyVault $($kv.VaultName) chosen." -type info
            return (Get-AzKeyVault -ResourceGroupName $kv.ResourceGroupName -VaultName $kv.VaultName)
        }
    }
}
function select-encryptionKey {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose Encryption Key resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $ek = select-KeyVault -autoSelectSingle|select-encryptionKey -autoSelectSingle
        
        choose Key Vault resource via out-gridview and if only one available - return unattended. then pass it to key choice.
    .INPUTS
        None.
    .OUTPUTS
        Encryption Key resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle,comment-based-help,valuebypipeline
    #>
    [cmdletbinding(DefaultParameterSetName='byObject')]
    param(
        #message shown before window popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "choose *Encryption Key*...",
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "choose Encryption Key",
        #changes return to terminal - define if cancel/empty is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #Vault name to choose key from
        [parameter(parameterSetName='byName',mandatory=$true,position=4)]
            [string]$vaultName,
        #Vault name to choose key from
        [parameter(parameterSetName='byObject',mandatory=$true,position=4,ValueFromPipeline)]
            [Microsoft.Azure.Commands.KeyVault.Models.PSKeyVault]$vault

    )
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $vaultName = $vault.VaultName
    }
    write-log $message
    $encKeyList = Get-AzKeyVaultKey -VaultName $vaultName | Select-Object name,enabled,Expires
    if($null -eq $encKeyList) {
        write-log "no keys in $vaultName"
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }    
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($encKeyList.count) ) { #when single object available, 'count' does not exist
        write-log "single KV available: $($encKeyList.Name) found. selecting." -type info
        return (Get-AzKeyVaultKey -VaultName $VaultName -Name $encKeyList.name)
    } else {
        $key = $encKeyList | Out-GridView -Title $title -OutputMode Single
        if($null -eq $key) {
            write-log "cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        } else {
            write-log "key $($key.name) chosen." -type info
            return (Get-AzKeyVaultKey -VaultName $vaultName -Name $key.name)
        }
    }
}
function select-recoveryVault {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose Recovery Vault resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $rv = select-RecoveryVault
        
        choose RecoveryVault using out-gridview
    .INPUTS
        None.
    .OUTPUTS
        Recovery Vault resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle,comment-based-help
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Backup Vault*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Select Backup Vault',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption
    )
    
    write-log $message
    $RVList = Get-AzRecoveryServicesVault | Select-Object name,ResourceGroupName,location,ID
    if([string]::isNullOrEmpty($RVList) ) {
        write-log "no Recovery Vaults found" -type warning
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($RVList.count) ) { #when single object available, 'count' does not exist
        write-log "single RV available: $($RVList.Name) found. selecting." -type info
        return (Get-AzRecoveryServicesVault -name $RVList.name)
    } else {
        $recoveryVault = $RVlist | Out-GridView -Title $title -OutputMode Single
        if($null -eq $recoveryVault) {
            write-host "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }
        $recoveryVault = Get-AzRecoveryServicesVault -name $recoveryVault.name
        write-log "Recovery Vault $($recoveryVault.name) chosen." -type info
        return $recoveryVault
    }
}
function select-recoveryContainer {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose Recovery Container resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        $rc = select-RecoveryVault -auto|select-RecoveryContainer -auto
        
        choose RecoveryContainer within RecoveryVault using out-gridview, and if sinlge available - unattended
    .INPUTS
        None.
    .OUTPUTS
        Recovery Container resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210302
            last changes
            - 210302 autoSelectSingle,comment-based-help
        
        TO|DO
        - currently only AzureVM supported - extend for different types
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Backup Container*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Select Backup Container',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption,
        #Vault object
        [Parameter(parameterSetName='byObject',mandatory=$true,position=4, ValueFromPipeline)]
            [Microsoft.Azure.Commands.RecoveryServices.ARSVault]$Vault,
        #VaultID
        [Parameter(parameterSetName='byName',mandatory=$true,position=4)]
            [string]$VaultID
    )

    write-log $message
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $vaultId = $vault.ID
    }

    $RCList = Get-AzRecoveryServicesBackupContainer -VaultId $vaultId -ContainerType AzureVM | Select-Object FriendlyName,ResourceGroupName, Status
    if([string]::isNullOrEmpty($RCList) ) {
        write-log "no Recovery Containers found in this vault" -type warning
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($RCList.count) ) { #when single object available, 'count' does not exist
        write-log "single RC available: $($RCList.Name) found. selecting." -type info
        return (Get-AzRecoveryServicesBackupContainer -VaultId $vaultId -ContainerType AzureVM -FriendlyName $RCList.FriendlyName)
    } else {
        $recoveryContainer = $RClist | Out-GridView -Title $title -OutputMode Single
        if($null -eq $recoveryContainer) {
            write-host "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }
        $recoveryContainer = Get-AzRecoveryServicesBackupContainer -VaultId $vaultId -ContainerType AzureVM -FriendlyName $recoveryContainer.FriendlyName
        write-log "RV container $($recoveryContainer.name) chosen." -type info
        return $recoveryContainer
    }
}
function select-LogAnalyticsWorkspace {
    <#
    .SYNOPSIS
        acceleration function enabling to visually choose Log Analytics resources. 
    .DESCRIPTION
        #todo
    .EXAMPLE
        #todo
    .INPUTS
        None.
    .OUTPUTS
        Log Analytics resource(s)
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210310
            last changes
            - 210310 init
        
        TO|DO
    #>
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Log Analytics Workspace*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Log Analytics Workspace',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical,
        #automatically select value if there is only single option
        [Parameter(mandatory=$false,position=3)]
            [switch]$autoSelectSingleOption
    )

    write-log $message

    $LAWList = get-azOperationalInsightsWorkspace 
    if([string]::isNullOrEmpty($LAWList) ) {
        write-log "no Log Analytincs Workspaces found" -type warning
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }
    if($autoSelectSingleOption.IsPresent -and [string]::isNullOrEmpty($RCList.count) ) { #when single object available, 'count' does not exist
        write-log "single LAW available: $($LAWList.Name) found. selecting." -type info
        return $LAWList
    } else {
        $LAWorkspace = $LAWList | Select-Object Name,ResourceGroupName,CustomerId | Out-GridView -Title 'select Log Analytics workspace' -OutputMode Single
        if($null -eq $LAWorkspace) {
            write-host "Cancelled."
            if($isCritical.IsPresent) {
                exit 0
            } else {
                return $null
            }
        }
        $LAWorkspace = get-azOperationalInsightsWorkspace -ResourceGroupName $LAWorkspace.ResourceGroupName -Name $LAWorkspace.name
        write-log "LAW $($LAWorkspace.name) chosen." -type info
        return $LAWorkspace
    }
}
Set-Alias -Name select-LAWorkspace -Value select-LogAnalyticsWorkspace

Export-ModuleMember -Function * -Alias 'select-NSG','select-LAWorkspace'