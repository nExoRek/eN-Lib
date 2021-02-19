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
    version 210215
        last changes
        - 210215 seelct-vnet and select-subnet fixes, select-recoveryContainer
        - 210208 select-recoveryVault, descriptions, fixes...
        - 210202 initialized alpha
    #TO|DO

#>
import-module eNLib
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
        None.
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210127
            last changes
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
            [switch]$isCritical
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
        write-log $message
        $sourceSubscription = $subscriptionList | out-GridView -title $title -OutputMode Single
        if($null -eq $sourceSubscription) {
            write-log "operation cancelled by the user."
            if($isCritical.IsPresent) {
                exit 0
            }        
            return 0
        }
        write-log "chosen subscription: $($sourceSubscription.name)" -type info
    } catch {
        write-log "error getting Subscription list $($_.exception)" -type error
        if($isCritical.IsPresent) {
            exit -3
        }        
        exit -3
    }
    
    try {
        $AzSourceContext = set-azContext -SubscriptionObject $sourceSubscription -force
        write-log "source context: $($AzSourceContext.Subscription.name)" -silent
    } catch {
        write-log "error changing context. $($_.Exception)" -type error
        if($isCritical.IsPresent) {
            exit -4
        }        
        exit -4
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
        version 210127
            last changes
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
            [switch]$isCritical
    )
    
    write-log $message
    $RGList = Get-AzResourceGroup | select-object ResourceGroupName, location, Tags
    if($null -eq $RGList) {
        write-log "No resource groups in this subscription." -type warning
        if($isCritical.IsPresent) {
            exit -1
        }
        return -1
    }
    $ResourceGroup = $RGList | Out-GridView -Title $title -OutputMode Single
    if ([string]::isNullOrEmpty($ResourceGroup)) {
        Write-log 'Cancelled by user'
        if($isCritical.IsPresent) {
            exit 0
        }
        return 0
    }
    Write-log "RG $($ResourceGroup.ResourceGroupName) chosen. " -type info
    return (Get-AzResourceGroup -Name $ResourceGroup.ResourceGroupName)
}
function select-StorageAccount {
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
    #Resource Group object
    [Parameter(parameterSetName='byObject',mandatory=$true,position=3,ValueFromPipeline=$true)]
        [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,
    #Parameter help description
    [Parameter(parameterSetName='byName',mandatory=$true,position=3)]
        [string]$ResourceGroupName
    )
    
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $ResourceGroupName = $ResourceGroup.ResourceGroupName
    }
    write-log $message
    $saList = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName |
        ? {$_.Sku.Name -match "^standard"} |
        select-object StorageAccountName,ResourceGroupName,PrimaryLocation,@{N='SkuName';E={$_.sku.name}}
    if([string]::isNullOrEmpty($saList) ) {
        write-host "there are no storage accounts in this Resource Group."
        if($isCritical.IsPresent) {
            exit -1
        } else {
            return -1
        }
    }        
    $storageAccount = $saList | Out-GridView -Title $title -OutputMode Single
    if([string]::isNullOrEmpty($storageAccount) ) {
        write-host "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }
    write-log "SA $($storageAccount.name) chosen." -type info
    return $storageAccount
}
function select-VM {
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
        version 210209
            last changes
            - 210209 initialized
    #>
    
    [CmdletBinding(DefaultParameterSetName='RGbyObj')]
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
        #resource group object containing VMs to list
        [Parameter(ValueFromPipeline=$true,parameterSetName='RGbyObj',mandatory=$false,position=3)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,  
        #resource group name containing VMs to list
        [Parameter(ValueFromPipeline=$true,parameterSetName='RGbyName',mandatory=$false,position=3)]
            [string]$ResourceGroupName,  
        #include backup information on the list
        [Parameter(mandatory=$false,position=4)]
            [switch]$status,
        #include backup information on the list
        [Parameter(mandatory=$false,position=5)]
            [switch]$includeBackupInfo
    )

    write-log $message 

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
            if( ![string]::isNullOrEmpty($backupStatus) ) {
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
    #ACTUAL VM Select
    $selectVM = $VMList | Out-GridView -Title $title -OutputMode Single
    if($null -eq $selectVM) {
        write-log "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }
    if($includeBackupInfo.IsPresent -and $selectVM.backup -ne 'no backup') {
        write-log "this machine is already backed up." -type warning
        if($isCritical.IsPresent) {
            exit -2
        } else {
            return -2
        }
    }
    if($status.IsPresent) {
        $powerstate = $selectVM.powerstate
        $VM = get-AzVM -Name $selectVM.name -ResourceGroupName $selectVM.ResourceGroupName | select-object *,@{N='powerstate';E={$powerstate}}
    } else {
        $VM = get-AzVM -Name $selectVM.name -ResourceGroupName $selectVM.ResourceGroupName 
    }
    write-log "$($VM.name) chosen type of $($VM.StorageProfile.OsDisk.OsType)." -type info
    return $VM

}
function select-vNet {
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
        #resource group object containing vNet
        [Parameter(parameterSetName='byObject',mandatory=$false,position=3,ValueFromPipeline=$true)]
            [Microsoft.Azure.Commands.ResourceManager.Cmdlets.SdkModels.PSResourceGroup]$ResourceGroup,
        #resource group name containing vNet
        [Parameter(parameterSetName='byName',mandatory=$false,position=3)]
            [string]$ResourceGroupName
    )
    if($PSCmdlet.ParameterSetName -eq 'byObject') {
        $ResourceGroupName=$ResourceGroup.ResourceGroupName
    }
    $azvNet=@{}
    if(![string]::isNullOrEmpty($ResourceGroupName) ) {
        $azvNet.add('ResourceGroupName',$ResourceGroupName)
    } 
    $vNetList = Get-AzVirtualNetwork @azvNet | select-object name,location, ResourceGroupName
    if($null -eq $vNetList) {
        write-log "there is no VNets in $ResourceGroup"
        if($isCritical.IsPresent) {
            exit -1
        }
        return -1
    }
    $vnetName = $vNetList | Out-GridView -Title $title -OutputMode Single
    if($null -eq $vnetName) {
        write-log "cancelled by the user."
        if($isCritical.IsPresent) {
            exit 0
        }
        return 0
    }
    write-log "vNet $($vnetName.name) chosen." -type info
    return (Get-AzVirtualNetwork -ResourceGroupName $vNetName.ResourceGroupName -Name $vnetName.Name)
}
function select-vSubnet {
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
        #vNet Object containing subnets to choose
        [Parameter(parameterSetName='byObject',mandatory=$true,position=3,ValueFromPipeline)]
            [Microsoft.Azure.Commands.Network.Models.PSVirtualNetwork]$vNet,
        #vNet name
        [Parameter(parameterSetName='byName',mandatory=$true,position=3,ValueFromPipeline)]
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
    $vSubnet = $vSubnetList | Out-GridView -Title $title -OutputMode Single
    if([string]::isNullOrEmpty($vSubnet) ) {
        write-host "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }
    write-log "vSubnet $($vSubnet.name) chosen." -type info
    return (Get-AzVirtualNetworkSubnetConfig -VirtualNetwork $vNet -Name $vSubnet.name)    
}
function select-KeyVault {
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = "choose *Key Vault*...",
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = "choose Key Vault",
        #changes return to terminal - define if cancel/empty is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical    
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
    $kv = $KVList | Out-GridView -Title $title -OutputMode Single
    if($null -eq $kv) {
        write-log "cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    } else {
        write-log "KeyVault $($kv.VaultName) chosen." -type info
        return (Get-AzKeyVault -ResourceGroupName $kv.ResourceGroupName -VaultName $kv.VaultName)
    }
}
function select-encryptionKey {
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
        #Vault name to choose key from
        [parameter(mandatory=$true,position=3)]
            [string]$vaultName

    )
    
    write-log $message
    $encKeyList = Get-AzKeyVaultKey -VaultName $vaultName | Select-Object name,enabled,Expires
    if($null -eq $encKeyList) {
        write-log "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }    
    $key = $encKeyList | Out-GridView -Title $title -OutputMode Single
    if($null -eq $key) {
        write-log "cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    } else {
        write-log "key $($key.name) chosen." -type info
        return (Get-AzKeyVaultKey -VaultName $vaultName -Name $key.name)
    }

}
function select-recoveryVault {
    param(
        #message displayed on the screen before windows popup
        [Parameter(mandatory=$false,position=0)]
            [string]$message = 'Select *Backup Vault*...',
        #short title message shown on GridView title bar
        [Parameter(mandatory=$false,position=1)]
            [string]$title = 'Select Backup Vault',
        #changes return to terminal - define if cancel/empty RG is critical and will result in Exit
        [Parameter(mandatory=$false,position=2)]
            [switch]$isCritical
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
    $recoveryVault = $RVlist | Out-GridView -Title $title -OutputMode Single
    if($null -eq $recoveryVault) {
        write-host "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }
    $recoveryVault = Get-AzRecoveryServicesVault -name $recoveryVault.name
    write-log "Recovery Vault $($recoveryVault.name) chosen." -type info
    return $recoveryVault
}
function select-recoveryContainer {
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
        #Vault object
        [Parameter(parameterSetName='byObject',mandatory=$true,position=3, ValueFromPipeline)]
            [Microsoft.Azure.Commands.RecoveryServices.ARSVault]$Vault,
        #VaultID
        [Parameter(parameterSetName='byName',mandatory=$true,position=3)]
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
    $recoveryContainer = $RClist | Out-GridView -Title $title -OutputMode Single
    if($null -eq $recoveryContainer) {
        write-host "Cancelled."
        if($isCritical.IsPresent) {
            exit 0
        } else {
            return 0
        }
    }
    $recoveryContainer = Get-AzRecoveryServicesBackupContainer -VaultId $vaultId -ContainerType AzureVM -FriendlyName $recoveryContainer.FriendlyName
    write-log "RV container $($recoveryContainer.name) chosen." -type info
    return $recoveryContainer
}

Export-ModuleMember -Function *