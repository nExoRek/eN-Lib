<#
.SYNOPSIS
    GUI for destroy-VMWithRelatedResources
.DESCRIPTION
    wizard function allowing to remove VM by first selecting it via out-gridview and confirming resoureces for deletion.
    Example script showing AzPseudoGUI usage and eNLib. both may be installed with install-module.
.EXAMPLE
    .\destroy-AzureVM.ps1

    run winzard allowing to choose VM and destroy it.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 210301
        last changes
        - 210301 checkbox list
        - 210215 return 0 fix
        - 210128 initialized
    
    TO|DO
#>
#requires -module AzPseudoGUI
[CmdletBinding()]
param()
function show-checkBoxList { 
    param()
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    $vShift = 20
    $allChkb = 1
    $nrOfDisks = $toRemoveManagedDisks.count
    $nrOfNICs = $toRemoveNICs.count
    $nrOfPIP = $toRemovePublicIPs.count
    $nrOfNSG = $toRemoveNSGs.count
    
    
    $chkForm = New-Object system.Windows.Forms.Form
    $chkForm.text = "Remove Resources"
    $chkForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
    $chkForm.AutoSize = $true
    #$chkForm.MaximumSize = New-Object System.Drawing.Size(250,500) 
    $chkForm.StartPosition = 'CenterScreen'
    $chkForm.FormBorderStyle = 'Fixed3D'
    $chkForm.margin = 0
    $chkForm.padding = 0
    $chkForm.Icon = [System.Drawing.SystemIcons]::Question
    $chkForm.Topmost = $true
    $chkForm.MaximizeBox = $false
    
    $chkVMBox = new-object System.Windows.Forms.GroupBox
    $chkVMBox.MinimumSize = New-Object System.Drawing.Size(200,100) 
    $chkVMBox.Location = New-Object System.Drawing.Point(10,10)
    $chkVMBox.Text = 'VM resources'
    $chkVMBox.AutoSize = $true
    $chkVMBox.margin = 0
    $chkVMBox.padding = 0
    #$chkVMBox.Anchor = 'left,top'
    $lastControl = $chkVMBox
    
    if($toRemoveBootDiagStorageContainer) { 
        #[resourceId('Microsoft.Storage/storageAccounts/blobServices/containers', <'storageAccountName'>, 'default', <'storageContainerName'>)]
        $chkbVMdiag = New-Object System.Windows.Forms.Checkbox 
        $chkbVMdiag.Location = New-Object System.Drawing.Point(10,20) 
        #$chkbVMdiag.Anchor = 'left,top'
        $chkbVMdiag.AutoSize = $true
        $chkbVMdiag.Text = "Boot Diag $($toRemoveBootDiagStorageContainer.Context.StorageAccountName)"
        $chkbVMdiag.Checked = $true
        $chkbVMdiag.TabIndex = $allChkb++
        $chkVMBox.Controls.Add($chkbVMdiag)

        $lastControl = $chkbVMdiag
    }

    $vLocation = $lastControl.bottom + $vShift
    $chkVMDisks = new-object System.Windows.Forms.GroupBox
    #$chkVMDisks.MinimumSize = New-Object System.Drawing.Size(180,20) 
    $chkVMDisks.AutoSize = $false
    $chkVMDisks.Height = ($nrOfDisks + 1) * $vShift + 10
    #$chkVMDisks.size = New-Object System.Drawing.Size( 180, [int](($nrOfDisks+2)*$vShift) ) 
    $chkVMDisks.Location = New-Object System.Drawing.Point(10,$vLocation)
    $chkVMDisks.Text = 'DISKs'
    #$chkVMDisks.Anchor = 'left,top'

    $processed=0
    foreach($disk in $toRemoveManagedDisks) {
        $chkbDisk = New-Object System.Windows.Forms.Checkbox 
        $chkbDisk.Location = New-Object System.Drawing.Point(10, ($vShift+($processed*$vShift)) ) 
        #$chkbDisk.Anchor = 'left,top'
        $chkbDisk.AutoSize = $true
        $chkbDisk.Tag = $disk.ID
        #$chkbDisk.size = New-Object System.Drawing.Size( 160, $vShift)
        $chkbDisk.margin = 0
        $chkbDisk.padding = 0
        $chkbDisk.Text = "$($disk.DiskSizeGB)GB, $($disk.OsType)"
        $chkbDisk.Checked = $True
        $chkbDisk.TabIndex = $allChkb++
        $chkVMDisks.Controls.Add($chkbDisk)
        $processed++
    }
    $chkVMBox.Controls.Add($chkVMDisks)
    $lastControl=$chkVMDisks
    
    if($toRemoveNICs) {
        $vLocation = $lastControl.Bottom
    
        $chkVMNICs = new-object System.Windows.Forms.GroupBox
        $chkVMNICs.AutoSize = $false
        #$chkVMNICs.size = New-Object System.Drawing.Size(180,[int](($nrOfNIC+2)*$vShift)) 
        $chkVMNICs.Height = ($nrOfNICs + 1) * $vShift + 10
        $chkVMNICs.Location = New-Object System.Drawing.Point(10,$vLocation)
        $chkVMNICs.Text = 'NICs'
        #$chkVMNICs.Anchor = 'left,top'
    
        $processed = 0
        foreach($nic in $toRemoveNICs) {
            $chkbNIC = New-Object System.Windows.Forms.Checkbox 
            $chkbNIC.Location = New-Object System.Drawing.Point(10, ($vShift+($processed*$vShift)) ) 
            #$chkbNIC.Anchor = 'left,top'
            $chkbNIC.AutoSize = $true
            #$chkbNIC.size = New-Object System.Drawing.Size( 160, $vShift) 
            $chkbNIC.Text = "$($nic.name): $($nic.IpConfigurations.PrivateIpAddress)"
            $chkbNIC.Tag = $nic.ID
            $chkbNIC.Checked = $true
            $chkbNIC.TabIndex = $allChkb++
            $chkVMNICs.Controls.Add($chkbNIC)
            $processed++
        }
        $chkVMBox.Controls.Add($chkVMNICs)
        $lastControl = $chkVMNICs
    }

    if($toRemovePublicIPs) {
        $vLocation = $lastControl.Bottom
    
        $chkVMPIPs = new-object System.Windows.Forms.GroupBox
        $chkVMPIPs.AutoSize = $false
        $chkVMPIPs.Height = ($nrOfPIP + 1) * $vShift + 10
        #$chkVMPIPs.size = New-Object System.Drawing.Size(180,[int](($nrOfPIP+2)*$vShift)) 
        $chkVMPIPs.Location = New-Object System.Drawing.Point(10,$vLocation)
        $chkVMPIPs.Text = 'Public IPs'
        #$chkVMNICs.Anchor = 'left,top'
    
        $processed = 0
        foreach($pip in $toRemovePublicIPs) {
            $chkbPIP = New-Object System.Windows.Forms.Checkbox 
            $chkbPIP.Location = New-Object System.Drawing.Point(10, ($vShift+($processed*$vShift)) ) 
            #$chkbNIC.Anchor = 'left,top'
            $chkbPIP.AutoSize = $true
            #$chkbPIP.size = New-Object System.Drawing.Size( 160, $vShift) 
            $chkbPIP.Text = "$($PIP.IpAddress)"
            $chkbPIP.Tag = $pip.ID
            $chkbPIP.Checked = $true
            $chkbPIP.TabIndex = $allChkb++
            $chkVMPIPs.Controls.Add($chkbPIP)
            $processed++
        }
        $chkVMBox.Controls.Add($chkVMPIPs)
        $lastControl = $chkVMPIPs
    }

    if($toRemoveNSGs) {
        $vLocation = $lastControl.Bottom
    
        $chkVMNSGs = new-object System.Windows.Forms.GroupBox
        $chkVMNSGs.AutoSize = $false
        $chkVMNSGs.Height = ($nrOfNSG + 1) * $vShift + 10
       #$chkVMNSGs.size = New-Object System.Drawing.Size(180,[int](($nrOfPIP+2)*$vShift)) 
        $chkVMNSGs.Location = New-Object System.Drawing.Point(10,$vLocation)
        $chkVMNSGs.Text = 'Empty NSGs'
        #$chkVMNSGs.Anchor = 'left,top'

        $processed = 0
        foreach($nsg in $toRemoveNSGs) {
            $chkbNSG = New-Object System.Windows.Forms.Checkbox 
            $chkbNSG.Location = New-Object System.Drawing.Point(10, ($vShift+($processed*$vShift)) ) 
            #$chkbNIC.Anchor = 'left,top'
            $chkbNSG.AutoSize = $true
            #$chkbNSG.size = New-Object System.Drawing.Size( 160, $vShift) 
            $chkbNSG.Text = $NSG.name
            $chkbNSG.Tag = $NSG.ID
            $chkbNSG.Checked = $true
            $chkbNSG.TabIndex = $allChkb++
            $chkVMNSGs.Controls.Add($chkbNSG)
            $processed++       
        }
        $chkVMBox.Controls.Add($chkVMNSGs)
        $lastControl = $chkVMNSGs
    }

    if($bckRecoveryItem) {
        $vLocation = $lastControl.bottom
        $chkVMBackup = new-object System.Windows.Forms.GroupBox
        $chkVMBackup.AutoSize = $false
        $chkVMBackup.Height = 2*$vShift+10
        $chkVMBackup.Location = New-Object System.Drawing.Point(10,$vLocation)
        $chkVMBackup.Text = 'Backup'
    
        $chkbVMrecoveryItem = New-Object System.Windows.Forms.Checkbox 
        $chkbVMrecoveryItem.Location = New-Object System.Drawing.Point(10,$vShift) 
       #$chkbVMrecoveryItem.Anchor = 'left,top'
        $chkbVMrecoveryItem.AutoSize = $true
        $chkbVMrecoveryItem.Text = "{0}, {1}" -f $bckRecoveryItem.ProtectionStatus, $bckRecoveryItem.LastBackupTime
        $chkbVMrecoveryItem.Tag = $bckRecoveryItem.ID
        $chkbVMrecoveryItem.TabIndex = $allChkb++
        
        $chkVMBackup.Controls.Add($chkbVMrecoveryItem)
        $chkVMBox.Controls.Add($chkVMbackup)

        $lastControl = $chkVMbackup
    }

    
        #$chkVMBox.size = new-object system.Drawing.size(180,[int]($chkbVMdiag.bottom + $vShift)) 
    #    $chkVMBox.AutoSize = $false
    
        $vLocation = $lastControl.bottom + 2*$vShift
        $btOK = New-Object System.Windows.Forms.Button
        $btOK.Location = New-Object System.Drawing.Size(15,$vLocation)
        $btOK.Size = New-Object System.Drawing.Size(70,20)
        $btOK.Text = "OK"
        $btOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        #$btOK.Anchor = 'left,bottom'
    
        $btCancel = New-Object System.Windows.Forms.Button
        $btCancel.Location = New-Object System.Drawing.Size(($chkForm.Width-90),$vLocation)
        $btCancel.Size = New-Object System.Drawing.Size(70,20)
        $btCancel.Text = "Cancel"
        $btCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        #$btCancel.Anchor = 'right'     
    
    $chkForm.AcceptButton = $btOK
    $chkForm.CancelButton = $btCancel
    $chkForm.Controls.AddRange(@($chkVMBox, $btOK, $btCancel))
    
    $result=$chkForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach($formsClt in $chkVMBox.controls) {
            if($formsClt.gettype().name -eq 'GroupBox') {
                foreach($ctl in $formsClt.controls) {
                    if($ctl.checked) {
                        #">> $($ctl.text) <<"
                        $ctl.tag
                    }
                }
            }
        }
        if($chkbVMrecoveryItem.Checked) {
            $script:removeOSDiag=$True
        }
    } else {
        write-host 'cancelled.'
    }
}
 
connect-Azure

#######################################################################################
#                            GATHER INFORMATION ON VM RESOURCES                       #
#######################################################################################

$myCtx = select-Subscription -isCritical -message "Choose Subscription containing VM to be destroyed..." -title "Choose Subscription containing VM to be destroyed..."

$VM = select-VM -isCritical -title 'select VM for destruction' -message 'select VM to be destroyed along with resources' -status
$ResourceGroupName=$VM.ResourceGroupName
$VMName=$VM.Name

write-log "*******gathering information on VM resources********"

write-log 'getting boot diagnostics information...' -type info
$removeOSDiag=$false #this is not a resource, requires additional flag
if ($vm.DiagnosticsProfile.bootDiagnostics.Enabled) {
    if([string]::isNullOrEmpty($vm.DiagnosticsProfile.bootDiagnostics.storageUri) ) {
        write-log "storageUri empty - not able to determine storage" -type error
    } else {
        $diagSa = [regex]::match($vm.DiagnosticsProfile.bootDiagnostics.storageUri, '^http[s]?://(.+?)\.').groups[1].value
        $vmId = $vm.VmId
        $diagSaRg = (Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $diagSa }).ResourceGroupName
        $toRemoveBootDiagStorageContainer = Get-AzStorageAccount -ResourceGroupName $diagSaRg -Name $diagSa |
            Get-AzStorageContainer | Where-Object { $_.Name -match "^bootdiagnostics-.*-$vmID" }    
        write-log "boot diagnostics enabled" -silent
    }

} else {
    write-log "boot diagnosticts not enabled" -silent
}

#diagnostic extention keeps data in Tables under Storage Account... but all machines in the same tables. 
#no easy way to delete information from only single machine. 
#write-log "getting diagnostic extention information..." -type info
#if($vm.Extensions | Where-Object publisher -eq 'Microsoft.Azure.Diagnostics') {
#    $toRemoveDiagExt = Get-AzVMDiagnosticsExtension -ResourceGroupName $ResourceGroupName -VMName $VMName
#    #$extDiagSAName = [regex]::match($toRemoveDiagExt.PublicSettings,'\"storageAccount\": \"([\w]*)\",').groups[1].value
#    write-log "diagnostic extensions enabled." -silent
#} else {
#    write-log "diagnostic extension not enabled." -silent
#}
    
write-log -type info -Message 'getting Azure network interfaces...'

$toRemoveNICs = Get-AzNetworkInterface -ResourceGroupName $ResourceGroupName -Name $VMName
if($toRemoveNICs) { write-log $toRemoveNICs -silent }

write-log -type info -message 'checking NSG...'
$toRemoveNSGs=@()
foreach($nic in $toRemoveNICs) {
    if([string]::isNullOrEmpty($nic.NetworkSecurityGroup.Id) ) {
        write-log "$($nic.name) not associated with any NSG" -type info
        continue
    }
    $NSG=($nic.NetworkSecurityGroup.Id).split('/')[-1]
    write-log "NSG for $($nic.name): $NSG" -silent
    if($toRemoveNSGs -notcontains $NSG) {
        $toRemoveNSGs+=Get-AzNetworkSecurityGroup -Name $NSG -ResourceGroupName $ResourceGroupName
    }
}
write-log "getting Public IPs..." -type info
try {
    $toRemovePublicIPs = Get-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -Name $VMName
    if($toRemovePublicIPs) { write-log $toRemovePublicIPs -silent }
} catch {
    write-log "error getting Public IPs $($_.exception)" -type error
}

write-log -type info  -Message 'getting OS disk...'
#OS unmanaged disk. not tested
if ('Uri' -in $vm.StorageProfile.OSDisk.Vhd) {
    $osDiskId = $vm.StorageProfile.OSDisk.Vhd.Uri
    #$osDiskContainerName = $osDiskId.Split('/')[-2]
    $toRemoveOSDisk = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $osDiskId.Split('/')[2].Split('.')[0] }
    write-log $toRemoveOSDisk -silent
} 

## All managed disks
write-log "getting managed disks..." -type info
$toRemoveManagedDisks = Get-AzDisk | Where-Object { $_.ManagedBy -eq $vm.Id }
write-log $toRemoveManagedDisks -silent

#check if VM is backed up
write-log "checking VM backup..." -type info
$backupStatus = Get-AzRecoveryServicesBackupStatus -ResourceGroupName $VM.ResourceGroupName -Name $VM.name -Type AzureVM
if($backupStatus.BackedUp) {
    $bckContainer = Get-AzRecoveryServicesBackupContainer -ContainerType "AzureVM" -Status registered -VaultId $backupStatus.VaultId -FriendlyName $vm.Name
    $bckRecoveryItem = Get-AzRecoveryServicesBackupItem -Container $bckContainer -WorkloadType AzureVM -VaultId $backupStatus.VaultId
}

#######################################################################################
#                               LIST AND ENSURE TO DELETE                             #
#######################################################################################

write-log "ARE YOU SURE YOU WANT TO REMOVE ALL LISTED RESOURCES?`n THIS CHANGE IS IRREVERISIBLE" -type error
write-host "VM name to be removed: $VMName"
$resourcesForDeletion = show-checkBoxList
if([string]::IsNullOrEmpty($resourcesForDeletion) ) {
    write-host 'cancelled.'
    exit 0
}
write-log "removing VM $VMname" -type info
try {
    $VM|Remove-AzVM -Force
    write-log "VM removed." -type OK
} catch {
    write-log "error removing VM: $($_.exception)"
    exit -1
}
foreach($resource in $resourcesForDeletion) {
    write-log "removing $resource" -type info
    $resourceSplited=$resource.split('/')
    try{
        remove-AzResource -ResourceID $resource -force
        write-log "$($resourceSplited[-2]): $($resourceSplited[-1]) removed." -type ok
    } catch { 
        write-log "error removing $($resourceSplited[-2]): $($resourceSplited[-1])." -type error
        write-log $_.exception -type error
    }
}
if($removeOSDiag) {
    write-log "removing OS Boot Diag container"
    try {
        $toRemoveBootDiagStorageContainer|Remove-AzStorageContainer -Force
        write-log "boot diag data container removed." -type OK
    } catch {
        write-log "error removing boot diag container: $($_.exception)" -type error
    }
}
write-log "done." -type ok