<#
.SYNOPSIS
    automated VM removal along with all depandand resources. 
.DESCRIPTION
    when removing VM all depand resources are being left. script is removing:
    - VM
    - OS disk
    - data disks
    - NICs
    - Public IPs 
    - boot diagnostics BLOB

    what is not deleted:
    - Storage Accounts for diagnostics boot/ext
    - unamanged data disks
    - exteded diagnostics data Tables
    - log analitics
    - backup

         *USE WITH CARE*
.EXAMPLE
    .\005-DestroyOldVMAndRelatedResources.ps1
    lanunches GUI-wizard allowing to choose VM for destruction.
.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 210112
        last changes
        - 210112 v1
        - 201212 initialized

  TO|DO
   - ability to choose which exact resources to be removed instead of ALL
   - multiple NIC VM not tested
   - unmanaged OS disk not tested
#>
#requires -module Az.Accounts,Az.Compute
[CmdletBinding()]
param()

function start-Logging {
    param(
        #create log in profile folder rather than script run path
        [Parameter(mandatory=$false,position=0)]
            [switch]$userProfilePath
    )
  
    $scriptBaseName = ([System.IO.FileInfo]$PSCommandPath).basename
    if($userProfilePath.IsPresent) {
        $logFolder = [Environment]::GetFolderPath("MyDocuments") + '\Logs'
    } else {
        $logFolder = "$PSScriptRoot\Logs"
    }
  
    if(-not (test-path $logFolder) ) {
        try{ 
            New-Item -ItemType Directory -Path $logFolder|Out-Null
            write-host "$LogFolder created."
        } catch {
            $_
            exit -2
        }
    }
  
    $script:logFile="{0}\_{1}-{2}.log" -f $logFolder,$scriptBaseName,$(Get-Date -Format yyMMddHHmm)
    write-Log "*logging initiated $(get-date)" -silent -skipTimestamp
    write-Log "*script parameters:" -silent -skipTimestamp
    if($script:PSBoundParameters.count -gt 0) {
        write-log $script:PSBoundParameters -silent -skipTimestamp
    } else {
        write-log "<none>" -silent -skipTimestamp
    }
    write-log "***************************************************" -silent -skipTimestamp
}
function write-log {
    param(
        #message to display - can be an object
        [parameter(mandatory=$false,position=0)]
                $message,
        #adds description and colour dependently on message type
        [parameter(mandatory=$false,position=1)]
            [string][validateSet('error','info','warning','ok')]$type,
        #do not output to a screen - logfile only
        [parameter(mandatory=$false,position=2)]
            [switch]$silent,
        # do not show timestamp with the message
        [Parameter(mandatory=$false,position=3)]
            [switch]$skipTimestamp
    )

    #ensure that whatever the type is - array, object.. - it will be output as string, add runtime
    if( [string]::IsNullOrEmpty($message) ) {
        $message='<NullOrEmpty>'
    } else {
        $message=($message|out-String).trim() 
    }

    try {
        $finalMessageString=@()
        if(-not $skipTimestamp) {
            $finalMessageString += "$(Get-Date -Format "hh:mm:ss>") "
        }
        if(-not [string]::IsNullOrEmpty( $type) ) { 
            $finalMessageString += $type.ToUpper()+": " 
        }
        $finalMessageString += $message
        $message=$finalMessageString -join ''
        Add-Content -Path $logFile -Value $message
        if(-not $silent) {
            switch($type) {
                'error' {
                    write-host -ForegroundColor Red $message
                }
                'info' {
                    Write-Host -ForegroundColor DarkGray $message
                }
                'warning' {
                    Write-Host -ForegroundColor Yellow $message
                }
                'ok' {
                    Write-Host -ForegroundColor Green $message
                }
                default {
                    Write-Host $message 
                }
            }
        }
    } catch {
        Write-Error 'not able to write to log. suggest to cancel the script run.'
        $_
    }    
}
function connect-Azure {
    try {
        $AzSourceContext=Get-AzContext
    } catch {
        write-log $_.exception -type error
        exit -1
    }
    if([string]::IsNullOrEmpty( $AzSourceContext ) ) {
            write-log "you need to be connected before running this script. use connect-AzAccount first." -type error
            exit -1
    }
    write-log "connected to $($AzSourceContext.Subscription.name) as $($AzSourceContext.account.id)" -silent -type info
    write-host "Your Azure connection:"
    write-host "  subscription: " -noNewLine
    write-host -foreground Yellow "$($AzSourceContext.Subscription.name)"
    write-host "  connected as: " -noNewLine 
    write-host -foreground Yellow "$($AzSourceContext.account.id)"
    Set-Item Env:\SuppressAzurePowerShellBreakingChangeWarnings "true"
}
function select-Subscription {
    try {
      $sourceSubscription=Get-AzSubscription|out-GridView -title "Select Subscription to host VM" -PassThru
      write-log "chosen sub: $($sourceSubscription.name)" -type info
    } catch {
      write-log "Error getting Source Subscription. Quitting." -type error
      write-log $_
      exit -3
    }
    
    try {
      $AzSourceContext=set-azContext -SubscriptionObject $sourceSubscription -force
      write-log "source context: $($AzSourceContext.Subscription.name)" -silent
    } catch {
      write-log "Error changing context $($_.Exception).`n Quitting." -type error
      exit -3
    }
}
function select-ResourceGroup {
    param(
    )
    write-log "Select *Resource Group* to place VM" -type warning
    $RG=Get-AzResourceGroup|select-object ResourceGroupName|out-gridview -title 'Select Resource Group' -OutputMode Single
    if([string]::isNullOrEmpty($RG)) {
        write-log 'cancelled by user. quitting' -type warning
        exit -3
    }
    write-host "$($RG.ResourceGroupName) chosen."
    return $RG.ResourceGroupName
}
function get-valueFromInputBox {
    <#
    .SYNOPSIS
        simple input message box function for PS GUI scripts
    .EXAMPLE
        $response = get-valueFromInputBox -title 'WARNING!' -text "type 'YES' to continue" -type Warning
        if($null -eq $response) {
            'cancelled'
            exit 0
        }
        if($response -ne 'YES') {
            'not correct. quitting'
            exit -1
        } else {
            "you agreed, let's continue"
        }
        write-host 'code to execute here'
    .INPUTS
        None.
    .OUTPUTS
        User Input or $null for cancel
    .LINK
        https://w-files.pl
    .NOTES
        nExoR ::))o-
        version 210113
            last changes
            - 210113 initialized
    #>
    
    param(
        [parameter(mandatory=$false,position=0)]
            [string]$title='input',
        [parameter(mandatory=$false,position=1)]
            [string]$text='put your input',
        [parameter(mandatory=$false,position=2)]
            [validateSet('Asterisk','Error','Exclamation','Hand','Information','None','Question','Stop','Warning')]
            [string]$type='Question'
    )
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $promptWindowForm = New-Object system.Windows.Forms.Form
    $promptWindowForm.ClientSize = '250,100'
    $promptWindowForm.text = $title
    $promptWindowForm.BackColor = "#ffffff"
    $promptWindowForm.AutoSize = $true
    $promptWindowForm.StartPosition = 'CenterScreen'
    $promptWindowForm.Icon = [System.Drawing.SystemIcons]::$type
    $promptWindowForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

    $lblPromptInfo = New-Object System.Windows.Forms.Label 
    $lblPromptInfo.Location = New-Object System.Drawing.Size(10,5) 
    $lblPromptInfo.Size = New-Object System.Drawing.Size(230,40)
    $lblPromptInfo.Text = $text

    $txtUserInput = New-Object system.Windows.Forms.TextBox
    $txtUserInput.multiline = $false
    $txtUserInput.ReadOnly = $false
    $txtUserInput.width = 230
    $txtUserInput.height = 25
    $txtUserInput.location = New-Object System.Drawing.Point(10, 50)

    $btOK = New-Object System.Windows.Forms.Button
    $btOK.Location = New-Object System.Drawing.Size(30,80) 
    $btOK.Size = New-Object System.Drawing.Size(70,20)
    $btOK.ForeColor = "Green"
    $btOK.Text = "Continue"
    $btOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btCancel = New-Object System.Windows.Forms.Button
    $btCancel.Location = New-Object System.Drawing.Size(150,80) 
    $btCancel.Size = New-Object System.Drawing.Size(70,20)
    $btCancel.ForeColor = "Red"
    $btCancel.Text = "Cancel"
    $btCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $promptWindowForm.AcceptButton=$btOK
    $promptWindowForm.CancelButton=$btCancel
    $promptWindowForm.Controls.AddRange(@($lblPromptInfo, $txtUserInput,$btOK,$btCancel))
    $promptWindowForm.Topmost = $true
    $promptWindowForm.Add_Shown( { $promptWindowForm.Activate();$txtUserInput.Select() })
    $result=$promptWindowForm.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $response = $txtUserInput.Text
        return $response
    }
    else {
        return $null
    }   
}
function get-ResourcesForDeletion {
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $chooseResourcesForm = New-Object system.Windows.Forms.Form
    $chooseResourcesForm.AutoSize = $true
    $chooseResourcesForm.text = "Choose resources to be deleted"
    $chooseResourcesForm.TopMost = $false
    $chooseResourcesForm.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)

    $gbResources = New-Object system.Windows.Forms.Groupbox
    $gbResources.AutoSize = $true
    $gbResources.text = "VM: $VMName"
    $gbResources.location = New-Object System.Drawing.Point(10, 10)

    $chbcheckBoxName = New-Object System.Windows.Forms.Checkbox 
    $chbcheckBoxName.Location = New-Object System.Drawing.Size(15,15) 
    $chbcheckBoxName.Size = New-Object System.Drawing.Size(20,20)
    $chbcheckBoxName.Text = "Text"
    [void]$chooseResourcesForm.ShowDialog()

}

start-Logging
connect-Azure

#######################################################################################
#                            GATHER INFORMATION ON VM RESOURCES                       #
#######################################################################################

write-log "Choose Subscription containing VM to be destroyed..." -type warning
select-Subscription

write-log "Choose VM to be destroyed..." -type warning
$selectVM=get-AzVM|Select-Object name,@{N='OsType';E={$_.StorageProfile.OsDisk.OsType}},@{N='VmSize';E={$_.HardwareProfile.VmSize}},Location,ResourceGroupName|Out-GridView -Title 'select VM for destruction' -OutputMode Single
$ResourceGroupName=$selectVM.ResourceGroupName
$VMName=$selectVM.Name
write-log "$VMName in RG $ResourceGroupName chosen." -type info
$vm = Get-AzVm -Name $VMName -ResourceGroupName $ResourceGroupName
    
write-log "gathering information on VM resources..." -type info

write-log -type info -Message 'getting boot diagnostics information...'
if ($vm.DiagnosticsProfile.bootDiagnostics.Enabled) {
    $diagSa = [regex]::match($vm.DiagnosticsProfile.bootDiagnostics.storageUri, '^http[s]?://(.+?)\.').groups[1].value
    $vmId = $vm.VmId
    $diagSaRg = (Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $diagSa }).ResourceGroupName
    $toRemoveBootDiagStorageContainer = Get-AzStorageAccount -ResourceGroupName $diagSaRg -Name $diagSa |
        Get-AzStorageContainer | Where-Object { $_.Name -match "^bootdiagnostics-.*-$vmID" }    
    write-log "boot diagnostics enabled" -silent
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
    $NSG=($nic.NetworkSecurityGroup.Id).split('/')[-1]
    write-log "NSG for $($nic.name): $NSG" -silent
    if($toRemoveNSGs -notcontains $NSG) {
        $toRemoveNSGs+=$NSG
    }
}
write-log "getting Public IPs..." -type info
$toRemovePublicIPs = Get-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -Name $VMName
if($toRemovePublicIPs) { write-log $toRemovePublicIPs -silent }

write-log -type info  -Message 'getting OS disk...'
#OS unmanaged disk. not tested
if ('Uri' -in $vm.StorageProfile.OSDisk.Vhd) {
    $osDiskId = $vm.StorageProfile.OSDisk.Vhd.Uri
    $osDiskContainerName = $osDiskId.Split('/')[-2]

    $toRemoveOSDisk = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $osDiskId.Split('/')[2].Split('.')[0] }
    write-log $toRemoveOSDisk -silent
} 

## All managed disks
write-log "getting managed disks..." -type info
$toRemoveManagedDisks=Get-AzDisk | Where-Object { $_.ManagedBy -eq $vm.Id }
write-log $toRemoveManagedDisks -silent

#######################################################################################
#                               LIST AND ENSURE TO DELETE                             #
#######################################################################################

write-log "ARE YOU SURE YOU WANT TO REMOVE ALL LISTED RESOURCES?`n THIS CHANGE IS IRREVERISIBLE" -type error
write-host "VM: $VMName"
if($toRemoveBootDiagStorageContainer) { 
    Write-Host "`tBoot Diag:"
    write-host "`t`t* Boot diag container enabled under SA $($toRemoveBootDiagStorageContainer.Context.StorageAccountName)"
}
if($toRemoveOSDisk) {
    write-host "`tunamanged OSDisk: $($toRemoveOSDisk.OsType), $($toRemoveOSDisk.diskSizeGB)GB, SKU $($toRemoveOSDisk.sku.name)"
}
if($toRemoveManagedDisks) {
    write-host "`tData disks ($($toRemoveManagedDisks.count)):"
    foreach($dataDisk in $toRemoveManagedDisks) {
        write-host "`t`t* $($dataDisk.name), $($dataDisk.DiskSizeGB)GB, $($dataDisk.OsType)"
    }
}
if($toRemoveNICs) {
    write-host "`tNetwork Interfaces ($($toRemoveNICs.count)):"
    foreach($NIC in $toRemoveNICs) {
        write-host "`t`t* $($nic.name): $($nic.IpConfigurations.PrivateIpAddress)"
    }
}
if($toRemovePublicIPs) {
    write-host "`tPublic IPs ($($toRemovePublicIPs.count)):"
    foreach($PIP in $toRemovePublicIPs) {
        write-host "`t`t* $($PIP.IpAddress)"
    }
}
if($toRemoveNSGs) {
    Write-Host "`tNetwork Security Groups ($($toRemoveNSGs.count)):"
    foreach($NSG in $toRemoveNSGs) {
        write-host "`t`t* $NSG"
    }
}

$deleteOrNot = get-valueFromInputBox -title 'WARNING!' -text 'All resources will be irrevocably removed. Type ''YES'' to continue' -type Warning

#######################################################################################
#                               ACUTAL REMOVAL OF RESOURCES                           #
#######################################################################################

if($deleteOrNot -eq 'YES') {

    #if($toRemoveDiagExt) { #Tables are nor disappearing. how to query them?
    #    write-log -type info "removing Extension Diagnostics storage account..."
    #    Remove-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $extDiagSAName -Force
    #}

    write-log -type info  -Message 'removing the Azure VM...'
    $null = $vm | Remove-AzVM -Force
    if($toRemoveBootDiagStorageContainer) {
        write-log -type info "removing OS diagnostics storage container..."
        $toRemoveBootDiagStorageContainer|Remove-AzStorageContainer -Force
    }

    if($toRemoveNICs) {
        write-log -type info "removing NICs..."
        foreach($nic in $toRemoveNICs) {
            Remove-AzNetworkInterface -Name $nic.Name -ResourceGroupName $ResourceGroupName -Force
        }
    }

    write-log -type info "removing Public IPs..."
    foreach($PIP in $toRemovePublicIPs) {
        Remove-AzPublicIpAddress -ResourceGroupName $ResourceGroupName -Name $PIP.Id.Split('/')[-1] -Force
    }
    if($toRemoveOSDisk) {
        write-log "removing unamanaged OS disk..."
        $toRemoveOSDisk | Remove-AzStorageBlob -Container $osDiskContainerName -Blob $osDiskId.Split('/')[-1]
        $toRemoveOSDisk | Get-AzStorageBlob -Container $osDiskContainerName -Blob "$($VMName)*.status" | Remove-AzStorageBlob
    }
    write-log "removing managed disks..." -type info
    $toRemoveManagedDisks| Remove-AzDisk -Force
    write-log "removing Network Secuirty Groups..." -type info
    foreach($NSGName in $toRemoveNSGs) {
        $NSG=Get-AzNetworkSecurityGroup -Name $NSGName
        if([string]::IsNullOrEmpty($NSG.NetworkInterfaces)) { 
            Remove-AzNetworkSecurityGroup -Name $NSG.name -ResourceGroupName $NSG.ResourceGroupName -Force
        } else {
            Write-log "$($NSG.name) is still being used. leaving." -type warning
        }
    }
    write-log "VM destroyed." -type ok
} else {
    write-log "cacelled." -type ok
}
