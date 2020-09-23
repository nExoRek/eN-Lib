<#
.SYNOPSIS
    PS GUI for Azure VM force shutdown.
.DESCRIPTION
    script prepared to show how to create interactive PS script with elements of GUI - mainly get-OutGridView
.EXAMPLE
    .\gui-NLibForceShutdownAzVM.ps1
.NOTES
    2o2o.o9.21 ::))o- 
#>
#requires -modules Az.Accounts, Az.Compute, Az.Resources
[cmdletbinding()]
param()

function select-RG {
    param(
       $RGList 
    )
    write-host -ForegroundColor Yellow "Select Resource Group VM resides on"
    $RG=$RGList|out-gridview -title 'Select Resource Group' -OutputMode Single
    if([string]::isNullOrEmpty($RG)) {
        write-host 'cancelled by user. quitting' -foreground Yellow
        exit -1
    }
    write-host "$($RG.ResourceGroupName) chosen. reading VMs..."
    return $RG.ResourceGroupName
}
function get-VMList {
    param(
        $RGname
    )
    return (
            get-azVM -resourceGroupName $RGName -status| `
        Select-Object name,powerstate,@{N='OS';E={$_.StorageProfile.OsDisk.OsType}}, `
            @{N='image';E={($_.StorageProfile.ImageReference.id -split '/')[-1]}}, `
            @{N='size';E={$_.HardwareProfile.vmsize}}
    )
}

if([string]::IsNullOrEmpty( (Get-AzContext) ) ) {
  write-host -ForegroundColor red "you need to be connected before running this script. use connect-AzAccount first."
  exit -1
}

#choose subscription
write-host -ForegroundColor Yellow "Select Subscription"
$subscription=get-AzSubscription|out-gridview -title 'Select Subscription' -OutputMode Single
if([string]::isNullOrEmpty($subscription)) {
    write-host 'cancelled by user. quitting' -foreground Yellow
    exit -1
}
set-AzContext -SubscriptionObject $subscription

#get all RG list
$RGsInThisSub=Get-AzResourceGroup

#first loop for the last question on VM
$choseAgain=$false
do {

    #second loop for RG choice 
    do {
        #choose RG
        $ResourceGroupName=select-RG -RGList $RGsInThisSub
        $VMlist=get-VMList -RGname $ResourceGroupName
        if([string]::isNullOrEmpty($VMlist)) {
            write-host "This Resource Group do not contain any VMs. Do you want to select another RG?"
            switch (
                [System.Windows.Forms.MessageBox]::show($this,"Choose another Resource Group?",'CONFIRM','OKCancel') 
            ) {
                'OK' {
                    #do nothing - VMList is null which will trigger loop
                }
                'Cancel' {
                    Write-Host -ForegroundColor Yellow "operation cancelled by the user. quitting."
                    exit -1
                }       
            }
        }
    } while( [string]::isNullOrEmpty($VMlist) )

    $VM=$VMlist|out-gridview -title 'select Virtual Machine' -OutputMode Single
    if([string]::isNullOrEmpty($VM)) {
        write-host 'cancelled by user. quitting.' -foreground Yellow
        exit -2
    } 

    if($vm.powerstate -notmatch 'running') {
        write-host "VM state is: $($vm.powerstate)"
        write-host -ForegroundColor Red "VM status is not 'running'. Can't shutdown machine in that state."
        write-host -ForegroundColor Red "Seems that you don't need to shutdown it any more..."
        write-host "Do you want to choose another VM in another RG?"
        switch(
            [System.Windows.Forms.MessageBox]::show($this,"Retry and chose another Resource Group?",'CONFIRM','OKCancel') 
        ) {
            'OK' {
                $choseAgain=$true
            }
            'Cancel' {
                write-host -ForegroundColor Yellow "operation cancelled by the user. quitting."
                exit -1
            }       
        }
    }
} while($choseAgain)

Switch(
  [System.Windows.Forms.MessageBox]::show($this,"Do you want to FORCE SHUTDOWN $($VM.name)?",'CONFIRM','YesNo') 
  ) {
    'Yes' {
        write-host "trying to shutdown $($VM.Name) forcibly (can take several minutes)..."
        try {
            stop-AzVM -ResourceGroupName $ResourceGroupName -name $VM.Name -force 
            write-host -ForegroundColor Yellow "KABOOOOM! You nailed it!"
        } catch {
            $_
            exit -8
        }
    }
    'No' {
        write-host -ForegroundColor Yellow "operation cancelled by the user. quitting."
        exit -1
    }
}

Get-AzVM -ResourceGroupName $ResourceGroupName -name $VM.Name -status
