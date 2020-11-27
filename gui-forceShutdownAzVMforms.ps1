<#
.SYNOPSIS
    Forms GUI to force shutdown Azure VM
.DESCRIPTION
.    script prepared to show how to create interactive PS script with elements of GUI - win Forms
.EXAMPLE
    .\gui-NLibForceShutdownAzVMforms.ps1
.NOTES
    2o2o.o9.21 ::))o- 
#>
#requires -modules Az.Accounts, Az.Compute, Az.Resources
[cmdletbinding()]
param()

#region Interface
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptGUI                  = New-Object system.Windows.Forms.Form
$scriptGUI.ClientSize       = '500,100'
$scriptGUI.StartPosition    = 'CenterScreen'
$scriptGUI.text             = "Az Force Shutdown"
$scriptGUI.BackColor        = "#ffffff"
$scriptGUI.AutoSize         = $true
#$scriptGUI.TopMost          = $true #this is 'always on front'
$scriptGUI.bri
$scriptGUI.Icon = [System.Drawing.SystemIcons]::Warning

$labelStatusLabel                = New-Object System.Windows.Forms.Label
$labelStatusLabel.text           = ""
$labelStatusLabel.AutoSize       = $false
$labelStatusLabel.width          = 350
$labelStatusLabel.height         = 30
$labelStatusLabel.location       = New-Object System.Drawing.Point(10,50)
$scriptGUI.controls.Add($labelStatusLabel)

$dropBoxSubscriptions             = New-Object system.windows.forms.ComboBox
$dropBoxSubscriptions.Text        = "choose subscription"
$dropBoxSubscriptions.Width       = 170
$dropBoxSubscriptions.Location    = New-Object System.Drawing.Point(5,10)
$dropBoxSubscriptions.AutoSize    = $true

$dropBoxRGs             = New-Object system.windows.forms.ComboBox
$dropBoxRGs.Text        = ""
$dropBoxRGs.Width       = 170
$dropBoxRGs.Location    = New-Object System.Drawing.Point(180,10)
$dropBoxRGs.AutoSize    = $true
$scriptGUI.controls.Add( $dropBoxRGs )

$dropBoxVMs             = New-Object system.windows.forms.ComboBox
$dropBoxVMs.Text        = ""
$dropBoxVMs.Width       = 170
$dropBoxVMs.Location    = New-Object System.Drawing.Point(360,10)
$dropBoxVMs.AutoSize    = $true
$scriptGUI.controls.Add( $dropBoxVMs )

$okButton               = New-Object System.Windows.Forms.Button
$okButton.Name          = "okButton"
$okButton.Size          = New-Object System.Drawing.Size(100,50)
#$okButton.UseVisualStyleBackColor = $True
$okButton.ForeColor     = [System.Drawing.Color]::Red
$okButton.Text          = "SHUTDOWN"
$okButton.Location      = New-Object System.Drawing.Point(390,40)
$okButton.DataBindings.DefaultDataSourceUpdateMode = 0

#endregion Interface

#region InterfaceActions
#dropdown list: subscription . when sub is chosen, change context and read Resource Groups to create RG list 
$dropBoxSubscriptions_SelectedIndexChanged= {
    [void] $dropBoxSubscriptions.Items.Remove("choose subscription")
    $labelStatusLabel.text = "changing context...."

    set-azContext -SubscriptionName $dropBoxSubscriptions.SelectedItem|out-null

    $labelStatusLabel.text = "reading Resource Groups...."

    #when sub selected - re-read Resource Groups
    $RGs=Get-AzResourceGroup
    [void] $dropBoxRGs.Items.Clear()
    [void] $dropBoxRGs.Items.Add("")
    $RGs|ForEach-Object{[void] $dropBoxRGs.Items.Add($_.ResourceGroupName)} 
    #$dropBoxRGs.SelectedIndex = 0
    $dropBoxRGs.Text = ""

    #clean VM list
    [void] $dropBoxVMs.Items.Clear()
    [void] $dropBoxVMs.Items.Add("")
    #$dropBoxVMs.SelectedIndex = 0
    $dropBoxVMs.Text = ""

    #just in case anything has left - disable shutdown until VM is chosen. 
    $scriptGUI.controls.Remove( $okButton )
    
    $labelStatusLabel.text = "found $($RGs.count) RGs."
}
$dropBoxSubscriptions.add_SelectedIndexChanged($dropBoxSubscriptions_SelectedIndexChanged)

#dropdown list: Resource Groups. when chosen - read available VMs.
$dropBoxRGs_SelectedIndexChanged={
    [void] $dropBoxRGs.Items.Remove("")

    #cleanup VMs
    [void] $dropBoxVMs.Items.Clear()
    [void] $dropBoxVMs.Items.Add("")
    $dropBoxVMs.Text = ""
    $scriptGUI.controls.Remove( $okButton )

    $script:RGName=$dropBoxRGs.SelectedItem
    if(-not [string]::IsNullOrEmpty($script:RGName) ) {

        $labelStatusLabel.text = "reading VMs...."

        $script:VMs=Get-AzVM -ResourceGroupName $script:RGName -status
        $script:VMs|ForEach-Object{[void] $dropBoxVMs.Items.Add($_.Name)} 
        #$dropBoxVMs.SelectedIndex = 0
    
        $labelStatusLabel.text = "found $($VMs.count) VMs"
    }
}
$dropBoxRGs.add_SelectedIndexChanged($dropBoxRGs_SelectedIndexChanged)

#dropdown list: VMs. when chosen, add 'OK' button. 
$dropBoxVMs_SelectedIndexChanged={
    [void] $dropBoxVMs.Items.Remove("")
    $script:VMName=$dropBoxVMs.SelectedItem
    if(-not [string]::IsNullOrEmpty($script:VMName) ) {
        $vmStatus=($script:VMs|Where-Object name -eq $script:VMName).powerstate
        $labelStatusLabel.text = "Machine powerstate is $vmStatus"
        if($vmStatus -match 'running') {
            $scriptGUI.controls.Add( $okButton )
        } else {
            $labelStatusLabel.text += "`ncan't shutdown this machine as it's not running"
        }
    }
}
$dropBoxVMs.add_SelectedIndexChanged($dropBoxVMs_SelectedIndexChanged)

$okButton.add_Click({
    Switch( [System.Windows.Forms.MessageBox]::show(
        $this,"Are you sure you want to force shutdown ""$VMName"" in RG ""$RGName""?",'CONFIRM','YesNo')
    ) {
        'Yes' {
            $labelStatusLabel.text = "force shutdown in progress...`n(this may take some time)"
            stop-AzVM -ResourceGroupName $RGName -name $VMName -force
            $labelStatusLabel.text = "KABOOM!"
            $scriptGUI.controls.Remove( $okButton )
        } 
        'No' {
            $labelStatusLabel.text = "canceled"
        }
    }
})
#endregion InterfaceActions

#######################################
#             INITIALIZE

$RGName=""
$VMName=""
$VMs=""

$AzSubs=Get-AzSubscription -ErrorAction SilentlyContinue
if([string]::IsNullOrEmpty($AzSubs)) {
    try {
        write-host "connecting..."
        connect-AzAccount 
        $Azsubs=Get-AzSubscription -ErrorAction Stop
    } catch {
        write-host $_
        exit -1
    }
}

$Azsubs|ForEach-Object{[void] $dropBoxSubscriptions.Items.Add($_.name)}
#$dropBoxSubscriptions.SelectedIndex = 0
$scriptGUI.controls.Add( $dropBoxSubscriptions )

# Display the form
#$scriptGUI.add_Load($form_OnLoadFunction)
[void]$scriptGUI.ShowDialog()

