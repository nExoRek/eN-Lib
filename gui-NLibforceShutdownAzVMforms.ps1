
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Connect-AzAccount
try {
    Get-AzContext
} catch {
    write-host 'must connect using connect-azaccount before running the script'
    exit -13
}
#Select-AzSubscription -SubscriptionId f2ec3daa-2b54-40a7-9e9f-4e85a449b412
#$labelStatus.Text=$azSubs -join ';'

$scriptGUI                  = New-Object system.Windows.Forms.Form
$scriptGUI.ClientSize       = '500,100'
$scriptGUI.text             = "Az Force Shutdown"
$scriptGUI.BackColor        = "#ffffff"
$scriptGUI.AutoSize         = $true
$scriptGUI.TopMost          = $true

$labelStatusLabel                = New-Object System.Windows.Forms.Label
$labelStatusLabel.text           = ""
$labelStatusLabel.AutoSize       = $false
$labelStatusLabel.width          = 350
$labelStatusLabel.height         = 30
$labelStatusLabel.location       = New-Object System.Drawing.Point(10,40)
$scriptGUI.controls.Add($labelStatusLabel)


$dropBoxSubscriptions             = New-Object system.windows.forms.ComboBox
$dropBoxSubscriptions.Text        = ""
$dropBoxSubscriptions.Width       = 170
$dropBoxSubscriptions.Location    = New-Object System.Drawing.Point(5,10)
$dropBoxSubscriptions.AutoSize    = $true
$Azsubs=Get-AzSubscription
[void] $dropBoxSubscriptions.Items.Add("chose sub")
$Azsubs|ForEach-Object{[void] $dropBoxSubscriptions.Items.Add($_.name)}
$dropBoxSubscriptions.SelectedIndex = 0
$scriptGUI.controls.Add( $dropBoxSubscriptions )

$dropBoxSubscriptions_SelectedIndexChanged= {
    [void] $dropBoxSubscriptions.Items.Remove("")
    $labelStatusLabel.text = "changing context...."

    set-azContext -SubscriptionName $dropBoxSubscriptions.SelectedItem|out-null

    $labelStatusLabel.text = "reading Resource Groups...."

    $RGs=Get-AzResourceGroup
    [void] $dropBoxRGs.Items.Clear()
    [void] $dropBoxRGs.Items.Add("")
    $RGs|ForEach-Object{[void] $dropBoxRGs.Items.Add($_.ResourceGroupName)} 
    #$dropBoxRGs.SelectedIndex = 0
    
    $scriptGUI.controls.Add( $dropBoxRGs )

    $labelStatusLabel.text = "found $($RGs.count) RGs."

}
$dropBoxSubscriptions.add_SelectedIndexChanged($dropBoxSubscriptions_SelectedIndexChanged)


$dropBoxRGs             = New-Object system.windows.forms.ComboBox
$dropBoxRGs.Text        = ""
$dropBoxRGs.Width       = 170
$dropBoxRGs.Location    = New-Object System.Drawing.Point(180,10)
$dropBoxRGs.AutoSize    = $true

$dropBoxRGs_SelectedIndexChanged={
    [void] $dropBoxRGs.Items.Remove("")

    $script:RGName=$dropBoxRGs.SelectedItem

    $labelStatusLabel.text = "reading VMs...."

    $VMs=Get-AzVM -ResourceGroupName $RGName -status
    [void] $dropBoxVMs.Items.Clear()
    [void] $dropBoxVMs.Items.Add("")
    $VMs|ForEach-Object{[void] $dropBoxVMs.Items.Add($_.Name)} 
    #$dropBoxVMs.SelectedIndex = 0

    $scriptGUI.controls.Add( $dropBoxVMs )

    $labelStatusLabel.text = "found $($VMs.count) VMs"
  
}
$dropBoxRGs.add_SelectedIndexChanged($dropBoxRGs_SelectedIndexChanged)

$dropBoxVMs             = New-Object system.windows.forms.ComboBox
$dropBoxVMs.Text        = ""
$dropBoxVMs.Width       = 170
$dropBoxVMs.Location    = New-Object System.Drawing.Point(360,10)
$dropBoxVMs.AutoSize    = $true
$dropBoxVMs_SelectedIndexChanged={
    [void] $dropBoxVMs.Items.Remove("")

    $script:VMName=$dropBoxVMs.SelectedItem

    $scriptGUI.controls.Add( $okButton )
  
}
$dropBoxVMs.add_SelectedIndexChanged($dropBoxVMs_SelectedIndexChanged)

$okButton               = New-Object System.Windows.Forms.Button
$okButton.Name          = "okButton"
$okButton.Size          = New-Object System.Drawing.Size(100,29)
#$okButton.UseVisualStyleBackColor = $True
$okButton.Text          = "SHUTDOWN"
$okButton.Location      = New-Object System.Drawing.Point(390,40)
$okButton.DataBindings.DefaultDataSourceUpdateMode = 0
$okButton.add_Click({
    Switch( [System.Windows.Forms.MessageBox]::show(
        $this,"Are you sure you want to force shutdown ""$VMName"" in RG ""$RGName""?",'CONFIRM','YesNo')
    ) {
        'Yes' {
            $labelStatusLabel.text = "force shutdown in progress..."
            stop-AzVM -ResourceGroupName $RGName -name $VMName -force
            $labelStatusLabel.text = "KABOOM!"
        } 
        'No' {
            $labelStatusLabel.text = "canceled"
        }
    }
})


$RGName=""
$VMName=""

# Display the form
#$scriptGUI.add_Load($form_OnLoadFunction)
[void]$scriptGUI.ShowDialog()

