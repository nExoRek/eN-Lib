<#
.SYNOPSIS
    simple wizard-script to kill idle timer. 
.DESCRIPTION
    if you have enough of your Windows GPO locking you screen too early - you need an Idle-killer.

    you can choose between mouse-move or key-press emulation. beware, that while timer is running and you chose 'key'
    it presses 'alt+\' which may interfere with what you do. mouse on the otherhand may be little annoying but it moving
    single pixel only. neithertheway the purpose is to be run when you're not at the screen.

    timer is set to 5sec by default and you can raise it up to 5min. 

.INPUTS
    None.
.OUTPUTS
    None.
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 220203
        last changes
        - 220203 initialized

    #TO|DO
    - hide from taskbar... but that's hard one.
#>
[CmdletBinding()]
param ()

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region GENERAL_OBJECTS
$script:STATUS=$true #+/- for the mouse
$script:INITIALIZE = $true

$timer=New-Object System.Windows.Forms.Timer
$timer.Interval = 5000 #every five seconds by default
$timer.add_Tick({
    if($chbMouseKey.TextAlign -eq "MiddleLeft") {   #mouse chosen
        $mousePosition = [Windows.Forms.Cursor]::Position
        if($script:STATUS) {
            $mousePosition.x++
            $mousePosition.y++
            $script:STATUS=$false
        } else {
            $mousePosition.x--
            $mousePosition.y--
            $script:STATUS=$true
        }
        [Windows.Forms.Cursor]::Position = $mousePosition
    } else {
        [System.Windows.Forms.SendKeys]::SendWait("%\") #keypress emu chosen
    }
    $timer.Interval = [int]$nudNumber.text * 1000
    Write-Verbose ("{0} {1} {2}" -f $timer.Interval, $script:STATUS, $nudNumber.Text)
})

$sysTray = New-Object System.Windows.Forms.NotifyIcon
$sysTray.Icon = [System.Drawing.SystemIcons]::Information
$sysTray.Text = "m0veR"
$sysTray.Visible = $false
$sysTray.add_mouseClick({
    #param($sender, $e)
    #$mainForm.Show();
    $mainForm.WindowState = 'Normal'
})

#endregion GENERAL_OBJECTS

#region MAINFORM
$mainForm = New-Object system.Windows.Forms.Form
$mainForm.AutoSize = $true
$mainForm.MaximumSize = New-Object System.Drawing.Size(240,100)
$mainForm.text = "m0veR"
#$mainForm.StartPosition = 'CenterScreen'
$mainForm.Location = New-Object System.Drawing.Point(100,100) 
$mainForm.Icon = [System.Drawing.SystemIcons]::Information
$mainForm.TopMost = $true
$mainForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8)
$mainForm.MaximizeBox = $false
$mainForm.MinimizeBox = $true
$mainForm.Dock = 'fill'

$layout = new-object system.windows.forms.tableLayoutPanel
#$layout.AutoSize = $false
$layout.ColumnCount = 5
$layout.RowCount = 2
[void]$layout.rowstyles.add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute), 30 ) )
[void]$layout.rowstyles.add( (new-object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute), 30 ) )
$layout.Dock = "fill"
$layout.BorderStyle = 'none' 

#region ROWnr1
$lblMouse = New-Object System.Windows.Forms.Label 
$lblMouse.Size = New-Object System.Drawing.Size(40,15)
$lblMouse.Text = "Mouse"
$lblMouse.Anchor = 'none'
$layout.Controls.Add($lblMouse,0,0)

$chbMouseKey = New-Object System.Windows.Forms.CheckBox
$chbMouseKey.Size = New-Object System.Drawing.Size(35,20)
$chbMouseKey.Appearance = 'Button'
$chbMouseKey.Text = "@"
$chbMouseKey.TextAlign = "MiddleLeft"
$chbMouseKey.Anchor = 'none'
$layout.Controls.Add($chbMouseKey,1,0)
$chbMouseKey.add_CheckedChanged({
    if($chbMouseKey.TextAlign -eq "MiddleLeft") {
        $chbMouseKey.TextAlign = "MiddleRight"
    } else {
        $chbMouseKey.TextAlign = "MiddleLeft"
    }
})

$lblKey = New-Object System.Windows.Forms.Label 
#$lblKey.Location = New-Object System.Drawing.Point(75,8) 
$lblKey.Size = New-Object System.Drawing.Size(30,15)
$lblKey.Text = "Key"
$lblKey.Anchor = 'none'
$layout.Controls.Add($lblKey,2,0)

$lblInterval = New-Object System.Windows.Forms.Label 
$lblInterval.Size = New-Object System.Drawing.Size(20,15)
$lblInterval.Text = "Int"
$lblInterval.Anchor = 'none'
$layout.Controls.Add($lblInterval,3,0)

$nudNumber = New-Object System.Windows.Forms.NumericUpDown
#$nudNumber.Location = New-Object System.Drawing.Point(125,5) 
$nudNumber.Size = New-Object System.Drawing.Size(50,15)
$nudNumber.Text = "5"
$nudNumber.Minimum = 1
$nudNumber.Maximum = 300
$nudNumber.Anchor = 'none'
$layout.Controls.Add($nudNumber,4,0)
$nudNumber.add_ValueChanged({
    write-verbose ([int]$nudNumber.text+1)
    #for some reason it doesn't work as expeted - as the results where skipped/not reflecetd
    #$timer.Interval = ([int]$nudNumber.text + 1) * 1000
})
#endregion ROWnr1

$btStartStop = New-Object System.Windows.Forms.Button
$btStartStop.Size = New-Object System.Drawing.Size(80,20)
$btStartStop.Text = "START"
$btStartStop.Anchor = 'none'
$btStartStop.add_Click({
    if($btStartStop.text -eq 'START') {
        $btStartStop.text = "STOP"
        $timer.Start()
    } else {    
        $btStartStop.text = "START"
        $timer.Stop()
    }
})
$layout.Controls.Add($btStartStop,0,1)
$layout.SetColumnSpan($btStartStop,5)

$mainForm.Controls.Add($layout)

$mainForm.add_shown({
})
$mainForm.add_Resize({
    if($mainForm.WindowState -eq 'Minimized') {
        $sysTray.Visible = $true
    } else {
        $sysTray.Visible = $false
    }
})
$mainForm.add_Closing({
    param($sender,$e)
    $timer.dispose()
    $sysTray.dispose()
    $mainForm.dispose()
    [System.GC]::Collect()
})
#endregion MAINFORM

$result = $mainForm.ShowDialog()
