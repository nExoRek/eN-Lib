<#
.SYNOPSIS
    multi-level text menu example. class + display
.EXAMPLE
    dot source the file to use the menu

    . .\MultiLevelTextMenu.lib.ps1
    
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 220322
        last changes
        - 220322 'TAG' & 'Description' instead of 'name'
        - 220306 initialized

    #TO|DO
    - create menu from JSON function
    - different color for 'back' and 'exit'
    - re-draw only particular lines - not all (performance)
#>

class MenuLevel {
    [string]$menuPrompt
    [string]$description
    [string]$tag
    [MenuLevel]$previousLevel
    [MenuLevel[]]$nextLevel
    MenuLevel() {
    }
    MenuLevel(
        [string]$title
    ) {
        $this.menuPrompt = $title
        $this.nextLevel += [MenuLevel]@{ 
            tag = 'exit';
            description = 'exit';
        }
    }
    #add additional menu level with subitems
    [MenuLevel[]] addMenuLevel([string]$tag,[string]$prompt) {
        $nLevel = [MenuLevel]::new()
        $nLevel.description = $tag
        $nLevel.tag = $tag
        $nLevel.menuPrompt = $prompt
        $nLevel.previousLevel += $this
        $this.nextLevel += $nLevel
        $back = [MenuLevel]::new()
        $back.description = 'back'
        $back.tag = 'back'
        $nLevel.nextLevel += $back
        return $nLevel
    }    
    #add leaf - for actual choice and execution
    [void] addLeafItem([string]$tag) {
        $leaf = [MenuLevel]::new()
        $leaf.description = $tag
        $leaf.tag = $tag
        $leaf.previousLevel += $this
        $this.nextLevel += $leaf
    }
    [void] addLeafItem([string]$tag,[string]$description) {
        $leaf = [MenuLevel]::new()
        $leaf.description = $description
        $leaf.tag = $tag
        $leaf.previousLevel += $this
        $this.nextLevel += $leaf
    }
    #print menu items from current level    
    [string[]] getMenuItems(){
        if($this.nextLevel) {
            return $this.nextLevel.tag
        } else {
            return $null
        }
    }
    #calculate maximum length of 'description' attribute values in current menu level
    [int] getLength() {
        $length = $this.menuPrompt.Length
        foreach($element in $this.nextLevel) {
            if( $element.description.Length -gt $length) { 
                $length = $element.description.Length
            }
        }
        return $length
    }
}

function Get-MenuSelection {
    <#
    .SYNOPSIS
        multi-level text-based menu.
    
    .DESCRIPTION
        alows to easily create multi-level (tree) text-based menu. menu allows to back to upper level, navigate down 
        or exit from main menu.
        menu itself is implemented with special object type of [MenuLevel] class. this class implements automatic 
        'exit' and 'back' options with proper constructor and methods usage.

        how to use the menu
        
        1.create menu object:
            $mainMenu = [MenuLevel]::new('select option') 
            constructor with single string value will create a new object, text will be menu title. it will automatically
            add 'exit' option
        2.leaf item is a choisable item that will quit the menu, returning the text value. to add leaf item:
            $mainMenu.addLeafItem('some choice')
        3.in order to add additional menu level:
            $nextLevel = $mainMenu.addMenuLevel('level2', 'level 2 title')
            this method returns a pointer to the next level menu and automatically adds 'back' option. it can now
            be easily populated with additional leaf items or levels:
            $nextLevel.addLeafItem('level 2 choice')
            $thirdLevel=$nextLevel.addMenuLevel('level3', 'level 3 title')
        4.now menu-as-object is ready, you can run the menu:
            $choice = Get-MenuSelection $mainMenu
            switch($choice) {
                [...]
            }
    .EXAMPLE
    #run below code to understand how to create menu:
        $mainMenu = [MenuLevel]::new('select option')
        $mainMenu.addLeafItem('terminating option') #terminating option
        $l2_1 = $mainMenu.addMenuLevel('submenu options 1', 'SUBMENU L2') #adding 2nd level menu
        $l3_1 = $l2_1.addMenuLevel('submenu option 2', 'SUBMENU L3') #adding next submenu - 3rd level
        $l2_1.addLeafItem('terminating option 2_1') #add terminating option to L2 menu
        $l2_1.addLeafItem('terminating option 2_2') #add terminating option to L2 menu

        $l3_1.addLeafItem('terminate op 3_1') 
        $l3_1.addLeafItem('terminate op 3_2') 
        $choice = Get-MenuSelection $mainMenu
        switch($choice) {
            'terminating option' { write-host "some logic there for option from main menu" } 
            'terminating option 2_1' { write-host "some logic there for option 1 from L2 menu" }
            'terminating option 2_2' { write-host "some logic there for option 2 from L2 menu"}
            'terminate op 3_1' { write-host "some logic there for option 1 from L3 menu" }
            'terminate op 3_2' { write-host "some logic there for option 2 from L3 menu" }
            default { write-host -ForegroundColor red "UNKNOWN OPTION" }
        }        
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        #List of menu items to display
        [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [MenuLevel]$MenuItems
    )
    function get-MenuPosition {
        param( [MenuLevel]$currentLevel )

        Clear-Host
        #position the cursor
        $winSize = $host.ui.RawUI.WindowSize
        $menuMinLength = $currentLevel.getLength()
        if(($winSize.Width - $menuMinLength) -ge 10 ) { #common scenario - calculate for '     <item.length>     '
            $menuItemLength = $menuMinLength + 10
        } elseif( ($winSize.Width - $menuMinLength) -lt 1 ) { #menu item length wider then screen size
            $menuItemLength = $winSize.Width
        } else { #there is less then 10 chars difference between longest element and screen size
            $menuItemLength = $menuMinLength + $nrOfSpaces
        }
        return [PSCustomObject]@{
            X = [int]( ($winSize.Width - $menuItemLength) / 2 )
            Y = 10
            menuItemLength = $menuItemLength
        }
    }
    function Show-Menu {
        param (
            [MenuLevel]$currentLevel,
            [int]$selectedItemIndex = 0,
            [PSCustomObject]$menuPosition,
            [validateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
                [string]$foregroundColor = "DarkGreen",
            [validateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
                [string]$backgroundColor = "Gray"
        )

        #set up position values
        $winWidth = $host.ui.RawUI.WindowSize.Width
        $Host.UI.RawUI.CursorPosition = [PSCustomObject]@{X=$menuPosition.X;Y=$menuPosition.Y}
        $nrOfSpaces = [int]( ($menuPosition.menuItemLength - $currentLevel.menuPrompt.Length)/2 )
        #print menu title
        Write-Host (" "*$nrOfSpaces + $currentLevel.menuPrompt) -ForegroundColor Green
        $y = $menuPosition.Y + 2

        #DRAW ACTUAL MENU ITEMS
        for ($item = 0; $item -lt $currentLevel.nextLevel.Count; $item++) {
            $currentDescription = $currentLevel.nextLevel[$item].description
            if($currentDescription.Length -gt $winWidth) {
                $currentDescription = $currentDescription.substring(0,($winWidth-1))
            }
            $nrOfSpaces = [int]( ($menuPosition.menuItemLength - $currentDescription.Length) / 2 )
            $Host.UI.RawUI.CursorPosition = [PSCustomObject]@{X=$menuPosition.X;Y=$y+$item}
            $itemText = (" "*$nrOfSpaces) +  $currentDescription + (" " * $nrOfSpaces)
            if ($selectedItemIndex -eq $item) {
#                if($currentDescription -eq 'back' -or $currentDescription -eq 'exit') {
    #need to create colors table and convert to [int]
#                    Write-Host $itemText -ForegroundColor (($foregroundColor + 1)%15) -BackgroundColor $backgroundColor
#                } else {
                    Write-Host $itemText -ForegroundColor $foregroundColor -BackgroundColor $backgroundColor
#                }
            } else {
                Write-Host $itemText
            }
        }
    }

    #show the menu
    $key = $null
    $itemNumber = 0
    $menuPosition = get-MenuPosition -currentLevel $MenuItems
    while ($key -ne 13) {
        Show-Menu -currentLevel $MenuItems -selectedItemIndex $itemNumber -menuPosition $menuPosition
        $press = $host.ui.rawui.readkey("NoEcho,IncludeKeyDown")
        $key = $press.virtualkeycode
        if ($key -eq 38 -and $itemNumber -gt 0) { #up arrow
            $itemNumber--
        }
        if ($key -eq 40 -and $itemNumber -lt $MenuItems.nextLevel.count-1) { #down arrow
            $itemNumber++
        }        
        if ($key -eq 8 -and $MenuItems.previousLevel) { #backspace
            $itemNumber = 0 #that assumes that 'back' is always on the first position.
            break   
        }        
    }
    #act on return
    if($MenuItems.nextLevel[$itemNumber].tag -eq 'exit') { #EXIT
        break
    } elseif($MenuItems.nextLevel[$itemNumber].tag -eq 'back') { #BACK
        Get-MenuSelection -MenuItems $MenuItems.previousLevel
    } elseif($MenuItems.nextLevel[$itemNumber].nextLevel) { #SHOW NEXT LEVEL
        Get-MenuSelection -MenuItems $MenuItems.nextLevel[$itemNumber]
    } else { #ACTUAL CHOICE - LEAF 
        return $MenuItems.nextLevel[$itemNumber].tag
    }
}
