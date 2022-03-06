<#
.SYNOPSIS
    multi-level text menu example. class + display
.EXAMPLE
    dot source the file to use the menu

    . .\menu.ps1
    
.LINK
    https://w-files.pl
.NOTES
    nExoR ::))o-
    version 220306
        last changes
        - 220306 initialized

    #TO|DO
#>

class MenuLevel {
    [string]$menuPrompt
    [string]$name
    [MenuLevel]$previousLevel
    [MenuLevel[]]$nextLevel
    MenuLevel() {
    }
    MenuLevel(
        [string]$title
    ) {
        $this.menuPrompt = $title
        $this.nextLevel += [MenuLevel]@{ name = 'exit' }
    }
    #add additional menu level with subitems
    [MenuLevel[]] addMenuLevel([string]$name,[string]$prompt) {
        $nLevel = [MenuLevel]::new()
        $nLevel.name = $name
        $nLevel.menuPrompt = $prompt
        $nLevel.previousLevel += $this
        $this.nextLevel += $nLevel
        $back = [MenuLevel]::new()
        $back.name = 'back'
        $nLevel.nextLevel += $back
        return $nLevel
    }    
    #add leaf - for actual choice and execution
    [void] addLeafItem([string]$name) {
        $leaf = [MenuLevel]::new()
        $leaf.name = $name
        $leaf.previousLevel += $this
        $this.nextLevel += $leaf
    }
    #print menu items from current level    
    [string[]] getMenuItems(){
        if($this.nextLevel) {
            return $this.nextLevel.name
        } else {
            return $null
        }
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
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        #List of menu items to display
        [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [MenuLevel[]]$MenuItems
    )
    
    function Show-Menu {
        param (
            [int]$selectedItemIndex = 0,
            [validateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
                [string]$foregroundColor = "DarkGreen",
            [validateSet("Black","DarkBlue","DarkGreen","DarkCyan","DarkRed","DarkMagenta","DarkYellow","Gray","DarkGray","Blue","Green","Cyan","Red","Magenta","Yellow","White")]
                [string]$backgroundColor = "Gray"
        )

        #position the cursor
        $maxLineLength = 40 
        $winSize = $host.ui.RawUI.WindowSize
        Clear-Host
        $x = ($winSize.width / 2) - 20
        $y = 10
        $Host.UI.RawUI.CursorPosition = [PSCustomObject]@{X=$x;Y=$y}
        #print menu title
        $nrOfSpaces = ($maxLineLength - $MenuItems.menuPrompt.Length) / 2
        Write-Host (" "*$nrOfSpaces + $MenuItems.menuPrompt) -ForegroundColor Green
        $y+=2

        #DRAW ACTUAL MENU ITEMS
        for ($item = 0; $item -lt $MenuItems.nextLevel.Count; $item++) {
            $nrOfSpaces = ($maxLineLength - $MenuItems.nextLevel[$item].name.Length) / 2
            $Host.UI.RawUI.CursorPosition = [PSCustomObject]@{X=$x;Y=$y+$item}
            $itemText = (" "*$nrOfSpaces) +  $MenuItems.nextLevel[$item].name + (" " * $nrOfSpaces)
            if ($selectedItemIndex -eq $item) {
                Write-Host $itemText -ForegroundColor $foregroundColor -BackgroundColor $backgroundColor
            } else {
                Write-Host $itemText
            }
        }
    }

    #show the menu
    $key = $null
    $itemNumber = 0
    while ($key -ne 13) {
        Show-Menu -selectedItemIndex $itemNumber
        $press = $host.ui.rawui.readkey("NoEcho,IncludeKeyDown")
        $key = $press.virtualkeycode
        if ($key -eq 38 -and $itemNumber -gt 0) {
            $itemNumber--
        }
        if ($key -eq 40 -and $itemNumber -lt $MenuItems.nextLevel.count) {
            $itemNumber++
        }        
    }
    #act on return
    if($MenuItems.nextLevel[$itemNumber].name -eq 'exit') { #EXIT
        break
    } elseif($MenuItems.nextLevel[$itemNumber].name -eq 'back') { #BACK
        Get-MenuSelection -MenuItems $MenuItems.previousLevel
    } elseif($MenuItems.nextLevel[$itemNumber].nextLevel) { #SHOW NEXT LEVEL
        Get-MenuSelection -MenuItems $MenuItems.nextLevel[$itemNumber]
    } else { #ACTUAL CHOICE - LEAF 
        return $MenuItems.nextLevel[$itemNumber].name
    }
}

