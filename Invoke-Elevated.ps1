<#
.SYNOPSIS
Meant to be used as a target of the 'Send to' context menu.
Opens the specified Path using the chosen Shell Open/Edit verb with elevated credentials.

.DESCRIPTION
Run script with the -Install option, then use via the 'Send to' file explorer context menu.

More details:

Using -Install will:
Create a shortcut in $env:APPDATA\Microsoft\Windows\SendTo with the name "InvokeElevated" and the following target:
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe <path>\Invoke-Elevated.ps1 -Path

Usage:
Right-click a file from File Explorer, select Send To > InvokeElevated.

.PARAMETER Path
Path to file to be opened.

.PARAMETER ElevatedPath
Path to file to be opened under elevated credentials.

.PARAMETER Install
Switch to install this script as target of Send To explorer context menu.
#>

param(
    [Parameter(Mandatory=$true, ParameterSetName='install')]
    [switch]
    $Install,

    [Parameter(Mandatory=$true, ParameterSetName='sendto')]
    [string]
    $Path,

    [Parameter(Mandatory=$true, ParameterSetName='elevated')]
    [string]
    $ElevatedPath
)

#region Internal Functions
function Wait-UserInput 
{
    [Diagnostics.CodeAnalysis.SuppressMessage("PSAvoidUsingWriteHost", "", Scope="Function")]
    param
    (
        [string]
        $Message = "Press any key to continue..."
    )

    if ((Test-Path variable:psISE) -and $psISE) {
        $shell = New-Object -ComObject "WScript.Shell"
        $null = $Shell.Popup("Click OK to continue...", 0, "Script Paused", 0)
    }
    else {     
        Write-Host -NoNewline $Message
        [void][System.Console]::ReadKey($true)
        Write-Host
    }
}

function New-Shortcut
{
    param ( 
        [Parameter(Mandatory=$true)]
        [string]
        $Path, 

        [string]
        $Arguments, 

        [Parameter(Mandatory=$true)]
        [string]
        $Destination
    )

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($Destination)
    $shortcut.TargetPath = $Path
    $shortcut.Arguments = $Arguments
    $shortcut.Save()
}
#endregion Internal Functions

#region MAIN SCRIPT
# Using this $root makes debugging from ISE much nicer.
if ($PSScriptRoot -eq "")
{
    $root = Split-Path -Parent $psISE.CurrentFile.FullPath
}
else
{
    $root = $PSScriptRoot
}

if ($Install)
{
    $target = "%windir%\System32\WindowsPowerShell\v1.0\powershell.exe"
    $arguments = "-WindowStyle Hidden -File `"$($PSCommandPath)`" -Path"
    $destination = Join-Path "$($env:APPDATA)\Microsoft\Windows\SendTo" "InvokeElevated.lnk"
    New-Shortcut -Path $target -Arguments $arguments -Destination $destination
}
elseif ($Path)
{
    $args = "-File `"$PSCommandPath`" -ElevatedPath `"$Path`""
    Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $args
}
else
{
    $app = New-Object -ComObject Shell.Application
    $folder = $app.NameSpace($(Split-Path $ElevatedPath -Parent))
    $file = $folder.ParseName($(Split-Path $ElevatedPath -Leaf))
    $verbs = $file.Verbs() | ? {$_.Name -imatch 'open' -or $_.Name -imatch 'edit'}
    $choiceVerbs = $verbs | % {$_.Name -replace '&',''}
    #$choiceVerbs
    $index = 0
    foreach ($choice in $choiceVerbs)
    {
        # add accelerator key
        $choiceVerbs[$index] = "(&$index) " + $choice
        $index++
    }
    #$choiceVerbs
    $options = [System.Management.Automation.Host.ChoiceDescription[]]$choiceVerbs
    $result = $Host.UI.PromptForChoice("Choose a command", $ElevatedPath, $options, 0)
    $verbs[$result].DoIt()
}

# for debugging...
#Wait-UserInput

#endregion MAIN SCRIPT
