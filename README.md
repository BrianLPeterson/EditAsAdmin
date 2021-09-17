# EditAsAdmin
Simple powershell script to Right click a file and open it as Administrator.   Useful for editing files such as the host file or a powershell script that needs admin privilege. 

## Description

Meant to be used as a target of the 'Send to' context menu of a file
Opens the specified Path using the chosen Shell Open/Edit verb with elevated credentials providing the UAC prompt

## Installation
Run script with the -Install option, then use via the 'Send to' file explorer context menu.

Using -Install will:
Create a shortcut in $env:APPDATA\Microsoft\Windows\SendTo with the name "InvokeElevated" and the following target:
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe <path>\Invoke-Elevated.ps1 -Path

## Usage
Right-click a file from File Explorer, select Send To > InvokeElevated.
