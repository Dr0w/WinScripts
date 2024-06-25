# Azure AD CSV Export
Prerequisites:

## Before you begin, ensure you have the following installed and configured:
- Microsoft Graph PowerShell Module
- PowerShell 7 (or later)
- .NET Framework 4.7.2 (or later)

## Upgrade to PowerShell 5.1 or Later
- Ensure your PowerShell version is 5.1 or later. You can check your PowerShell version by running:
$PSVersionTable.PSVersion

## Powershell distribution site:
- Installing PowerShell on Windows
https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4#installing-the-msi-package

## You can download and install the .NET Framework from the official Microsoft website:
- Download .NET Framework 4.7.2
https://dotnet.microsoft.com/en-us/download/dotnet-framework/net472

## Install Microsoft Graph PowerShell Module
- Follow the instructions here:
https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0

## Steps to Run the Script
- Open PowerShell as Administrator.
- Install the Microsoft.Graph Module:
```
Install-Module -Name Microsoft.Graph -Scope CurrentUser
```
- Copy the script to directory you prefer.
- Navigate to the directory where the script is saved:
```
cd "C:\path\to\your\script\directory"
```
- Execute the script:
```
.\Import_AzureAD.ps1
```
- In case you will receive a message like “.ps1 is not digitally signed. The script will not execute on the system.” , run:
```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```
>This command sets the execution policy to bypass for only the current PowerShell session after the window is closed, the next PowerShell session will open running with the default execution policy. “Bypass” means nothing is blocked and no warnings, prompts, or messages will be displayed.