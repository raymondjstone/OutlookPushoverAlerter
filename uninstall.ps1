#Requires -Version 5.1
<#
.SYNOPSIS
	Removes the OutlookPushoverAlerter add-in registration from this user profile.
.DESCRIPTION
	Deletes the registry key Outlook uses to load the add-in.
	Close Outlook before running; restart it afterwards.
#>

$ErrorActionPreference = 'Stop'

$regKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPushoverAlerter'

if (Test-Path $regKey) {
	Remove-Item -Path $regKey -Recurse -Force
	Write-Host 'Uninstalled. Restart Outlook.' -ForegroundColor Green
} else {
	Write-Host 'OutlookPushoverAlerter is not currently registered.' -ForegroundColor Yellow
}
