#Requires -Version 5.1
<#
.SYNOPSIS
	Builds OutlookPushoverAlerter in Release and registers it as an Outlook add-in.
.DESCRIPTION
	Compiles the project, then writes the four registry values Outlook needs under
	HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookPushoverAlerter.
	Close Outlook before running; restart it afterwards.
#>

$ErrorActionPreference = 'Stop'
$projectDir = $PSScriptRoot

# -- 1. Locate MSBuild --------------------------------------------------------
$msb = Get-ChildItem 'C:\Program Files\Microsoft Visual Studio' `
		-Recurse -Filter 'MSBuild.exe' -ErrorAction SilentlyContinue |
	   Where-Object { $_.FullName -notmatch 'amd64' } |
	   Select-Object -First 1 -ExpandProperty FullName

if (-not $msb) { throw 'MSBuild.exe not found. Is Visual Studio installed?' }

# -- 2. Build Release ---------------------------------------------------------
Write-Host 'Building Release...' -ForegroundColor Cyan
& $msb "$projectDir\OutlookPushoverAlerter.sln" /t:Build /p:Configuration=Release /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }

# -- 3. Verify the VSTO manifest exists ---------------------------------------
$vstoPath = "$projectDir\bin\Release\OutlookPushoverAlerter.vsto"
if (-not (Test-Path $vstoPath)) {
	throw "VSTO manifest not found at: $vstoPath"
}

# -- 4. Write registry entries ------------------------------------------------
$regKey = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookPushoverAlerter'
New-Item -Path $regKey -Force | Out-Null
Set-ItemProperty -Path $regKey -Name 'Description'  -Value 'Sends a Pushover notification when watched senders email you'
Set-ItemProperty -Path $regKey -Name 'FriendlyName' -Value 'Outlook Pushover Alerter'
Set-ItemProperty -Path $regKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $regKey -Name 'Manifest'     -Value "$vstoPath|vstolocal"

Write-Host "Done. Start Outlook - the 'Pushover Alerts' tab should appear." -ForegroundColor Green
