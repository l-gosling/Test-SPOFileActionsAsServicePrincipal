<#
.SYNOPSIS
    Retrieves the drive IDs for a specified SharePoint Online site.

.DESCRIPTION
    This script retrieves the drive IDs for a specified SharePoint Online site using Microsoft Graph API.
    It first sets the script path as the current location, imports necessary functions, and gets the configuration from an XML file.
    Then, it obtains an access token using the service principal credentials and retrieves the drive IDs.

.EXAMPLE
    .\01_Get-SPOSiteDriveIDs.ps1
    Retrieves the drive IDs for the specified SharePoint Online site as configured in the _scriptConfig.xml file.

.NOTES
    Author      : Lukas Gosling
    CHANGELOG   : 09.01.2025 Script creation
#>

#get script path and set it as current location
[string]$ScriptPath = Switch ($host.name){
    'Visual Studio Code Host' { Split-Path $PsEditor.GetEditorContext().CurrentFile.Path;$TempPathHost = $PsEditor.GetEditorContext().CurrentFile.Path.split("\");[string]$ScriptName=$TempPathHost[($TempPathHost.Length -1)] }
    'Windows PowerShell ISE Host' {  Split-Path -Path $psISE.CurrentFile.FullPath;[string]$ScriptName=$psISE.CurrentFile.DisplayName }
    'ConsoleHost' { $PsScriptRoot;[string]$ScriptName=$MyInvocation.MyCommand.Name}
}
Set-Location $ScriptPath

#import functions
foreach($PSScriptFile in (Get-ChildItem -Path .\functions -Filter "*.ps1"))
{
    Invoke-Expression -Command ". '$($PSScriptFile.FullName)'"
}

#get config from xml
[xml]$scriptConfig = Get-Content -Path .\_scriptConfig.xml

#get access token
$header = Get-HeaderForMGGraphAuth -TenantId $scriptConfig.Configuration.ServicePrincipal.TenantId -AppId $scriptConfig.Configuration.ServicePrincipal.AppId -ClientSecret $scriptConfig.Configuration.ServicePrincipal.ClientSecret

#clear screen for better readability
Clear-Host
Start-Sleep -Seconds 1

#get drive id
((Invoke-WebRequest -Method Get -Headers $header  -Uri "https://graph.microsoft.com/v1.0/sites/$($scriptConfig.Configuration.SPO.SiteId)/drives").Content | ConvertFrom-Json).Value | Select-Object Name,Id,webUrl,driveType
