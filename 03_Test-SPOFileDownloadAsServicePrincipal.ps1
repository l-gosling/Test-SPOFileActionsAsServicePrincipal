<#
.SYNOPSIS
    Tests downloading a file from SharePoint Online as a service principal.

.DESCRIPTION
    This script tests the ability to download a file from SharePoint Online using a service principal.
    It sets the script path as the current location, imports necessary functions, and retrieves the drive ID and file details from SharePoint Online using Microsoft Graph API.
    If you want to download a other file like the test_download.xlsx, you have to change the path in the variable $pathFileFromDownload.

.EXAMPLE
    .\03_Test-SPOFileDownloadAsServicePrincipal.ps1
    Tests downloading a file from SharePoint Online as a service principal.

.COMPONENT
    PowerShell 7.0 or later

.NOTES
    Author      : Lukas Gosling
    CHANGELOG   : 09.01.2025 Script creation
#>

#region prepare execution

#get script path and set it as current location
[string]$ScriptPath = Switch ($host.name){
    'Visual Studio Code Host' { Split-Path $PsEditor.GetEditorContext().CurrentFile.Path;$TempPathHost = $PsEditor.GetEditorContext().CurrentFile.Path.split("\");[string]$ScriptName=$TempPathHost[($TempPathHost.Length -1)] }
    'Windows PowerShell ISE Host' {  Split-Path -Path $psISE.CurrentFile.FullPath;[string]$ScriptName=$psISE.CurrentFile.DisplayName }
    'ConsoleHost' { $PsScriptRoot;[string]$ScriptName=$MyInvocation.MyCommand.Name}
}
Set-Location $scriptPath

#import functions
foreach($PSScriptFile in (Get-ChildItem -Path .\functions -Filter "*.ps1"))
{
    Invoke-Expression -Command ". '$($PSScriptFile.FullName)'"
}

#get config from xml
[xml]$scriptConfig = Get-Content -Path .\_scriptConfig.xml

#endregion prepare execution
#region declare variables

#file path of downloaded file
$pathFileFromDownload = "$scriptPath\files\test_download.txt"
#$pathFileFromDownload = "$scriptPath\files\test_download.xlsx"

#spo path to file
$downloadPathFolder = [string]$scriptConfig.Configuration.SPO.FilePath
$downloadPathFileName = (Split-Path -Path $pathFileFromDownload -Leaf).Replace("_download","")
$downloadPath = $downloadPathFolder + "/" + $downloadPathFileName

#remove leading slashes if multiple are present
if ($downloadPath.StartsWith("//")) {
    $downloadPath = $downloadPath.Substring(1)
}

#uri to download file
$fileUri = "https://graph.microsoft.com/v1.0/sites/$($scriptConfig.Configuration.SPO.SiteId)/drives/$($scriptConfig.Configuration.SPO.DriveId)/root:$($downloadPath)"

#endregion declare variables
#region download file

#clear screen for better readability
Clear-Host
Start-Sleep -Seconds 1

#delte old file
if (Test-Path $pathFileFromDownload) {
    Write-Output "Delete old file '$($pathFileFromDownload)'"
    Remove-Item $pathFileFromDownload
}

#get access token
$header = Get-HeaderForMGGraphAuth -TenantId $scriptConfig.Configuration.ServicePrincipal.TenantId -AppId $scriptConfig.Configuration.ServicePrincipal.AppId -ClientSecret $scriptConfig.Configuration.ServicePrincipal.ClientSecret

#get download uri
Write-Output "Get Download URI for: $fileUri"
$downloadUrl = ((Invoke-WebRequest -Method Get -Headers $header -Uri $fileUri -ContentType 'multipart/form-data').Content | ConvertFrom-Json).'@microsoft.graph.downloadUrl'

#download file
Write-Output "Download file to '$pathFileFromDownload'"
Invoke-WebRequest -Uri $downloadUrl -OutFile $pathFileFromDownload
#endregion download file
