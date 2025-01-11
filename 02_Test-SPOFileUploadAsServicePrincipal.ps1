<#
.SYNOPSIS
    Tests downloading a file from SharePoint Online as a service principal.

.DESCRIPTION
    This script tests the ability to upload a file to SharePoint Online using a service principal.
    It sets the script path as the current location, imports necessary functions, and retrieves the drive ID and file details from SharePoint Online using Microsoft Graph API.
    The script creates an upload session and uploads the specified file.
    If you want to upload a other file like the test_upload.xlsx, you have to change the path in the variable $pathFileToUpload.

.EXAMPLE
    .\02_Test-SPOFileUploadAsServicePrincipal.ps1
    Tests uploading a file to SharePoint Online as a service principal.

.COMPONENT
    PowerShell 7.0 or later

.NOTES
    You can upload the entire file, or split the file into multiple byte ranges, as 
    long as the maximum bytes in any given request is less than 60 MiB.

    If your app splits a file into multiple byte ranges, 
    the size of each byte range MUST be a multiple of 320 KiB (327,680 bytes). 
    Using a fragment size that does not divide evenly by 320 KiB will result in 
    errors committing some files.

.NOTES
    https://pnp.github.io/script-samples/graph-upload-file-to-sharepoint/README.html?tabs=azure-cli
    https://learn.microsoft.com/graph/api/driveitem-createuploadsession

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

#file path of file to upload
$pathFileToUpload = "$scriptPath\files\test_upload.txt"
#$pathFileToUpload = "$scriptPath\files\test_upload.xlsx"

#spo destination path 
$uploadPathFolder = [string]$scriptConfig.Configuration.SPO.FilePath #"/"
$uploadPathFileName = (Split-Path -Path $pathFileToUpload -Leaf).Replace("_upload","")
$uploadPath = $uploadPathFolder + "/" + $uploadPathFileName

#remove leading slashes if multiple are present
if ($uploadPath.StartsWith("//")) {
    $uploadPath = $uploadPath.Substring(1)
}

#uri to download file
$uploadSessionUri = "https://graph.microsoft.com/v1.0/sites/$($scriptConfig.Configuration.SPO.SiteId)/drives/$($scriptConfig.Configuration.SPO.DriveId)/root:$($uploadPath):/createUploadSession"

#endregion declare variables
#region upload file

#clear screen for better readability
Clear-Host
Start-Sleep -Seconds 1

#get access token
$header = Get-HeaderForMGGraphAuth -TenantId $scriptConfig.Configuration.ServicePrincipal.TenantId -AppId $scriptConfig.Configuration.ServicePrincipal.AppId -ClientSecret $scriptConfig.Configuration.ServicePrincipal.ClientSecret

#create upload session
Write-Output "Uploade file '$($pathFileToUpload)'"
Write-Output "Upload Session URI: $uploadSessionUri"
$uploadSession = Invoke-WebRequest -Method Post -Headers $header -Uri $uploadSessionUri | ConvertFrom-Json

#get local file
Write-Output "Getting local file..."
$fileInBytes = [System.IO.File]::ReadAllBytes($pathFileToUpload)
$fileLength = $fileInBytes.Length

#calc for file upload
$partSizeBytes = 320 * 1024 * 4  #Uploads 1.31MiB at a time.
$index = 0
$start = 0
$end = 0

$maxloops = [Math]::Round([Math]::Ceiling($fileLength / $partSizeBytes))

#file upload loop
while ($fileLength -gt ($end + 1)) {
    $start = $index * $partSizeBytes
    if (($start + $partSizeBytes - 1 ) -lt $fileLength) {
        $end = ($start + $partSizeBytes - 1)
    }
    else {
        $end = ($start + ($fileLength - ($index * $partSizeBytes)) - 1)
    }
    [byte[]]$body = $fileInBytes[$start..$end]
    $headers = @{    
        'Content-Range' = "bytes $start-$end/$fileLength"
    }
    Write-Output "bytes $start-$end/$fileLength | Index: $index and ChunkSize: $partSizeBytes"
    Invoke-WebRequest -Method Put -Uri $uploadSession.uploadUrl -Body $body -Headers $headers -SkipHeaderValidation | Out-Null
    $index++
    Write-Output "Percentage Completed: $([Math]::Ceiling($index/$maxloops*100)) %"
}

#endregion upload file
