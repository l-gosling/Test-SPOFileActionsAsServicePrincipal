# get script path and set it as current location
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

#get access token
$header = Get-HeaderForMGGraphAuth -TenantId $scriptConfig.Configuration.ServicePrincipal.TenantId -AppId $scriptConfig.Configuration.ServicePrincipal.AppId -ClientSecret $scriptConfig.Configuration.ServicePrincipal.ClientSecret

#clear screen for better readability
Clear-Host
Start-Sleep -Seconds 1

#define uri
$uri = "https://graph.microsoft.com/v1.0/drives/b!BR7-jqcqHEapO6dHOAuaNW30A0wNaI5MvhlrVPv1atDC9mXxmRuSS6Kr7mf3CyLP/root:/Export:/children"

#call ali
$respone = Invoke-WebRequest -Method Get -Headers $header -Uri $uri -ContentType 'multipart/form-data'

#output
$respone




