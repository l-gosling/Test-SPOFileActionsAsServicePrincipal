<#
.SYNOPSIS
    Retrieves the authorization header for Microsoft Graph API using service principal credentials.

.DESCRIPTION
    This function retrieves the authorization header required for accessing Microsoft Graph API.
    It uses the service principal credentials (tenant ID, app ID, and client secret) to obtain an access token from Azure AD.
    The access token is then formatted into an authorization header.

.PARAMETER TenantId
    The tenant ID of the Azure AD directory.

.PARAMETER AppId
    The application ID (client ID) of the service principal.

.PARAMETER ClientSecret
    The client secret of the service principal.

.OUTPUTS
    A hashtable containing the authorization header with the access token.

.EXAMPLE
    $header = Get-HeaderForMGGraphAuth -TenantId "your-tenant-id" -AppId "your-app-id" -ClientSecret "your-client-secret"
    Retrieves the authorization header for Microsoft Graph API using the specified service principal credentials.

.NOTES
    Author      : Lukas Gosling
    CHANGELOG   : 09.01.2025 Function creation
#>
function Get-HeaderForMGGraphAuth {
    param (
        [Parameter(mandatory=$true)]
        [string]$TenantId,

        [Parameter(mandatory=$true)]
        [string]$AppId,

        [Parameter(mandatory=$true)]
        [string]$ClientSecret
    )
    #region get access token

    $scopes =  "https://graph.microsoft.com/.default"
    $loginURL = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    Remove-Variable body -ErrorAction SilentlyContinue
    $body = @{grant_type="client_credentials";client_id=$appId;client_secret=$clientSecret;scope=$scopes}

    $Token = Invoke-RestMethod -Method Post -Uri $loginURL -Body $body
    #$Token.access_token  #expires after one hour
    $headerParams  = @{'Authorization'="$($Token.token_type) $($Token.access_token)"}

    return $headerParams

    #endregion get access token
}