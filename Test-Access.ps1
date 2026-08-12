# Store your client secret in an environment variable before running, e.g.:
#   $env:GRAPH_CLIENT_SECRET = "your_secret_here"
$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf"
$ClientId     = "5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b"
$ClientSecret = $env:GRAPH_CLIENT_SECRET
$Resource     = "https://meshsystemsus.sharepoint.com"

if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "ERROR: Set the environment variable GRAPH_CLIENT_SECRET before running." -ForegroundColor Red
    exit 1
}

# Get token via client credentials flow
$tokenUrl  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "$Resource/.default"
}

$tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $tokenBody
$token = $tokenResponse.access_token

Write-Host "Token acquired!" -ForegroundColor Green

# Decode and verify audience
$parts   = $token.Split('.')
$payload = [System.Text.Encoding]::UTF8.GetString(
    [Convert]::FromBase64String(
        $parts[1].PadRight($parts[1].Length + (4 - $parts[1].Length % 4) % 4, '=')
    )
) | ConvertFrom-Json

Write-Host "token : $token"
Write-Host "aud : $($payload.aud)"
Write-Host "appid: $($payload.appid)"
Write-Host "tid : $($payload.tid)"


# Use your existing Graph token
$regionCheck = Invoke-RestMethod `
    -Uri "https://graph.microsoft.com/beta/sites?`$filter=siteCollection/root ne null&`$select=webUrl,siteCollection" `
    -Headers $headers `
    -Method Get

$regionCheck.value | ForEach-Object {
    Write-Host "Site: $($_.webUrl)"
    Write-Host "Region: $($_.siteCollection.dataLocationCode)"
}