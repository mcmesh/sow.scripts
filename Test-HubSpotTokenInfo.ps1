# ══════════════════════════════════════════════════════════════════════════════
#  Test-HubSpotTokenInfo.ps1
#
#  Calls the HubSpot Private App token-info endpoint to display the Hub ID,
#  App ID, and every scope currently granted to your access token.
#
#  API reference:
#    POST https://api.hubapi.com/oauth/v2/private-apps/get/access-token-info
#
#  Usage:
#    $env:HUBSPOT_API_KEY = "pat-na1-..."
#    .\Test-HubSpotTokenInfo.ps1
# ══════════════════════════════════════════════════════════════════════════════

$ApiKey = $env:HUBSPOT_API_KEY

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Host "ERROR: HUBSPOT_API_KEY environment variable is not set." -ForegroundColor Red
    Write-Host '       $env:HUBSPOT_API_KEY = "pat-na1-..."' -ForegroundColor Yellow
    exit 1
}

$Uri  = "https://api.hubapi.com/oauth/v2/private-apps/get/access-token-info"
$Body = @{ tokenKey = $ApiKey } | ConvertTo-Json

Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  HubSpot Token Info"                      -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Endpoint : $Uri"
Write-Host ""

try {
    $response = Invoke-RestMethod `
        -Uri         $Uri `
        -Method      Post `
        -Body        $Body `
        -ContentType "application/json"

    Write-Host "  Hub ID   : $($response.hubId)"   -ForegroundColor Green
    Write-Host "  App ID   : $($response.appId)"   -ForegroundColor Green
    Write-Host "  User ID  : $($response.userId)"  -ForegroundColor Green
    Write-Host ""
    Write-Host "  Scopes ($($response.scopes.Count)):" -ForegroundColor Cyan
    foreach ($scope in ($response.scopes | Sort-Object)) {
        Write-Host "    - $scope" -ForegroundColor White
    }

    # Highlight whether the scopes needed by Export-HubSpotEmails.ps1 are present
    Write-Host ""
    Write-Host "  Required scope check for Export-HubSpotEmails.ps1:" -ForegroundColor Cyan
    foreach ($required in @("crm.objects.emails.read", "sales-email-read", "crm.objects.companies.read", "crm.objects.deals.read", "crm.objects.contacts.read")) {
        if ($response.scopes -contains $required) {
            Write-Host "    [PRESENT] $required" -ForegroundColor Green
        } else {
            Write-Host "    [MISSING] $required" -ForegroundColor Red
        }
    }
}
catch {
    $statusCode = $null
    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
    }
    Write-Host "ERROR: Request failed (HTTP $statusCode)" -ForegroundColor Red
    if (-not [string]::IsNullOrWhiteSpace($_.ErrorDetails.Message)) {
        Write-Host "HubSpot error detail: $($_.ErrorDetails.Message)" -ForegroundColor Red
    } else {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    exit 1
}

Write-Host ""
Write-Host "══════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
