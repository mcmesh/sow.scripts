#Requires -Version 7.0
<#
.SYNOPSIS
    Exports all Teams meeting transcripts for a given organizer via Microsoft Graph API.

.DESCRIPTION
    Authenticates using client credentials (app-only), fetches all transcripts for
    the specified organizer, handles pagination, and downloads transcript content.
    Falls back to OneDrive DriveItem API if the transcript content endpoint returns
    3004 (meeting resource expired).

.NOTES
    Required App Registration Permissions (Application, with Admin Consent):
        - OnlineMeetings.Read.All
        - OnlineMeetingTranscript.Read.All
        - User.Read.All
        - Files.Read.All  <-- needed for OneDrive fallback

    Set your client secret as an environment variable before running:
        $env:GRAPH_CLIENT_SECRET = "your-secret-here"

.EXAMPLE
    # List transcripts only
    .\MC-Export-TeamsTranscripts_TEST.ps1

    # Download all transcripts
    .\MC-Export-TeamsTranscripts_TEST.ps1 -DownloadTranscripts

    # Download to a specific folder
    .\MC-Export-TeamsTranscripts_TEST.ps1 -DownloadTranscripts -OutputFolder "C:\Transcripts"
#>

param(
    [string]$OrganizerId  = "27872551-e5ef-430a-a7da-468fd277e124", #"a5280b34-0b66-4a5c-8bbf-bc4c14694b1f", #
    [string]$ClientId     = "9051ff18-d122-4e52-a378-838f688889e8",
    [string]$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf",
    [switch]$DownloadTranscripts,
    [string]$OutputFolder = ".\TranscriptExport"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($env:GRAPH_CLIENT_SECRET)) {
    throw "Set GRAPH_CLIENT_SECRET to the app registration's client secret."
}

# ============================================================
# STEP 1: Authenticate
# ============================================================
Write-Host "`n[1] Authenticating..." -ForegroundColor Cyan

$tokenResponse = Invoke-RestMethod `
    -Method Post `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
        client_id     = $ClientId
        client_secret = $env:GRAPH_CLIENT_SECRET
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

$headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }
Write-Host "  Token acquired" -ForegroundColor Green

# ============================================================
# STEP 2: Fetch all transcript metadata (with pagination)
# ============================================================
Write-Host "`n[2] Fetching transcript list..." -ForegroundColor Cyan

$allTranscripts = [System.Collections.Generic.List[object]]::new()
$uri = "https://graph.microsoft.com/v1.0/users/$OrganizerId/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='$OrganizerId')"

do {
    $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    if ($response.value) { $allTranscripts.AddRange($response.value) }
    $uri = $response.'@odata.nextLink'
} while ($uri)

Write-Host "  Total transcripts found: $($allTranscripts.Count)" -ForegroundColor Green
$allTranscripts | Format-List

if ($allTranscripts.Count -eq 0) {
    Write-Warning "No transcripts found. Exiting."
    exit 0
}

# ============================================================
# STEP 3: Download transcripts (if requested)
# ============================================================
if (-not $DownloadTranscripts) {
    Write-Host "`nTip: Re-run with -DownloadTranscripts to save all transcripts as .vtt files." -ForegroundColor Yellow
    exit 0
}

# Create output folder
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "`nCreated output folder: $OutputFolder" -ForegroundColor Gray
}

Write-Host "`n[3] Downloading $($allTranscripts.Count) transcripts..." -ForegroundColor Cyan

# Get the user's OneDrive drive ID once (used for fallback)
Write-Host "  Resolving OneDrive drive ID for fallback..." -ForegroundColor Gray
try {
    $drive = Invoke-RestMethod `
        -Method Get `
        -Uri "https://graph.microsoft.com/v1.0/users/$OrganizerId/drive" `
        -Headers $headers
    $driveId = $drive.id
    Write-Host "  Drive ID: $driveId" -ForegroundColor Gray
} catch {
    Write-Warning "  Could not resolve OneDrive drive — fallback will be unavailable: $($_.Exception.Message)"
    $driveId = $null
}

$index     = 0
$succeeded = 0
$failed    = 0
$fallbacks = 0

foreach ($transcript in $allTranscripts) {
    $index++

    $meetingId    = $transcript.meetingId
    $transcriptId = $transcript.id
    $created      = $transcript.createdDateTime -replace '[T]', '_' -replace '[:\.]', '-' -replace 'Z', ''
    $fileName     = "transcript_${created}_${index}.vtt"
    $outPath      = Join-Path $OutputFolder $fileName

    Write-Host "  [$index/$($allTranscripts.Count)] $fileName" -ForegroundColor Gray

    # --- Primary: Transcript content endpoint ---
    $contentUri = "https://graph.microsoft.com/v1.0/users/$OrganizerId/onlineMeetings/$meetingId/transcripts/$transcriptId/content"

    try {
        
        Write-Host "uri: $contentUri" -ForegroundColor Yellow
        Invoke-RestMethod `
            -Method Get `
            -Uri $contentUri `
            -Headers @{
                Authorization = "Bearer $($tokenResponse.access_token)"
                Accept        = "text/vtt"
            } `
            -OutFile $outPath

        Write-Host "    Saved (primary): $outPath" -ForegroundColor Green
        $succeeded++
        continue   # Move to next transcript

    } catch {
        $errorBody  = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
        $innerCode  = $errorBody.error.innerError.code
        $statusCode = $_.Exception.Response.StatusCode.value__

        if ($statusCode -eq 404 -and $driveId) {
            # 3004: Meeting resource expired — fall back to OneDrive DriveItem search
            Write-Warning "    Primary failed (meeting expired) — trying OneDrive fallback..."
        } elseif ($innerCode -eq "SpeakerAttributionNotAllowed") {
            # Retry with unattributed format
            Write-Warning "    Speaker attribution disabled — retrying without speaker info..."
            try {
                Invoke-RestMethod `
                    -Method Get `
                    -Uri $contentUri `
                    -Headers @{
                        Authorization = "Bearer $($tokenResponse.access_token)"
                        Accept        = "application/vnd.microsoft.graph.transcript+text"
                    } `
                    -OutFile $outPath
                Write-Host "    Saved (unattributed): $outPath" -ForegroundColor Yellow
                $succeeded++
                continue
            } catch {
                Write-Warning "    Retry also failed: $($_.Exception.Message)"
                $failed++
                continue
            }
        } else {
            Write-Warning "    Failed [$statusCode] $innerCode : $($errorBody.error.message)"
            $failed++
            continue
        }
    }

    # --- Fallback: Search OneDrive Recordings folder by createdDateTime ---
    # Transcripts are stored as .vtt files alongside the .mp4 recording in OneDrive
    if ($driveId) {
        try {
            # List all items in the Recordings folder
            $recordingItems = Invoke-RestMethod `
                -Method Get `
                -Uri "https://graph.microsoft.com/v1.0/users/$OrganizerId/drive/root:/Documents/Recordings:/children?`$select=id,name,size,lastModifiedDateTime,@microsoft.graph.downloadUrl" `
                -Headers $headers

            # Match by looking for .vtt files modified close to the transcript createdDateTime
            $transcriptDate = [datetime]$transcript.createdDateTime
            $matchedVtt = $recordingItems.value | Where-Object {
                $_.name -like "*.vtt" -and
                [math]::Abs(([datetime]$_.lastModifiedDateTime - $transcriptDate).TotalHours) -lt 24
            } | Select-Object -First 1

            if ($matchedVtt) {
                # Use the pre-authenticated download URL — no auth header needed
                $downloadUrl = $matchedVtt.'@microsoft.graph.downloadUrl'
                Invoke-RestMethod -Method Get -Uri $downloadUrl -OutFile $outPath
                Write-Host "    Saved (OneDrive fallback): $outPath  [$($matchedVtt.name)]" -ForegroundColor Yellow
                $succeeded++
                $fallbacks++
            } else {
                # Last resort: list all VTT files so we can see what's available
                $allVtts = $recordingItems.value | Where-Object { $_.name -like "*.vtt" }
                if ($allVtts.Count -gt 0) {
                    Write-Warning "    No date-matched VTT found. Available VTT files in Recordings folder:"
                    $allVtts | ForEach-Object { Write-Warning "      $($_.name) [$($_.lastModifiedDateTime)]" }
                } else {
                    Write-Warning "    No VTT files found in OneDrive Recordings folder."
                }
                $failed++
            }

        } catch {
            Write-Warning "    OneDrive fallback failed: $($_.Exception.Message)"
            $failed++
        }
    }
}

# ============================================================
# SUMMARY
# ============================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Download complete" -ForegroundColor Cyan
Write-Host "  Succeeded  : $succeeded" -ForegroundColor Green
if ($fallbacks -gt 0) {
    Write-Host "  Via OneDrive fallback: $fallbacks" -ForegroundColor Yellow
}
if ($failed -gt 0) {
    Write-Host "  Failed     : $failed" -ForegroundColor Red
}
Write-Host "  Output     : $OutputFolder" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan