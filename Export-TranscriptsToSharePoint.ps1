#Requires -Version 7.0
<#
.SYNOPSIS
    Reads Attention transcript JSON files and uploads them as formatted .docx files
    to the matching SharePoint "Call Transcripts" folder.

.DESCRIPTION
    For each *.json file in -TranscriptDir (skipping files that start with "_"):
      1. Reads the customer and project fields from the JSON.
      2. Searches Graph for a SharePoint site whose display name contains both
         the customer and project strings.
      3. Checks (and optionally creates) the -FolderName folder at the root of the
         site's Documents library.
      4. Builds a formatted .docx from the transcript using raw Open XML packaged
         via System.IO.Compression.ZipArchive — no Word installation required.
      5. Uploads the .docx to the "Call Transcripts" folder via Graph PUT.

    Use -WhatIf to preview all decisions (source file → matched site → target path)
    without writing any files or calling any mutating Graph endpoints.

    Credentials must be supplied via environment variable — do NOT hardcode secrets:
      $env:GRAPH_CLIENT_SECRET = "your_secret"

.PARAMETER TranscriptDir
    Folder containing Attention transcript JSON files.
    Default: C:\Temp\AttentionTranscripts

.PARAMETER FolderName
    Name of the target folder at the root of each site's Documents library.
    Default: "Call Transcripts"

.PARAMETER CreateFolderIfMissing
    When present, creates the -FolderName folder if it doesn't exist.
    Requires Files.ReadWrite.All on the app registration.

.PARAMETER WhatIf
    Dry run. Prints the full plan for each transcript — site match, folder status,
    target path — without creating any files or uploading anything.

.EXAMPLE
    # Preview what would happen
    pwsh ./Export-TranscriptsToSharePoint.ps1 -WhatIf

.EXAMPLE
    # Upload transcripts; skip any site whose Call Transcripts folder is missing
    pwsh ./Export-TranscriptsToSharePoint.ps1

.EXAMPLE
    # Upload and auto-create the folder if it's absent
    pwsh ./Export-TranscriptsToSharePoint.ps1 -CreateFolderIfMissing
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string] $TranscriptDir = "C:\Temp\AttentionTranscripts",

    [Parameter(Mandatory = $false)]
    [string] $FolderName = "Call Transcripts",

    [Parameter(Mandatory = $false)]
    [switch] $CreateFolderIfMissing,

    [Parameter(Mandatory = $false)]
    [switch] $WhatIf
)

# ── Config ─────────────────────────────────────────────────────────────────────
$TenantId     = "d49110b2-6f26-4c66-b723-1729cdb9a3cf"
$ClientId     = "5b5e9181-c9a9-4c5c-8036-8e58f5c7b10b"
$ClientSecret = $env:GRAPH_CLIENT_SECRET
$SearchRegion = "NAM"
$SearchPage   = 25
# ──────────────────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
    Write-Host "ERROR: Set the environment variable GRAPH_CLIENT_SECRET before running." -ForegroundColor Red
    exit 1
}

# ══════════════════════════════════════════════════════════════════════════════
#  Helper functions
# ══════════════════════════════════════════════════════════════════════════════

# Replace characters that are illegal in Windows filenames
function Get-SafeFilename([string]$name) {
    return ($name -replace '[\\/:*?"<>|]', '_').Trim()
}

# XML-escape a string for embedding in attribute values or text nodes
function ConvertTo-XmlText([string]$text) {
    return $text `
        -replace '&',  '&amp;'  `
        -replace '<',  '&lt;'   `
        -replace '>',  '&gt;'   `
        -replace '"',  '&quot;' `
        -replace "'", '&apos;'
}

# Parse the flat transcript string into an ordered list of speaker/text pairs.
# Speaker label heuristic: a line that starts with a non-whitespace character,
# contains only a name (max 50 chars) and ends with ":", with nothing else on
# the same line. This avoids false positives like "here's the thing:" mid-speech.
function ConvertTo-SpeakerTurns([string]$transcript) {
    $turns   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $lines   = $transcript -split "`n"
    $speaker = $null
    $buffer  = [System.Collections.Generic.List[string]]::new()

    foreach ($raw in $lines) {
        $line = $raw.Trim()
        # Match: non-whitespace start, name-length content (1–50 chars), trailing colon only
        if ($line -match '^\S[^:\n]{1,50}:\s*$') {
            if ($null -ne $speaker -and $buffer.Count -gt 0) {
                $turns.Add([PSCustomObject]@{
                    Speaker = $speaker
                    Text    = ($buffer -join ' ').Trim()
                })
                $buffer.Clear()
            }
            $speaker = $line.TrimEnd(':').Trim()
        }
        elseif ($line.Length -gt 0) {
            $buffer.Add($line)
        }
    }

    # Flush the final turn
    if ($null -ne $speaker -and $buffer.Count -gt 0) {
        $turns.Add([PSCustomObject]@{
            Speaker = $speaker
            Text    = ($buffer -join ' ').Trim()
        })
    }

    return ,$turns   # comma forces List return; avoids PS unwrapping single-item lists
}

# Build a .docx file at $OutPath using raw Open XML + ZipArchive.
# No Word installation required — only System.IO.Compression (built into .NET).
function New-TranscriptDocx {
    param(
        [string] $Title,
        [string] $DateStr,
        [string] $TranscriptText,
        [string] $OutPath
    )

    $turns = ConvertTo-SpeakerTurns -transcript $TranscriptText

    $titleXml = ConvertTo-XmlText $Title
    $dateXml  = ConvertTo-XmlText $DateStr

    # Build one pair of paragraphs per speaker turn:
    #   <bold>  Speaker:
    #   <normal> Their text
    $turnXml = [System.Text.StringBuilder]::new()
    foreach ($turn in $turns) {
        $spkXml  = ConvertTo-XmlText "$($turn.Speaker):"
        $textXml = ConvertTo-XmlText $turn.Text

        $null = $turnXml.Append("    <w:p>
      <w:r><w:rPr><w:b/></w:rPr>
        <w:t xml:space=""preserve"">$spkXml</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space=""preserve"">$textXml</w:t></w:r>
    </w:p>`n")
    }

    # ── word/document.xml ─────────────────────────────────────────────────
    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>$titleXml</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>
      <w:r><w:t>$dateXml</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>
        </w:pBdr>
        <w:spacing w:after="120"/>
      </w:pPr>
    </w:p>
$($turnXml.ToString())    <w:sectPr>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"@

    # ── word/styles.xml ───────────────────────────────────────────────────
    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- Normal (default paragraph) -->
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="160"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- Heading 1 -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr>
      <w:outlineLvl w:val="0"/>
      <w:spacing w:before="240" w:after="60"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>
      <w:b/><w:sz w:val="32"/><w:szCs w:val="32"/>
      <w:color w:val="2F5496"/>
    </w:rPr>
  </w:style>

  <!-- Subtitle -->
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="subtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:spacing w:before="0" w:after="200"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
      <w:color w:val="595959"/>
    </w:rPr>
  </w:style>

</w:styles>
'@

    # ── Package relationships ─────────────────────────────────────────────
    $contentTypesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
'@

    $relsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>
'@

    $wordRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>
'@

    # ── Write ZIP (.docx) ─────────────────────────────────────────────────
    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }

    $stream = [System.IO.File]::Open($OutPath, [System.IO.FileMode]::CreateNew)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new(
            $stream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $enc = [System.Text.Encoding]::UTF8

            $parts = [ordered]@{
                '[Content_Types].xml'           = $contentTypesXml
                '_rels/.rels'                   = $relsXml
                'word/_rels/document.xml.rels'  = $wordRelsXml
                'word/document.xml'             = $documentXml
                'word/styles.xml'               = $stylesXml
            }

            foreach ($name in $parts.Keys) {
                $entry  = $zip.CreateEntry($name)
                $writer = $entry.Open()
                try {
                    $bytes = $enc.GetBytes($parts[$name])
                    $writer.Write($bytes, 0, $bytes.Length)
                }
                finally { $writer.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $stream.Dispose() }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Authenticate
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "Authenticating..." -ForegroundColor Cyan

try {
    $tokenResponse = Invoke-RestMethod `
        -Uri     "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Method  Post `
        -Body    @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
        }
    $token = $tokenResponse.access_token
    Write-Host "Token acquired." -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Failed to get token — $($_.ErrorDetails.Message)" -ForegroundColor Red
    exit 1
}

$jsonHeaders = @{
    Authorization  = "Bearer $token"
    "Content-Type" = "application/json"
}

# ══════════════════════════════════════════════════════════════════════════════
#  Graph helpers (scoped to use $jsonHeaders / $token from parent scope)
# ══════════════════════════════════════════════════════════════════════════════

# Search for SharePoint sites matching $query; return all paged results
function Search-Sites([string]$query) {
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $from    = 0
    $total   = $null

    do {
        $body = @{
            requests = @(@{
                entityTypes = @("site")
                query       = @{ queryString = $query }
                fields      = @("displayName", "webUrl")
                from        = $from
                size        = $SearchPage
                region      = $SearchRegion
            })
        } | ConvertTo-Json -Depth 5

        try {
            $resp = Invoke-RestMethod `
                -Uri     "https://graph.microsoft.com/v1.0/search/query" `
                -Headers $jsonHeaders `
                -Method  Post `
                -Body    $body
        }
        catch {
            Write-Host "  WARN: Search failed — $($_.ErrorDetails.Message)" -ForegroundColor Yellow
            return ,$results
        }

        $container = $resp.value[0].hitsContainers[0]
        if ($null -eq $total) { $total = $container.total }
        if ($total -eq 0) { break }

        foreach ($hit in $container.hits) {
            $results.Add([PSCustomObject]@{
                DisplayName = $hit.resource.displayName
                WebUrl      = $hit.resource.webUrl
            })
        }
        $from += $SearchPage

    } while ($from -lt $total)

    return ,$results
}

# Resolve the Documents drive and check whether $FolderName exists at its root.
# Returns a PSCustomObject with DriveId, DriveName, FolderFound, FolderWebUrl.
# Returns $null on any unrecoverable error.
function Resolve-DriveAndFolder([string]$siteUrl) {
    $uri      = [System.Uri]$siteUrl
    $hostname = $uri.Host
    $path     = $uri.AbsolutePath

    # Step 1: site ID
    try {
        $siteRes = Invoke-RestMethod `
            -Uri     "https://graph.microsoft.com/v1.0/sites/${hostname}:${path}" `
            -Headers $jsonHeaders `
            -Method  Get
        $siteId = $siteRes.id
    }
    catch {
        $status = $_.Exception.Response.StatusCode.value__
        Write-Host "  WARN: Could not resolve site ID ($status) for $siteUrl" -ForegroundColor Yellow
        return $null
    }

    # Step 2: Documents drive
    try {
        $drivesRes = Invoke-RestMethod `
            -Uri     "https://graph.microsoft.com/v1.0/sites/$siteId/drives" `
            -Headers $jsonHeaders `
            -Method  Get
    }
    catch {
        Write-Host "  WARN: Could not list drives ($($_.Exception.Response.StatusCode.value__))" -ForegroundColor Yellow
        return $null
    }

    $drive = $drivesRes.value |
        Where-Object { $_.driveType -eq "documentLibrary" -and $_.name -eq "Documents" } |
        Select-Object -First 1

    if (-not $drive) {
        $drive = $drivesRes.value |
            Where-Object { $_.driveType -eq "documentLibrary" } |
            Select-Object -First 1
    }

    if (-not $drive) {
        Write-Host "  WARN: No document library found on site." -ForegroundColor Yellow
        return $null
    }

    # Step 3: walk root children (with nextLink paging) to find the folder
    $folderFound  = $false
    $folderWebUrl = $null
    $nextLink     = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root/children"

    do {
        try {
            $childrenRes = Invoke-RestMethod `
                -Uri     $nextLink `
                -Headers $jsonHeaders `
                -Method  Get
        }
        catch {
            Write-Host "  WARN: Failed to list drive root ($($_.Exception.Response.StatusCode.value__))" -ForegroundColor Yellow
            break
        }

        foreach ($item in $childrenRes.value) {
            if ($null -ne $item.folder -and $item.name -ieq $FolderName) {
                $folderFound  = $true
                $folderWebUrl = $item.webUrl
                break
            }
        }
        $nextLink = $childrenRes.'@odata.nextLink'

    } while (-not $folderFound -and $null -ne $nextLink)

    return [PSCustomObject]@{
        DriveId      = $drive.id
        DriveName    = $drive.name
        FolderFound  = $folderFound
        FolderWebUrl = $folderWebUrl
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Enumerate transcript files
# ══════════════════════════════════════════════════════════════════════════════
Write-Host "`nScanning '$TranscriptDir'..." -ForegroundColor Cyan

if (-not (Test-Path $TranscriptDir)) {
    Write-Host "ERROR: Directory not found: $TranscriptDir" -ForegroundColor Red
    exit 1
}

$files = Get-ChildItem -Path $TranscriptDir -Filter "*.json" |
    Where-Object { $_.Name -notmatch '^_' }

Write-Host "  Found $($files.Count) transcript file(s)." -ForegroundColor Green

if ($files.Count -eq 0) {
    Write-Host "Nothing to do." -ForegroundColor Yellow
    exit 0
}

if ($WhatIf) {
    Write-Host "`n[WHATIF MODE] No files will be written or uploaded.`n" -ForegroundColor Magenta
}

# ══════════════════════════════════════════════════════════════════════════════
#  Counters
# ══════════════════════════════════════════════════════════════════════════════
$cUploaded  = 0
$cExists    = 0
$cNoMatch   = 0
$cAmbiguous = 0
$cNoFolder  = 0
$cErrors    = 0

# ══════════════════════════════════════════════════════════════════════════════
#  Per-transcript loop
# ══════════════════════════════════════════════════════════════════════════════
foreach ($file in $files) {
    Write-Host "`n── $($file.Name)" -ForegroundColor Cyan
    Write-Host "   ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray

    # ── Parse JSON ──────────────────────────────────────────────────────────
    try {
        # Use .NET ReadAllText so UTF-8 BOM is handled automatically
        $t = [System.IO.File]::ReadAllText($file.FullName) | ConvertFrom-Json
    }
    catch {
        Write-Host "  [ERROR] Cannot parse JSON: $($_.Exception.Message)" -ForegroundColor Red
        $cErrors++
        continue
    }

    $customer = $t.customer
    $project  = $t.project
    $title    = $t.title

    $dateStr  = try {
        [datetime]::Parse($t.createdAt).ToUniversalTime().ToString("MMMM d, yyyy 'at' h:mm tt 'UTC'")
    }
    catch { $t.createdAt }

    $safeTitle = Get-SafeFilename $title
    $docxName  = "$safeTitle.docx"
    $tempDocx  = Join-Path $env:TEMP $docxName

    Write-Host "  Title    : $title"
    Write-Host "  Customer : $customer"
    Write-Host "  Project  : $project"
    Write-Host "  Date     : $dateStr"

    # ── Site matching ────────────────────────────────────────────────────────
    $allSites = Search-Sites -query $customer

    $matched = @($allSites | Where-Object {
        $_.DisplayName -imatch [regex]::Escape($customer) -and
        $_.DisplayName -imatch [regex]::Escape($project)
    })

    if ($matched.Count -eq 0) {
        Write-Host "  [NO MATCH] No site found containing both '$customer' and '$project' in display name." -ForegroundColor Yellow
        $cNoMatch++
        continue
    }

    if ($matched.Count -gt 1) {
        Write-Host "  [AMBIGUOUS] $($matched.Count) sites matched — skipping. Matches:" -ForegroundColor Yellow
        $matched | ForEach-Object { Write-Host "    - $($_.DisplayName)  $($_.WebUrl)" }
        $cAmbiguous++
        continue
    }

    $site = $matched[0]
    Write-Host "  Site     : $($site.DisplayName)" -ForegroundColor Green
    Write-Host "  SiteUrl  : $($site.WebUrl)"

    # ── Resolve drive + folder ───────────────────────────────────────────────
    $driveInfo = Resolve-DriveAndFolder -siteUrl $site.WebUrl

    if ($null -eq $driveInfo) {
        Write-Host "  [ERROR] Could not resolve drive — skipping." -ForegroundColor Red
        $cErrors++
        continue
    }

    $targetPath = "Documents/$FolderName/$docxName"

    $folderStatus = if ($driveInfo.FolderFound) {
        "[EXISTS]  $($driveInfo.FolderWebUrl)"
    }
    elseif ($CreateFolderIfMissing) {
        "[MISSING — would create]"
    }
    else {
        "[MISSING — would skip (pass -CreateFolderIfMissing to create it)]"
    }

    # ── WhatIf report — stop here ────────────────────────────────────────────
    if ($WhatIf) {
        Write-Host ""
        Write-Host "  [WHATIF]" -ForegroundColor Magenta -NoNewline
        Write-Host " $title"
        Write-Host "    Source  : $($file.Name)"
        Write-Host "    Site    : $($site.DisplayName)"
        Write-Host "    SiteUrl : $($site.WebUrl)"
        Write-Host "    Target  : $targetPath"
        Write-Host "    Folder  : $folderStatus"
        $cUploaded++
        continue
    }

    # ── Ensure folder exists (real run only) ─────────────────────────────────
    if (-not $driveInfo.FolderFound) {
        if (-not $CreateFolderIfMissing) {
            Write-Host "  [NO FOLDER] '$FolderName' not found. Pass -CreateFolderIfMissing to create it." -ForegroundColor Yellow
            $cNoFolder++
            continue
        }

        Write-Host "  Creating folder '$FolderName'..." -ForegroundColor Cyan

        $newFolderBody = @{
            name   = $FolderName
            folder = @{}
            "@microsoft.graph.conflictBehavior" = "fail"
        } | ConvertTo-Json -Depth 3

        try {
            $created = Invoke-RestMethod `
                -Uri     "https://graph.microsoft.com/v1.0/drives/$($driveInfo.DriveId)/items/root/children" `
                -Headers $jsonHeaders `
                -Method  Post `
                -Body    $newFolderBody
            Write-Host "  [CREATED] $($created.webUrl)" -ForegroundColor Green
        }
        catch {
            $status = $_.Exception.Response.StatusCode.value__
            switch ($status) {
                409 {
                    Write-Host "  [INFO] Folder already exists (race condition — continuing)." -ForegroundColor Gray
                }
                403 {
                    Write-Host "  [ERROR] Permission denied creating folder." -ForegroundColor Red
                    Write-Host "    Ensure 'Files.ReadWrite.All' (application permission) is granted"
                    Write-Host "    and admin consent has been applied to app $ClientId."
                    $cErrors++
                    continue
                }
                default {
                    Write-Host "  [ERROR] Failed to create folder ($status)." -ForegroundColor Red
                    Write-Host "    $($_.ErrorDetails.Message)"
                    $cErrors++
                    continue
                }
            }
        }
    }

    # ── Build .docx ──────────────────────────────────────────────────────────
    Write-Host "  Building $docxName..." -ForegroundColor Cyan

    try {
        New-TranscriptDocx `
            -Title          $title `
            -DateStr        $dateStr `
            -TranscriptText $t.transcript `
            -OutPath        $tempDocx
        Write-Host "  Docx written: $tempDocx" -ForegroundColor Green
    }
    catch {
        Write-Host "  [ERROR] Failed to build docx: $($_.Exception.Message)" -ForegroundColor Red
        $cErrors++
        continue
    }

    # ── Upload to SharePoint ─────────────────────────────────────────────────
    Write-Host "  Uploading to SharePoint..." -ForegroundColor Cyan

    $encodedName = [uri]::EscapeDataString($docxName)
    $uploadUri   = "https://graph.microsoft.com/v1.0/drives/$($driveInfo.DriveId)" +
                   "/items/root:/$([uri]::EscapeDataString($FolderName))/$($encodedName):/content"

    $uploadHeaders = @{
        Authorization  = "Bearer $token"
        "Content-Type" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    try {
        $uploaded = Invoke-RestMethod `
            -Uri     $uploadUri `
            -Headers $uploadHeaders `
            -Method  Put `
            -InFile  $tempDocx
        Write-Host "  [UPLOADED] $($uploaded.webUrl)" -ForegroundColor Green
        $cUploaded++
    }
    catch {
        $status = $_.Exception.Response.StatusCode.value__
        switch ($status) {
            409 {
                Write-Host "  [EXISTS] File already exists in SharePoint." -ForegroundColor Gray
                $cExists++
            }
            403 {
                Write-Host "  [ERROR] Permission denied uploading file." -ForegroundColor Red
                Write-Host "    Ensure 'Files.ReadWrite.All' is granted and admin-consented."
                $cErrors++
            }
            default {
                Write-Host "  [ERROR] Upload failed ($status): $($_.ErrorDetails.Message)" -ForegroundColor Red
                $cErrors++
            }
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  Summary
# ══════════════════════════════════════════════════════════════════════════════
$whatIfLabel = if ($WhatIf) { " (WhatIf)" } else { "" }
$uploadLabel = "  Uploaded$whatIfLabel"

# Pad width = longest label across all possible states (  Ambiguous site match = 24 chars)
$padWidth = 24

Write-Host "`n── Summary ─────────────────────────────────────────────────────" -ForegroundColor Cyan
Write-Host ("{0,-$padWidth} : {1}" -f "  Transcripts scanned",   $files.Count)
Write-Host ("{0,-$padWidth} : {1}" -f $uploadLabel,               $cUploaded)
Write-Host ("{0,-$padWidth} : {1}" -f "  Already exists",         $cExists)
Write-Host ("{0,-$padWidth} : {1}" -f "  No site match",          $cNoMatch)
Write-Host ("{0,-$padWidth} : {1}" -f "  Ambiguous site match",   $cAmbiguous)
Write-Host ("{0,-$padWidth} : {1}" -f "  No folder (skipped)",    $cNoFolder)
Write-Host ("{0,-$padWidth} : {1}" -f "  Errors",                 $cErrors)
Write-Host "────────────────────────────────────────────────────────────────" -ForegroundColor Cyan
