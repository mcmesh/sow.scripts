#Requires -Version 7.0

[CmdletBinding()]
param(
    [string]$CustomerName,
    [string[]]$CustomerAliases = @(),
    [string]$SourceFolder,
    [string]$OutputPath,
    [ValidateSet("Auto", "Prompt")]
    [string]$SelectionMode,
    [bool]$Recurse = $true,
    [switch]$Force,
    [switch]$NonInteractive,
    [switch]$AllowIncompleteCoverage,
    [string]$DecisionFile
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

# create: no report exists; update: preserve the existing report and add changes; replace: -Force or overwriteOutput
$mode = "create"

function Write-Diagnostic {
    param([string]$Message)
    [Console]::Error.WriteLine($Message)
}

function Get-StableId {
    param(
        [string]$Prefix,
        [string]$Value
    )

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
    $hash = [System.Security.Cryptography.SHA256]::HashData($bytes)
    $shortHash = [Convert]::ToHexString($hash).ToLowerInvariant().Substring(0, 12)
    return "$Prefix-$shortHash"
}

function ConvertTo-NormalizedText {
    param([AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    $decomposed = $Value.Normalize([System.Text.NormalizationForm]::FormD)
    $characters = foreach ($character in $decomposed.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($character) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            $character
        }
    }
    $withoutDiacritics = (-join $characters).Normalize([System.Text.NormalizationForm]::FormC)
    return (($withoutDiacritics.ToLowerInvariant() -replace '[^\p{L}\p{Nd}]+', ' ').Trim() -replace '\s+', ' ')
}

function Write-ResultAndExit {
    param(
        [hashtable]$Result,
        [int]$ExitCode
    )

    [Console]::Out.WriteLine(($Result | ConvertTo-Json -Depth 30))
    exit $ExitCode
}

function New-BaseResult {
    param([string]$Status)

    return [ordered]@{
        schemaVersion      = "1.0"
        status             = $Status
        mode               = $mode
        customer           = [ordered]@{
            requestedName = $CustomerName
            aliases       = @($CustomerAliases)
            canonicalName = $CustomerName
        }
        sourceFolder       = $SourceFolder
        output             = [ordered]@{
            path                    = $OutputPath
            force                   = [bool]$Force
            allowIncompleteCoverage = [bool]$AllowIncompleteCoverage
        }
        coverage           = [ordered]@{
            enumerated = 0
            selected   = 0
            excluded   = 0
            failed     = 0
            unresolved = 0
        }
        selectedDocuments  = @()
        rejectedCandidates = @()
        excludedFiles      = @()
        failedFiles        = @()
        unresolvedGroups   = @()
        warnings           = @()
    }
}

function Get-ZipEntryText {
    param([System.IO.Compression.ZipArchiveEntry]$Entry)

    $stream = $Entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-WordNodeText {
    param(
        [System.Xml.XmlNode]$Node,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $parts = foreach ($textNode in $Node.SelectNodes('.//w:t[not(ancestor::w:del)]', $NamespaceManager)) {
        $textNode.InnerText
    }
    return (($parts -join "") -replace '\s+', ' ').Trim()
}

function Get-DocxExtraction {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $documentEntry = $archive.GetEntry("word/document.xml")
        if ($null -eq $documentEntry) {
            throw "The DOCX archive does not contain word/document.xml."
        }

        $sections = [System.Collections.Generic.List[object]]::new()
        $allText = [System.Collections.Generic.List[string]]::new()
        $trackedChanges = $false
        $paragraphNumber = 0
        $tableNumber = 0

        $documentXmlText = Get-ZipEntryText -Entry $documentEntry
        if ($documentXmlText -match '<w:(ins|del)(\s|>)') {
            $trackedChanges = $true
        }

        [xml]$documentXml = $documentXmlText
        $namespaceManager = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
        $namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        foreach ($node in $documentXml.SelectNodes('//w:body/*', $namespaceManager)) {
            if ($node.LocalName -eq "p") {
                $paragraphNumber++
                $text = Get-WordNodeText -Node $node -NamespaceManager $namespaceManager
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $styleNode = $node.SelectSingleNode('./w:pPr/w:pStyle', $namespaceManager)
                    $heading = $null
                    if ($null -ne $styleNode) {
                        $heading = $styleNode.GetAttribute("val", $namespaceManager.LookupNamespace("w"))
                    }
                    $sections.Add([pscustomobject]@{
                        locator = "document paragraph $paragraphNumber"
                        heading = $heading
                        text    = $text
                    })
                    $allText.Add($text)
                }
            }
            elseif ($node.LocalName -eq "tbl") {
                $tableNumber++
                $rowNumber = 0
                foreach ($row in $node.SelectNodes('./w:tr', $namespaceManager)) {
                    $rowNumber++
                    $cells = foreach ($cell in $row.SelectNodes('./w:tc', $namespaceManager)) {
                        Get-WordNodeText -Node $cell -NamespaceManager $namespaceManager
                    }
                    $text = (($cells | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " | ").Trim()
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        $sections.Add([pscustomobject]@{
                            locator = "table $tableNumber row $rowNumber"
                            heading = "Table $tableNumber"
                            text    = $text
                        })
                        $allText.Add($text)
                    }
                }
            }
        }

        foreach ($entry in $archive.Entries | Where-Object { $_.FullName -match '^word/(header|footer)\d+\.xml$' }) {
            $partName = [System.IO.Path]::GetFileNameWithoutExtension($entry.Name)
            [xml]$partXml = Get-ZipEntryText -Entry $entry
            $partNamespaceManager = [System.Xml.XmlNamespaceManager]::new($partXml.NameTable)
            $partNamespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            $partParagraph = 0
            foreach ($node in $partXml.SelectNodes('//w:p', $partNamespaceManager)) {
                $partParagraph++
                $text = Get-WordNodeText -Node $node -NamespaceManager $partNamespaceManager
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $sections.Add([pscustomobject]@{
                        locator = "$partName paragraph $partParagraph"
                        heading = $partName
                        text    = $text
                    })
                    $allText.Add($text)
                }
            }
        }

        $title = $null
        $coreEntry = $archive.GetEntry("docProps/core.xml")
        if ($null -ne $coreEntry) {
            [xml]$coreXml = Get-ZipEntryText -Entry $coreEntry
            $titleNode = $coreXml.SelectSingleNode("//*[local-name()='title']")
            if ($null -ne $titleNode -and -not [string]::IsNullOrWhiteSpace($titleNode.InnerText)) {
                $title = $titleNode.InnerText.Trim()
            }
        }

        $mediaCount = @($archive.Entries | Where-Object { $_.FullName -like 'word/media/*' }).Count
        if ($sections.Count -eq 0) {
            throw "No readable text was found in the DOCX document."
        }

        return [pscustomobject]@{
            title          = $title
            text           = ($allText -join [Environment]::NewLine)
            sections       = $sections.ToArray()
            status         = "complete"
            extractor      = "PowerShell Open XML"
            version        = $PSVersionTable.PSVersion.ToString()
            trackedChanges = $trackedChanges
            mediaCount     = $mediaCount
        }
    }
    finally {
        $archive.Dispose()
    }
}

function Get-PdfTool {
    $command = Get-Command "pdftotext" -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -eq $command) {
        return $null
    }

    $versionText = (& $command.Source -v 2>&1 | Select-Object -First 1).ToString()
    return [pscustomobject]@{
        path    = $command.Source
        version = $versionText
    }
}

function Get-PdfExtraction {
    param(
        [string]$Path,
        [pscustomobject]$PdfTool
    )

    if ($null -eq $PdfTool) {
        throw "pdftotext was not found on PATH."
    }

    $temporaryOutput = [System.IO.Path]::GetTempFileName()
    $temporaryError = [System.IO.Path]::GetTempFileName()
    try {
        & $PdfTool.path -layout -enc UTF-8 $Path $temporaryOutput 2> $temporaryError
        if ($LASTEXITCODE -ne 0) {
            $detail = (Get-Content -LiteralPath $temporaryError -Raw -ErrorAction SilentlyContinue).Trim()
            throw "pdftotext failed with exit code $LASTEXITCODE. $detail"
        }

        $text = Get-Content -LiteralPath $temporaryOutput -Raw
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw "No text was extracted; the PDF may be scanned or image-only."
        }

        $sections = [System.Collections.Generic.List[object]]::new()
        $pages = $text -split "`f"
        for ($index = 0; $index -lt $pages.Count; $index++) {
            $pageText = ($pages[$index] -replace '[ \t]+\r?\n', [Environment]::NewLine).Trim()
            if (-not [string]::IsNullOrWhiteSpace($pageText)) {
                $sections.Add([pscustomobject]@{
                    locator = "page $($index + 1)"
                    heading = $null
                    text    = $pageText
                })
            }
        }

        if ($sections.Count -eq 0) {
            throw "The PDF extractor returned no readable pages."
        }

        return [pscustomobject]@{
            title          = $null
            text           = $text.Trim()
            sections       = $sections.ToArray()
            status         = "complete"
            extractor      = $PdfTool.path
            version        = $PdfTool.version
            trackedChanges = $false
            mediaCount     = 0
        }
    }
    finally {
        Remove-Item -LiteralPath $temporaryOutput -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $temporaryError -Force -ErrorAction SilentlyContinue
    }
}

function Get-CustomerMatch {
    param(
        [System.IO.FileInfo]$File,
        [pscustomobject]$Extraction,
        [string[]]$Names
    )

    $evidence = [System.Collections.Generic.List[object]]::new()
    $score = 0
    $normalizedFileName = ConvertTo-NormalizedText -Value $File.BaseName
    $normalizedFolder = ConvertTo-NormalizedText -Value $File.DirectoryName
    $normalizedTitle = ConvertTo-NormalizedText -Value $Extraction.title
    $normalizedText = ConvertTo-NormalizedText -Value $Extraction.text

    foreach ($name in $Names) {
        $normalizedName = ConvertTo-NormalizedText -Value $name
        if ([string]::IsNullOrWhiteSpace($normalizedName)) {
            continue
        }

        if ($normalizedFileName.Contains($normalizedName)) {
            $score += 45
            $evidence.Add([pscustomobject]@{ reason = "Filename matched '$name'"; score = 45; locator = "filename" })
        }
        if ($normalizedFolder.Contains($normalizedName)) {
            $score += 30
            $evidence.Add([pscustomobject]@{ reason = "Folder matched '$name'"; score = 30; locator = "folder" })
        }
        if ($normalizedTitle.Contains($normalizedName)) {
            $score += 35
            $evidence.Add([pscustomobject]@{ reason = "Document title matched '$name'"; score = 35; locator = "document properties" })
        }
        if ($normalizedText.Contains($normalizedName)) {
            $score += 20
            $matchedSection = $Extraction.sections | Where-Object {
                (ConvertTo-NormalizedText -Value $_.text).Contains($normalizedName)
            } | Select-Object -First 1
            $locator = if ($null -ne $matchedSection) { $matchedSection.locator } else { "document text" }
            $evidence.Add([pscustomobject]@{ reason = "Document text matched '$name'"; score = 20; locator = $locator })
        }

        $partyPattern = "(customer|client|between|party|agreement with).{0,100}$([regex]::Escape($normalizedName))|$([regex]::Escape($normalizedName)).{0,100}(customer|client|between|party)"
        if ($normalizedText -match $partyPattern) {
            $score += 50
            $matchedSection = $Extraction.sections | Where-Object {
                (ConvertTo-NormalizedText -Value $_.text) -match $partyPattern
            } | Select-Object -First 1
            $locator = if ($null -ne $matchedSection) { $matchedSection.locator } else { "document text" }
            $evidence.Add([pscustomobject]@{ reason = "Contracting-party context matched '$name'"; score = 50; locator = $locator })
        }
    }

    return [pscustomobject]@{
        score    = $score
        evidence = $evidence.ToArray()
    }
}

function Get-SignatureAssessment {
    param(
        [System.IO.FileInfo]$File,
        [pscustomobject]$Extraction
    )

    $evidence = [System.Collections.Generic.List[object]]::new()
    $rawText = $Extraction.text
    $normalizedText = ConvertTo-NormalizedText -Value $rawText
    $normalizedName = ConvertTo-NormalizedText -Value $File.BaseName
    $score = 0

    $explicitExecution = $rawText -match '(?im)^\s*(status\s*:\s*)?fully executed\b|this\s+(agreement|statement of work|sow|document)\s+(is|has been)\s+fully executed\b|certificate of completion|all parties have signed|signing (is )?complete|completed electronic signature'
    $partialEvidence = $normalizedText -match '(signed by|signature date|docusign|adobe sign|electronic signature)'
    $negativeEvidence = $normalizedName -match '\b(draft|redline|template|unsigned|working|review)\b' -or
        $normalizedText -match '\b(unsigned|draft for review|not for execution)\b' -or
        $rawText -match '(?im)^\s*(signature|signed by)\s*:\s*(_+|\s*)$'

    if ($explicitExecution) {
        $score += 70
        $section = $Extraction.sections | Where-Object {
            $_.text -match '(?im)^\s*(status\s*:\s*)?fully executed\b|this\s+(agreement|statement of work|sow|document)\s+(is|has been)\s+fully executed\b|certificate of completion|all parties have signed|signing (is )?complete|completed electronic signature'
        } | Select-Object -First 1
        $evidence.Add([pscustomobject]@{
            reason  = "Explicit execution-completion text detected"
            score   = 70
            locator = if ($null -ne $section) { $section.locator } else { "document text" }
        })
    }
    elseif ($partialEvidence) {
        $score += 35
        $section = $Extraction.sections | Where-Object {
            (ConvertTo-NormalizedText -Value $_.text) -match '(signed by|signature date|docusign|adobe sign|electronic signature)'
        } | Select-Object -First 1
        $evidence.Add([pscustomobject]@{
            reason  = "Partial signature-related text detected"
            score   = 35
            locator = if ($null -ne $section) { $section.locator } else { "document text" }
        })
    }

    if ($normalizedName -match '\b(signed|executed)\b') {
        $score += 20
        $evidence.Add([pscustomobject]@{ reason = "Filename contains signed/executed marker"; score = 20; locator = "filename" })
    }
    elseif ($normalizedName -match '\b(final|approved)\b') {
        $score += 10
        $evidence.Add([pscustomobject]@{ reason = "Filename contains final/approved marker"; score = 10; locator = "filename" })
    }

    if ($Extraction.mediaCount -gt 0) {
        $evidence.Add([pscustomobject]@{
            reason  = "Document contains embedded images; image signatures are not verified"
            score   = 0
            locator = "DOCX media"
        })
    }

    if ($negativeEvidence) {
        $score -= 100
        $evidence.Add([pscustomobject]@{ reason = "Draft, review, or unsigned marker detected"; score = -100; locator = "filename or document text" })
    }

    $status = if ($negativeEvidence) {
        "Unsigned"
    }
    elseif ($explicitExecution) {
        "Execution evidence detected"
    }
    elseif ($partialEvidence) {
        "Partially signed"
    }
    else {
        "Unverified"
    }

    return [pscustomobject]@{
        status   = $status
        score    = $score
        evidence = $evidence.ToArray()
    }
}

function Get-DateAssessment {
    param([pscustomobject]$Extraction)

    $pattern = '(?im)\b(effective date|effective as of)\s*[:\-]?\s*([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4}|\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4})'
    $match = [regex]::Match($Extraction.text, $pattern)
    if (-not $match.Success) {
        return [pscustomobject]@{ normalized = $null; original = $null; locator = $null }
    }

    $original = $match.Groups[2].Value
    $parsed = [datetime]::MinValue
    $normalized = $null
    if ([datetime]::TryParse($original, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
        $normalized = $parsed.ToString("yyyy-MM-dd")
    }
    $section = $Extraction.sections | Where-Object { $_.text -match [regex]::Escape($original) } | Select-Object -First 1

    return [pscustomobject]@{
        normalized = $normalized
        original   = $original
        locator    = if ($null -ne $section) { $section.locator } else { "document text" }
    }
}

function Get-AgreementType {
    param(
        [System.IO.FileInfo]$File,
        [pscustomobject]$Extraction
    )

    $combined = ConvertTo-NormalizedText -Value "$($File.BaseName) $($Extraction.title)"
    if ($combined -match '\b(amendment|addendum)\b') {
        return "Amendment"
    }
    if ($combined -match '\b(change order|change request)\b') {
        return "Change order"
    }
    return "SOW"
}

function Get-GroupKey {
    param(
        [System.IO.FileInfo]$File,
        [pscustomobject]$Extraction,
        [string]$AgreementType
    )

    $stem = ConvertTo-NormalizedText -Value $File.BaseName
    $stem = ($stem -replace '\b(signed|executed|final|approved|draft|redline|unsigned|copy|version|ver|v)\s*\d*\b', ' ')
    $stem = ($stem -replace '\b\d{4}[- ]\d{2}[- ]\d{2}\b', ' ' -replace '\s+', ' ').Trim()
    $title = ConvertTo-NormalizedText -Value $Extraction.title
    if (-not [string]::IsNullOrWhiteSpace($title)) {
        return "$AgreementType|$title"
    }
    return "$AgreementType|$stem"
}

function Read-Decisions {
    if ([string]::IsNullOrWhiteSpace($DecisionFile)) {
        return $null
    }
    if (-not (Test-Path -LiteralPath $DecisionFile -PathType Leaf)) {
        throw "Decision file does not exist: $DecisionFile"
    }

    $decision = Get-Content -LiteralPath $DecisionFile -Raw | ConvertFrom-Json
    if ($decision.schemaVersion -ne "1.0") {
        throw "Decision file schemaVersion must be 1.0."
    }
    return $decision
}

function Get-DecisionForGroup {
    param(
        [pscustomobject]$Decisions,
        [string]$GroupId
    )

    if ($null -eq $Decisions -or $null -eq $Decisions.groupDecisions) {
        return $null
    }
    return $Decisions.groupDecisions | Where-Object { $_.groupId -eq $GroupId } | Select-Object -First 1
}

function Get-SignatureRank {
    param([string]$Status)

    switch ($Status) {
        "Execution evidence detected" { return 3 }
        "Partially signed" { return 2 }
        "Unverified" { return 1 }
        default { return 0 }
    }
}

function Read-ExistingReport {
    param([string]$Path)

    $sources = [System.Collections.Generic.List[object]]::new()
    $skippedGroupIds = [System.Collections.Generic.List[string]]::new()
    $excludedFileIds = [System.Collections.Generic.List[string]]::new()
    $section = ""
    $generated = $null
    $allPaths = [System.Collections.Generic.List[string]]::new()

    foreach ($line in Get-Content -LiteralPath $Path) {
        foreach ($match in [regex]::Matches($line, '`([^`]+)`')) {
            $allPaths.Add($match.Groups[1].Value)
        }
        if ($line -match '^##\s+(.+?)\s*$') {
            $section = $Matches[1]
            continue
        }
        if ($null -eq $generated -and $line -match '^\|\s*Report generated\s*\|\s*([^|]+?)\s*\|') {
            $generated = $Matches[1]
        }
        if ($section -eq "Sources" -and $line -match '^\|\s*(\d+)\s*/\s*(doc-[0-9a-f]{12})\b') {
            $paths = @([regex]::Matches($line, '`([^`]+)`') | ForEach-Object { $_.Groups[1].Value })
            $sources.Add([pscustomobject]@{
                documentId = $Matches[2]
                sequence   = [int]$Matches[1]
                paths      = $paths
            })
        }
        elseif ($section -eq "Review Notes" -and $line -match '^\s*-\s' -and $line -match '(?i)\b(excluded|skipped)\b') {
            foreach ($match in [regex]::Matches($line, '\bgrp-[0-9a-f]{12}\b')) {
                $skippedGroupIds.Add($match.Value)
            }
            foreach ($match in [regex]::Matches($line, '\bfile-[0-9a-f]{12}\b')) {
                $excludedFileIds.Add($match.Value)
            }
        }
    }

    return [pscustomobject]@{
        sha256          = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
        generated       = $generated
        sources         = $sources.ToArray()
        allPaths        = @($allPaths | Select-Object -Unique)
        skippedGroupIds = @($skippedGroupIds | Select-Object -Unique)
        excludedFileIds = @($excludedFileIds | Select-Object -Unique)
    }
}

function Get-DecisionForFile {
    param(
        [pscustomobject]$Decisions,
        [string]$CandidateId
    )

    if ($null -eq $Decisions -or $null -eq $Decisions.fileDecisions) {
        return $null
    }
    return $Decisions.fileDecisions | Where-Object { $_.candidateId -eq $CandidateId } | Select-Object -First 1
}

$missingInputs = [System.Collections.Generic.List[string]]::new()
if ([string]::IsNullOrWhiteSpace($CustomerName)) {
    if ($NonInteractive) { $missingInputs.Add("CustomerName") } else { $CustomerName = Read-Host "Customer name" }
}
if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
    if ($NonInteractive) { $missingInputs.Add("SourceFolder") } else { $SourceFolder = Read-Host "Source folder" }
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    if ($NonInteractive) { $missingInputs.Add("OutputPath") } else { $OutputPath = Read-Host "Markdown output path" }
}
if ([string]::IsNullOrWhiteSpace($SelectionMode)) {
    if ($NonInteractive) {
        $missingInputs.Add("SelectionMode")
    }
    else {
        $SelectionMode = Read-Host "Selection mode (Auto or Prompt)"
        if ($SelectionMode -notin @("Auto", "Prompt")) {
            $SelectionMode = "Auto"
        }
    }
}

if ($missingInputs.Count -gt 0) {
    $result = New-BaseResult -Status "needs_input"
    $result.warnings = @("Missing required inputs: $($missingInputs -join ', ')")
    Write-ResultAndExit -Result $result -ExitCode 2
}

try {
    $SourceFolder = (Resolve-Path -LiteralPath $SourceFolder -ErrorAction Stop).Path
    if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
        throw "SourceFolder is not a directory: $SourceFolder"
    }

    if ([System.IO.Path]::GetExtension($OutputPath) -ne ".md") {
        throw "OutputPath must end in .md."
    }
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    $outputParent = [System.IO.Path]::GetDirectoryName($OutputPath)
    if ([string]::IsNullOrWhiteSpace($outputParent) -or -not (Test-Path -LiteralPath $outputParent -PathType Container)) {
        throw "The output parent directory does not exist: $outputParent"
    }

    $decisions = Read-Decisions
    $allowIncomplete = [bool]$AllowIncompleteCoverage
    if ($null -ne $decisions -and $decisions.allowIncompleteCoverage) {
        $allowIncomplete = $true
    }

    $overwriteApproved = [bool]$Force
    if ($null -ne $decisions -and $decisions.overwriteOutput) {
        $overwriteApproved = $true
    }
    $existing = $null
    if (Test-Path -LiteralPath $OutputPath -PathType Container) {
        throw "OutputPath is a directory: $OutputPath"
    }
    if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
        if ($overwriteApproved) {
            $mode = "replace"
        }
        else {
            $mode = "update"
            $existing = Read-ExistingReport -Path $OutputPath
            if ($existing.sources.Count -eq 0) {
                $result = New-BaseResult -Status "needs_input"
                $result.sourceFolder = $SourceFolder
                $result.output.path = $OutputPath
                $result.warnings = @("Existing output has no recognizable Sources table, so it cannot be updated: $OutputPath. Choose another output path or authorize replacement with -Force.")
                Write-ResultAndExit -Result $result -ExitCode 2
            }
        }
    }
    $previousIds = if ($null -ne $existing) { @($existing.sources | ForEach-Object { $_.documentId }) } else { @() }

    $pdfTool = Get-PdfTool
    $names = @($CustomerName) + @($CustomerAliases | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $files = @(Get-ChildItem -LiteralPath $SourceFolder -File -Recurse:$Recurse | Where-Object {
        $_.Extension.ToLowerInvariant() -in @(".docx", ".pdf") -and
        -not $_.Name.StartsWith("~`$") -and
        -not ($_.Attributes -band [System.IO.FileAttributes]::Hidden) -and
        (ConvertTo-NormalizedText -Value $_.BaseName) -notmatch '\b(template|customer engagement summary)\b' -and
        $_.FullName -ne $OutputPath
    })

    if ($files.Count -eq 0) {
        throw "No eligible DOCX or PDF files were found."
    }

    Write-Diagnostic "Scanning $($files.Count) DOCX/PDF files for '$CustomerName'."
    $candidates = [System.Collections.Generic.List[object]]::new()
    $failedFiles = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    foreach ($file in $files) {
        $candidateId = Get-StableId -Prefix "file" -Value $file.FullName
        try {
            $extraction = if ($file.Extension.ToLowerInvariant() -eq ".docx") {
                Get-DocxExtraction -Path $file.FullName
            }
            else {
                Get-PdfExtraction -Path $file.FullName -PdfTool $pdfTool
            }

            $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            $customerMatch = Get-CustomerMatch -File $file -Extraction $extraction -Names $names
            if ($customerMatch.score -lt 20) {
                continue
            }

            $signature = Get-SignatureAssessment -File $file -Extraction $extraction
            $date = Get-DateAssessment -Extraction $extraction
            $agreementType = Get-AgreementType -File $file -Extraction $extraction
            $groupKey = Get-GroupKey -File $file -Extraction $extraction -AgreementType $agreementType
            $groupId = Get-StableId -Prefix "grp" -Value $groupKey
            $documentId = Get-StableId -Prefix "doc" -Value "$hash|$groupId"

            $candidateWarnings = [System.Collections.Generic.List[string]]::new()
            if ($extraction.trackedChanges) {
                $candidateWarnings.Add("Tracked changes were detected; deleted text was excluded.")
            }
            if ($extraction.mediaCount -gt 0) {
                $candidateWarnings.Add("Embedded images were detected and were not interpreted as signatures.")
            }

            $candidates.Add([pscustomobject]@{
                candidateId        = $candidateId
                documentId         = $documentId
                groupId            = $groupId
                groupKey           = $groupKey
                path               = $file.FullName
                sourcePaths        = @($file.FullName)
                sha256             = $hash
                title              = if ($extraction.title) { $extraction.title } else { $file.BaseName }
                agreementType      = $agreementType
                parentGroupId      = $null
                modifiedUtc        = $file.LastWriteTimeUtc.ToString("o")
                customerScore      = $customerMatch.score
                matchEvidence      = @($customerMatch.evidence)
                signatureStatus    = $signature.status
                signatureScore     = $signature.score
                selectionEvidence  = @($signature.evidence)
                effectiveDate      = $date.normalized
                effectiveDateText  = $date.original
                effectiveDateWhere = $date.locator
                extraction         = [pscustomobject]@{
                    status    = $extraction.status
                    extractor = $extraction.extractor
                    version   = $extraction.version
                }
                sections           = @($extraction.sections)
                warnings           = $candidateWarnings.ToArray()
            })
        }
        catch {
            $pathText = ConvertTo-NormalizedText -Value "$($file.DirectoryName) $($file.BaseName)"
            $pathMatch = $false
            foreach ($name in $names) {
                if ($pathText.Contains((ConvertTo-NormalizedText -Value $name))) {
                    $pathMatch = $true
                    break
                }
            }
            $failedFiles.Add([pscustomobject]@{
                candidateId = $candidateId
                path        = $file.FullName
                reason      = $_.Exception.Message
                relevant    = $pathMatch
                groupId     = $null
            })
        }
    }

    if ($candidates.Count -eq 0 -and $failedFiles.Count -eq 0) {
        throw "No documents reliably matched the requested customer."
    }

    $deduplicated = [System.Collections.Generic.List[object]]::new()
    foreach ($hashGroup in $candidates | Group-Object sha256) {
        $primary = $hashGroup.Group | Sort-Object customerScore -Descending | Select-Object -First 1
        $primary.sourcePaths = @($hashGroup.Group.path | Sort-Object -Unique)
        $deduplicated.Add($primary)
    }
    $candidates = $deduplicated

    $baseGroups = @($candidates | Where-Object { $_.agreementType -eq "SOW" } | Select-Object -ExpandProperty groupId -Unique)
    if ($baseGroups.Count -eq 1) {
        foreach ($candidate in $candidates | Where-Object { $_.agreementType -ne "SOW" }) {
            $candidate.parentGroupId = $baseGroups[0]
        }
    }
    elseif (@($candidates | Where-Object { $_.agreementType -ne "SOW" }).Count -gt 0) {
        $warnings.Add("One or more amendments/change orders could not be linked to a unique parent SOW.")
    }

    $selected = [System.Collections.Generic.List[object]]::new()
    $rejected = [System.Collections.Generic.List[object]]::new()
    $unresolved = [System.Collections.Generic.List[object]]::new()
    $knownDocumentIds = @($candidates | ForEach-Object { $_.documentId })
    $knownGroupIds = @($candidates | ForEach-Object { $_.groupId } | Select-Object -Unique)

    if ($null -ne $decisions) {
        foreach ($decision in @($decisions.groupDecisions)) {
            if ($decision.groupId -notin $knownGroupIds) {
                throw "Decision file contains an unknown groupId: $($decision.groupId)"
            }
            if ($decision.action -eq "select" -and $decision.documentId -notin $knownDocumentIds) {
                throw "Decision file contains an unknown documentId: $($decision.documentId)"
            }
        }
    }

    foreach ($group in $candidates | Group-Object groupId) {
        $ranked = @($group.Group | Sort-Object @{
            Expression = {
                $signatureRank = switch ($_.signatureStatus) {
                    "Execution evidence detected" { 1000 }
                    "Partially signed" { 500 }
                    "Unverified" { 100 }
                    default { 0 }
                }
                $signatureRank + $_.signatureScore + $_.customerScore
            }
            Descending = $true
        }, @{ Expression = "modifiedUtc"; Descending = $true })

        $groupId = $group.Name
        $decision = Get-DecisionForGroup -Decisions $decisions -GroupId $groupId
        $chosen = $null
        $selectionMethod = $null

        $previous = $null
        $keepPrevious = $false
        if ($mode -eq "update") {
            $previous = $ranked | Where-Object { $_.documentId -in $previousIds } | Select-Object -First 1
            if ($null -ne $previous) {
                $previousRank = Get-SignatureRank -Status $previous.signatureStatus
                $revisions = @($ranked | Where-Object {
                    $_.documentId -ne $previous.documentId -and
                    (Get-SignatureRank -Status $_.signatureStatus) -ge [Math]::Max(1, $previousRank)
                })
                $keepPrevious = $revisions.Count -eq 0
            }
        }

        if ($null -ne $decision) {
            if ($decision.action -eq "select") {
                $chosen = $ranked | Where-Object { $_.documentId -eq $decision.documentId } | Select-Object -First 1
                if ($null -eq $chosen) {
                    throw "Selected documentId $($decision.documentId) is not in group $groupId."
                }
                $selectionMethod = "User selected"
            }
            elseif ($decision.action -eq "skip") {
                if ([string]::IsNullOrWhiteSpace($decision.reason)) {
                    throw "Skipping group $groupId requires a reason."
                }
                if (-not $allowIncomplete) {
                    throw "Skipping group $groupId requires allowIncompleteCoverage."
                }
                $unresolved.Add([pscustomobject]@{
                    groupId   = $groupId
                    action    = "skipped"
                    reason    = $decision.reason
                    candidates = @($ranked | Select-Object documentId, path, title, signatureStatus)
                })
                continue
            }
            else {
                throw "Unsupported group decision action '$($decision.action)' for $groupId."
            }
        }
        elseif ($keepPrevious) {
            $chosen = $previous
            $selectionMethod = "Previously reported"
        }
        elseif ($mode -eq "update" -and $null -eq $previous -and $groupId -in $existing.skippedGroupIds) {
            $unresolved.Add([pscustomobject]@{
                groupId        = $groupId
                action         = "skipped"
                reason         = "Carried forward: the existing report records this group as skipped or excluded."
                carriedForward = $true
                candidates     = @($ranked | Select-Object documentId, path, title, signatureStatus)
            })
            continue
        }
        elseif ($SelectionMode -eq "Auto") {
            $top = $ranked[0]
            $runnerUp = if ($ranked.Count -gt 1) { $ranked[1] } else { $null }
            $uniqueExecutionWinner = $top.signatureStatus -eq "Execution evidence detected" -and
                ($null -eq $runnerUp -or $runnerUp.signatureStatus -ne "Execution evidence detected")
            if ($uniqueExecutionWinner) {
                $chosen = $top
                $selectionMethod = "Auto"
            }
        }
        elseif (-not $NonInteractive) {
            Write-Diagnostic ""
            Write-Diagnostic "Select a candidate for group ${groupId}:"
            for ($index = 0; $index -lt $ranked.Count; $index++) {
                Write-Diagnostic "[$($index + 1)] $($ranked[$index].path) - $($ranked[$index].signatureStatus)"
            }
            Write-Diagnostic "[S] Skip this group"
            $choice = Read-Host "Selection"
            if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $ranked.Count) {
                $chosen = $ranked[[int]$choice - 1]
                $selectionMethod = "User selected"
            }
            elseif ($choice -match '^(s|skip)$') {
                $reason = Read-Host "Reason for skipping"
                if (-not $allowIncomplete) {
                    throw "Skipping requires -AllowIncompleteCoverage."
                }
                $unresolved.Add([pscustomobject]@{
                    groupId    = $groupId
                    action     = "skipped"
                    reason     = $reason
                    candidates = @($ranked | Select-Object documentId, path, title, signatureStatus)
                })
                continue
            }
        }

        if ($null -eq $chosen) {
            $unresolvedReason = if ($null -ne $previous) {
                "A new candidate may revise the previously reported document $($previous.documentId). Choose the previous document to keep it or a new one to supersede it."
            }
            else {
                "No unique high-confidence executed candidate was selected."
            }
            $unresolved.Add([pscustomobject]@{
                groupId    = $groupId
                action     = "unresolved"
                reason     = $unresolvedReason
                candidates = @($ranked | Select-Object documentId, path, title, agreementType, signatureStatus, signatureScore, customerScore, selectionEvidence)
            })
            continue
        }

        $confidence = if ($chosen.signatureStatus -eq "Execution evidence detected" -and $selectionMethod -eq "Auto") {
            "High"
        }
        elseif ($chosen.signatureStatus -eq "Execution evidence detected") {
            "Medium"
        }
        else {
            "Low"
        }

        $updateStatus = $null
        $supersedes = $null
        $reportSequence = $null
        if ($mode -eq "update") {
            if ($chosen.documentId -in $previousIds) {
                $updateStatus = "unchanged"
                $reportSequence = ($existing.sources | Where-Object { $_.documentId -eq $chosen.documentId } | Select-Object -First 1).sequence
            }
            elseif ($null -ne $previous) {
                $updateStatus = "supersedes"
                $supersedes = $previous.documentId
                $reportSequence = ($existing.sources | Where-Object { $_.documentId -eq $previous.documentId } | Select-Object -First 1).sequence
            }
            else {
                $updateStatus = "new"
            }
        }

        $sections = @($chosen.sections)
        if ($updateStatus -eq "unchanged") {
            $sections = @()
        }

        $selected.Add([pscustomobject]@{
            documentId         = $chosen.documentId
            groupId            = $chosen.groupId
            updateStatus       = $updateStatus
            supersedesDocumentId = $supersedes
            reportSequence     = $reportSequence
            sourcePaths        = @($chosen.sourcePaths)
            sha256             = $chosen.sha256
            agreementType      = $chosen.agreementType
            parentGroupId      = $chosen.parentGroupId
            title              = $chosen.title
            dates              = [pscustomobject]@{
                effective         = $chosen.effectiveDate
                effectiveOriginal = $chosen.effectiveDateText
                effectiveLocator  = $chosen.effectiveDateWhere
                signed            = $null
                plannedStart      = $null
                plannedEnd        = $null
            }
            signatureStatus     = $chosen.signatureStatus
            selectionConfidence = $confidence
            selectionMethod     = $selectionMethod
            matchEvidence       = @($chosen.matchEvidence)
            selectionEvidence   = @($chosen.selectionEvidence)
            extraction          = $chosen.extraction
            sections            = $sections
            warnings            = @($chosen.warnings)
        })

        foreach ($candidate in $ranked | Where-Object { $_.documentId -ne $chosen.documentId }) {
            $rejected.Add([pscustomobject]@{
                candidateId = $candidate.candidateId
                documentId  = $candidate.documentId
                groupId     = $candidate.groupId
                path        = $candidate.path
                reason      = "A different candidate was selected for this agreement group."
                signatureStatus = $candidate.signatureStatus
            })
        }
    }

    $excludedFiles = [System.Collections.Generic.List[object]]::new()
    $blockingFailures = [System.Collections.Generic.List[object]]::new()
    foreach ($failure in $failedFiles) {
        $fileDecision = Get-DecisionForFile -Decisions $decisions -CandidateId $failure.candidateId
        if ($null -ne $fileDecision -and $fileDecision.action -eq "exclude") {
            if ([string]::IsNullOrWhiteSpace($fileDecision.reason)) {
                throw "Excluding file $($failure.candidateId) requires a reason."
            }
            if (-not $allowIncomplete) {
                throw "Excluding file $($failure.candidateId) requires allowIncompleteCoverage."
            }
            $excludedFiles.Add([pscustomobject]@{
                candidateId = $failure.candidateId
                path        = $failure.path
                reason      = $fileDecision.reason
            })
        }
        elseif ($mode -eq "update" -and $failure.candidateId -in $existing.excludedFileIds) {
            $excludedFiles.Add([pscustomobject]@{
                candidateId    = $failure.candidateId
                path           = $failure.path
                reason         = "Carried forward: the existing report records this file as excluded."
                carriedForward = $true
            })
        }
        elseif ($failure.relevant) {
            $blockingFailures.Add($failure)
        }
    }

    $chronology = @{}
    if ($null -ne $decisions -and $null -ne $decisions.chronology) {
        foreach ($item in $decisions.chronology) {
            if ($item.groupId -notin @($selected | ForEach-Object { $_.groupId })) {
                throw "Chronology contains an unknown or unselected groupId: $($item.groupId)"
            }
            if ([int]$item.order -lt 1 -or $chronology.ContainsValue([int]$item.order)) {
                throw "Chronology order values must be unique positive integers."
            }
            $chronology[$item.groupId] = [int]$item.order
        }
    }

    if ($selected.Count -gt 1 -and $chronology.Count -ne $selected.Count) {
        $knownDates = @($selected | ForEach-Object { $_.dates.effective })
        $changed = @($selected | Where-Object { $_.updateStatus -ne "unchanged" })
        $hasMissingDate = @($changed | Where-Object { [string]::IsNullOrWhiteSpace($_.dates.effective) }).Count -gt 0
        $duplicateDates = @($knownDates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        $hasDuplicateDate = @($changed | Where-Object { $_.dates.effective -in $duplicateDates }).Count -gt 0
        if ($hasMissingDate -or $hasDuplicateDate) {
            $unresolved.Add([pscustomobject]@{
                groupId    = "chronology"
                action     = "unresolved"
                reason     = "Agreement order cannot be established from unique effective dates. Supply chronology decisions for every selected group."
                candidates = @($selected | Select-Object groupId, documentId, title, agreementType, @{ Name = "effectiveDate"; Expression = { $_.dates.effective } })
            })
        }
    }

    $activeUnresolved = @($unresolved | Where-Object { $_.action -ne "skipped" })
    $requiresInput = $activeUnresolved.Count -gt 0 -or $blockingFailures.Count -gt 0
    if ($selected.Count -eq 0 -and -not $requiresInput) {
        throw "No agreements were selected for the report."
    }

    $orderedSelected = @($selected | Sort-Object @{
        Expression = {
            if ($chronology.ContainsKey($_.groupId)) { $chronology[$_.groupId] } else { [int]::MaxValue }
        }
    }, @{
        Expression = {
            if ($_.dates.effective) { $_.dates.effective } else { "9999-12-31" }
        }
    }, @{
        Expression = {
            if ($null -ne $_.reportSequence) { [int]$_.reportSequence } else { [int]::MaxValue }
        }
    }, groupId)

    $result = New-BaseResult -Status $(if ($requiresInput) { "needs_input" } else { "ready" })
    $result.customer.canonicalName = $CustomerName
    $result.sourceFolder = $SourceFolder
    $result.output = [ordered]@{
        path                    = $OutputPath
        force                   = $overwriteApproved
        allowIncompleteCoverage = $allowIncomplete
    }
    $result.coverage = [ordered]@{
        enumerated = $files.Count
        selected   = $orderedSelected.Count
        excluded   = $excludedFiles.Count + @($unresolved | Where-Object { $_.action -eq "skipped" }).Count
        failed     = $blockingFailures.Count
        unresolved = $activeUnresolved.Count
    }
    $result.selectedDocuments = $orderedSelected
    $result.rejectedCandidates = $rejected.ToArray()
    $result.excludedFiles = $excludedFiles.ToArray()
    $result.failedFiles = $blockingFailures.ToArray()
    $result.unresolvedGroups = $unresolved.ToArray()
    $result.warnings = $warnings.ToArray()

    if ($mode -eq "update") {
        $missing = @($existing.sources | Where-Object { $_.documentId -notin $knownDocumentIds } | ForEach-Object {
            [pscustomobject]@{
                documentId     = $_.documentId
                sequence       = $_.sequence
                paths          = @($_.paths)
                pathsStillExist = @($_.paths | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }).Count -gt 0
            }
        })
        $unchanged = @($orderedSelected | Where-Object { $_.updateStatus -eq "unchanged" })
        $newDocuments = @($orderedSelected | Where-Object { $_.updateStatus -eq "new" })
        $superseding = @($orderedSelected | Where-Object { $_.updateStatus -eq "supersedes" })
        $newCopies = @($unchanged | ForEach-Object {
            $document = $_
            $added = @($document.sourcePaths | Where-Object { $_ -notin $existing.allPaths })
            if ($added.Count -gt 0) {
                [pscustomobject]@{ documentId = $document.documentId; newSourcePaths = $added }
            }
        })
        foreach ($entry in $missing) {
            $result.warnings += "Previously reported document $($entry.documentId) was not found in the source folder$(if ($entry.pathsStillExist) { '; its file exists at the old path but its content changed' } else { '' }). It is preserved in the report."
        }
        $result.update = [ordered]@{
            existingReport = [ordered]@{
                path            = $OutputPath
                sha256          = $existing.sha256
                reportGenerated = $existing.generated
                sourceCount     = $existing.sources.Count
            }
            backupPath          = "$OutputPath.bak-$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))"
            hasChanges          = ($newDocuments.Count + $superseding.Count + $newCopies.Count) -gt 0
            newDocumentIds      = @($newDocuments | ForEach-Object { $_.documentId })
            supersedingDocumentIds = @($superseding | ForEach-Object { $_.documentId })
            unchangedDocumentIds = @($unchanged | ForEach-Object { $_.documentId })
            unchangedWithNewSourcePaths = $newCopies
            missingFromFolder   = $missing
            carriedForward      = [ordered]@{
                groupIds = @($unresolved | Where-Object { $_.PSObject.Properties["carriedForward"] } | ForEach-Object { $_.groupId })
                fileIds  = @($excludedFiles | Where-Object { $_.PSObject.Properties["carriedForward"] } | ForEach-Object { $_.candidateId })
            }
        }
    }

    Write-ResultAndExit -Result $result -ExitCode $(if ($requiresInput) { 2 } else { 0 })
}
catch {
    $result = New-BaseResult -Status "failed"
    $result.warnings = @($_.Exception.Message)
    Write-ResultAndExit -Result $result -ExitCode 1
}
