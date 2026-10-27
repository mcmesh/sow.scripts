#Requires -Version 7.0

[CmdletBinding()]
param()

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$skillRoot = Split-Path -Parent $PSScriptRoot
$helperPath = Join-Path $PSScriptRoot "New-SowCustomerSummaryInput.ps1"
$fixtureGenerator = Join-Path $skillRoot "fixtures/New-SowFixtures.ps1"
$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) "sow-skill-$([guid]::NewGuid().ToString('N'))"

function Assert-Equal {
    param(
        $Actual,
        $Expected,
        [string]$Message
    )
    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Invoke-HelperProcess {
    param([string[]]$Arguments)

    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = (Get-Command pwsh -CommandType Application | Select-Object -First 1).Source
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    [void]$startInfo.ArgumentList.Add("-NoProfile")
    [void]$startInfo.ArgumentList.Add("-File")
    [void]$startInfo.ArgumentList.Add($helperPath)
    foreach ($argument in $Arguments) {
        [void]$startInfo.ArgumentList.Add($argument)
    }

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    [void]$process.Start()
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    $json = $null
    if (-not [string]::IsNullOrWhiteSpace($stdout)) {
        $json = $stdout | ConvertFrom-Json
    }
    return [pscustomobject]@{
        exitCode = $process.ExitCode
        stdout   = $stdout
        stderr   = $stderr
        json     = $json
    }
}

try {
    [System.IO.Directory]::CreateDirectory($temporaryRoot) | Out-Null
    $fixtureJson = & $fixtureGenerator -DestinationPath (Join-Path $temporaryRoot "fixtures")
    $fixtures = $fixtureJson | ConvertFrom-Json
    $outputDirectory = [System.IO.Directory]::CreateDirectory((Join-Path $temporaryRoot "output")).FullName

    $readyResult = Invoke-HelperProcess -Arguments @(
        "-CustomerName", "Contosó",
        "-SourceFolder", $fixtures.ready,
        "-OutputPath", (Join-Path $outputDirectory "contoso.md"),
        "-SelectionMode", "Auto",
        "-NonInteractive"
    )
    Assert-Equal $readyResult.exitCode 0 "Ready fixture exit code."
    Assert-Equal $readyResult.json.status "ready" "Ready fixture status."
    Assert-Equal $readyResult.json.coverage.selected 2 "Ready fixture selection count."
    Assert-True (@($readyResult.json.selectedDocuments | Where-Object { $_.agreementType -eq "SOW" }).Count -eq 1) "Expected one selected base SOW."
    Assert-True (@($readyResult.json.selectedDocuments | Where-Object { $_.agreementType -eq "Change order" }).Count -eq 1) "Expected one selected change order."
    $baseSow = $readyResult.json.selectedDocuments | Where-Object { $_.agreementType -eq "SOW" } | Select-Object -First 1
    Assert-Equal @($baseSow.sourcePaths).Count 2 "Hash-identical copies should preserve both paths."
    Assert-Equal $baseSow.signatureStatus "Execution evidence detected" "Signed SOW status."
    Assert-True (@($baseSow.warnings | Where-Object { $_ -match "Tracked changes" }).Count -eq 1) "Tracked changes warning was not retained."

    $ambiguousResult = Invoke-HelperProcess -Arguments @(
        "-CustomerName", "Contoso",
        "-SourceFolder", $fixtures.ambiguous,
        "-OutputPath", (Join-Path $outputDirectory "ambiguous.md"),
        "-SelectionMode", "Auto",
        "-NonInteractive"
    )
    Assert-Equal $ambiguousResult.exitCode 2 "Ambiguous fixture exit code."
    Assert-Equal $ambiguousResult.json.status "needs_input" "Ambiguous fixture status."
    Assert-Equal $ambiguousResult.json.coverage.unresolved 1 "Ambiguous group count."

    $group = $ambiguousResult.json.unresolvedGroups[0]
    $selectedCandidate = $group.candidates[0]
    $decisionPath = Join-Path $temporaryRoot "decisions.json"
    @{
        schemaVersion = "1.0"
        allowIncompleteCoverage = $false
        overwriteOutput = $false
        groupDecisions = @(
            @{
                groupId = $group.groupId
                action = "select"
                documentId = $selectedCandidate.documentId
            }
        )
        fileDecisions = @()
        chronology = @(@{ groupId = $group.groupId; order = 1 })
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $decisionPath -Encoding utf8

    $decidedResult = Invoke-HelperProcess -Arguments @(
        "-CustomerName", "Contoso",
        "-SourceFolder", $fixtures.ambiguous,
        "-OutputPath", (Join-Path $outputDirectory "ambiguous.md"),
        "-SelectionMode", "Auto",
        "-NonInteractive",
        "-DecisionFile", $decisionPath
    )
    Assert-Equal $decidedResult.exitCode 0 "Decided fixture exit code."
    Assert-Equal $decidedResult.json.status "ready" "Decided fixture status."
    Assert-Equal $decidedResult.json.selectedDocuments[0].selectionMethod "User selected" "Decision selection method."

    $chronologyResult = Invoke-HelperProcess -Arguments @(
        "-CustomerName", "Contoso",
        "-SourceFolder", $fixtures.chronology,
        "-OutputPath", (Join-Path $outputDirectory "chronology.md"),
        "-SelectionMode", "Auto",
        "-NonInteractive"
    )
    Assert-Equal $chronologyResult.exitCode 2 "Chronology fixture exit code."
    Assert-Equal $chronologyResult.json.status "needs_input" "Chronology fixture status."
    $chronologyIssue = $chronologyResult.json.unresolvedGroups | Where-Object { $_.groupId -eq "chronology" } | Select-Object -First 1
    Assert-True ($null -ne $chronologyIssue) "Missing dates should require a chronology decision."

    $chronologyDecisionPath = Join-Path $temporaryRoot "chronology-decisions.json"
    $chronologyOrder = 0
    @{
        schemaVersion = "1.0"
        allowIncompleteCoverage = $false
        overwriteOutput = $false
        groupDecisions = @()
        fileDecisions = @()
        chronology = @($chronologyResult.json.selectedDocuments | ForEach-Object {
            $chronologyOrder++
            @{ groupId = $_.groupId; order = $chronologyOrder }
        })
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $chronologyDecisionPath -Encoding utf8

    $orderedResult = Invoke-HelperProcess -Arguments @(
        "-CustomerName", "Contoso",
        "-SourceFolder", $fixtures.chronology,
        "-OutputPath", (Join-Path $outputDirectory "chronology.md"),
        "-SelectionMode", "Auto",
        "-NonInteractive",
        "-DecisionFile", $chronologyDecisionPath
    )
    Assert-Equal $orderedResult.exitCode 0 "Chronology decision exit code."
    Assert-Equal $orderedResult.json.status "ready" "Chronology decision status."

    $pdfTool = Get-Command pdftotext -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -ne $pdfTool) {
        $pdfSuccessResult = Invoke-HelperProcess -Arguments @(
            "-CustomerName", "Contoso",
            "-SourceFolder", $fixtures.pdfSuccess,
            "-OutputPath", (Join-Path $outputDirectory "pdf-success.md"),
            "-SelectionMode", "Auto",
            "-NonInteractive"
        )
        Assert-Equal $pdfSuccessResult.exitCode 0 "PDF extraction success exit code."
        Assert-Equal $pdfSuccessResult.json.status "ready" "PDF extraction success status."
        Assert-Equal $pdfSuccessResult.json.selectedDocuments[0].extraction.status "complete" "PDF extraction status."
        Assert-True ($pdfSuccessResult.json.selectedDocuments[0].extraction.extractor -match "pdftotext") "PDF extractor path was not retained."

        $pdfFailureResult = Invoke-HelperProcess -Arguments @(
            "-CustomerName", "Contoso",
            "-SourceFolder", $fixtures.pdfFailure,
            "-OutputPath", (Join-Path $outputDirectory "pdf-failure.md"),
            "-SelectionMode", "Auto",
            "-NonInteractive"
        )
        Assert-Equal $pdfFailureResult.exitCode 2 "Invalid PDF exit code."
        Assert-Equal $pdfFailureResult.json.coverage.failed 1 "Invalid PDF failure count."
    }
    else {
        $pdfResult = Invoke-HelperProcess -Arguments @(
            "-CustomerName", "Contoso",
            "-SourceFolder", $fixtures.pdfSuccess,
            "-OutputPath", (Join-Path $outputDirectory "pdf.md"),
            "-SelectionMode", "Auto",
            "-NonInteractive"
        )
        Assert-Equal $pdfResult.exitCode 2 "Missing PDF extractor exit code."
        Assert-Equal $pdfResult.json.coverage.failed 1 "Missing PDF extractor failure count."
        Assert-True ($pdfResult.json.failedFiles[0].reason -match "pdftotext") "Missing PDF extractor reason was not preserved."
    }

    $missingResult = Invoke-HelperProcess -Arguments @("-NonInteractive")
    Assert-Equal $missingResult.exitCode 2 "Missing-input exit code."
    Assert-Equal $missingResult.json.status "needs_input" "Missing-input status."

    function New-ReportStub {
        param(
            [string]$Path,
            [string[]]$DocumentIds,
            [string[]]$ReviewNotes = @()
        )
        $lines = @("# Customer Engagement Summary: Contoso", "", "| Report detail | Value |", "|---|---|", "| Report generated | 2026-01-01 |", "", "## Sources", "", "| Sequence / source ID | Agreement | Source path |", "|---|---|---|")
        for ($index = 0; $index -lt $DocumentIds.Count; $index++) {
            $lines += "| $($index + 1) / $($DocumentIds[$index]) | Agreement | ``/old/path-$index.docx`` |"
        }
        $lines += @("", "## Review Notes", "")
        $lines += $ReviewNotes
        Set-Content -LiteralPath $Path -Value $lines -Encoding utf8
    }

    $baseDocument = $readyResult.json.selectedDocuments | Where-Object { $_.agreementType -eq "SOW" } | Select-Object -First 1
    $changeOrder = $readyResult.json.selectedDocuments | Where-Object { $_.agreementType -eq "Change order" } | Select-Object -First 1
    $updateArguments = @(
        "-CustomerName", "Contosó",
        "-SourceFolder", $fixtures.ready,
        "-SelectionMode", "Auto",
        "-NonInteractive"
    )

    $updatePath = Join-Path $outputDirectory "update.md"
    New-ReportStub -Path $updatePath -DocumentIds @($baseDocument.documentId)
    $originalReport = Get-Content -LiteralPath $updatePath -Raw
    $updateResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath))
    Assert-Equal $updateResult.exitCode 0 "Update mode exit code."
    Assert-Equal $updateResult.json.mode "update" "Existing report should trigger update mode."
    Assert-True $updateResult.json.update.hasChanges "Update should detect the new change order."
    Assert-Equal @($updateResult.json.update.newDocumentIds).Count 1 "Update new document count."
    Assert-Equal $updateResult.json.update.newDocumentIds[0] $changeOrder.documentId "Update new document id."
    Assert-Equal @($updateResult.json.update.unchangedDocumentIds).Count 1 "Update unchanged document count."
    $unchangedDocument = $updateResult.json.selectedDocuments | Where-Object { $_.documentId -eq $baseDocument.documentId }
    Assert-Equal $unchangedDocument.updateStatus "unchanged" "Previously reported document status."
    Assert-Equal $unchangedDocument.selectionMethod "Previously reported" "Previously reported selection method."
    Assert-Equal @($unchangedDocument.sections).Count 0 "Unchanged documents should omit extracted sections."
    Assert-True ($updateResult.json.update.backupPath -like "$updatePath.bak-*") "Backup path was not suggested."
    Assert-Equal (Get-Content -LiteralPath $updatePath -Raw) $originalReport "The helper must not modify the existing report."

    New-ReportStub -Path $updatePath -DocumentIds @($baseDocument.documentId, $changeOrder.documentId) -ReviewNotes @("- Additional source paths: " + ((@($baseDocument.sourcePaths) + @($changeOrder.sourcePaths) | ForEach-Object { "``$_``" }) -join ", "))
    $noChangeResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath))
    Assert-Equal $noChangeResult.exitCode 0 "No-change update exit code."
    Assert-True (-not $noChangeResult.json.update.hasChanges) "A complete report should report no changes."

    New-ReportStub -Path $updatePath -DocumentIds @($baseDocument.documentId, "doc-000000000000")
    $missingDocumentResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath))
    Assert-Equal @($missingDocumentResult.json.update.missingFromFolder).Count 1 "Missing previously reported document count."
    Assert-Equal $missingDocumentResult.json.update.missingFromFolder[0].documentId "doc-000000000000" "Missing document id."

    New-ReportStub -Path $updatePath -DocumentIds @($baseDocument.documentId) -ReviewNotes @("- Helper group ``$($changeOrder.groupId)`` was skipped by the user.")
    $carriedResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath))
    Assert-Equal $carriedResult.exitCode 0 "Carried-forward skip exit code."
    Assert-Equal @($carriedResult.json.update.newDocumentIds).Count 0 "A skipped group must not reappear as new."
    Assert-Equal $carriedResult.json.update.carriedForward.groupIds[0] $changeOrder.groupId "Carried-forward group id."

    Set-Content -LiteralPath $updatePath -Value "Not a report." -Encoding utf8
    $unrecognizedResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath))
    Assert-Equal $unrecognizedResult.exitCode 2 "Unrecognized existing report exit code."

    $replaceResult = Invoke-HelperProcess -Arguments ($updateArguments + @("-OutputPath", $updatePath, "-Force"))
    Assert-Equal $replaceResult.exitCode 0 "Replace mode exit code."
    Assert-Equal $replaceResult.json.mode "replace" "Force should select replace mode."

    $revisionPath = Join-Path $outputDirectory "revision.md"
    $previousCandidate = $group.candidates[0]
    $newCandidate = $group.candidates[1]
    New-ReportStub -Path $revisionPath -DocumentIds @($previousCandidate.documentId)
    $revisionArguments = @(
        "-CustomerName", "Contoso",
        "-SourceFolder", $fixtures.ambiguous,
        "-OutputPath", $revisionPath,
        "-SelectionMode", "Auto",
        "-NonInteractive"
    )
    $revisionResult = Invoke-HelperProcess -Arguments $revisionArguments
    Assert-Equal $revisionResult.exitCode 2 "A competing executed revision should require a decision."
    Assert-True ($revisionResult.json.unresolvedGroups[0].reason -match "revise") "Revision reason was not reported."

    $supersedePath = Join-Path $temporaryRoot "supersede-decisions.json"
    @{
        schemaVersion = "1.0"
        allowIncompleteCoverage = $false
        overwriteOutput = $false
        groupDecisions = @(@{ groupId = $group.groupId; action = "select"; documentId = $newCandidate.documentId })
        fileDecisions = @()
        chronology = @(@{ groupId = $group.groupId; order = 1 })
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $supersedePath -Encoding utf8
    $supersedeResult = Invoke-HelperProcess -Arguments ($revisionArguments + @("-DecisionFile", $supersedePath))
    Assert-Equal $supersedeResult.exitCode 0 "Supersede decision exit code."
    Assert-Equal $supersedeResult.json.selectedDocuments[0].updateStatus "supersedes" "Superseding revision status."
    Assert-Equal $supersedeResult.json.selectedDocuments[0].supersedesDocumentId $previousCandidate.documentId "Superseded document id."

    Write-Host "All SOW customer summary checks passed."
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
