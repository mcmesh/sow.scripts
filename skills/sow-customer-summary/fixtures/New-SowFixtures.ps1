#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DestinationPath
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

function ConvertTo-XmlEscapedText {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function Add-ZipTextEntry {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$Name,
        [string]$Content
    )

    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($Content)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function New-SyntheticDocx {
    param(
        [string]$Path,
        [string]$Title,
        [string[]]$Paragraphs,
        [switch]$TrackedChange
    )

    $paragraphXml = foreach ($paragraph in $Paragraphs) {
        $escaped = ConvertTo-XmlEscapedText -Value $paragraph
        "<w:p><w:r><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>"
    }
    if ($TrackedChange) {
        $paragraphXml += '<w:p><w:del><w:r><w:delText>Deleted draft language</w:delText></w:r></w:del><w:ins><w:r><w:t>Approved replacement language</w:t></w:r></w:ins></w:p>'
    }

    $documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    $($paragraphXml -join [Environment]::NewLine)
  </w:body>
</w:document>
"@
    $escapedTitle = ConvertTo-XmlEscapedText -Value $Title
    $coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>$escapedTitle</dc:title>
</cp:coreProperties>
"@
    $contentTypes = @"
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>
"@

    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipTextEntry -Archive $archive -Name "[Content_Types].xml" -Content $contentTypes
        Add-ZipTextEntry -Archive $archive -Name "word/document.xml" -Content $documentXml
        Add-ZipTextEntry -Archive $archive -Name "docProps/core.xml" -Content $coreXml
    }
    finally {
        $archive.Dispose()
    }
}

function New-SyntheticPdf {
    param(
        [string]$Path,
        [string[]]$Lines
    )

    $escapedLines = foreach ($line in $Lines) {
        $line.Replace('\', '\\').Replace('(', '\(').Replace(')', '\)')
    }
    $textCommands = [System.Collections.Generic.List[string]]::new()
    $textCommands.Add("BT /F1 12 Tf 72 720 Td")
    for ($index = 0; $index -lt $escapedLines.Count; $index++) {
        if ($index -gt 0) {
            $textCommands.Add("0 -18 Td")
        }
        $textCommands.Add("($($escapedLines[$index])) Tj")
    }
    $textCommands.Add("ET")
    $stream = $textCommands -join "`n"

    $objects = @(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        "<< /Length $([System.Text.Encoding]::ASCII.GetByteCount($stream)) >>`nstream`n$stream`nendstream"
    )

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append("%PDF-1.4`n")
    $offsets = [System.Collections.Generic.List[int]]::new()
    for ($index = 0; $index -lt $objects.Count; $index++) {
        $offsets.Add([System.Text.Encoding]::ASCII.GetByteCount($builder.ToString()))
        [void]$builder.Append("$($index + 1) 0 obj`n$($objects[$index])`nendobj`n")
    }

    $xrefOffset = [System.Text.Encoding]::ASCII.GetByteCount($builder.ToString())
    [void]$builder.Append("xref`n0 $($objects.Count + 1)`n")
    [void]$builder.Append("0000000000 65535 f `n")
    foreach ($offset in $offsets) {
        [void]$builder.Append(("{0:D10} 00000 n `n" -f $offset))
    }
    [void]$builder.Append("trailer`n<< /Size $($objects.Count + 1) /Root 1 0 R >>`nstartxref`n$xrefOffset`n%%EOF`n")
    [System.IO.File]::WriteAllText($Path, $builder.ToString(), [System.Text.Encoding]::ASCII)
}

$root = [System.IO.Path]::GetFullPath($DestinationPath)
[System.IO.Directory]::CreateDirectory($root) | Out-Null
$ready = [System.IO.Directory]::CreateDirectory((Join-Path $root "ready")).FullName
$ambiguous = [System.IO.Directory]::CreateDirectory((Join-Path $root "ambiguous")).FullName
$chronology = [System.IO.Directory]::CreateDirectory((Join-Path $root "chronology")).FullName
$pdfSuccess = [System.IO.Directory]::CreateDirectory((Join-Path $root "pdf-success")).FullName
$pdfFailure = [System.IO.Directory]::CreateDirectory((Join-Path $root "pdf-failure")).FullName

$draftPath = Join-Path $ready "Contoso-Implementation-Draft.docx"
New-SyntheticDocx -Path $draftPath -Title "Contoso Implementation" -Paragraphs @(
    "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
    "Effective Date: January 1, 2025",
    "DRAFT FOR REVIEW - NOT FOR EXECUTION",
    "The planned scope includes discovery, configuration, and user enablement."
)

$signedPath = Join-Path $ready "Contoso-Implementation-Signed.docx"
New-SyntheticDocx -Path $signedPath -Title "Contoso Implementation" -TrackedChange -Paragraphs @(
    "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
    "Effective Date: January 1, 2025",
    "FULLY EXECUTED. All parties have signed this agreement.",
    "The contracted scope includes discovery, configuration, data migration, and user enablement.",
    "Deliverables include a design document, configured solution, migration report, and training sessions."
)
Copy-Item -LiteralPath $signedPath -Destination (Join-Path $ready "Contoso-Implementation-Executed-Copy.docx")

New-SyntheticDocx -Path (Join-Path $ready "Contoso-Change-Order-Signed.docx") -Title "Contoso Change Order" -Paragraphs @(
    "Change Order between Contoso Ltd (Customer) and Example Consulting.",
    "Effective Date: February 15, 2025",
    "FULLY EXECUTED. Signing complete for all parties.",
    "This change adds dashboard development and administrator knowledge transfer.",
    "Deliverables include three dashboards and two administrator workshops."
)

New-SyntheticDocx -Path (Join-Path $ready "Fabrikam-Implementation-Signed.docx") -Title "Fabrikam Implementation" -Paragraphs @(
    "Statement of Work between Fabrikam Inc (Customer) and Example Consulting.",
    "Effective Date: January 10, 2025",
    "FULLY EXECUTED. All parties have signed this agreement.",
    "The scope covers an unrelated implementation."
)

foreach ($version in 1..2) {
    New-SyntheticDocx -Path (Join-Path $ambiguous "Contoso-Discovery-Version-$version-Signed.docx") -Title "Contoso Discovery" -Paragraphs @(
        "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
        "Effective Date: March $version, 2025",
        "FULLY EXECUTED. All parties have signed this agreement.",
        "Revision $version includes a materially different discovery scope."
    )
}

New-SyntheticDocx -Path (Join-Path $chronology "Contoso-Foundation-Signed.docx") -Title "Contoso Foundation" -Paragraphs @(
    "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
    "FULLY EXECUTED. All parties have signed this agreement.",
    "The scope establishes the solution foundation."
)
New-SyntheticDocx -Path (Join-Path $chronology "Contoso-Optimization-Signed.docx") -Title "Contoso Optimization" -Paragraphs @(
    "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
    "FULLY EXECUTED. All parties have signed this agreement.",
    "The scope covers solution optimization."
)

New-SyntheticPdf -Path (Join-Path $pdfSuccess "Contoso-PDF-Signed.pdf") -Lines @(
    "Statement of Work between Contoso Ltd (Customer) and Example Consulting.",
    "Effective Date: April 1, 2025",
    "FULLY EXECUTED. All parties have signed this agreement.",
    "The contracted scope includes a PDF-based readiness assessment.",
    "Deliverables include an assessment report and prioritized recommendations."
)

[System.IO.File]::WriteAllText(
    (Join-Path $pdfFailure "Contoso-Scanned-Signed.pdf"),
    "%PDF-1.4`n% Synthetic invalid PDF used to exercise extraction failure.`n",
    [System.Text.Encoding]::ASCII
)

[pscustomobject]@{
    root       = $root
    ready      = $ready
    ambiguous  = $ambiguous
    chronology = $chronology
    pdfSuccess = $pdfSuccess
    pdfFailure = $pdfFailure
} | ConvertTo-Json
