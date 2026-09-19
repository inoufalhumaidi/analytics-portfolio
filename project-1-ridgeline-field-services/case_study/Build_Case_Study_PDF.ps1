<#
=============================================================================
Project 1 -- Ridgeline Field Services
Script:  Build_Case_Study_PDF.ps1
Purpose: Re-print Project1_Ridgeline_Case_Study.html to PDF with headless Edge.

WHY THIS EXISTS
    The HTML case study is hand-authored, but the PDF beside it was produced
    once by hand. So every correction to the HTML left a stale PDF sitting next
    to it saying something different -- and the PDF is the artefact a reader is
    most likely to open. Projects 2 and 3 regenerate their PDF as part of
    building the case study; this gives Project 1 the same guarantee for the
    one step that was actually manual.

    The HTML itself stays hand-authored here: unlike Projects 2 and 3 there is
    no markdown source to render from, and rewriting it into one would change
    the document rather than make it reproducible.

USAGE
    powershell case_study/Build_Case_Study_PDF.ps1

DATA DISCLOSURE: Ridgeline Field Services is fictional; all data is synthetic.
No confidential data and no production system is involved.
=============================================================================
#>

$ErrorActionPreference = "Stop"

$Html = Join-Path $PSScriptRoot 'Project1_Ridgeline_Case_Study.html'
$Pdf  = Join-Path $PSScriptRoot 'Project1_Ridgeline_Case_Study.pdf'

if (-not (Test-Path $Html)) { throw "Case study HTML not found at $Html" }

$edge = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $edge) {
    Write-Host "Edge not found -- PDF skipped. The HTML is still the source of truth." -ForegroundColor Yellow
    return
}

if (Test-Path $Pdf) { Remove-Item $Pdf -Force }
$profileDir = Join-Path $env:TEMP ("edgepdf_" + [guid]::NewGuid().ToString('N'))
$edgeArgs = @(
    '--headless', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer',
    "--user-data-dir=$profileDir",
    "--print-to-pdf=$Pdf",
    ([Uri]$Html).AbsoluteUri
)

# Edge writes harmless renderer warnings to stderr. In Windows PowerShell,
# redirecting a native executable's stderr wraps each line in an ErrorRecord,
# which $ErrorActionPreference = 'Stop' then treats as fatal -- so the build
# fails on a warning from a tool that succeeded. Let it write, and judge the
# result by whether the PDF appeared.
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try { & $edge @edgeArgs | Out-Null } finally { $ErrorActionPreference = $prevEAP }

# Edge returns before the file is flushed; wait for it rather than assuming.
$deadline = (Get-Date).AddSeconds(60)
while (-not (Test-Path $Pdf) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 400 }
if (Test-Path $profileDir) { Remove-Item $profileDir -Recurse -Force -ErrorAction SilentlyContinue }

if (Test-Path $Pdf) {
    $kb = [math]::Round((Get-Item $Pdf).Length / 1KB, 1)
    Write-Host "PDF written: $Pdf ($kb KB)" -ForegroundColor Green
} else {
    throw "PDF was not produced. The HTML is unchanged and still valid."
}
