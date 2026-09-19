<#
================================================================================
Project 4 -- Talon Robotics: Payload Deployment Delivery
Script:  Build_Case_Study.ps1
Purpose: Render the case study from Markdown into a styled HTML page, then print
         that page to PDF with headless Edge.

WHY RENDER RATHER THAN HAND-WRITE THE HTML
    The Markdown is the source of truth. Keeping a separate hand-written HTML
    copy guarantees the two drift, and the one a reader opens is usually the
    stale one.

SCOPE OF THE CONVERTER
    Deliberately small: headings, tables, fenced code, blockquotes, lists,
    horizontal rules, bold, italic, inline code and links. That is everything
    the case study actually uses. A general Markdown parser would be more code
    and more ways to be wrong.

DATA DISCLOSURE: Talon Robotics is fictional; all data is synthetic.
================================================================================
#>

param(
    [string]$Source = (Join-Path $PSScriptRoot 'Talon_Payload_Delivery_Case_Study.md'),
    [string]$Html   = (Join-Path $PSScriptRoot 'Talon_Payload_Delivery_Case_Study.html'),
    [string]$Pdf    = (Join-Path $PSScriptRoot 'Talon_Payload_Delivery_Case_Study.pdf')
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Source)) { throw "Source not found: $Source" }

function ConvertTo-Inline {
    param([string]$s)
    if ($null -eq $s) { return '' }
    # escape first, so that markup we add below is not escaped in turn
    $s = $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    # inline code before emphasis: a backtick span must not have its * interpreted
    $s = [regex]::Replace($s, '`([^`]+)`', '<code>$1</code>')
    $s = [regex]::Replace($s, '\[([^\]]+)\]\(([^)]+)\)', '<a href="$2">$1</a>')
    $s = [regex]::Replace($s, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    $s = [regex]::Replace($s, '(?<![\*\w])\*([^*]+)\*(?!\*)', '<em>$1</em>')
    $s = $s -replace '---','&mdash;' -replace '\s--\s',' &mdash; '
    return $s
}

$lines = Get-Content $Source -Encoding UTF8
$out   = New-Object System.Text.StringBuilder
$i     = 0
$inCode = $false
$listOpen = $false

function Close-List { if ($script:listOpen) { [void]$out.AppendLine('</ul>'); $script:listOpen = $false } }

while ($i -lt $lines.Count) {
    $line = $lines[$i]

    # ---- fenced code -------------------------------------------------------
    if ($line -match '^\s*```') {
        Close-List
        if (-not $inCode) { [void]$out.AppendLine('<pre><code>'); $inCode = $true }
        else              { [void]$out.AppendLine('</code></pre>'); $inCode = $false }
        $i++; continue
    }
    if ($inCode) {
        [void]$out.AppendLine(($line -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'))
        $i++; continue
    }

    # ---- table -------------------------------------------------------------
    if ($line -match '^\s*\|' -and ($i + 1) -lt $lines.Count -and $lines[$i+1] -match '^\s*\|[\s\-:\|]+\|\s*$') {
        Close-List
        $headerCells = ($line.Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
        # alignment row drives the column alignment, as Markdown intends
        $alignCells  = ($lines[$i+1].Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
        $aligns = @()
        foreach ($a in $alignCells) {
            if     ($a -match '^:.*:$') { $aligns += 'center' }
            elseif ($a -match ':$')     { $aligns += 'right'  }
            else                        { $aligns += 'left'   }
        }
        [void]$out.AppendLine('<table><thead><tr>')
        for ($c = 0; $c -lt $headerCells.Count; $c++) {
            $al = if ($c -lt $aligns.Count) { $aligns[$c] } else { 'left' }
            [void]$out.AppendLine("<th style=""text-align:$al"">$(ConvertTo-Inline $headerCells[$c])</th>")
        }
        [void]$out.AppendLine('</tr></thead><tbody>')
        $i += 2
        while ($i -lt $lines.Count -and $lines[$i] -match '^\s*\|') {
            $cells = ($lines[$i].Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            [void]$out.AppendLine('<tr>')
            for ($c = 0; $c -lt $cells.Count; $c++) {
                $al = if ($c -lt $aligns.Count) { $aligns[$c] } else { 'left' }
                [void]$out.AppendLine("<td style=""text-align:$al"">$(ConvertTo-Inline $cells[$c])</td>")
            }
            [void]$out.AppendLine('</tr>')
            $i++
        }
        [void]$out.AppendLine('</tbody></table>')
        continue
    }

    # ---- blockquote --------------------------------------------------------
    if ($line -match '^\s*>\s?(.*)$') {
        Close-List
        $buf = @()
        while ($i -lt $lines.Count -and $lines[$i] -match '^\s*>\s?(.*)$') {
            $buf += $Matches[1]
            $i++
        }
        [void]$out.AppendLine("<blockquote><p>$(ConvertTo-Inline ($buf -join ' '))</p></blockquote>")
        continue
    }

    # ---- headings ----------------------------------------------------------
    if ($line -match '^(#{1,6})\s+(.*)$') {
        Close-List
        $lvl = $Matches[1].Length
        [void]$out.AppendLine("<h$lvl>$(ConvertTo-Inline $Matches[2])</h$lvl>")
        $i++; continue
    }

    # ---- horizontal rule ---------------------------------------------------
    if ($line -match '^\s*---+\s*$') { Close-List; [void]$out.AppendLine('<hr/>'); $i++; continue }

    # ---- list --------------------------------------------------------------
    if ($line -match '^\s*(?:[-*]|\d+\.)\s+(.*)$') {
        if (-not $listOpen) { [void]$out.AppendLine('<ul>'); $listOpen = $true }
        [void]$out.AppendLine("<li>$(ConvertTo-Inline $Matches[1])</li>")
        $i++; continue
    }

    # ---- paragraph / blank -------------------------------------------------
    if ($line.Trim() -eq '') { Close-List; $i++; continue }
    Close-List
    $para = @()
    while ($i -lt $lines.Count -and $lines[$i].Trim() -ne '' -and
           $lines[$i] -notmatch '^\s*(#|>|\||---+\s*$|```)' -and
           $lines[$i] -notmatch '^\s*(?:[-*]|\d+\.)\s') {
        $para += $lines[$i].Trim()
        $i++
    }
    if ($para.Count) { [void]$out.AppendLine("<p>$(ConvertTo-Inline ($para -join ' '))</p>") }
}
Close-List

$css = @'
:root{
  --ink:#1b2430; --muted:#5b6b7f; --rule:#dfe5ec; --navy:#1f2d3d;
  --band:#f4f6f8; --accent:#c47920; --red:#9c0006; --redbg:#ffc7ce;
  --amber:#9c6500; --amberbg:#ffeb9c; --green:#006100; --greenbg:#c6efce;
}
*{box-sizing:border-box}
body{
  font-family:"Segoe UI",-apple-system,Helvetica,Arial,sans-serif;
  color:var(--ink); line-height:1.62; max-width:960px;
  margin:0 auto; padding:56px 40px 80px; font-size:15px;
}
h1{font-size:30px;line-height:1.25;color:var(--navy);margin:0 0 6px;letter-spacing:-.3px}
h2{font-size:21px;color:var(--navy);margin:38px 0 12px;padding-bottom:7px;border-bottom:2px solid var(--rule)}
h3{font-size:16.5px;color:var(--accent);margin:26px 0 8px}
h4{font-size:15px;margin:20px 0 6px}
p{margin:0 0 12px}
hr{border:0;border-top:1px solid var(--rule);margin:30px 0}
a{color:var(--accent)}
code{
  font-family:Consolas,"SF Mono",Menlo,monospace; font-size:.88em;
  background:var(--band); padding:1px 5px; border-radius:3px; color:#324a63;
}
pre{
  background:#1b2430; color:#e6edf3; padding:16px 18px; border-radius:6px;
  overflow-x:auto; font-size:12.5px; line-height:1.5;
}
pre code{background:none;color:inherit;padding:0;font-size:inherit}
blockquote{
  margin:16px 0; padding:11px 18px; background:var(--band);
  border-left:4px solid var(--accent); color:var(--muted);
}
blockquote p{margin:0}
table{border-collapse:collapse;width:100%;margin:16px 0;font-size:13.5px}
th{
  background:var(--navy); color:#fff; font-weight:600; text-align:left;
  padding:9px 11px; border:1px solid var(--navy);
}
td{padding:8px 11px;border:1px solid var(--rule);vertical-align:top}
/* numbers and their units belong on one line; a wrapped "60.81 / d" reads as two values */
td[style*="right"],th[style*="right"],td[style*="center"],th[style*="center"]{white-space:nowrap}
tbody tr:nth-child(even){background:#fafbfd}
ul{margin:0 0 14px;padding-left:22px}
li{margin:4px 0}
strong{color:#0f1a26}
@media print{
  body{max-width:none;padding:0 12mm;font-size:11.2pt}
  h1{font-size:23pt} h2{font-size:15pt;margin-top:22px} h3{font-size:12.5pt}
  table{font-size:9.6pt; page-break-inside:auto}
  tr{page-break-inside:avoid}
  h2,h3{page-break-after:avoid}
  pre{font-size:8.8pt;background:#f2f4f7;color:#1b2430;border:1px solid #dfe5ec}
  blockquote{background:#f2f4f7}
}
'@

$page = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Talon Robotics &mdash; Payload Deployment Delivery Case Study</title>
<style>$css</style>
</head>
<body>
$($out.ToString())
</body>
</html>
"@

[System.IO.File]::WriteAllText($Html, $page, (New-Object System.Text.UTF8Encoding $false))
Write-Host "  HTML written: $Html" -ForegroundColor Green

# ---- print to PDF with headless Edge ---------------------------------------
$edge = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $edge) {
    Write-Host "  Edge not found -- HTML written, PDF skipped." -ForegroundColor Yellow
    return
}

if (Test-Path $Pdf) { Remove-Item $Pdf -Force }
$profileDir = Join-Path $env:TEMP ("edgepdf_" + [guid]::NewGuid().ToString('N'))
$args = @(
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
try { & $edge @args | Out-Null } finally { $ErrorActionPreference = $prevEAP }

# Edge returns before the file is flushed; wait for it rather than assuming.
$deadline = (Get-Date).AddSeconds(60)
while (-not (Test-Path $Pdf) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 400 }
if (Test-Path $profileDir) { Remove-Item $profileDir -Recurse -Force -ErrorAction SilentlyContinue }

if (Test-Path $Pdf) {
    $kb = [math]::Round((Get-Item $Pdf).Length / 1KB, 1)
    Write-Host "  PDF written:  $Pdf ($kb KB)" -ForegroundColor Green
} else {
    Write-Host "  PDF was not produced -- the HTML is still valid." -ForegroundColor Yellow
}
