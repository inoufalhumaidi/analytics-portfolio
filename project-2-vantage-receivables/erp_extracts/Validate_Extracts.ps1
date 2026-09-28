<#
================================================================================
Validate_Extracts.ps1 -- the SAME file in Projects 2, 3 and 4.
Purpose: Prove the committed CSV extracts are exactly what SQL produces today.

WHY THIS EXISTS
    The Excel workbook and the Streamlit demo read these files, not SQL. A fix
    that reaches the SQL views but not the extracts leaves both of them
    reporting the old number with nothing to say so -- the way Project 1's
    utilization fix reached SQL and Excel but not the .pbix. Nothing compared
    the extracts with SQL. This does.

HOW
    It runs this folder's own Export_Extracts.ps1 into a temporary folder, so
    the queries exist once and cannot drift from a second copy, then compares
    every regenerated file with the committed one, line by line. The export is
    deterministic (every query has an ORDER BY), so the rule is exact equality:
    a looser test such as row counts or a tolerance would pass an extract whose
    values had moved. Line endings and the UTF-8 byte-order mark are ignored,
    because git may check a file out with either line ending; nothing else is.

    A committed CSV that no query produces fails too: it would be a file the
    workbook or the app reads that nothing keeps current.

    Only SELECTs reach the database. Export_Extracts.ps1 reads views, functions
    and read-only procedures, and writes nothing but the CSV files, which here
    go to the temporary folder and are deleted afterwards.

USAGE
    powershell -File erp_extracts/Validate_Extracts.ps1 [-ServerInstance <server\instance>]
    Exit code 0 = every extract matches SQL; 1 = any difference; 2 = not run
    (SQL Server unreachable -- which proves nothing either way).

DATA DISCLOSURE: all companies in this portfolio are fictional and all data is
synthetic. No confidential data and no production system is involved.
================================================================================
#>

param(
    [string]$ServerInstance = 'localhost\TEW_SQLEXPRESS'
)

$ErrorActionPreference = 'Stop'

$export = Join-Path $PSScriptRoot 'Export_Extracts.ps1'
if (-not (Test-Path $export)) { throw "Export_Extracts.ps1 not found beside this script." }

function Get-Lines([string]$path) {
    # -Encoding UTF8 drops the byte-order mark; splitting on any newline makes
    # CRLF and LF checkouts compare equal.
    @((Get-Content -LiteralPath $path -Raw -Encoding UTF8) -split "\r?\n")
}

# A window of a long CSV line that starts shortly before the first character at
# which the two lines differ -- a difference past a fixed cut-off would print two
# lines that look identical.
function Get-Window([string]$s, [int]$at) {
    if ($null -eq $s) { return "(end of file)" }
    $start = [math]::Max(0, [math]::Min($at, $s.Length) - 50)
    $len = [math]::Min(140, $s.Length - $start)
    $w = $s.Substring($start, $len)
    if ($start -gt 0) { $w = "..." + $w }
    if ($start + $len -lt $s.Length) { $w = $w + "..." }
    return $w
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("extract_check_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp | Out-Null
$code = 0
try {
    try {
        & $export -ServerInstance $ServerInstance -OutputPath $tmp *> $null
    } catch {
        Write-Host "NOT RUN: could not regenerate the extracts from SQL -- $($_.Exception.Message)" -ForegroundColor Yellow
        $code = 2
    }

    if ($code -eq 0) {
        $fresh     = @(Get-ChildItem -LiteralPath $tmp -Filter *.csv | Sort-Object Name)
        $committed = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter *.csv | Sort-Object Name)
        if ($fresh.Count -eq 0) { throw "Export_Extracts.ps1 produced no CSV files." }

        Write-Host ""
        Write-Host "Extracts in $PSScriptRoot against SQL on $ServerInstance" -ForegroundColor Cyan
        Write-Host ("-" * 78)
        $pass = 0; $failures = @()
        foreach ($f in $fresh) {
            $mine = Join-Path $PSScriptRoot $f.Name
            if (-not (Test-Path -LiteralPath $mine)) {
                $failures += "$($f.Name): SQL produces it, but no committed file exists"
                Write-Host ("  {0,-34} {1}" -f $f.Name, "MISSING") -ForegroundColor Red
                continue
            }
            $a = Get-Lines $mine
            $b = Get-Lines $f.FullName
            $first = -1
            for ($i = 0; $i -lt [math]::Max($a.Count, $b.Count); $i++) {
                $x = if ($i -lt $a.Count) { $a[$i] } else { $null }
                $y = if ($i -lt $b.Count) { $b[$i] } else { $null }
                if ($x -cne $y) { $first = $i; break }
            }
            # the last element is the empty string after the final newline
            $rows = [math]::Max(0, $b.Count - 2)
            if ($first -lt 0) {
                $pass++
                Write-Host ("  {0,-34} {1,8} rows  MATCH" -f $f.Name, $rows) -ForegroundColor Green
            } else {
                $x = if ($first -lt $a.Count) { $a[$first] } else { $null }
                $y = if ($first -lt $b.Count) { $b[$first] } else { $null }
                $at = 0
                if ($null -ne $x -and $null -ne $y) {
                    $lim = [math]::Min($x.Length, $y.Length)
                    while ($at -lt $lim -and $x[$at] -ceq $y[$at]) { $at++ }
                }
                $failures += ("{0}: committed {1} lines, SQL now {2}; first difference at line {3}, character {4}`n        committed: {5}`n        SQL now:   {6}" -f
                    $f.Name, ($a.Count - 1), ($b.Count - 1), ($first + 1), ($at + 1), (Get-Window $x $at), (Get-Window $y $at))
                Write-Host ("  {0,-34} {1,8} rows  DIFFERS at line {2}" -f $f.Name, $rows, ($first + 1)) -ForegroundColor Red
            }
        }
        foreach ($c in $committed) {
            if (-not ($fresh | Where-Object { $_.Name -eq $c.Name })) {
                $failures += "$($c.Name): committed, but no query in Export_Extracts.ps1 produces it"
                Write-Host ("  {0,-34} {1}" -f $c.Name, "NOT PRODUCED BY ANY QUERY") -ForegroundColor Red
            }
        }

        Write-Host ("-" * 78)
        if ($failures.Count -eq 0) {
            Write-Host "EXTRACTS MATCH: $pass of $pass files are exactly what SQL produces today." -ForegroundColor Green
        } else {
            Write-Host "EXTRACTS DIFFER: $pass matched, $($failures.Count) did not." -ForegroundColor Red
            $failures | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
            Write-Host "Re-run Export_Extracts.ps1 if SQL is right, then rebuild what reads the extracts." -ForegroundColor Red
            $code = 1
        }
    }
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
exit $code
