# =============================================================================
# Sync_Measures.ps1 -- the SAME file in every project with a Power BI model
# (Projects 1, 3 and 4). It learns everything project-specific from the
# project's own Build_PowerBI_Model.ps1: the measures, and the table each is
# built on. Keep the copies identical; a per-project fork would drift exactly
# the way this tool exists to catch.
#
# Checks that the DAX measures agree in all three places they live:
#   1. Build_PowerBI_Model.ps1  -- the source of truth
#   2. DAX_Measures.md          -- the annotated copy a reader pastes from
#   3. the .pbix open in Power BI Desktop
# and, with -Apply, brings the open .pbix into line with the script.
#
# WHY THIS EXISTS
#   In Project 1 the weekday-utilization fix (b1447a6) reached SQL, Excel, the
#   exports and the documents -- but not the .pbix, which went on reporting
#   70.455% where everything else said 65.891%. DAX_Measures.md drifted the same way: it kept
#   the defective [Worked Minutes] and was missing five of the 31 measures.
#   Nothing compared either of them with the script, so nothing noticed.
#   This does.
#
# USAGE
#   powershell -File Sync_Measures.ps1 -DocOnly   script vs DAX_Measures.md (no Power BI needed)
#   powershell -File Sync_Measures.ps1            ...and vs the .pbix open in Power BI Desktop
#   powershell -File Sync_Measures.ps1 -Apply     update the open model to match the script,
#                                                 then File > Save in Power BI Desktop
#   Exit code 0 = script, document and model all agree; 1 = any difference.
#
# WHAT COUNTS AS A DIFFERENCE
#   Any measure missing, extra, or defined differently -- in the document or in
#   the model, on any table. Both are compared as a stream of DAX tokens, so
#   layout is ignored (line breaks, indentation, spacing, and --, // and /* */
#   comments) while every token counts: a "string", a [name] and a 'table name'
#   are each one token compared exactly, "&&" is not "& &", and names are
#   case-sensitive. The document is read the way GitHub renders it: every dax
#   code block -- backtick or tilde fence, in a list or a blockquote -- must
#   define a measure as "Name = expression", or the check fails naming its line;
#   a block inside an HTML comment does not render, so it does not count.
#   "Defined" means what the build script sets: expression, format string and
#   display folder. Properties it never sets -- hidden, description, a dynamic
#   format string added in Desktop -- are not compared.
#
# -Apply changes MEASURES ONLY -- no table, column or relationship -- so there
# is no refresh and no credential prompt. It edits a measure IN PLACE, on
# whichever table it lives, so every visual bound to it keeps working, and adds
# a missing one to the table the build script builds it on. It never deletes: an extra measure fails
# the check until a person removes it or adds it to the script. It cannot fix
# the document; edit DAX_Measures.md by hand.
#
# Unlike Build_PowerBI_Model.ps1, this is safe to run against a finished report.
# The build script rebuilds the data model and cannot recreate report pages;
# saving its output over a finished .pbix would discard them.
# =============================================================================
param(
    [switch]$Apply,
    [switch]$DocOnly,
    [int]$Port = 0,
    [string]$DocPath = (Join-Path $PSScriptRoot "DAX_Measures.md")
)
$ErrorActionPreference = "Stop"

# -----------------------------------------------------------------------------
# 1. The definitions, read from the build script itself
#
#    The block between the MEASURES markers is run with a New-Measure that only
#    records its arguments, so the definitions come from the very code that
#    builds the model -- here-strings, shared format variables and all. It runs
#    under StrictMode, so a variable the block needs but does not define (say,
#    $pct moved above the marker) is an error rather than a silently empty
#    format string. A New-Measure call outside the markers is an error too:
#    otherwise it would build a measure this check never sees.
# -----------------------------------------------------------------------------
$buildScript = Join-Path $PSScriptRoot "Build_PowerBI_Model.ps1"
# -Encoding UTF8 is not optional: Windows PowerShell 5.1 reads a file without a
# BOM as Windows-1252, which turns a UTF-8 non-breaking space into two wrong characters.
$text = Get-Content $buildScript -Raw -Encoding UTF8
$begin = $text.IndexOf("# ---- MEASURES BEGIN")
$end = $text.IndexOf("# ---- MEASURES END")
if ($begin -lt 0 -or $end -le $begin) { throw "MEASURES BEGIN / END markers not found in $buildScript" }
# Calls are found by PowerShell's own parser, not a regex: a call written as
#   if (...) { New-Measure ... }   or   $null = New-Measure ...
# is still a call, while a comment or string that mentions the name is not.
$tokens = $null; $parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors) { throw "$buildScript does not parse: $($parseErrors[0].Message) (line $($parseErrors[0].Extent.StartLineNumber))" }
$calls = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-Measure' }, $true)
$outside = @($calls | Where-Object { $_.Extent.StartOffset -lt $begin -or $_.Extent.StartOffset -ge $end })
if ($outside) {
    $lines = ($outside | ForEach-Object { $_.Extent.StartLineNumber }) -join ', '
    throw "$($outside.Count) New-Measure call(s) in $buildScript sit outside the MEASURES markers (line $lines) -- move them inside."
}

# The block passes table objects ($F, $fact, ...) that the script assigns ABOVE
# the marker as  $X = $model.Tables["Name"].  Each becomes a stand-in carrying
# just the table's name, so every measure records the table it is built on.
$script:defs = New-Object System.Collections.Generic.List[object]
function New-Measure($table, $name, $expr, $fmt, $folder) {
    $tableName = if ($table) { $table.Name } else { $null }
    $script:defs.Add([pscustomobject]@{ Name = $name; Table = $tableName; Expression = $expr; FormatString = $fmt; DisplayFolder = $folder })
}
$standIns = @(foreach ($m in [regex]::Matches($text.Substring(0, $begin), '(?m)^[ \t]*\$(?<var>\w+)\s*=\s*\$model\.Tables\["(?<tbl>[^"]+)"\]\s*$')) {
    [pscustomobject]@{ Var = $m.Groups['var'].Value; Table = $m.Groups['tbl'].Value } })
$block = [scriptblock]::Create($text.Substring($begin, $end - $begin))
# Stand-ins and the block live in a CHILD scope: a build script whose table
# variable happened to be called $text or $begin must not overwrite this tool's.
& {
    Set-StrictMode -Version Latest
    foreach ($si in $standIns) { Set-Variable -Name $si.Var -Value ([pscustomobject]@{ Name = $si.Table }) }
    . $block
}

$dupes = @($script:defs | Group-Object Name | Where-Object Count -gt 1 | ForEach-Object Name)
if ($dupes) { throw "The build script defines these measures more than once: $($dupes -join ', ')" }
$homeless = @($script:defs | Where-Object { -not $_.Table } | ForEach-Object Name)
if ($homeless) { throw "Could not tell which table these measures are built on (their table variable is not assigned as `$X = `$model.Tables[""Name""] above the marker): $($homeless -join ', ')" }
$script:tables = @($script:defs | ForEach-Object Table | Sort-Object -Unique)
function Get-Def([string]$name) { $script:defs | Where-Object { $_.Name -ceq $name } | Select-Object -First 1 }
function Test-ScriptName([string]$name) { [bool]($script:defs | Where-Object { $_.Name -ieq $name }) }
Write-Host "Build_PowerBI_Model.ps1 defines $($script:defs.Count) measures on $($script:tables.Count) table(s): $($script:tables -join ', ')."

# DAX as a token stream. Order matters: a "string", a [name] (with ]] escapes)
# and a 'table name' are matched before anything inside them can be mistaken
# for a comment or an operator; comments are matched before operators; the
# two-character operators before the one-character ones. Whitespace is never a
# token, which is what makes layout irrelevant and nothing else.
$script:daxToken = [regex]('"(?:[^"]|"")*"|\[(?:[^\]]|\]\])*\]|''(?:[^'']|'''')*''' +
    '|--[^\r\n]*|//[^\r\n]*|/\*[\s\S]*?\*/' +
    '|&&|\|\||<=|>=|<>|==' +
    '|\d+(?:\.\d+)?(?:[eE][-+]?\d+)?|\w+|\S')
function ConvertTo-Canonical([string]$dax) {
    if ($null -eq $dax) { return "" }
    $kept = foreach ($m in $script:daxToken.Matches($dax)) {
        $t = $m.Value
        if ($t.StartsWith('--') -or $t.StartsWith('//') -or $t.StartsWith('/*')) { continue }
        $t
    }
    return ($kept -join ' ')
}

# Show WHERE two canonical forms part, not their first 170 characters.
function Show-Difference([string]$label, [string]$a, [string]$b) {
    $i = 0; $n = [Math]::Min($a.Length, $b.Length)
    while ($i -lt $n -and $a[$i] -ceq $b[$i]) { $i++ }
    $from = [Math]::Max(0, $i - 45)
    Write-Host "      script: ...$($a.Substring($from, [Math]::Min(100, $a.Length - $from)))"
    Write-Host ("      {0,-6}: ...{1}" -f $label, $b.Substring($from, [Math]::Min(100, $b.Length - $from)))
    Write-Host ("              " + (' ' * ($i - $from + 3)) + "^ first difference")
}

# -----------------------------------------------------------------------------
# 2. DAX_Measures.md, read the way GitHub renders it (CommonMark fences)
# -----------------------------------------------------------------------------
$docDrift = 0
$doc = New-Object 'System.Collections.Generic.Dictionary[string,string]'   # case-sensitive
$raw = Get-Content $DocPath -Raw -Encoding UTF8
$opener = [regex]'^(?<quote>(?:[ \t]{0,3}>[ \t]?)*)[ \t]*(?:(?:[-*+]|\d{1,9}[.)])[ \t]+)?(?<fence>`{3,}|~{3,})[ \t]*(?<info>.*)$'
# A dax block inside an HTML comment is invisible on the rendered page but still
# sits in the file -- which is exactly how a stale definition lingers unseen. It
# is a failure, not a thing to skip.
foreach ($c in [regex]::Matches($raw, '<!--[\s\S]*?-->')) {
    foreach ($l in ($c.Value -split "\r?\n")) {
        $o = $opener.Match($l)
        if ($o.Success -and (($o.Groups['info'].Value.Trim() -split '\s+')[0] -ieq 'dax')) {
            $at = [regex]::Matches($raw.Substring(0, $c.Index), "\n").Count + 1
            Write-Host "  DOC    a dax block is hidden inside an HTML comment (from line $at) -- un-comment it or delete it" -ForegroundColor Red
            $docDrift++; break
        }
    }
}
# Then blank every comment out, keeping its line breaks so the line numbers in
# later messages still point at the file.
$raw = [regex]::Replace($raw, '<!--[\s\S]*?-->', { param($m) [regex]::Replace($m.Value, '[^\r\n]', '') })
$lines = $raw -split "\r?\n"
$inBlock = $false
for ($ln = 0; $ln -lt $lines.Count; $ln++) {
    $line = $lines[$ln]
    if (-not $inBlock) {
        $o = $opener.Match($line)
        if (-not $o.Success) { continue }
        $fence = $o.Groups['fence'].Value; $info = $o.Groups['info'].Value.Trim()
        if ($fence[0] -eq '`' -and $info.Contains('`')) { continue }      # inline ```code``` in prose
        $inBlock = $true; $startLine = $ln + 1; $body = @()
        $depth = ($o.Groups['quote'].Value.ToCharArray() | Where-Object { $_ -eq '>' }).Count
        $isDax = (($info -split '\s+')[0] -ieq 'dax')
        $unquote = [regex]("^(?:[ \t]{0,3}>[ \t]?){0,$depth}")
        $closer = [regex]('^[ \t]*' + [regex]::Escape([string]$fence[0]) + '{' + $fence.Length + ',}[ \t]*$')
        continue
    }
    $content = $unquote.Replace($line, '', 1)
    if (-not $closer.IsMatch($content)) { $body += $content; continue }
    $inBlock = $false
    if (-not $isDax) { continue }
    $code = @($body | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^\s*(--|//)' })
    # The name may sit alone on its line with "=" opening the next -- valid DAX
    # layout -- so match across the joined block, but never let a name span lines.
    $head = if ($code) { [regex]::Match(($code -join "`n"), '^\s*(?<name>[^=\r\n]+?)\s*=\s*(?<rest>[\s\S]*)$') } else { $null }
    if (-not $head -or -not $head.Success) {
        Write-Host "  DOC    the dax block at line $startLine does not define a measure as 'Name = expression'" -ForegroundColor Red
        $docDrift++; continue
    }
    $name = $head.Groups['name'].Value.Trim()
    if ($doc.ContainsKey($name)) { Write-Host "  DOC    '$name' is defined twice (second at line $startLine)" -ForegroundColor Red; $docDrift++ }
    $doc[$name] = $head.Groups['rest'].Value
}
if ($inBlock) { Write-Host "  DOC    the code block opened at line $startLine is never closed" -ForegroundColor Red; $docDrift++ }

$docOk = 0
foreach ($d in $script:defs) {
    if (-not $doc.ContainsKey($d.Name)) { Write-Host "  DOC    missing: $($d.Name)" -ForegroundColor Red; $docDrift++; continue }
    $a = ConvertTo-Canonical $d.Expression; $b = ConvertTo-Canonical $doc[$d.Name]
    if ($a -cne $b) { Write-Host "  DOC    differs: $($d.Name)" -ForegroundColor Red; Show-Difference "doc" $a $b; $docDrift++ }
    else { $docOk++ }
}
foreach ($n in $doc.Keys) {
    if (-not (Get-Def $n)) { Write-Host "  DOC    not in the script: $n" -ForegroundColor Red; $docDrift++ }
}
Write-Host "DAX_Measures.md: $docOk of $($script:defs.Count) match."

if ($DocOnly) {
    if ($docDrift) { Write-Host "DRIFT: $docDrift difference(s) in DAX_Measures.md -- edit the document to match the script." -ForegroundColor Red; exit 1 }
    Write-Host "OK: the document matches the script." -ForegroundColor Green; exit 0
}

# -----------------------------------------------------------------------------
# 3. The .pbix open in Power BI Desktop
# -----------------------------------------------------------------------------
$dllDir = Join-Path $PSScriptRoot ".amo_libs\lib\net45"
if (-not (Test-Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll"))) {
    throw "Analysis Services client libraries not found in .amo_libs\ -- run Build_PowerBI_Model.ps1 once (it downloads them), or use -DocOnly."
}
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Core.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.dll")
Add-Type -Path (Join-Path $dllDir "Microsoft.AnalysisServices.Tabular.dll")

# Each open Desktop window runs its own engine. Find the ones holding this model,
# and name them by their window title so the choice is never a guess.
$candidates = @()
foreach ($p in @(Get-Process -Name msmdsrv -ErrorAction SilentlyContinue)) {
    $listen = Get-NetTCPConnection -OwningProcess $p.Id -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $listen -or ($Port -and $listen.LocalPort -ne $Port)) { continue }
    $parentId = (Get-CimInstance Win32_Process -Filter "ProcessId = $($p.Id)").ParentProcessId
    $title = (Get-Process -Id $parentId -ErrorAction SilentlyContinue).MainWindowTitle
    $s = New-Object Microsoft.AnalysisServices.Tabular.Server
    try {
        $s.Connect("localhost:$($listen.LocalPort)")
        $mdl = if ($s.Databases.Count) { $s.Databases[0].Model } else { $null }
        if ($mdl -and -not @($script:tables | Where-Object { -not $mdl.Tables.ContainsName($_) })) {
            $candidates += [pscustomobject]@{ Port = $listen.LocalPort; Title = $title }
        }
    } finally { $s.Disconnect() }
}
if ($candidates.Count -eq 0) { throw "No open Power BI Desktop model contains all of: $($script:tables -join ', '). Open the .pbix first, or use -DocOnly." }
if ($candidates.Count -gt 1) {
    $list = ($candidates | ForEach-Object { "port $($_.Port) '$($_.Title)'" }) -join '; '
    throw "More than one open model contains this project's tables ($list). Pass -Port to choose."
}
$target = $candidates[0]
Write-Host "Model: '$($target.Title)' on port $($target.Port)."

$modelDrift = 0
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect("localhost:$($target.Port)")
try {
    $model = $server.Databases[0].Model
    $norm = { param($v) if ($null -eq $v) { "" } else { ($v -replace "`r`n", "`n").Trim() } }

    # A measure is found by name on ANY table (measure names are unique across
    # the model, ignoring case), so one moved to a _Measures table is compared
    # where it lives rather than reported missing and added a second time.
    function Find-Measure([string]$name) {
        foreach ($t in $model.Tables) { foreach ($m in $t.Measures) { if ($m.Name -ieq $name) { return $m } } }
        return $null
    }
    function Get-ModelDrift {
        $out = @()
        foreach ($d in $script:defs) {
            $m = Find-Measure $d.Name
            if (-not $m) { $out += [pscustomobject]@{ Name = $d.Name; What = "missing" }; continue }
            $what = @()
            if ($m.Name -cne $d.Name) { $what += "name case ('$($m.Name)')" }
            if ((ConvertTo-Canonical $m.Expression) -cne (ConvertTo-Canonical $d.Expression)) { $what += "expression" }
            if ((& $norm $m.FormatString) -cne (& $norm $d.FormatString)) { $what += "format" }
            if ((& $norm $m.DisplayFolder) -cne (& $norm $d.DisplayFolder)) { $what += "folder" }
            if ($what) { $out += [pscustomobject]@{ Name = $d.Name; What = ($what -join ", ") } }
        }
        return $out
    }
    function Get-Extras {
        # A name the script defines in another case is a RENAME, reported once by
        # Get-ModelDrift -- not also an extra whose "remedy" would delete it.
        @(foreach ($t in $model.Tables) { foreach ($m in $t.Measures) {
            if (-not (Test-ScriptName $m.Name)) { "$($t.Name)[$($m.Name)]" } } })
    }

    $changes = @(Get-ModelDrift)
    foreach ($x in $changes) {
        Write-Host "  MODEL  $($x.What): $($x.Name)" -ForegroundColor Yellow
        if ($x.What -like "*expression*") {
            Show-Difference "model" (ConvertTo-Canonical (Get-Def $x.Name).Expression) (ConvertTo-Canonical (Find-Measure $x.Name).Expression)
        }
    }
    foreach ($d in $script:defs) {
        $m = Find-Measure $d.Name
        if ($m -and $m.Table -and $m.Table.Name -ne $d.Table) { Write-Host "  MODEL  note: '$($d.Name)' lives on $($m.Table.Name); the script builds it on $($d.Table) (compared where it lives)" }
    }
    Write-Host "Model: $($script:defs.Count - $changes.Count) of $($script:defs.Count) match."

    if ($Apply -and $changes.Count) {
        foreach ($x in $changes) {
            $d = Get-Def $x.Name
            $m = Find-Measure $d.Name
            if (-not $m) {
                $m = New-Object Microsoft.AnalysisServices.Tabular.Measure
                $m.Name = $d.Name
                $m.LineageTag = [guid]::NewGuid().ToString()
                $model.Tables[$d.Table].Measures.Add($m) | Out-Null
            }
            $m.Name = $d.Name
            $m.Expression = $d.Expression
            $m.FormatString = $d.FormatString
            $m.DisplayFolder = $d.DisplayFolder
        }
        $model.SaveChanges() | Out-Null
        # The engine validates every expression on save. Read the verdict back for
        # the SCRIPT's measures, then re-compare: "applied" means the model now
        # matches, not that we tried.
        $broken = @($script:defs | ForEach-Object { Find-Measure $_.Name } |
                    Where-Object { $_ -and $_.State -ne [Microsoft.AnalysisServices.Tabular.ObjectState]::Ready })
        foreach ($b in $broken) { Write-Host "  BROKEN $($b.Name): $($b.ErrorMessage)" -ForegroundColor Red }
        $after = @(Get-ModelDrift)
        if ($broken.Count -or $after.Count) {
            Write-Host "APPLY FAILED: $($after.Count) still differ, $($broken.Count) not valid. The open session HAS been changed -- close it WITHOUT saving." -ForegroundColor Red
            exit 1
        }
        Write-Host "Applied $($changes.Count) change(s); the model's measures now match the script and are all valid." -ForegroundColor Green
        Write-Host "Now save the file in Power BI Desktop (File > Save) -- until then only the open session has changed."
        $changes = @()
    }
    $modelDrift += $changes.Count
    $extras = @(Get-Extras)
    foreach ($e in $extras) { Write-Host "  MODEL  not in the script: $e" -ForegroundColor Yellow }
    $modelDrift += $extras.Count
} finally { $server.Disconnect() }

if ($docDrift -or $modelDrift) {
    if ($docDrift) { Write-Host "DRIFT: $docDrift difference(s) in DAX_Measures.md -- edit the document to match the script." -ForegroundColor Red }
    if ($changes.Count) { Write-Host "DRIFT: $($changes.Count) measure(s) in the model differ -- re-run with -Apply." -ForegroundColor Red }
    if ($extras.Count) { Write-Host "DRIFT: $($extras.Count) measure(s) in the model are not in the script -- remove them in Power BI Desktop, or add them to Build_PowerBI_Model.ps1." -ForegroundColor Red }
    exit 1
}
Write-Host "OK: the script, the document and the open model agree." -ForegroundColor Green
exit 0
