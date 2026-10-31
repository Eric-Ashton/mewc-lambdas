# classify_mutation_harness.ps1
#
# Robustness check for classify_rows (vba/setup_workbook_subs.bas): runs the
# classifier on the Case sheet of one or more real case workbooks, then again after
# each of a list of small "author variation" mutations (points/answer headers in
# other languages, accented "Bonus", misspelt "Example", blank or textual Level
# column, blank rows after headers, dropped header rows, ...) and reports, per run,
# the number of levels, the bonus question count, the questions and example found
# for each level, and the number of column-A labels that differ from the unmutated
# baseline.
#
# Usage (from a PowerShell prompt in the repo root; Excel must be installed and
# "Trust access to the VBA project object model" must be on):
#   & tools\classify_mutation_harness.ps1 -Books @('C:\path\case one.xlsm','C:\path\case two.xlsm') `
#       [-Only @('baseline','Answer->Resposta')] [-DumpRows 20-80]
# Call it in-process like this (not via powershell.exe -File), so the array
# parameters bind correctly.
#
# Each workbook needs a sheet named "Case". The module is injected into a scratch
# workbook in a hidden Excel instance (classify_rows is made Public for the run),
# so the repo's working-tree code is what gets tested. Nothing is saved.
#
# Reading the output: every line should show the baseline's level count and
# "L<n>:<questions>e" per level; "<<<<< LEVELS CHANGED" or labeldiff>0 marks a
# variation the classifier does not survive. Known, accepted limits: game numbers
# or points written as text with letters ("Q1", "Game 1", "10 pts") and tables with
# no header row at all.
param(
    [Parameter(Mandatory = $true)] [string[]] $Books,
    [string[]] $Only = @(),
    [string] $DumpRows = ""
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$code = ([IO.File]::ReadAllLines((Join-Path $repo 'vba\setup_workbook_subs.bas')) | ? { -not $_.StartsWith("Attribute ") }) -join "`r`n"
$code = $code.Replace("Private Sub classify_rows()", "Public Sub classify_rows()")
$code += "`r`n`r`nPublic Function zz_levels() As Long`r`n    zz_levels = number_of_levels`r`nEnd Function`r`nPublic Sub zz_set_case(ByVal ws As Worksheet)`r`n    Set attempt_workbook = ThisWorkbook`r`n    Set case_worksheet = ws`r`nEnd Sub`r`nPublic Function zz_run_classify() As String`r`n    On Error Resume Next`r`n    Err.Clear`r`n    number_of_levels = 0`r`n    classify_rows`r`n    zz_run_classify = Err.Number & `" `" & Err.Description`r`nEnd Function`r`n"

function IsNum($v) { return ($v -is [double] -or $v -is [int] -or $v -is [decimal]) }
function Txt($v) { if ($v -eq $null) { return "" }; return [string]$v }
# cell replacers operate on a cloned [object[,]]; $cB,$cC,$cD,$cE are column indices inside the array
function Replace-Cells($a, $pred, $fn) {
    $r1 = $a.GetLowerBound(0); $r2 = $a.GetUpperBound(0); $c1 = $a.GetLowerBound(1); $c2 = $a.GetUpperBound(1)
    for ($r = $r1; $r -le $r2; $r++) { for ($c = $c1; $c -le $c2; $c++) { $v = $a[$r, $c]; if (& $pred $v $r $c) { $a[$r, $c] = (& $fn $v $r $c) } } }
    return , $a
}
function Is-HeaderRow($a, $r) { $d = (Txt $a[$r, $script:cD]).ToLower(); return ($d -match "point|pts|score|marks") }
function Is-QuestionRow($a, $r) { return ((IsNum $a[$r, $script:cB]) -and (IsNum $a[$r, $script:cD]) -and ($a[$r, $script:cD] -ne 0)) }
function Delete-Rows($a, $rowsToDrop) {
    $r1 = $a.GetLowerBound(0); $r2 = $a.GetUpperBound(0); $c1 = $a.GetLowerBound(1); $c2 = $a.GetUpperBound(1)
    $keep = @(); for ($r = $r1; $r -le $r2; $r++) { if (-not ($rowsToDrop -contains $r)) { $keep += $r } }
    $n = [Array]::CreateInstance([object], [int[]]@($keep.Count, $c2), [int[]]@(1, 1)); $i = 1
    foreach ($r in $keep) { for ($c = $c1; $c -le $c2; $c++) { $n[$i, $c] = $a[$r, $c] }; $i++ }
    return , $n
}
function Insert-BlankAfter($a, $rowsAfter, $count) {
    $r1 = $a.GetLowerBound(0); $r2 = $a.GetUpperBound(0); $c1 = $a.GetLowerBound(1); $c2 = $a.GetUpperBound(1)
    $total = $r2 + $rowsAfter.Count * $count; $n = [Array]::CreateInstance([object], [int[]]@($total, $c2), [int[]]@(1, 1)); $i = 1
    for ($r = $r1; $r -le $r2; $r++) { for ($c = $c1; $c -le $c2; $c++) { $n[$i, $c] = $a[$r, $c] }; $i++; if ($rowsAfter -contains $r) { $i += $count } }
    return , $n
}
function Same($v, $t) { return ((Txt $v).Trim() -ieq $t) }
function Sw($a, $from, $to) { return Replace-Cells $a { param($v) Same $v $from } { $to } }

$mutations = [ordered]@{
    "baseline"             = { param($a) , $a }
    "Points->Pontos"       = { param($a) Sw $a "Points" "Pontos" }
    "Answer->Resposta"     = { param($a) Sw $a "Answer" "Resposta" }
    "Pontos+Resposta"      = { param($a) $a = Sw $a "Points" "Pontos"; Sw $a "Answer" "Resposta" }
    "Pontos+Jogo"          = { param($a) $a = Sw $a "Points" "Pontos"; Replace-Cells $a { param($v) (Same $v "Game #") -or (Same $v "Question #") } { "Jogo #" } }
    "all-Portuguese-hdr"   = { param($a) $a = Sw $a "Points" "Pontos"; $a = Sw $a "Answer" "Resposta"; $a = Sw $a "Level" ("N" + [char]0xED + "vel"); Replace-Cells $a { param($v) (Same $v "Game #") -or (Same $v "Question #") } { "Jogo #" } }
    "Punkte+Antwort"       = { param($a) $a = Sw $a "Points" "Punkte"; Sw $a "Answer" "Antwort" }
    "Puntos+Respuesta"     = { param($a) $a = Sw $a "Points" "Puntos"; Sw $a "Answer" "Respuesta" }
    "Points->Value"        = { param($a) Sw $a "Points" "Value" }
    "hdr-E-blank"          = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cE) -and (Same $v "Answer") } { $null } }
    "Bonus->Bonus(accent)" = { param($a) Replace-Cells $a { param($v) (Txt $v) -match "^Bonus \d" } { param($v) (Txt $v).Replace("Bonus", "B" + [char]0xF4 + "nus") } }
    "Example->Exemplo"     = { param($a) Replace-Cells $a { param($v) (Txt $v) -match "^Example\d" } { param($v) (Txt $v).Replace("Example", "Exemplo") } }
    "Example->Ejemplo"     = { param($a) Replace-Cells $a { param($v) (Txt $v) -match "^Example\d" } { param($v) (Txt $v).Replace("Example", "Ejemplo") } }
    "Example typo Exanple" = { param($a) Replace-Cells $a { param($v) (Txt $v) -match "^Example\d" } { param($v) (Txt $v).Replace("Example", "Exanple") } }
    "Example->'Ex 1'"      = { param($a) Replace-Cells $a { param($v) (Txt $v) -match "^Example\d" } { param($v) (Txt $v).Replace("Example", "Ex ") } }
    "example points blank" = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cD) -and ((Txt $a[$r, $script:cB]) -match "^Example\d") } { $null } }
    "levelcol-C-blank"     = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cC) -and (Is-QuestionRow $a $r) } { $null } }
    "levelcol-C-'Level n'" = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cC) -and (Is-QuestionRow $a $r) } { param($v) "Level " + (Txt $v) } }
    "levelcol-C=points"    = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cC) -and (Is-QuestionRow $a $r) } { param($v, $r) $a[$r, $script:cD] } }
    "levelcol-C=game#"     = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cC) -and (Is-QuestionRow $a $r) } { param($v, $r) $a[$r, $script:cB] } }
    "game#-as-text"        = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cB) -and (Is-QuestionRow $a $r) } { param($v) [string]$v } }
    "game#-'Q1'"           = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cB) -and (Is-QuestionRow $a $r) } { param($v) "Q" + [string]$v } }
    "game#-'Game 1'"       = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cB) -and (Is-QuestionRow $a $r) } { param($v) "Game " + [string]$v } }
    "points-as-text"       = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cD) -and (Is-QuestionRow $a $r) } { param($v) [string]$v } }
    "points-'10 pts'"      = { param($a) Replace-Cells $a { param($v, $r, $c) ($c -eq $script:cD) -and (Is-QuestionRow $a $r) } { param($v) ([string]$v) + " pts" } }
    "full-Portuguese"      = { param($a) $a = Sw $a "Points" "Pontos"; $a = Sw $a "Answer" "Resposta"; $a = Sw $a "Level" ("N" + [char]0xED + "vel"); $a = Replace-Cells $a { param($v) (Same $v "Game #") -or (Same $v "Question #") } { "Jogo #" }; $a = Replace-Cells $a { param($v) (Txt $v) -match "^Bonus \d" } { param($v) (Txt $v).Replace("Bonus", "B" + [char]0xF4 + "nus") }; Replace-Cells $a { param($v) (Txt $v) -match "^Example\d" } { param($v) (Txt $v).Replace("Example", "Exemplo") } }
    "unknown-lang-hdr"     = { param($a) $a = Sw $a "Points" "Taskai"; $a = Sw $a "Answer" "Atsakymas"; $a = Sw $a "Level" "Lygis"; Replace-Cells $a { param($v) (Same $v "Game #") -or (Same $v "Question #") } { "Zaidimas" } }
    "16-blank-after-hdr"   = { param($a) $h = @(); for ($r = 1; $r -le $a.GetUpperBound(0); $r++) { if (Is-HeaderRow $a $r) { $h += $r } }; Insert-BlankAfter $a $h 16 }
    "drop-bonus-header"    = { param($a) $h = @(); for ($r = 1; $r -le $a.GetUpperBound(0); $r++) { if (Is-HeaderRow $a $r) { $h += $r; break } }; Delete-Rows $a $h }
    "drop-bonus-section"   = { param($a) $d = @(); $first = 0; for ($r = 1; $r -le $a.GetUpperBound(0); $r++) { if (($first -eq 0) -and (Is-HeaderRow $a $r)) { $first = $r; $d += $r }; if ((Txt $a[$r, $script:cB]) -match "^B.nus \d") { $d += $r } }; Delete-Rows $a $d }
    "drop-level-headers"   = { param($a) $h = @(); for ($r = 1; $r -le $a.GetUpperBound(0); $r++) { if (Is-HeaderRow $a $r) { $h += $r } }; if ($h.Count -gt 1) { $h = $h[1..($h.Count - 1)] }; Delete-Rows $a $h }
    "drop-'Level n' title" = { param($a) $h = @(); for ($r = 1; $r -le $a.GetUpperBound(0); $r++) { if ((Txt $a[$r, $script:cB]) -match "^Level \d+$") { $h += $r } }; Delete-Rows $a $h }
}

$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 3
try {
    $wb = $xl.Workbooks.Add(); $xl.Calculation = -4135
    $m = $wb.VBProject.VBComponents.Add(1); $m.Name = "setup_workbook_subs"; $m.CodeModule.AddFromString($code)
    $q = "'" + $wb.Name + "'!"
    try { $null = $xl.Run($q + "setup_workbook_subs.SheetExists", $wb, "Sheet1") } catch { "MODULE DOES NOT COMPILE: " + $_.Exception.Message; throw }
    foreach ($path in $Books) {
        $src = $xl.Workbooks.Open($path, 0, $true)
        $ur = $src.Worksheets.Item("Case").UsedRange; $vals = $ur.Value2; $r0 = $ur.Row; $c0 = $ur.Column
        $src.Close($false)
        $script:cB = 2 - $c0 + 1; $script:cC = $script:cB + 1; $script:cD = $script:cB + 2; $script:cE = $script:cB + 3
        $cs = $wb.Worksheets.Add([Type]::Missing, $wb.Sheets.Item($wb.Sheets.Count)); $cs.Name = "Case_" + $wb.Sheets.Count
        $null = $xl.Run($q + "setup_workbook_subs.zz_set_case", $cs)
        "`n######## " + (Split-Path -Leaf $path) + "  (used range starts R" + $r0 + "C" + $c0 + ", " + $vals.GetUpperBound(0) + " rows)"
        $base = $null; $baseLv = -1
        foreach ($mname in $mutations.Keys) {
            if ($Only.Count -gt 0 -and -not ($Only -contains $mname)) { continue }
            $arr = $vals.Clone(); $arr = & $mutations[$mname] $arr
            $null = $cs.Cells.Clear()
            $cs.Range($cs.Cells.Item($r0, $c0), $cs.Cells.Item($r0 + $arr.GetUpperBound(0) - 1, $c0 + $arr.GetUpperBound(1) - 1)).Value2 = $arr
            $err = $xl.Run($q + "setup_workbook_subs.zz_run_classify")
            $notes = @(); try { $el = $wb.Worksheets.Item("ErrorLog"); for ($r = 2; $r -le 12; $r++) { $t = $el.Cells.Item($r, 4).Text; if ($t -ne "") { $notes += $t } }; $null = $el.Delete() } catch {}
            $cc = $wb.Worksheets.Item("case copy"); $lastNew = $cc.Cells.Item($cc.Rows.Count, 1).End(-4162).Row
            $lv = $xl.Run($q + "setup_workbook_subs.zz_levels")
            $labels = @{}; $hdrs = @(); $q_by = @{}; $ex_by = @{}; $nb = 0; $nUnlabelled = 0
            for ($r = 1; $r -le $lastNew; $r++) {
                $t = [string]$cc.Cells.Item($r, 1).Text; $labels[$r] = $t
                if ($t -like "*Header") { $hdrs += ("r$r " + $t.Replace(" Header", "")) }
                if ($t -match "^Level (\d+) Level Question$") { $q_by[[int]$matches[1]] = 1 + [int]$q_by[[int]$matches[1]] }
                if ($t -match "^Level (\d+) Example Question$") { $ex_by[[int]$matches[1]] = 1 }
                if ($t -eq "Bonus Question") { $nb++ }
                if ($t -eq "Level Question" -or $t -eq "Example Question") { $nUnlabelled++ }
            }
            $per = @(); for ($l = 1; $l -le $lv; $l++) { $per += ("L$l" + ":" + [int]$q_by[$l] + $(if ($ex_by[$l]) { "e" } else { "-" })) }
            $diff = ""; if ($base -ne $null -and $mname -notmatch "^(16-blank|drop-)") { $n = 0; foreach ($k in $labels.Keys) { if ($labels[$k] -ne $base[$k]) { $n++ } }; $diff = "  labeldiff=$n" }
            $flag = $(if ($baseLv -ge 0 -and $mname -ne "baseline" -and $lv -ne $baseLv) { " <<<<< LEVELS CHANGED" } else { "" })
            if ($mname -eq "baseline") { $base = $labels; $baseLv = $lv }
            ("{0,-22} levels={1} bonusQ={2} unlabelledQ={3} {4}{5}{6}" -f $mname, $lv, $nb, $nUnlabelled, ($per -join " "), $diff, $flag)
            if ($err -ne "0 ") { "        classify_rows raised: " + $err }
            if ($notes.Count -gt 0) { "        notes: " + (($notes | Select-Object -First 3) -join " || ") }
            if ($mname -eq "baseline" -or $flag -ne "") { "        headers: " + ($hdrs -join ", ") }
            if ($DumpRows -ne "" -and $Only.Count -gt 0) {
                $lo, $hi = $DumpRows.Split("-")
                for ($r = [int]$lo; $r -le [int]$hi; $r++) {
                    $lb = [string]$cc.Cells.Item($r, 1).Text; $bb = [string]$cc.Cells.Item($r, 2).Text
                    if ($lb -ne "Spacing Row" -or $bb -ne "") { "   r{0,-4} {1,-26} B={2,-22} C={3,-12} D={4,-10} E={5}" -f $r, $lb, $bb.Substring(0, [Math]::Min(22, $bb.Length)), [string]$cc.Cells.Item($r, 3).Text, [string]$cc.Cells.Item($r, 4).Text, [string]$cc.Cells.Item($r, 5).Text }
                }
            }
            $null = $cc.Delete()
        }
        $null = $cs.Delete()
    }
    $wb.Close($false)
} finally { $xl.Quit() }
