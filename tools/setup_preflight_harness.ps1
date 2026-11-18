# setup_preflight_harness.ps1
#
# Runs the working-tree setup_workbook (vba/setup_workbook_subs.bas) inside scratch
# copies of a real template, once per "user error" scenario plus a few good runs, and
# checks what each run left behind. A blocked run must leave the workbook exactly as
# it was (same file, same sheets, no external links, nothing new in the case folder,
# Excel's alerts back on); a good run must produce a built model file.
#
# Message boxes cannot be shown in a hidden Excel, so the module is injected with every
# MsgBox call rerouted to a recorder that returns a scripted answer; the recorded
# prompts are what the PASS/FAIL checks read.
#
# Usage (PowerShell prompt in the repo root; Excel installed; "Trust access to the VBA
# project object model" on):
#   & tools\setup_preflight_harness.ps1 -Template 'C:\path\MEWC Lambdas and VBA.xlsm' `
#       -CaseFile 'C:\path\some case.xlsx' [-Only 'xlsx target','valid']
# Nothing outside a temp folder is written; the template and case are only copied.
param(
    [Parameter(Mandatory = $true)] [string] $Template,
    [Parameter(Mandatory = $true)] [string] $CaseFile,
    [string[]] $Only = @()
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$vbYes = 6; $vbNo = 7

# ---- module source with MsgBox rerouted
$lines = [IO.File]::ReadAllLines((Join-Path $repo 'vba\setup_workbook_subs.bas')) | ? { -not $_.StartsWith("Attribute ") }
$out = New-Object System.Collections.Generic.List[string]
foreach ($l in $lines) {
    $out.Add(([regex]::Replace($l, '\bMsgBox\b', 'zz_MsgBox')))
    if ($l.StartsWith("Option Private Module")) { $out.Add("Public zz_log As String"); $out.Add("Public zz_answer As Long") }
}
$code = ($out -join "`r`n") + @"

Public Function zz_MsgBox(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    zz_log = zz_log & "[[" & prompt & "]]"
    If (buttons And 7) = 0 Then zz_MsgBox = 1 Else zz_MsgBox = zz_answer
End Function
Public Sub zz_reset(ByVal answer As Long)
    zz_log = "": zz_answer = answer
End Sub
Public Function zz_getlog() As String
    zz_getlog = zz_log
End Function
"@

$caseName = Split-Path -Leaf $CaseFile
$caseBase = [IO.Path]::GetFileNameWithoutExtension($caseName)
$root = Join-Path $env:TEMP ("setup_preflight_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
$null = New-Item -ItemType Directory -Path $root

# Each scenario: name, prep = scriptblock returning @{B1;B2;B3} given the case dir, answer, expect ('blocked'|'built'),
# msg = text the recorded prompts must contain, optional before/after hooks, built = expected model file name.
$S = @(
    @{ n = 'xlsx target';             b3 = 'test.xlsx';            expect = 'blocked'; msg = 'must end in .xlsm' }
    @{ n = 'xls target';              b3 = 'test.xls';             expect = 'blocked'; msg = 'must end in .xlsm' }
    @{ n = 'bad character';           b3 = 'te?st.xlsm';           expect = 'blocked'; msg = 'does not allow' }
    @{ n = 'path in B3';              b3 = 'sub\test.xlsm';        expect = 'blocked'; msg = 'does not allow' }
    @{ n = 'only an extension';       b3 = '.xlsm';                expect = 'blocked'; msg = 'needs a name' }
    @{ n = 'reserved name';           b3 = 'CON.xlsm';             expect = 'blocked'; msg = 'reserves' }
    @{ n = 'blank B3, answer No';     b3 = '';                     expect = 'blocked'; msg = 'is empty'; answer = $vbNo }
    @{ n = 'wrong folder';            b1 = '{dir}\nope';           b3 = 'test.xlsm'; expect = 'blocked'; msg = 'Case folder not found' }
    @{ n = 'blank folder';            b1 = '';                     b3 = 'test.xlsm'; expect = 'blocked'; msg = 'B1' }
    @{ n = 'wrong case file';         b2 = 'no such case.xlsx';    b3 = 'test.xlsm'; expect = 'blocked'; msg = 'Case file not found'; msg2 = $caseName }
    @{ n = 'case file w/o extension'; b2 = $caseBase;              b3 = 'test.xlsm'; expect = 'blocked'; msg = 'Did you mean' }
    @{ n = 'blank case file';         b2 = '';                     b3 = 'test.xlsm'; expect = 'blocked'; msg = 'B2' }
    @{ n = 'case file is not Excel';  b2 = 'notes.txt';            b3 = 'test.xlsm'; expect = 'blocked'; msg = 'not an Excel workbook'; pre = { param($d) Set-Content (Join-Path $d 'notes.txt') 'x' } }
    @{ n = 'full path in B2';         b2 = '{dir}\' + $caseName;   b3 = 'test.xlsm'; expect = 'blocked'; msg = 'just the file name' }
    @{ n = 'model = case file';       b2 = 'samename.xlsm';        b3 = 'samename.xlsm'; expect = 'blocked'; msg = 'same as the case file'; pre = { param($d) Copy-Item (Join-Path $d $caseName) (Join-Path $d 'samename.xlsm') } }
    @{ n = 'model = open workbook';   b3 = 'other.xlsm';           expect = 'blocked'; msg = 'already open'; openOther = $true }
    @{ n = 'model = template name';   b3 = '{tpl}';                expect = 'blocked'; msg = "template's own name"; tplInCaseDir = $true }
    @{ n = 'overwrite declined';      b3 = 'exists.xlsm';          expect = 'blocked'; msg = 'already exists'; answer = $vbNo; pre = { param($d) Set-Content (Join-Path $d 'exists.xlsm') 'old' }; keep = 'exists.xlsm' }
    @{ n = 'save fails (read-only)';  b3 = 'locked.xlsm';          expect = 'blocked'; msg = 'Could not save the model file'; answer = $vbYes; pre = { param($d) $f = Join-Path $d 'locked.xlsm'; Set-Content $f 'old'; (Get-Item $f).IsReadOnly = $true }; keep = 'locked.xlsm' }
    @{ n = 'no case sheet, answer No'; b2 = 'odd.xlsx';            b3 = 'odd attempt.xlsm'; expect = 'aborted'; msg = 'No case sheet was identified'; answer = $vbNo; oddCase = $true }
    @{ n = 'valid';                   b3 = 'attempt.xlsm';         expect = 'built'; built = 'attempt.xlsm'; rerun = $true }
    @{ n = 'no extension + quotes';   b1 = '"{dir}\"'; b2 = '"' + $caseName + '"'; b3 = '  attempt2 '; expect = 'built'; built = 'attempt2.xlsm' }
    @{ n = 'dots in the name';        b3 = 'attempt 2026.11.18';   expect = 'built'; built = 'attempt 2026.11.18.xlsm' }
    @{ n = 'unknown suffix';          b3 = 'attempt.final';        expect = 'built'; built = 'attempt.final.xlsm' }
    # B1/B2 on the case file's REAL folder (e.g. under OneDrive) must be accepted: the run gets as far as the B3 check.
    @{ n = 'real case folder';        b1 = (Split-Path -Parent $CaseFile); b3 = 'te?st.xlsm'; expect = 'blocked'; msg = 'does not allow' }
    # A failure before the case sheet exists aborts; one after it is only logged and the build goes on (as it always has).
    @{ n = 'early import error';      b3 = 'early.xlsm'; expect = 'aborted'; msg = 'could not import';
       mutate = @('Call TimeMark("03 open case workbook")', 'Err.Raise 513, , "simulated early failure"') }
    @{ n = 'late import error';       b3 = 'late.xlsm';  expect = 'built'; built = 'late.xlsm'; nomsg = 'could not import';
       mutate = @('Call TimeMark("06 junk-sheet cleanup")', 'Err.Raise 513, , "simulated late failure"') }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1; $xl.AskToUpdateLinks = $false
$pass = 0; $fail = 0
function SheetNames($wb) { $a = @(); foreach ($s in $wb.Sheets) { $a += $s.Name }; return ($a -join '|') }
function Links($wb) { $l = $wb.LinkSources(1); if ($l -eq $null) { return 0 } else { return @($l).Count } }
try {
    $i = 0
    foreach ($sc in $S) {
        if ($Only.Count -gt 0 -and -not ($Only -contains $sc.n)) { continue }
        $i++
        $work = Join-Path $root ("s$i"); $tdir = Join-Path $work 'T'; $cdir = Join-Path $work 'C'
        $null = New-Item -ItemType Directory -Path $tdir, $cdir
        $tplName = 'Template.xlsm'
        $tplPath = if ($sc.tplInCaseDir) { Join-Path $cdir $tplName } else { Join-Path $tdir $tplName }
        Copy-Item $Template $tplPath; (Get-Item $tplPath).IsReadOnly = $false
        Copy-Item $CaseFile (Join-Path $cdir $caseName); (Get-Item (Join-Path $cdir $caseName)).IsReadOnly = $false
        if ($sc.pre) { & $sc.pre $cdir }
        if ($sc.oddCase) { $ob = $xl.Workbooks.Add(); $ob.Worksheets.Item(1).Name = 'Foo'; $ob.Worksheets.Item(1).Range('B2').Value2 = 'hello'; $ob.SaveAs((Join-Path $cdir 'odd.xlsx'), 51); $ob.Close($false) }
        $other = $null
        if ($sc.openOther) { $other = $xl.Workbooks.Add(); $other.SaveAs((Join-Path $tdir 'other.xlsm'), 52) }

        $wb = $xl.Workbooks.Open($tplPath)
        $cm = $wb.VBProject.VBComponents.Item('setup_workbook_subs').CodeModule
        $src = $code
        if ($sc.mutate) { $src = $code.Replace($sc.mutate[0], $sc.mutate[1]); if ($src -eq $code) { throw "mutation target not found: " + $sc.mutate[0] } }
        $cm.DeleteLines(1, $cm.CountOfLines); $cm.AddFromString($src)
        $q = "'" + $wb.Name + "'!setup_workbook_subs."
        $b1 = if ($sc.ContainsKey('b1')) { $sc.b1.Replace('{dir}', $cdir) } else { $cdir }
        $b2 = if ($sc.ContainsKey('b2')) { $sc.b2.Replace('{dir}', $cdir) } else { $caseName }
        $b3 = $sc.b3.Replace('{tpl}', $tplName)
        $prep = $wb.Worksheets.Item('Prep')
        $prep.Range('B1').Value2 = $b1; $prep.Range('B2').Value2 = $b2; $prep.Range('B3').Value2 = $b3
        $before = SheetNames $wb; $nameBefore = $wb.Name
        $filesBefore = (Get-ChildItem $cdir -Recurse -File | % { $_.FullName.Substring($cdir.Length) + ':' + $_.Length }) -join '|'
        $answer = if ($sc.answer) { $sc.answer } else { $vbYes }
        $null = $xl.Run($q + 'zz_reset', $answer)
        $xl.DisplayAlerts = $true
        $t = Get-Date
        try { $null = $xl.Run($q + 'setup_workbook') } catch { $runErr = $_.Exception.Message }
        $secs = [Math]::Round(((Get-Date) - $t).TotalSeconds, 1)
        $q = "'" + $wb.Name + "'!setup_workbook_subs."     # a successful SaveAs renamed the workbook
        $log = [string]$xl.Run($q + 'zz_getlog')
        $after = SheetNames $wb; $nameAfter = $wb.Name; $links = Links $wb
        $state = "alerts=" + $xl.DisplayAlerts + " screen=" + $xl.ScreenUpdating + " calc=" + $xl.Calculation + " events=" + $xl.EnableEvents
        $stateOk = $xl.DisplayAlerts -and $xl.ScreenUpdating -and ($xl.Calculation -eq -4105) -and $xl.EnableEvents
        $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
        $openCount = $xl.Workbooks.Count
        $filesAfter = (Get-ChildItem $cdir -Recurse -File | % { $_.FullName.Substring($cdir.Length) + ':' + $_.Length }) -join '|'

        $why = @()
        if (-not $stateOk) { $why += "app state not restored ($state)" }
        if ($links -ne 0) { $why += "external links: $links" }
        if ($sc.msg -and -not $log.Contains($sc.msg)) { $why += "message lacks '" + $sc.msg + "'" }
        if ($sc.msg2 -and -not $log.Contains($sc.msg2)) { $why += "message lacks '" + $sc.msg2 + "'" }
        if ($sc.nomsg -and $log.Contains($sc.nomsg)) { $why += "unexpected message '" + $sc.nomsg + "'" }
        $expectOpen = if ($sc.openOther) { 2 } else { 1 }
        if ($openCount -ne $expectOpen) { $why += "open workbooks: $openCount" }
        switch ($sc.expect) {
            'blocked' {
                if ($after -ne $before) { $why += "sheets changed: $before -> $after" }
                if ($nameAfter -ne $nameBefore) { $why += "workbook renamed to $nameAfter" }
                if ($filesAfter -ne $filesBefore -and -not $sc.keep) { $why += "case folder changed" }
                if ($sc.keep) { $kf = Join-Path $cdir $sc.keep; if ((Get-Content $kf -Raw).Trim() -ne 'old') { $why += "existing file was overwritten" } }
            }
            'aborted' {   # the model file was created, but nothing may be built on top of it
                foreach ($bad in '_L1', 'AG', 'Hints', 'case copy') { if (("|" + $after + "|").Contains("|" + $bad + "|")) { $why += "built $bad after an abort" } }
            }
            'built' {
                if ($nameAfter -ne $sc.built) { $why += "workbook is $nameAfter, expected " + $sc.built }
                foreach ($need in 'Case', 'case copy', '_L1', 'AG', 'Hints') { if (-not ("|" + $after + "|").Contains("|" + $need + "|")) { $why += "missing sheet $need" } }
                if (-not (Test-Path (Join-Path $cdir $sc.built))) { $why += "model file not on disk" }
            }
        }
        if ($sc.rerun -and $why.Count -eq 0) {   # pressing Setup again inside the built workbook, answering No
            $null = $xl.Run($q + 'zz_reset', $vbNo); $xl.DisplayAlerts = $true; $xl.EnableEvents = $true
            $null = $xl.Run($q + 'setup_workbook'); $log2 = [string]$xl.Run($q + 'zz_getlog')
            if (-not $log2.Contains('ALREADY been set up')) { $why += "re-run was not questioned" }
            if ((SheetNames $wb) -ne $after) { $why += "re-run (answer No) changed the sheets" }
            if (-not $xl.DisplayAlerts) { $why += "re-run left alerts off" }
            $xl.DisplayAlerts = $false; $xl.EnableEvents = $false
        }
        if ($why.Count -eq 0) { $pass++; $verdict = 'PASS' } else { $fail++; $verdict = 'FAIL' }
        ("{0}  {1,-26} {2,-8} {3,5}s  {4}" -f $verdict, $sc.n, $sc.expect, $secs, $(if ($why.Count) { ($why -join '; ') } else { '' }))
        if ($why.Count -gt 0 -or $Only.Count -gt 0) { "        prompts: " + $log.Replace("`n", ' / ').Substring(0, [Math]::Min(400, $log.Length)) }
        $wb.Close($false); if ($other) { $other.Close($false) }
        while ($xl.Workbooks.Count -gt 0) { $xl.Workbooks.Item(1).Close($false) }
    }
    "`n$pass passed, $fail failed"
} finally {
    $xl.Quit()
    Get-ChildItem $root -Recurse -File -ErrorAction SilentlyContinue | % { try { $_.IsReadOnly = $false } catch {} }
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}
