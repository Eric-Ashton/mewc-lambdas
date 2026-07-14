Attribute VB_Name = "vba_tests"
' deploy: test
'==============================================================================
' vba_tests  -  unit-test harness for the repo's VBA functions/subs.
'
'   run_vba_tests   Runs every test group, writes a PASS/FAIL grid to the
'                   "vba_tests" sheet (created if missing), and pops a summary.
'
' The test CASES live here as text (versioned in vba/vba_tests.bas). The
' "vba_tests" worksheet is generated OUTPUT - regenerated on each run - so the
' committed sheet is just the latest green snapshot, not the source of truth.
'
' Assertion model:
'   grp "name"                 - start a test group (label for following checks)
'   chk case, expected, actual - Describe(actual) must equal the expected string
'   chkTrue case, condition    - condition must be True
'
' Describe() canonicalises any value to a stable string:
'   scalar     -> its text            ("#FF0000", "5", "TRUE")
'   Range      -> its address         ("$B$4")   / "Nothing"
'   error      -> its name            ("#REF!")  (best-effort, cosmetic)
'   2D array   -> "RxC[a|b|c|...]"    (row-major, includes bounds' size)
'   1D array   -> "Nx1D[a|b|...]"
'
' Fixtures: tests that need real cells/sheets create temp "zz_*" sheets and
' delete them, so a run leaves the workbook as it found it.
'
' run_vba_tests is Public so it shows in Alt+F8; that's intended for a dev
' workbook. (Macro-list cleanup is a separate to-do.)
'==============================================================================
Option Explicit

Private gGroup As String
Private gCount As Long
Private gGroups() As String
Private gCases() As String
Private gExpected() As String
Private gActual() As String
Private gPass() As Boolean

' ---- entry point ------------------------------------------------------------
Public Sub run_vba_tests()
    Dim prevSU As Boolean, prevDA As Boolean
    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    Application.ScreenUpdating = False

    gCount = 0
    ReDim gGroups(1 To 2000)
    ReDim gCases(1 To 2000)
    ReDim gExpected(1 To 2000)
    ReDim gActual(1 To 2000)
    ReDim gPass(1 To 2000)

    ' Each group guards its own fixtures; a crash records a failure and moves on.
    RunGuarded "unicode_split"
    RunGuarded "get_color"
    RunGuarded "sheet_names"
    RunGuarded "sheet_data"
    RunGuarded "xconvert"
    RunGuarded "maze_solver"
    RunGuarded "select_sheets"
    RunGuarded "write_bg_color"

    write_results

    Application.ScreenUpdating = prevSU
    Application.DisplayAlerts = prevDA

    Dim passN As Long, i As Long
    For i = 1 To gCount
        If gPass(i) Then passN = passN + 1
    Next i
    MsgBox passN & " / " & gCount & " VBA tests passed." & _
           IIf(passN = gCount, "", vbLf & vbLf & "See the vba_tests sheet for the failures."), _
           IIf(passN = gCount, vbInformation, vbExclamation), "run_vba_tests"
End Sub

Private Sub RunGuarded(ByVal which As String)
    On Error GoTo Failed
    Select Case which
        Case "unicode_split": test_unicode_split
        Case "get_color": test_get_color
        Case "sheet_names": test_sheet_names
        Case "sheet_data": test_sheet_data
        Case "xconvert": test_xconvert
        Case "maze_solver": test_maze_solver
        Case "select_sheets": test_select_sheets
        Case "write_bg_color": test_write_bg_color
    End Select
    Exit Sub
Failed:
    gGroup = which
    record "(group crashed)", "no error", "Error " & Err.Number & ": " & Err.Description, False
    ' best-effort fixture cleanup
    KillSheet "zz_color": KillSheet "zz_ind": KillSheet "zz_fx1": KillSheet "zz_fx2"
    KillSheet "zz_conv": KillSheet "zz_maze": KillSheet "zz_bgc"
End Sub

' ---- assertions -------------------------------------------------------------
Private Sub grp(ByVal name As String)
    gGroup = name
End Sub

Private Sub chk(ByVal caseName As String, ByVal expected As String, ByVal actual As Variant)
    Dim a As String
    a = Describe(actual)
    record caseName, expected, a, (a = expected)
End Sub

Private Sub chkTrue(ByVal caseName As String, ByVal cond As Boolean)
    record caseName, "TRUE", IIf(cond, "TRUE", "FALSE"), cond
End Sub

Private Sub record(ByVal caseName As String, ByVal expected As String, _
                   ByVal actual As String, ByVal passed As Boolean)
    gCount = gCount + 1
    gGroups(gCount) = gGroup
    gCases(gCount) = caseName
    gExpected(gCount) = expected
    gActual(gCount) = actual
    gPass(gCount) = passed
End Sub

' ---- Describe: canonical string for any value -------------------------------
Private Function Describe(ByVal v As Variant) As String
    Dim r As Long, c As Long, rL As Long, rU As Long, cL As Long, cU As Long
    Dim is2D As Boolean, parts() As String, n As Long

    If IsObject(v) Then
        If v Is Nothing Then
            Describe = "Nothing"
        ElseIf TypeOf v Is Range Then
            Describe = v.Address(True, True)
        Else
            Describe = "Object"
        End If
        Exit Function
    End If

    If IsError(v) Then
        Describe = ErrToStr(v)
        Exit Function
    End If

    If IsArray(v) Then
        On Error Resume Next
        rL = LBound(v, 1): rU = UBound(v, 1)
        cL = 0: cU = -1
        cL = LBound(v, 2)
        is2D = (Err.Number = 0)
        If is2D Then cU = UBound(v, 2)
        Err.Clear
        On Error GoTo 0
        If is2D Then
            ReDim parts(0 To (rU - rL + 1) * (cU - cL + 1) - 1)
            n = 0
            For r = rL To rU
                For c = cL To cU
                    parts(n) = ElemStr(v(r, c)): n = n + 1
                Next c
            Next r
            Describe = (rU - rL + 1) & "x" & (cU - cL + 1) & "[" & Join(parts, "|") & "]"
        Else
            ReDim parts(0 To rU - rL)
            n = 0
            For r = rL To rU
                parts(n) = ElemStr(v(r)): n = n + 1
            Next r
            Describe = (rU - rL + 1) & "x1D[" & Join(parts, "|") & "]"
        End If
        Exit Function
    End If

    Describe = ElemStr(v)
End Function

Private Function ElemStr(ByVal e As Variant) As String
    If IsError(e) Then
        ElemStr = ErrToStr(e)
    ElseIf IsEmpty(e) Then
        ElemStr = "<empty>"
    ElseIf IsNull(e) Then
        ElemStr = "<null>"
    ElseIf VarType(e) = vbBoolean Then
        ElemStr = IIf(e, "TRUE", "FALSE")
    Else
        ElemStr = CStr(e)
    End If
End Function

Private Function ErrToStr(ByVal e As Variant) As String
    Dim s As String
    On Error Resume Next
    Select Case True
        Case e = CVErr(xlErrNA): s = "#N/A"
        Case e = CVErr(xlErrValue): s = "#VALUE!"
        Case e = CVErr(xlErrRef): s = "#REF!"
        Case e = CVErr(xlErrName): s = "#NAME?"
        Case e = CVErr(xlErrNum): s = "#NUM!"
        Case e = CVErr(xlErrDiv0): s = "#DIV/0!"
        Case e = CVErr(xlErrNull): s = "#NULL!"
        Case Else: s = "#ERR"
    End Select
    On Error GoTo 0
    If s = "" Then s = "#ERR"
    ErrToStr = s
End Function

' ---- fixture helpers --------------------------------------------------------
Private Function AddSheet(ByVal nm As String) As Worksheet
    KillSheet nm
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = nm
    Set AddSheet = ws
End Function

Private Sub KillSheet(ByVal nm As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(nm).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' ---- test groups ------------------------------------------------------------
Private Sub test_unicode_split()
    grp "UNICODE_SPLIT"
    chk "abc -> 3 rows", "3x1[a|b|c]", UNICODE_SPLIT("abc")
    chk "ab -> 2 rows", "2x1[a|b]", UNICODE_SPLIT("ab")
    chk "empty -> ''", "", UNICODE_SPLIT("")
End Sub

Private Sub test_get_color()
    grp "get_color"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_color")
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    chk "solid red", "#FF0000", get_color(ws.Range("A1"))
    ws.Range("A2").Interior.Pattern = xlPatternNone
    chk "unfilled -> white", "#FFFFFF", get_color(ws.Range("A2"))
    ws.Range("A3").Interior.Color = RGB(0, 0, 0)
    chk "solid black", "#000000", get_color(ws.Range("A3"))
    KillSheet "zz_color"
End Sub

Private Sub test_sheet_names()
    grp "sheet_names"
    Dim res As Variant
    res = sheet_names()
    chkTrue "count matches workbook", (UBound(res, 1) - LBound(res, 1) + 1) = ThisWorkbook.Worksheets.count
    chkTrue "vertical (1 column)", (UBound(res, 2) - LBound(res, 2) + 1) = 1
    chkTrue "first element = sheet 1", CStr(res(LBound(res, 1), LBound(res, 2))) = ThisWorkbook.Worksheets(1).Name
End Sub

Private Sub test_sheet_data()
    grp "sheet_data"
    Dim a As Worksheet, b As Worksheet
    Set a = AddSheet("zz_fx1")
    Set b = AddSheet("zz_fx2")
    a.Range("B2").Value = 1: a.Range("C2").Value = 2
    a.Range("B3").Value = 3: a.Range("C3").Value = 4
    b.Range("B2").Value = 5: b.Range("C2").Value = 6
    b.Range("B3").Value = 7: b.Range("C3").Value = 8

    chk "stacked 2 sheets x B2:C3", _
        "5x4[Sheet|Row|B|C|zz_fx1|2|1|2|zz_fx1|3|3|4|zz_fx2|2|5|6|zz_fx2|3|7|8]", _
        sheet_data(Array("zz_fx1", "zz_fx2"), "B2:C3", 0)

    chkTrue "invalid sheet name -> error", IsError(sheet_data(Array("nope_zz"), "B2:C3", 0))

    KillSheet "zz_fx1"
    KillSheet "zz_fx2"
End Sub

Private Sub test_xconvert()
    grp "xconvert"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_conv")
    ' 3-column table: From | Factor | To
    ws.Range("A1").Value = "m":  ws.Range("B1").Value = 100:  ws.Range("C1").Value = "cm"
    ws.Range("A2").Value = "km": ws.Range("B2").Value = 1000: ws.Range("C2").Value = "m"
    Dim tbl As Range
    Set tbl = ws.Range("A1:C2")

    chk "direct m->cm (2)", "200", xconvert(tbl, "m", 2, "cm")
    chk "reverse cm->m (200)", "2", xconvert(tbl, "cm", 200, "m")
    chk "same unit m->m (5)", "5", xconvert(tbl, "m", 5, "m")
    chk "multi-hop km->cm (2)", "200000", xconvert(tbl, "km", 2, "cm")
    chkTrue "no path -> error", IsError(xconvert(tbl, "m", 1, "xyz"))
    chk "malformed 2-col -> #VALUE!", "#VALUE!", xconvert(ws.Range("A1:B2"), "m", 1, "cm")

    KillSheet "zz_conv"
End Sub

Private Sub test_maze_solver()
    grp "maze_solver"
    Dim ws As Worksheet, rng As Range

    ' no-diagonal, 3x3 all open, start top-left
    Set ws = AddSheet("zz_maze"): ws.Activate
    Set rng = ws.Range("A1:C3")
    rng.Interior.Color = RGB(255, 255, 255)
    ws.Range("A1").Value = 0
    rng.Select
    Application.Run "maze_solver_color_no_diagonal"
    chk "no-diag 3x3", "3x3[0|1|2|1|2|3|2|3|4]", rng.Value
    KillSheet "zz_maze"

    ' with-diagonal, 3x3 all open (Chebyshev distances)
    Set ws = AddSheet("zz_maze"): ws.Activate
    Set rng = ws.Range("A1:C3")
    rng.Interior.Color = RGB(255, 255, 255)
    ws.Range("A1").Value = 0
    rng.Select
    Application.Run "maze_solver_color_with_diagonal"
    chk "with-diag 3x3", "3x3[0|1|2|1|1|2|2|2|2]", rng.Value
    KillSheet "zz_maze"

    ' no-diagonal with a wall at B1:B2 (different color -> impassable, stay empty)
    Set ws = AddSheet("zz_maze"): ws.Activate
    Set rng = ws.Range("A1:C3")
    rng.Interior.Color = RGB(255, 255, 255)
    ws.Range("B1:B2").Interior.Color = RGB(0, 0, 0)
    ws.Range("A1").Value = 0
    rng.Select
    Application.Run "maze_solver_color_no_diagonal"
    chk "no-diag wall", "3x3[0|<empty>|6|1|<empty>|5|2|3|4]", rng.Value
    KillSheet "zz_maze"
End Sub

Private Sub test_select_sheets()
    grp "select_sheets"
    Application.Run "select_first_sheet"
    chkTrue "first sheet active", ActiveSheet.Name = ThisWorkbook.Worksheets(1).Name
    Application.Run "select_last_sheet"
    chkTrue "last sheet active", _
        ActiveSheet.Name = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count).Name
End Sub

Private Sub test_write_bg_color()
    grp "write_background_color"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_bgc"): ws.Activate
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    ws.Range("A2").Interior.Color = RGB(0, 0, 0)
    ws.Range("A1:A2").Select
    Application.Run "write_background_color"
    chk "red -> #FF0000", "#FF0000", ws.Range("A1").Value
    chk "black -> #000000", "#000000", ws.Range("A2").Value
    KillSheet "zz_bgc"
End Sub

' ---- results writer ---------------------------------------------------------
Private Sub write_results()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("vba_tests")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "vba_tests"
    End If

    ws.Cells.Clear
    Dim passN As Long, i As Long
    For i = 1 To gCount
        If gPass(i) Then passN = passN + 1
    Next i

    ws.Range("A1").Value = "VBA Unit Tests"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = passN & " / " & gCount & " PASS"
    ws.Range("A4").Value = "#"
    ws.Range("B4").Value = "Group"
    ws.Range("C4").Value = "Case"
    ws.Range("D4").Value = "Expected"
    ws.Range("E4").Value = "Actual"
    ws.Range("F4").Value = "Result"
    ws.Range("A4:F4").Font.Bold = True

    ' force Expected/Actual columns to text so "#REF!" etc. don't become errors
    ws.Columns("D:E").NumberFormat = "@"

    For i = 1 To gCount
        ws.Cells(4 + i, 1).Value = i
        ws.Cells(4 + i, 2).Value = gGroups(i)
        ws.Cells(4 + i, 3).Value = gCases(i)
        ws.Cells(4 + i, 4).Value = gExpected(i)
        ws.Cells(4 + i, 5).Value = gActual(i)
        ws.Cells(4 + i, 6).Value = IIf(gPass(i), "PASS", "FAIL")
        ws.Cells(4 + i, 6).Interior.Color = IIf(gPass(i), RGB(198, 239, 206), RGB(255, 199, 206))
    Next i

    ws.Columns("A:F").AutoFit
    ws.Activate
End Sub
