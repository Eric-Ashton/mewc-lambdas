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
' Fixtures: each test that needs real cells/sheets builds a persistent "zz_*"
' sheet (refreshed in place each run, added at the end of the tab list) so you
' can inspect what it set up and produced - kept in the workbook like the lambda
' test sheets, next to the "vba_tests" results sheet.
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
Private gMapGroup() As String   ' group name -> the first fixture sheet it built
Private gMapSheet() As String
Private gMapN As Long

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
    ReDim gMapGroup(1 To 200)
    ReDim gMapSheet(1 To 200)
    gMapN = 0

    ' Each group guards its own fixtures; a crash records a failure and moves on.
    RunGuarded "unicode_split"
    RunGuarded "get_color"
    RunGuarded "sheet_names"
    RunGuarded "sheet_data"
    RunGuarded "xconvert"
    RunGuarded "maze_solver"
    RunGuarded "select_sheets"
    RunGuarded "write_bg_color"
    RunGuarded "fill_color"
    RunGuarded "flood_fill"
    RunGuarded "frequency_table"
    RunGuarded "frequency_table_by_char"
    RunGuarded "tall_board"
    RunGuarded "col_num_to_letter"
    RunGuarded "sheet_exists"
    RunGuarded "sanitize_row"
    RunGuarded "count_error_cells"
    RunGuarded "last_used_row_col"
    RunGuarded "find_yellow"
    RunGuarded "get_return_column_number"
    RunGuarded "make_level_data_table"

    write_results
    write_group_blocks        ' per-group Case/Expected/Actual/Result on each sheet

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
        Case "fill_color": test_fill_color
        Case "flood_fill": test_flood_fill
        Case "frequency_table": test_frequency_table
        Case "frequency_table_by_char": test_frequency_table_by_char
        Case "tall_board": test_tall_board
        Case "col_num_to_letter": test_col_num_to_letter
        Case "sheet_exists": test_sheet_exists
        Case "sanitize_row": test_sanitize_row
        Case "count_error_cells": test_count_error_cells
        Case "last_used_row_col": test_last_used_row_col
        Case "find_yellow": test_find_yellow
        Case "get_return_column_number": test_get_return_column_number
        Case "make_level_data_table": test_make_level_data_table
    End Select
    Exit Sub
Failed:
    gGroup = which
    record "(group crashed)", "no error", "Error " & Err.Number & ": " & Err.Description, False
    ' fixtures are left in place (persistent) so a crash can be inspected
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
' Get (or create at the end of the tab list) the persistent fixture sheet, then
' clear it so each run refreshes it in place. The sheet is kept in the workbook.
Private Function bare_sheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set bare_sheet = ws
End Function

' As bare_sheet, but also remembers which sheet the current group built so its
' Case/Expected/Actual/Result block can be written onto it afterwards.
Private Function AddSheet(ByVal nm As String) As Worksheet
    Set AddSheet = bare_sheet(nm)
    record_group_sheet nm
End Function

Private Sub record_group_sheet(ByVal nm As String)
    If Len(gGroup) = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To gMapN
        If gMapGroup(i) = gGroup Then Exit Sub      ' keep the FIRST sheet per group
    Next i
    If gMapN >= 200 Then Exit Sub
    gMapN = gMapN + 1
    gMapGroup(gMapN) = gGroup
    gMapSheet(gMapN) = nm
End Sub

' Deletes a sheet outright (no re-create). Used for the output sheets that
' frequency_table / frequency_table_by_char / tall_board create under their
' own auto-numbered names (Freqs1, FreqsChar1, tall_board, ...) - unlike the
' zz_* fixtures above, we don't want those piling up on every re-run.
Private Sub DeleteSheet(ByVal nm As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(nm).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' Snapshot of current sheet names, and the sheet (if any) that appeared since
' that snapshot was taken - used to find the auto-named output sheet a sub
' just created without hard-coding its name.
Private Function SnapshotSheetNames() As Object
    Dim d As Object, ws As Worksheet
    Set d = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Worksheets
        d(ws.Name) = True
    Next ws
    Set SnapshotSheetNames = d
End Function

Private Function NewSheetSince(before As Object) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not before.Exists(ws.Name) Then
            Set NewSheetSince = ws
            Exit Function
        End If
    Next ws
End Function

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
    ws.Range("C1").Value = "get_color: A1 red / A2 unfilled / A3 black"
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

    ws.Range("A4").Value = "xconvert conversion table (From | Factor | To)"
End Sub

Private Sub test_maze_solver()
    grp "maze_solver"
    ' All three cases laid out on ONE sheet so keep-mode shows them together.
    ' 0 = start, white = path, black = wall. Each grid is solved on its own selection.
    Dim ws As Worksheet
    Set ws = AddSheet("zz_maze"): ws.Activate
    ws.Range("A1").Value = "no-diag open (A2:C4)"
    ws.Range("E1").Value = "with-diag open (E2:G4)"
    ws.Range("A6").Value = "no-diag + wall at B7:B8 (A7:C9)"

    ' Case 1 - 4-directional, open 3x3 -> Manhattan-style BFS distances
    ws.Range("A2:C4").Interior.Color = RGB(255, 255, 255)
    ws.Range("A2").Value = 0
    ws.Range("A2:C4").Select
    Application.Run "maze_solver_color_no_diagonal"
    chk "no-diag open 3x3", "3x3[0|1|2|1|2|3|2|3|4]", ws.Range("A2:C4").Value

    ' Case 2 - 8-directional, open 3x3 -> Chebyshev distances
    ws.Range("E2:G4").Interior.Color = RGB(255, 255, 255)
    ws.Range("E2").Value = 0
    ws.Range("E2:G4").Select
    Application.Run "maze_solver_color_with_diagonal"
    chk "with-diag open 3x3", "3x3[0|1|2|1|1|2|2|2|2]", ws.Range("E2:G4").Value

    ' Case 3 - 4-directional with a wall (B7:B8 black -> impassable, stay empty)
    ws.Range("A7:C9").Interior.Color = RGB(255, 255, 255)
    ws.Range("B7:B8").Interior.Color = RGB(0, 0, 0)
    ws.Range("A7").Value = 0
    ws.Range("A7:C9").Select
    Application.Run "maze_solver_color_no_diagonal"
    chk "no-diag wall", "3x3[0|<empty>|6|1|<empty>|5|2|3|4]", ws.Range("A7:C9").Value
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
    ws.Range("C1").Value = "write_background_color: A1 was red, A2 was black"
End Sub

Private Sub test_fill_color()
    grp "fill_color"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_fillcolor"): ws.Activate

    ws.Range("A1").Value = "X"
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    ws.Range("A2").Interior.Color = RGB(255, 0, 0)     ' blank, same color as source -> filled
    ws.Range("A3").Interior.Color = RGB(0, 0, 255)     ' blank, different color -> stays blank
    ws.Range("B1").Formula = "="""""
    ws.Range("B1").Interior.Color = RGB(255, 0, 0)     ' blank-valued formula cell -> must survive

    ws.Range("A1").Select
    Application.Run "fill_color"

    chk "same-color blank filled with source value", "X", ws.Range("A2").Value
    chk "different-color blank left untouched", "<empty>", ws.Range("A3").Value
    chkTrue "formula cell not overwritten", ws.Range("B1").HasFormula

    ws.Range("D1").Value = _
        "fill_color: A1=""X""/red (source, selected); A2 red (filled); A3 blue (untouched); B1 red formula (untouched)"
End Sub

Private Sub test_flood_fill()
    grp "flood_fill"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_flood"): ws.Activate

    ' basic flood: A1 seed spreads through matching red cells, stops at the blue column
    ws.Range("A1").Value = "X"
    ws.Range("A1:A3").Interior.Color = RGB(255, 0, 0)
    ws.Range("B1:B3").Interior.Color = RGB(0, 0, 255)
    ws.Range("A1:B3").Select
    Application.Run "flood_fill"
    chk "flood-fills a same-color blank neighbor", "X", ws.Range("A2").Value
    chk "flood-fills through a chain of same-color blanks", "X", ws.Range("A3").Value
    chk "stops at a color-mismatched neighbor", "<empty>", ws.Range("B1").Value

    ' border stop: matching color but a border between D1/D2 blocks the flood
    ws.Range("D1").Value = "Y"
    ws.Range("D1:D2").Interior.Color = RGB(0, 255, 0)
    ws.Range("D1").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("D1:D2").Select
    Application.Run "flood_fill"
    chk "stops at a border even with matching color", "<empty>", ws.Range("D2").Value

    ws.Range("F1").Value = "flood_fill: A1:B3 is the color-boundary case; D1:D2 is the border-boundary case"
End Sub

Private Sub test_frequency_table()
    grp "frequency_table"
    Dim ws As Worksheet, newWs As Worksheet, before As Object
    Set ws = AddSheet("zz_freqsrc"): ws.Activate
    ws.Range("A1").Value = "a"
    ws.Range("A2").Value = "a"
    ws.Range("A3").Value = "b"
    ws.Range("A1:A3").Select

    Set before = SnapshotSheetNames()
    Application.Run "frequency_table"
    Set newWs = NewSheetSince(before)

    chkTrue "creates a new Freqs* sheet", Not newWs Is Nothing
    If Not newWs Is Nothing Then
        chkTrue "new sheet name starts with Freqs", Left$(newWs.Name, 5) = "Freqs"
        chk "header A1", "Unique Values", newWs.Range("A1").Value
        chk "header B1", "Counts", newWs.Range("B1").Value
        chk "row2 (sorted desc) value", "a", newWs.Range("A2").Value
        chk "row2 count", "2", newWs.Range("B2").Value
        chk "row3 value", "b", newWs.Range("A3").Value
        chk "row3 count", "1", newWs.Range("B3").Value
        DeleteSheet newWs.Name   ' auto-named output sheet - don't let it pile up on re-runs
    End If
End Sub

Private Sub test_frequency_table_by_char()
    grp "frequency_table_by_char"
    Dim ws As Worksheet, newWs As Worksheet, before As Object
    Set ws = AddSheet("zz_freqcharsrc"): ws.Activate
    ws.Range("A1").Value = "aab"
    ws.Range("A1").Select

    Set before = SnapshotSheetNames()
    Application.Run "frequency_table_by_char"
    Set newWs = NewSheetSince(before)

    chkTrue "creates a new FreqsChar* sheet", Not newWs Is Nothing
    If Not newWs Is Nothing Then
        chkTrue "new sheet name starts with FreqsChar", Left$(newWs.Name, 9) = "FreqsChar"
        chk "row2 char (sorted desc)", "a", newWs.Range("A2").Value
        chk "row2 count", "2", newWs.Range("B2").Value
        chk "row2 unicode (AscW of 'a')", "97", newWs.Range("C2").Value
        chk "row3 char", "b", newWs.Range("A3").Value
        chk "row3 count", "1", newWs.Range("B3").Value
        chk "row3 unicode (AscW of 'b')", "98", newWs.Range("C3").Value
        DeleteSheet newWs.Name
    End If
End Sub

Private Sub test_tall_board()
    grp "tall_board"
    Dim ws As Worksheet, newWs As Worksheet, before As Object
    Set ws = AddSheet("zz_tallsrc"): ws.Activate
    ws.Range("A1").Value = "hi"
    ws.Range("A1").Interior.Color = RGB(255, 0, 0)
    ws.Range("A1").Select

    Set before = SnapshotSheetNames()
    Application.Run "tall_board"
    Set newWs = NewSheetSince(before)

    chkTrue "creates a new tall_board* sheet", Not newWs Is Nothing
    If Not newWs Is Nothing Then
        chkTrue "new sheet name starts with tall_board", Left$(newWs.Name, 10) = "tall_board"
        chk "header A1", "Address", newWs.Range("A1").Value
        chk "row2 Address", "A1", newWs.Range("A2").Value
        chk "row2 Value", "hi", newWs.Range("B2").Value
        chk "row2 Background Color", "#FF0000", newWs.Range("C2").Value
        ' Font Color / Font Name (columns D:E) depend on the workbook's default
        ' font, so they're not asserted here - smoke-tested only (see PR notes).
        DeleteSheet newWs.Name
    End If
End Sub

Private Sub test_col_num_to_letter()
    grp "ColNumToLetter"
    chk "column 1 -> A", "A", ColNumToLetter(1)
    chk "column 26 -> Z", "Z", ColNumToLetter(26)
    chk "column 27 -> AA", "AA", ColNumToLetter(27)
    chk "column 28 -> AB", "AB", ColNumToLetter(28)
End Sub

Private Sub test_sheet_exists()
    grp "SheetExists"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_exists")
    chkTrue "existing sheet -> True", SheetExists(ThisWorkbook, "zz_exists")
    chkTrue "nonexistent sheet -> False", Not SheetExists(ThisWorkbook, "zz_definitely_not_a_sheet")
End Sub

Private Sub test_sanitize_row()
    grp "SanitizeRow"
    Dim arr(1 To 1, 1 To 3) As Variant
    arr(1, 1) = "ok": arr(1, 2) = CVErr(xlErrValue): arr(1, 3) = 5
    SanitizeRow arr
    chk "2D row: error element -> ''", "1x3[ok||5]", arr

    Dim v As Variant
    v = CVErr(xlErrDiv0)
    SanitizeRow v
    chk "scalar error -> ''", "", v
End Sub

Private Sub test_count_error_cells()
    grp "CountErrorCells"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_errcells")
    ws.Range("A1").Value = "ok"
    ws.Range("A2").Formula = "=1/0"        ' #DIV/0!
    ws.Range("B1").Formula = "=NA()"       ' #N/A
    ws.Range("B2").Value = 5
    ws.Calculate
    chkTrue "counts the 2 error cells in UsedRange", CountErrorCells(ws) = 2
End Sub

Private Sub test_last_used_row_col()
    grp "GetLastUsedRow/Col"
    ' Both helpers use an unqualified [a1] as the Find "After" anchor, which
    ' resolves to the ACTIVE sheet's A1 - so ws must be active or Find errors
    ' (After must be a cell within the range being searched). Activating the
    ' fixture here documents that requirement.
    Dim ws As Worksheet
    Set ws = AddSheet("zz_lastused"): ws.Activate
    ws.Range("B2").Value = "x"
    ws.Range("D5").Value = "y"
    chkTrue "last used row = 5", GetLastUsedRow(ws) = 5
    chkTrue "last used col = D (4)", GetLastUsedCol(ws) = 4
End Sub

Private Sub test_find_yellow()
    grp "find_yellow"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_yellow")
    ws.Range("B5").Value = "x"
    ws.Range("B5").Interior.Color = vbYellow
    chkTrue "finds the yellow cell's row in column B", find_yellow(ws) = 5
End Sub

Private Sub test_get_return_column_number()
    grp "GetReturnColumnNumber"
    Dim ws As Worksheet
    Set ws = AddSheet("zz_retcol")
    ' Text-format the cells before writing so Excel stores the "=..." text
    ' literally instead of parsing/calculating it as a live formula.
    ' GetReturnColumnNumber only ever reads .Formula as a STRING - it never
    ' needs the referenced sheet/range to actually exist or resolve, and a
    ' live XLOOKUP referencing a "Case!" sheet this workbook doesn't have
    ' triggered an "Update Values" external-link prompt when it WAS a real
    ' formula. Plain text sidesteps that entirely.
    ws.Range("C1:C2").NumberFormat = "@"
    ws.Range("C1").Value = "=XLOOKUP(A2,Case!C1:C369,Case!H1:H369)"
    ' LIKELY BUG (button_subs.bas, ~line 235): "colLetters = colLetters Like
    ' ")*" Or colLetters Like """"" assigns a Boolean to the String colLetters,
    ' which VBA coerces to the literal text "False" (or "True") - clobbering
    ' whatever the real column letters were. Only the uppercase "F" in "False"
    ' survives the letters-only filter, so for any normal formula this
    ' currently returns column F (6) regardless of the real target column.
    ' Characterizing the CURRENT (buggy) behavior as a baseline, not the
    ' intended one - see the PR notes.
    chk "XLOOKUP formula (real target is H) -> currently always F (bug)", "6", GetReturnColumnNumber(ws, 1)

    ws.Range("C2").Value = "=SUM(A1)"      ' fewer than 2 colons -> 0, unaffected by the bug above
    chk "formula with <2 colons -> 0", "0", GetReturnColumnNumber(ws, 2)

    ws.Range("E1").Value = "GetReturnColumnNumber: C1 has a realistic XLOOKUP formula; C2 has too few colons"
End Sub

Private Sub test_make_level_data_table()
    grp "make_level_data_table"
    Dim ws As Worksheet, newWs As Worksheet, before As Object
    Set ws = AddSheet("zz_tmsrc"): ws.Activate
    ws.Range("A1").Value = "Game": ws.Range("B1").Value = "Answer"
    ws.Range("A2").Value = 101: ws.Range("B2").Value = "foo"
    ws.Range("A1:B2").Select

    Set before = SnapshotSheetNames()
    Application.Run "make_level_data_table"
    Set newWs = NewSheetSince(before)

    chkTrue "creates a new _Tn sheet", Not newWs Is Nothing
    If Not newWs Is Nothing Then
        chkTrue "new sheet name starts with _T", Left$(newWs.Name, 2) = "_T"
        chk "header A1 copied", "Game", newWs.Range("A1").Value
        chk "header B1 copied", "Answer", newWs.Range("B1").Value
        chk "A2 = first game number", "101", newWs.Range("A2").Value
        chk "B4 yellow fill (drop target formula here)", "#FFFF00", get_color(newWs.Range("B4"))
        chk "A5 = game numbers list starts here", "101", newWs.Range("A5").Value
        ' Excel only quotes a sheet name in stored formula text when the name
        ' needs escaping (spaces, etc.) - "zz_tmsrc" doesn't, so despite the
        ' sub writing the quoted form, .Formula2 reads back unquoted.
        chk "B2 XLOOKUP formula spills the matching row", _
            "=XLOOKUP(A2,zz_tmsrc!$A$2,zz_tmsrc!$B$2)", newWs.Range("B2").Formula2
        DeleteSheet newWs.Name
    End If
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

' ---- per-group result blocks ------------------------------------------------
' Give every group its own Case | Expected | Actual | Result table on its sheet,
' so each sheet reads like a lambda test sheet. The block is anchored just to the
' right of whatever the fixture already used, so it can never overwrite it.
Private Sub write_group_blocks()
    Dim i As Long, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 1 To gCount
        If Not seen.Exists(gGroups(i)) Then
            seen.Add gGroups(i), True
            write_one_group gGroups(i)
        End If
    Next i
End Sub

Private Sub write_one_group(ByVal groupName As String)
    Dim ws As Worksheet
    Set ws = group_sheet(groupName)
    If ws Is Nothing Then Exit Sub

    Dim anchor As Long
    anchor = ws.UsedRange.Column + ws.UsedRange.Columns.count + 1
    If anchor < 1 Then anchor = 1

    ws.Cells(1, anchor).Value = groupName
    ws.Cells(1, anchor).Font.Bold = True
    ws.Cells(2, anchor).Value = "Case"
    ws.Cells(2, anchor + 1).Value = "Expected"
    ws.Cells(2, anchor + 2).Value = "Actual"
    ws.Cells(2, anchor + 3).Value = "Result"
    ws.Range(ws.Cells(2, anchor), ws.Cells(2, anchor + 3)).Font.Bold = True
    ' keep "#REF!" and friends as text, not errors
    ws.Range(ws.Cells(1, anchor + 1), ws.Cells(1, anchor + 2)).EntireColumn.NumberFormat = "@"

    Dim i As Long, r As Long, passN As Long, totN As Long
    r = 3
    For i = 1 To gCount
        If gGroups(i) = groupName Then
            ws.Cells(r, anchor).Value = gCases(i)
            ws.Cells(r, anchor + 1).Value = gExpected(i)
            ws.Cells(r, anchor + 2).Value = gActual(i)
            ws.Cells(r, anchor + 3).Value = IIf(gPass(i), "PASS", "FAIL")
            ws.Cells(r, anchor + 3).Interior.Color = _
                IIf(gPass(i), RGB(198, 239, 206), RGB(255, 199, 206))
            totN = totN + 1
            If gPass(i) Then passN = passN + 1
            r = r + 1
        End If
    Next i

    ws.Cells(1, anchor + 3).Value = passN & " / " & totN & " PASS"
    ws.Cells(1, anchor + 3).Font.Bold = True
    ws.Range(ws.Cells(1, anchor), ws.Cells(1, anchor + 3)).EntireColumn.AutoFit
End Sub

' The sheet a group built (via AddSheet). Groups with no fixture of their own get
' a results-only sheet named after the group.
Private Function group_sheet(ByVal groupName As String) As Worksheet
    Dim i As Long
    For i = 1 To gMapN
        If gMapGroup(i) = groupName Then
            On Error Resume Next
            Set group_sheet = ThisWorkbook.Worksheets(gMapSheet(i))
            On Error GoTo 0
            If Not group_sheet Is Nothing Then Exit Function
        End If
    Next i
    Set group_sheet = bare_sheet("zz_" & safe_sheet_name(groupName))
End Function

' Excel sheet names can't contain : \ / ? * [ ] and cap at 31 characters.
Private Function safe_sheet_name(ByVal s As String) As String
    Dim bad As Variant, v As Variant
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each v In bad
        s = Replace(s, CStr(v), "_")
    Next v
    If Len(s) > 28 Then s = Left$(s, 28)
    safe_sheet_name = s
End Function
