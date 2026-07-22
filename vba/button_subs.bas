Attribute VB_Name = "button_subs"
' deploy: shared
'========================
' Subroutine: what_if_on
'========================
' Applies a What-If Data Table starting from the yellow row.
'
' Assumes the input cell is 5 rows above the yellow row (same column B).

Public Sub what_if_on(Optional ByVal ignore As Variant)   ' Optional arg hides it from Alt+F8; still button/Run-callable
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim yellow_cell_row_n As Long
    Dim last_what_if_row As Long
    Dim table_range As Range
    Dim input_cell As Range
    
    ThisWorkbook.Save
    yellow_cell_row_n = find_yellow(ws)
    If yellow_cell_row_n = 0 Then Exit Sub

    ' Find bottom row of data table
    With ws.Cells(yellow_cell_row_n, "B").Offset(1, -1)
        last_what_if_row = IIf(IsEmpty(.Value), .Row, .End(xlDown).Row)
    End With

    ' Define and apply the What-If Data Table
    Set table_range = ws.Range("A" & yellow_cell_row_n & ":B" & last_what_if_row)
    Set input_cell = ws.Cells(yellow_cell_row_n - 5, 2)
    table_range.Table ColumnInput:=input_cell
    
End Sub

'========================
' Subroutine: what_if_off
'========================
' Clears the output column (B) of the What-If Data Table starting from the yellow row + 1
' down to the end of contiguous data.

Public Sub what_if_off(Optional ByVal ignore As Variant)   ' Optional arg hides it from Alt+F8; still button/Run-callable
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim yellow_cell_row_n As Long
    Dim clear_start As Range
    Dim last_row As Long

    yellow_cell_row_n = find_yellow(ws)
    If yellow_cell_row_n = 0 Then Exit Sub

    Set clear_start = ws.Cells(yellow_cell_row_n + 1, "B")
    last_row = clear_start.End(xlDown).Row
    ws.Range(ws.Cells(clear_start.Row, "B"), ws.Cells(last_row, "B")).ClearContents
    
End Sub


'========================
' Subroutine: copy_previous
'========================
' Build rectangles from the SOURCE (previous level) sheet using the shared
' anchor (case_copy!Z1). Copy to the DESTINATION (current) sheet at the
' exact same addresses. Prompts before overwrite if destination has any
' real values (ignores a single space " ").

Public Sub copy_previous(Optional ByVal ignore As Variant)   ' Optional arg hides it from Alt+F8; still button/Run-callable
    Dim ws As Worksheet: Set ws = ActiveSheet                 ' destination (current level)
    Dim ws_prev As Worksheet                                  ' source (previous level)
    Dim level_number As Long, prev_level_name As String

    Dim anchorText As String
    Dim srcAnchor As Range                                    ' anchor on source sheet
    Dim yRow As Long, hdrRow As Long, hdrCol As Long

    Dim lastRowS As Long, lastColS As Long                    ' TRUE last row/col on source
    Dim r1Src As Range, r2Src As Range, r3Src As Range        ' source rectangles
    Dim r1Dst As Range, r2Dst As Range, r3Dst As Range        ' destination rectangles

    Dim hasData As Boolean, resp As VbMsgBoxResult

    ' --- find previous level sheet (source) ---
    On Error GoTo InvalidLevel
    level_number = CLng(Mid$(ws.Name, 3))
    If level_number = 1 Then
        MsgBox "There is no previous level to copy from (you are on Level 1).", vbExclamation
        Exit Sub
    End If
    prev_level_name = "_L" & (level_number - 1)
    On Error GoTo 0

    On Error Resume Next
    Set ws_prev = ThisWorkbook.Sheets(prev_level_name)
    On Error GoTo 0
    If ws_prev Is Nothing Then
        MsgBox "Previous level sheet '" & prev_level_name & "' not found.", vbCritical
        Exit Sub
    End If

    ' --- resolve the shared anchor ON THE SOURCE SHEET ---
    On Error Resume Next
    If case_copy Is Nothing Then Set case_copy = ThisWorkbook.Worksheets("case copy")
    On Error GoTo 0
    If Not case_copy Is Nothing Then anchorText = CStr(case_copy.Range("Z1").Value)

    If Len(anchorText) > 0 Then
        Set srcAnchor = ws_prev.Range(anchorText)
    Else
        ' Fallback: compute from the source sheet geometry
        yRow = find_yellow(ws_prev)
        If yRow = 0 Then
            MsgBox "Can't find yellow anchor on the previous sheet; run setup/align first.", vbExclamation
            Exit Sub
        End If
        hdrRow = yRow - 6
        hdrCol = Application.Max(1, last_col_all_levels + 1)
        Set srcAnchor = ws_prev.Cells(hdrRow, hdrCol)
    End If

    ' --- TRUE last row/col on SOURCE sheet ---
    Call TrueLastRowCol(ws_prev, lastRowS, lastColS)
    If lastColS < srcAnchor.Column Then lastColS = srcAnchor.Column
    If lastRowS < (srcAnchor.Row + 6) Then lastRowS = srcAnchor.Row + 6

    ' --- Build rectangles ON SOURCE ---
    ' rect1: starts at anchor, 2 rows tall, extends to true last column
    Set r1Src = ws_prev.Range(ws_prev.Cells(srcAnchor.Row, srcAnchor.Column), _
                              ws_prev.Cells(srcAnchor.Row + 1, lastColS))
    ' rect2: column A, 4 rows tall, directly below rect1, extends to true last column
    Set r2Src = ws_prev.Range(ws_prev.Cells(srcAnchor.Row + 2, 1), _
                              ws_prev.Cells(srcAnchor.Row + 5, lastColS))
    ' rect3: column C, from below rect2 to true last row/col
    Set r3Src = ws_prev.Range(ws_prev.Cells(srcAnchor.Row + 6, 3), _
                              ws_prev.Cells(lastRowS, lastColS))

    ' --- Mirror rectangles ON DESTINATION using the SAME addresses ---
    Set r1Dst = ws.Range(r1Src.Address(False, False))
    Set r2Dst = ws.Range(r2Src.Address(False, False))
    Set r3Dst = ws.Range(r3Src.Address(False, False))

    ' --- overwrite check on destination (ignoring single-space " ") ---
    hasData = HasRealValues(r1Dst) Or HasRealValues(r2Dst) Or HasRealValues(r3Dst)
    If hasData Then
        resp = MsgBox("The solution area already has content." & vbCrLf & _
                      "Overwrite it?", vbYesNo + vbQuestion)
        If resp <> vbYes Then Exit Sub
    End If

    ' --- Copy: formats + formulas ---
    Application.CutCopyMode = False

    r1Src.Copy: r1Dst.PasteSpecial Paste:=xlPasteAll
    r2Src.Copy: r2Dst.PasteSpecial Paste:=xlPasteAll
    r3Src.Copy: r3Dst.PasteSpecial Paste:=xlPasteAll

    Application.CutCopyMode = False

    ' --- Also copy the yellow-cell formula (same row, column B) ---
    yRow = find_yellow(ws_prev)
    If yRow > 0 Then ws.Cells(yRow, 2).Formula = ws_prev.Cells(yRow, 2).Formula

    ' --- Put the cursor at the start of rect1 ---
    ws.Activate
    r1Dst.Cells(1, 1).Select
    Exit Sub

InvalidLevel:
    MsgBox "Could not determine level number from sheet name: '" & ws.Name & "'", vbCritical
End Sub




'========================
' Function: find_yellow (helper)
'========================


Function find_yellow(ws As Worksheet) As Long
    Dim foundCell As Range
    
    ' Clear previous format criteria
    Application.FindFormat.Clear
    ' Set search criteria to Yellow fill
    Application.FindFormat.Interior.Color = vbYellow
    
    ' Find the cell in Column B
    Set foundCell = ws.Columns("B").Find(What:="*", _
                                         SearchDirection:=xlNext, _
                                         SearchFormat:=True)
                                         
    ' Fallback: standard Find sometimes misses empty cells if looking for "*",
    ' so if that failed, try searching for empty string with format
    If foundCell Is Nothing Then
         Set foundCell = ws.Columns("B").Find(What:="", _
                                         SearchDirection:=xlNext, _
                                         SearchFormat:=True)
    End If

    If Not foundCell Is Nothing Then
        find_yellow = foundCell.Row
    Else
        MsgBox "Could not find a yellow cell in Column B.", vbExclamation
        find_yellow = 0
    End If
    
    ' Clean up
    Application.FindFormat.Clear
End Function


'========================
' Function: GetReturnColumnNumber (Helper)
'========================
' Takes an XLOOKUP string from column C of the given row and returns the column number of the return column

Function GetReturnColumnNumber(ws As Worksheet, xlookup_row As Long) As Long
    Dim f As String
    Dim parts As Variant
    Dim refPart As String
    Dim colLetters As String
    
    ' Get the formula from column C of the given row
    f = ws.Cells(xlookup_row, "C").Formula
    
    ' Example formula contains "...Case!C1:H369..."
    ' Split by ":" ? we want the 2nd item after the second colon
    parts = Split(f, ":")
    If UBound(parts) < 2 Then
        GetReturnColumnNumber = 0
        Exit Function
    End If
    
    ' Take the right side of the second colon (like "H369)")
    refPart = parts(2)
    
    ' Clean it: remove row numbers and any trailing ) or quotes
    colLetters = UCase$(refPart)
    colLetters = colLetters Like ")*" Or colLetters Like """" ' extra cleaning
    colLetters = colLetters
    colLetters = Replace(colLetters, ")", "")
    colLetters = Replace(colLetters, """", "")
    colLetters = Replace(colLetters, " ", "")
    
    ' Remove digits (row numbers), keep only letters
    Dim i As Long, ch As String, lettersOnly As String
    For i = 1 To Len(colLetters)
        ch = Mid$(colLetters, i, 1)
        If ch >= "A" And ch <= "Z" Then
            lettersOnly = lettersOnly & ch
        End If
    Next
    
    ' Convert letters to number
    GetReturnColumnNumber = Range(lettersOnly & "1").Column
End Function

' Helper: Returns True if any cell has non-empty *content* (ignores pure whitespace like " ")
Private Function HasRealValues(ByVal rng As Range) As Boolean
    Dim c As Range
    If rng Is Nothing Then Exit Function
    For Each c In rng.Cells
        If Not IsError(c.Value) Then
            If Len(c.Value) > 0 Then
                If Trim$(CStr(c.Value)) <> "" Then
                    HasRealValues = True
                    Exit Function
                End If
            End If
        Else
            ' an error is still "real content"
            HasRealValues = True
            Exit Function
        End If
    Next c
End Function

' Helper: True last row/col with any real content (constants or formulas).
' Ignores formatting-only used-range bloat; counts formula-blanks ("").
Private Sub TrueLastRowCol(ByVal ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim f As Range

    lastRow = 1
    lastCol = 1

    ' Last row that has anything (value or formula)
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious)
    If Not f Is Nothing Then lastRow = f.Row

    ' Last column that has anything (value or formula)
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                          SearchDirection:=xlPrevious)
    If Not f Is Nothing Then lastCol = f.Column
End Sub




'========================
' Subroutine: done
'========================
Public Sub done(Optional ByVal ignore As Variant)   ' Optional arg hides it from Alt+F8; still button/Run-callable
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim case_sheet As Worksheet
    Dim yellow_cell_row_n As Long
    Dim answers_start As Range
    Dim last_row As Long
    Dim first_game_number As Long
    Dim case_start_cell As Range
    Dim answer_range As Range
    Dim answers_array As Variant
    Dim level_number As Long
    Dim next_sheet_name As String
    Dim next_sheet As Worksheet
    Dim cell As Range
    Dim first_val As Variant
    Dim all_same As Boolean

    ' ThisWorkbook.Save ' Optional: Saving often slows things down too, disable if confident
    Application.CutCopyMode = False

    ' Find the yellow cell
    yellow_cell_row_n = find_yellow(ws)
    If yellow_cell_row_n = 0 Then Exit Sub

    ' If the cell below the yellow cell is empty, run what_if_on
    If IsEmpty(ws.Cells(yellow_cell_row_n + 1, "B").Value) Then
        what_if_on
    End If

    ' Get the first game number
    first_game_number = ws.Cells(yellow_cell_row_n, "B").Offset(1, -1).Value

    ' Define answer start and last row
    Set answers_start = ws.Cells(yellow_cell_row_n + 1, "B")
    last_row = answers_start.End(xlDown).Row

    ' Get the "Case" sheet
    On Error Resume Next
    Set case_sheet = ThisWorkbook.Sheets("Case")
    On Error GoTo 0
    If case_sheet Is Nothing Then
        MsgBox "Sheet named 'Case' not found.", vbCritical
        Exit Sub
    End If

    ' Find the match in 'Case' sheet
    Set case_start_cell = case_sheet.Columns("B").Find(What:=first_game_number, LookIn:=xlValues, LookAt:=xlWhole)
    If case_start_cell Is Nothing Then
        MsgBox "Could not find game number (" & first_game_number & ") in 'Case' sheet.", vbCritical
        Exit Sub
    End If

    ' Validation checks
    If case_start_cell.Offset(1, 0).Value <> first_game_number + 1 Then
        MsgBox "Expected " & (first_game_number + 1) & " below " & first_game_number & " in 'Case' sheet.", vbCritical
        Exit Sub
    End If

    ' Capture answers into memory (Reading is fast)
    Set answer_range = ws.Range(ws.Cells(answers_start.Row, "B"), ws.Cells(last_row, "B"))
    answers_array = answer_range.Value

    ' Check for errors and sameness
    first_val = answer_range.Cells(1, 1).Value
    all_same = True
    For Each cell In answer_range
        If IsError(cell.Value) Then
            MsgBox "Error(s) found in answers", vbCritical
            Exit Sub
        End If
        If cell.Value <> first_val Then
            all_same = False
        End If
    Next cell

    If all_same Then
        MsgBox "All answers are the same. Check your model.", vbCritical
        Exit Sub
    End If
    
    ' Scroll to answers and pause for 2 seconds
    ActiveWindow.ScrollRow = yellow_cell_row_n
    Application.Wait Now + TimeValue("00:00:02")

    
    ' 1. Freeze Calculation
    ' This prevents the "Write" operation below from triggering a recalc of the heavy Data Table
    Application.Calculation = xlCalculationManual
    
    ' 2. Paste values into Case sheet.
    '    Copy + PasteSpecial xlPasteValues, NOT ".Value = answers_array":
    '    assigning a string like "12 - 16 - 18" to a General-formatted cell makes
    '    Excel re-parse it, which silently turned that answer into the date
    '    12/16/2018. PasteSpecial transfers the stored value without re-parsing,
    '    and pastes values only so the Case sheet keeps its own formatting.
    answer_range.Copy
    case_sheet.Cells(case_start_cell.Row, "E").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' 3. Kill the Data Table (Now instantaneous because Calc is off)
    what_if_off
    
    ' 4. Restore Calculation
    ' Since the Data Table is now gone (cleared by step 3), this will not trigger the 1-minute wait.
    Application.Calculation = xlCalculationAutomatic
    

    ' Copy to clipboard (Must be done AFTER restoring Calc, as switching modes can clear clipboard)
    case_sheet.Cells(case_start_cell.Row, "E").Resize(UBound(answers_array, 1), 1).Copy

    ' Navigation and visuals
    ws.Tab.Color = RGB(0, 100, 0) ' Green for done
    
    On Error GoTo SkipNavigation
    ' Parse level number from sheet name ("_L1", "_L12", "_L20", ...)
    ' Use Mid$(ws.Name, 3) to take everything after the "_L" prefix so we
    ' correctly handle multi-digit level numbers up to _L20 (previously a
    ' vestigial Right(ws.Name, 1) call returned only the last digit and
    ' would have produced the wrong next-sheet name for levels > 9).
    level_number = CLng(Mid$(ws.Name, 3))

    next_sheet_name = "_L" & (level_number + 1)

    ' Activate next sheet
    Set next_sheet = ThisWorkbook.Sheets(next_sheet_name)
    next_sheet.Tab.Color = RGB(255, 255, 0) ' Yellow for active
    next_sheet.Select
    
    Exit Sub

SkipNavigation:
    ' If next sheet fails, just ensure we have the copy
    case_sheet.Cells(case_start_cell.Row, "E").Resize(UBound(answers_array, 1), 1).Copy
End Sub
