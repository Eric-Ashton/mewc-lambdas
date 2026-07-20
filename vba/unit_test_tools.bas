Attribute VB_Name = "unit_test_tools"
' deploy: test
'==============================================================================
' unit_test_tools  -  tooling for the Excel unit-test workbook (the UPSTREAM /
'                     dev side of the repo). The old manage_lambda module was
'                     folded in here so the whole test loop lives in one place.
'
'   lambda_update       Push the Lamb sheet into the Name Manager (delete + re-add
'                       each name from its Code cell, set the Name-Manager comment
'                       from Comment) and THEN rewrite the test-sheet formulas as
'                       dynamic arrays. One button for the test loop: edit lambdas
'                       -> lambda_update -> re-tested.
'
'   link_test_headers   Make the Lamb sheet the single source of truth. On every
'                       test sheet it replaces the hard-coded Signature (B1),
'                       Description (I1) and Code (B2) with formulas that look
'                       them up from the Lamb sheet by the tab name (= lambda
'                       name). Re-run after adding new test sheets.
'   fit_test_row_heights  Size row 1 (Description, I1) and row 2 (Code, B2) on
'                       each test sheet to fit their content, since merged cells
'                       don't auto-fit. Run after link_test_headers.
'   fix_test_formulas   Rewrite every test-sheet formula as a dynamic array
'                       (via .Formula2): strips the implicit-intersection "@"
'                       and converts legacy CSE {array} formulas to spilling.
'   sort_test_tabs      Restore tab order: Prep, Lamb, lambda_tests, vba_tests
'                       pinned first, everything else alphabetical.
'
'   Which sheets count as lambda test sheets is decided in one place,
'   is_lambda_test_sheet - the four fixed sheets and the zz_* VBA fixtures are
'   not test sheets and must never get header lookups written onto them.
'
'   The header lookup: CELL("filename",A1) gives the tab name; that name + "("
'   is wildcard-matched against the Lamb Signature column, so `set` finds
'   `set(...)` and never `set_d(...)`. Returns Lamb col A / D / C.
'
'   NOTE: this is the *upstream* test workbook's tooling (lambda_update also
'   refreshes the test sheets). The downstream template workbook is a consumer
'   only and will get its own, simpler update sub later.
'==============================================================================
Option Explicit

Private Const LAMB_SHEET As String = "Lamb"
Private Const COL_SIG As Long = 1
Private Const COL_COMMENT As Long = 2
Private Const COL_CODE As Long = 3

'==============================================================================
' Lambda library management  (merged in from manage_lambda)
'==============================================================================

' Read every row of "Lamb", (re)create the defined name from its Code cell, set
' the Name-Manager comment from Comment, then rewrite the test-sheet formulas so
' they exercise the just-updated lambdas.
Public Sub lambda_update()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim names() As String, codes() As String, cmts() As String
    Dim added() As Boolean
    Dim addedCount As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LAMB_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & LAMB_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_SIG).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ReDim names(2 To lastRow)
    ReDim codes(2 To lastRow)
    ReDim cmts(2 To lastRow)
    ReDim added(2 To lastRow)

    ' --- collect rows; delete any existing name with the same name ---
    For r = 2 To lastRow
        Dim nm As String, code As String
        nm = LambdaName(CStr(ws.Cells(r, COL_SIG).Value))
        code = CStr(ws.Cells(r, COL_CODE).Value)
        If Left$(code, 1) = "'" Then code = Mid$(code, 2)   ' drop a literal leading apostrophe if present

        If Len(nm) > 0 And Left$(LTrim$(code), 1) = "=" Then
            names(r) = nm
            codes(r) = code
            cmts(r) = CStr(ws.Cells(r, COL_COMMENT).Value)
            On Error Resume Next
            ThisWorkbook.names(nm).Delete
            On Error GoTo 0
        Else
            names(r) = ""       ' blank/invalid row - skip
            added(r) = True
        End If
    Next r

    ' --- add names; repeat passes so a lambda that calls another still resolves ---
    Dim pass As Long, progress As Boolean
    For pass = 1 To 25
        progress = False
        For r = 2 To lastRow
            If Not added(r) And Len(names(r)) > 0 Then
                On Error Resume Next
                Err.Clear
                ThisWorkbook.names.Add Name:=names(r), RefersTo:=codes(r)
                If Err.Number = 0 Then
                    added(r) = True
                    progress = True
                    addedCount = addedCount + 1
                End If
                On Error GoTo 0
            End If
        Next r
        If Not progress Then Exit For
    Next pass

    ' --- set the Name Manager comment on each name that added ---
    For r = 2 To lastRow
        If added(r) And Len(names(r)) > 0 And Len(cmts(r)) > 0 Then
            On Error Resume Next
            ThisWorkbook.names(names(r)).Comment = Left$(cmts(r), 255)
            On Error GoTo 0
        End If
    Next r

    ' --- list any that never added (genuine formula errors) ---
    Dim failed As String
    For r = 2 To lastRow
        If Not added(r) And Len(names(r)) > 0 Then
            failed = failed & vbLf & "  " & names(r)
        End If
    Next r

    ' --- refresh the test sheets against the just-updated lambdas ---
    Dim fixed As Long
    fixed = rewrite_test_formulas()

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    Dim msg As String
    msg = addedCount & " lambda(s) updated; " & fixed & " test formula(s) rewritten."
    If Len(failed) > 0 Then
        MsgBox msg & vbLf & vbLf & "Could not add (check the Code):" & failed, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
End Sub

' Name = text before "(" in the signature (whole string if there is no "(").
Private Function LambdaName(ByVal sig As String) As String
    Dim p As Long
    sig = Trim$(sig)
    p = InStr(sig, "(")
    If p > 0 Then sig = Left$(sig, p - 1)
    LambdaName = Trim$(sig)
End Function

'==============================================================================
' Test-sheet headers  (single-source-of-truth lookups against Lamb)
'==============================================================================

' True only for a lambda's own test sheet. Excludes the four fixed sheets and
' the zz_* VBA fixtures, which have no lambda of that name and whose B1/B2/I1
' are real fixture data - writing header lookups onto them would clobber it.
' Centralised so the callers below cannot drift apart.
Public Function is_lambda_test_sheet(ByVal sheet_name As String) As Boolean
    Select Case sheet_name
        Case "Prep", "Lamb", "lambda_tests", "vba_tests"
            is_lambda_test_sheet = False
        Case Else
            is_lambda_test_sheet = (LCase$(Left$(sheet_name, 3)) <> "zz_")
    End Select
End Function

Public Sub link_test_headers()
    Dim ws As Worksheet, n As Long
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If is_lambda_test_sheet(ws.Name) Then
            set_hdr ws.Range("B1"), "A"   ' Signature   <- Lamb col A
            set_hdr ws.Range("I1"), "D"   ' Description  <- Lamb col D
            set_hdr ws.Range("B2"), "C"   ' Code (text)  <- Lamb col C
            n = n + 1
        End If
    Next ws
    Application.ScreenUpdating = True
    MsgBox "Linked headers on " & n & " test sheet(s) to the Lamb sheet.", _
           vbInformation, "link_test_headers"
End Sub

Private Sub set_hdr(target As Range, ByVal lambCol As String)
    target.NumberFormat = "General"          ' drop any text/quote-prefix state
    target.Formula2 = hdr_formula(lambCol)   ' Formula2 -> enters cleanly, no "@"
End Sub

Private Function hdr_formula(ByVal lambCol As String) As String
    Dim q As String
    q = Chr$(34)                             ' a double-quote character
    hdr_formula = _
        "=LET(fn, CELL(" & q & "filename" & q & ",$A$1), " & _
        "nm, MID(fn, FIND(" & q & "]" & q & ", fn) + 1, 255), " & _
        "XLOOKUP(nm & " & q & "(*" & q & ", Lamb!$A$2:$A$1000, " & _
        "Lamb!$" & lambCol & "$2:$" & lambCol & "$1000, " & q & q & ", 2))"
End Function

Public Sub fit_test_row_heights()
    ' Merged cells don't auto-fit, so size row 1 (Description in I1) and row 2
    ' (Code in B2) by counting their hard line breaks x 21.
    Const LINE_H As Double = 21
    Dim ws As Worksheet, n As Long
    Application.ScreenUpdating = False
    Application.Calculate                        ' make sure the lookups are current
    For Each ws In ThisWorkbook.Worksheets
        If is_lambda_test_sheet(ws.Name) Then
            set_row_height ws.Rows(1), LINE_H * display_lines(ws.Range("I1"))
            set_row_height ws.Rows(2), LINE_H * display_lines(ws.Range("B2"))
            n = n + 1
        End If
    Next ws
    Application.ScreenUpdating = True
    MsgBox "Sized row 1 (description) & row 2 (code) on " & n & " test sheet(s).", _
           vbInformation, "fit_test_row_heights"
End Sub

Private Function display_lines(target As Range) As Long
    ' Count hard line breaks in the cell (+1). The merged header cells are wide
    ' enough that soft wrapping is rare; estimating it over-counted the height.
    If IsError(target.Value) Then display_lines = 1: Exit Function
    Dim s As String
    s = Replace(CStr(target.Value), vbCr, "")
    If Len(s) = 0 Then display_lines = 1: Exit Function
    display_lines = Len(s) - Len(Replace(s, vbLf, "")) + 1
End Function

Private Sub set_row_height(r As Range, ByVal h As Double)
    ' Excel's max row height is 409.5 pt; clamp so long code doesn't 1004-error.
    If h < 15 Then h = 15
    If h > 409 Then h = 409
    On Error Resume Next          ' skip a protected/odd sheet rather than halt
    r.RowHeight = h
    On Error GoTo 0
End Sub

'==============================================================================
' Fix test formulas  (dynamic-array rewrite)
'==============================================================================

Public Sub fix_test_formulas()
    Dim changed As Long
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    changed = rewrite_test_formulas()

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    MsgBox "Rewrote " & changed & " formula(s) as dynamic arrays.", _
           vbInformation, "fix_test_formulas"
End Sub

' Core worker shared by fix_test_formulas and lambda_update. Rewrites each
' test-sheet formula via .Formula2 (stripping "@" and un-CSE'ing arrays) and
' returns the number of anchors rewritten. The CALLER owns Calculation /
' ScreenUpdating / messaging (both callers set Calculation manual first).
Private Function rewrite_test_formulas() As Long
    Dim ws As Worksheet
    Dim fcells As Range, cell As Range, a As Range, an As Range
    Dim anchors As Collection, seen As Object
    Dim ftext As String
    Dim changed As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Prep" And ws.Name <> "Lamb" Then

            Set fcells = Nothing
            On Error Resume Next
            Set fcells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0

            If Not fcells Is Nothing Then
                Set anchors = New Collection
                Set seen = CreateObject("Scripting.Dictionary")
                For Each cell In fcells
                    If cell.HasArray Then
                        Set a = cell.CurrentArray.Cells(1, 1)
                    Else
                        Set a = cell
                    End If
                    If Not seen.Exists(a.Address) Then
                        seen.Add a.Address, True
                        anchors.Add a
                    End If
                Next cell

                For Each an In anchors
                    ftext = an.Formula2
                    If Left$(ftext, 1) <> "=" Then ftext = an.FormulaArray
                    If Left$(ftext, 1) <> "=" Then ftext = an.Formula
                    ftext = Replace(ftext, "@", "")

                    If Left$(ftext, 1) = "=" Then
                        On Error Resume Next
                        If an.HasArray Then
                            an.CurrentArray.ClearContents
                        Else
                            an.ClearContents
                        End If
                        an.Formula2 = ftext
                        On Error GoTo 0
                        changed = changed + 1
                    End If
                Next an
            End If

        End If
    Next ws

    rewrite_test_formulas = changed
End Function

'==============================================================================
' Tab order
'==============================================================================

' Put the tabs back in order: Prep, Lamb, lambda_tests, vba_tests pinned as the
' first four, everything else alphabetical (case-insensitive). Run after adding
' a test sheet - Excel drops a copied sheet next to its source, and the VBA
' harness appends new zz_ fixtures at the end.
'
' zz_ fixtures need no special handling: they sort to the back on their own.
Public Sub sort_test_tabs()
    Dim pinned As Variant
    pinned = Array("Prep", "Lamb", "lambda_tests", "vba_tests")

    Dim others() As String
    Dim n As Long
    ReDim others(1 To ThisWorkbook.Worksheets.count)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsError(Application.Match(ws.Name, pinned, 0)) Then
            n = n + 1
            others(n) = ws.Name
        End If
    Next ws

    ' Insertion sort - n is small and this avoids depending on any sort helper.
    Dim i As Long, j As Long, tmp As String
    For i = 2 To n
        tmp = others(i)
        j = i - 1
        Do While j >= 1
            If LCase$(others(j)) <= LCase$(tmp) Then Exit Do
            others(j + 1) = others(j)
            j = j - 1
        Loop
        others(j + 1) = tmp
    Next i

    Application.ScreenUpdating = False

    ' Place each sheet after the previous one. Pinned sheets that don't exist
    ' are skipped rather than raising, so this still works on a partial workbook.
    Dim prev As String
    Dim k As Long
    For k = LBound(pinned) To UBound(pinned)
        If SheetThere(CStr(pinned(k))) Then
            MoveSheet CStr(pinned(k)), prev
            prev = CStr(pinned(k))
        End If
    Next k
    For i = 1 To n
        MoveSheet others(i), prev
        prev = others(i)
    Next i

    Application.ScreenUpdating = True
    MsgBox "Sorted " & (n + 4) & " tabs: 4 pinned, " & n & " alphabetical.", _
           vbInformation, "sort_test_tabs"
End Sub

Private Function SheetThere(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetThere = Not ws Is Nothing
End Function

' Move nm directly after after_nm, or to the front when after_nm is empty.
Private Sub MoveSheet(ByVal nm As String, ByVal after_nm As String)
    If Len(after_nm) = 0 Then
        ThisWorkbook.Worksheets(nm).Move Before:=ThisWorkbook.Sheets(1)
    Else
        ThisWorkbook.Worksheets(nm).Move After:=ThisWorkbook.Worksheets(after_nm)
    End If
End Sub
