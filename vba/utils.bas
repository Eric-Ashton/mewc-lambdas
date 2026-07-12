Attribute VB_Name = "utils"
'==============================================================================
' utils  -  general spreadsheet utilities for MEWC work.
'
'   convert_cse_to_dynamic   Finds every legacy CSE (Ctrl+Shift+Enter,
'                            {curly-brace}) array formula in the ACTIVE workbook
'                            and re-enters it as a modern dynamic-array formula
'                            (via .Formula2) so it spills normally. Leaves plain
'                            and already-dynamic formulas untouched. Handy when a
'                            case arrives with CSE formulas that misbehave.
'
'   clean_function_refs      Deletes broken hidden "_xleta.*" function-reference
'                            names. These are created when you use a bare built-in
'                            as a lambda (e.g. =BYROW(rng, min)); a round-trip
'                            through external tooling can corrupt them to #NAME?,
'                            which then makes that bare function fail everywhere
'                            (while others still work). After running, re-enter or
'                            full-recalc (Ctrl+Alt+F9) any affected formula and
'                            Excel rebuilds a clean reference.
'==============================================================================
Option Explicit

Public Sub convert_cse_to_dynamic()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fcells As Range, cell As Range, an As Range
    Dim anchors As Collection, seen As Object
    Dim ftext As String
    Dim converted As Long, failed As Long

    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In wb.Worksheets

        Set fcells = Nothing
        On Error Resume Next
        Set fcells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If Not fcells Is Nothing Then
            ' --- collect the anchor (top-left) of each CSE array, de-duplicated ---
            Set anchors = New Collection
            Set seen = CreateObject("Scripting.Dictionary")
            For Each cell In fcells
                ' CSE = part of an array formula but NOT a dynamic spill
                If cell.HasArray And Not cell.HasSpill Then
                    Set an = cell.CurrentArray.Cells(1, 1)
                    If Not seen.Exists(ws.Name & "!" & an.Address) Then
                        seen.Add ws.Name & "!" & an.Address, True
                        anchors.Add an
                    End If
                End If
            Next cell

            ' --- convert each: clear the CSE block, re-enter anchor as dynamic ---
            For Each an In anchors
                ftext = an.FormulaArray
                If Left$(ftext, 1) <> "=" Then ftext = an.Formula2
                If Left$(ftext, 1) = "=" Then
                    On Error Resume Next
                    Err.Clear
                    an.CurrentArray.ClearContents
                    an.Formula2 = ftext
                    If Err.Number = 0 Then converted = converted + 1 Else failed = failed + 1
                    On Error GoTo 0
                End If
            Next an
        End If

    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Converted " & converted & " CSE array formula(s) to dynamic in '" _
        & wb.Name & "'." & IIf(failed > 0, vbLf & failed & " could not be converted " _
        & "(likely blocked by neighbouring data - would #SPILL).", ""), _
        vbInformation, "convert_cse_to_dynamic"
End Sub


Public Sub clean_function_refs()
    Dim nm As Name
    Dim kill As Collection
    Dim s As Variant
    Dim removed As Long, refers As String

    Set kill = New Collection
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.Name, "_xleta.", vbTextCompare) > 0 Then
            kill.Add nm.Name        ' collect first; don't delete mid-iteration
        End If
    Next nm

    For Each s In kill
        On Error Resume Next
        ThisWorkbook.Names(s).Delete
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
        On Error GoTo 0
    Next s

    If removed > 0 Then
        Application.CalculateFull       ' let Excel rebuild clean references
    End If

    MsgBox "Removed " & removed & " broken function-reference name(s)." & vbLf & _
        "Re-enter (or Ctrl+Alt+F9) any formula using a bare function such as " & _
        "=BYROW(rng, min) to rebuild a clean reference.", _
        vbInformation, "clean_function_refs"
End Sub
