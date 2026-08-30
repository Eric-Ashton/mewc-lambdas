Attribute VB_Name = "formula_shortcuts"
' deploy: shared
'==============================================================================
' formula_shortcuts - keyboard shortcuts for authoring. Two drop a starter
' formula skeleton into the ACTIVE cell (only when it is empty, so they can
' never overwrite existing work); one runs the split_cols utility.
'
'   Ctrl+Shift+J  ->  a SCAN(...) skeleton
'   Ctrl+Shift+K  ->  a scan2(...) skeleton
'   Ctrl+Shift+Q  ->  split_cols (split the selected spilled 2D formula into one CHOOSECOLS formula per column)
'
' Both skeletons keep their line breaks, and use placeholder names (a_init,
' x_arr, new_a, ...) that you replace - so a freshly-inserted cell shows #NAME?
' until you fill it in, which is expected.
'
' The bindings are set on workbook open (Auto_Open) and cleared on close
' (Auto_Close). Application.OnKey bindings are per Excel session and app-global,
' so if one ever stops responding, re-arm it by running
' register_formula_shortcuts from Alt+F8 (unregister_formula_shortcuts clears
' them).
'==============================================================================
Option Explicit

' --- the skeletons (Public + as functions so the harness can check them) ------

' Ctrl+Shift+J skeleton.
Public Function scan_template() As String
    scan_template = _
        "=SCAN(a_init, x_arr, LAMBDA(a, x, " & vbLf & _
        "new_a" & vbLf & _
        "))"
End Function

' Ctrl+Shift+K skeleton. The comma sits at the END of the x_arr line (rather
' than the start of the next line as first sketched) because Excel absorbs a
' line break placed immediately BEFORE a comma - trailing it keeps the break, so
' each LAMBDA still lands on its own line exactly as intended.
Public Function scan2_template() As String
    scan2_template = _
        "=scan2(a_init, b_init, x_arr," & vbLf & _
        "LAMBDA(a,b,x," & vbLf & _
        "new_a" & vbLf & _
        "),LAMBDA(a,b,x," & vbLf & _
        "new_b" & vbLf & _
        "))"
End Function

' --- the action: fill only when empty ----------------------------------------

' Put startFormula into target's top-left cell, but only if that cell is empty -
' existing content is never overwritten. Best-effort: a protected sheet or a
' Nothing target just does nothing rather than raising.
Public Sub fill_if_empty(ByVal target As Range, ByVal startFormula As String)
    On Error GoTo done
    If target Is Nothing Then Exit Sub
    Dim cell As Range
    Set cell = target.Cells(1, 1)
    If Not IsEmpty(cell.Value2) Then Exit Sub   ' occupied -> leave it alone
    cell.Formula2 = startFormula
done:
End Sub

' --- the shortcut entry points (bound by Application.OnKey) -------------------

Public Sub insert_scan_template()
    fill_if_empty ActiveCell, scan_template()
End Sub

Public Sub insert_scan2_template()
    fill_if_empty ActiveCell, scan2_template()
End Sub

' --- (un)register the key bindings --------------------------------------------

Public Sub register_formula_shortcuts()
    On Error Resume Next
    Application.OnKey "^+j", "'" & ThisWorkbook.Name & "'!insert_scan_template"
    Application.OnKey "^+k", "'" & ThisWorkbook.Name & "'!insert_scan2_template"
    Application.OnKey "^+q", "'" & ThisWorkbook.Name & "'!split_cols"
    On Error GoTo 0
End Sub

Public Sub unregister_formula_shortcuts()
    On Error Resume Next
    Application.OnKey "^+j"
    Application.OnKey "^+k"
    Application.OnKey "^+q"
    On Error GoTo 0
End Sub

' Runs automatically when the workbook is opened / closed through the Excel UI.
Public Sub Auto_Open()
    ' These are authoring shortcuts for the working template. Don't rebind the
    ' tester's keys just because the unit-test workbook was opened (the module
    ' is still deployed there so its logic can be unit-tested).
    If InStr(1, ThisWorkbook.Name, "Unit Tests", vbTextCompare) > 0 Then Exit Sub
    register_formula_shortcuts
End Sub

Public Sub Auto_Close()
    unregister_formula_shortcuts
End Sub
