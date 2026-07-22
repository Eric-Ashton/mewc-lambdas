Attribute VB_Name = "gc_buttons"
' deploy: template
Option Private Module   ' button entry points - hidden from Alt+F8, still callable by their buttons
'==============================================================================
' gc_buttons  -  feedback-button entry points for a guess-and-check (_GCn) sheet.
'
' create_gc_sheet (module guess_and_check) lays down a 4x2 grid of buttons for
' "0..7 MORE correct", an "8+ pts" fallback, and an "Undo" button. The 0..7
' buttons are captioned with the ABSOLUTE score the platform will show once that
' many more are right - (confirmed + j) x P - and re-captioned after every round.
'
' Each wrapper forwards the ACTIVE hit count (correct among the currently-tested
' guesses, i.e. excluding already-confirmed answers) to guess_and_check.gc_feedback.
' Option Private Module keeps these out of Alt+F8 while their buttons still call
' them; create_gc_sheet stays visible in module guess_and_check.
'==============================================================================
Option Explicit

Public Sub gc_fb0()
    gc_feedback 0
End Sub
Public Sub gc_fb1()
    gc_feedback 1
End Sub
Public Sub gc_fb2()
    gc_feedback 2
End Sub
Public Sub gc_fb3()
    gc_feedback 3
End Sub
Public Sub gc_fb4()
    gc_feedback 4
End Sub
Public Sub gc_fb5()
    gc_feedback 5
End Sub
Public Sub gc_fb6()
    gc_feedback 6
End Sub
Public Sub gc_fb7()
    gc_feedback 7
End Sub

' Fallback for 8+ more correct: the operator types the ABSOLUTE points the
' platform showed into FB_CELL; active hits = points/P - (already-confirmed).
Public Sub gc_fbN()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim v As Variant: v = ws.Range(FB_CELL).Value
    If Not IsNumeric(v) Or Trim$(CStr(v)) = "" Then
        MsgBox "Type the POINTS the platform showed into cell " & FB_CELL & _
               " first, then click this.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim p As Double
    If IsNumeric(ws.Range("B3").Value) Then p = CDbl(ws.Range("B3").Value)
    If p <= 0 Then
        MsgBox "Points Per Game (B3) is not set, so points can't be converted to a count.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim totalCorrect As Double: totalCorrect = CDbl(v) / p
    If Abs(totalCorrect - CLng(totalCorrect)) > 0.000001 Then
        MsgBox CDbl(v) & " points is not a whole multiple of " & p & _
               " points per game. Check the number.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim confirmed As Long: confirmed = Application.WorksheetFunction.Count(ws.Range("B14:B100000"))
    gc_feedback CLng(totalCorrect) - confirmed
End Sub

Public Sub gc_undo()
    gc_apply_undo ActiveSheet
End Sub
