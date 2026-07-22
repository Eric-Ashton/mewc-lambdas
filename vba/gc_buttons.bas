Attribute VB_Name = "gc_buttons"
' deploy: template
Option Private Module   ' button entry points - hidden from Alt+F8, still callable by their buttons
'==============================================================================
' gc_buttons  -  feedback-button entry points for a guess-and-check (_GCn) sheet.
'
' create_gc_sheet (module guess_and_check) lays down one form-control button per
' "number correct" 0..7, captioned with the POINTS the platform shows for that
' many, plus an "8+ pts" fallback. Each button's OnAction points at the matching
' wrapper below; every wrapper just forwards the exact count to
' guess_and_check.gc_feedback.
'
' These wrappers must take no argument (so a button can call them), which would
' normally list them in Alt+F8. Option Private Module is what keeps them out of
' that list while leaving them callable by their buttons - the same trick the sync
' modules use. create_gc_sheet itself stays in module guess_and_check, where it is
' the one visible Alt+F8 entry point.
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

' Fallback for 8+ correct: the operator types the POINTS the platform showed into
' the entry cell (FB_CELL); we divide by points-per-game (B3) to recover the count.
Public Sub gc_fbN()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim v As Variant: v = ws.Range(FB_CELL).Value
    If Not IsNumeric(v) Or Trim$(CStr(v)) = "" Then
        MsgBox "For 8 or more correct, type the POINTS the platform showed into cell " & _
               FB_CELL & " first, then click this.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim pts As Double, k As Long
    If IsNumeric(ws.Range("B3").Value) Then pts = CDbl(ws.Range("B3").Value)
    If pts > 0 Then
        k = CLng(CDbl(v) / pts)         ' points -> count
    Else
        k = CLng(v)                     ' no points-per-game known: treat entry as a raw count
    End If
    If k < 0 Then
        MsgBox "That comes out to a negative count.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    gc_feedback k
End Sub
