Attribute VB_Name = "guess_and_check"
' deploy: template
'==============================================================================
' guess_and_check  -  build and drive a guess-and-check sheet for a level.
'
' Strategy: the live MEWC platform scores a whole level at once, telling you HOW
' MANY of your answers are right but not WHICH. For levels whose answers are
' small integers it can be faster to guess and check against that feedback than
' to solve the case. See docs/guess-and-check-business-rules.md for the full
' spec; this module implements it.
'
'   create_gc_sheet   PUBLIC. The only sub shown in Alt+F8. Run it FROM a level
'                     tab (_L1.._L7): it reads the active sheet's name to pick
'                     the level, parses that level's example / hints / points /
'                     game numbers, infers significance and whether negatives
'                     are allowed, and builds a new "_GCn" sheet (or _GCn(2),
'                     _GCn(3), ... if one already exists) wired up with three
'                     feedback buttons.
'
' The feedback subs (gc_zero_right / gc_one_right / gc_two_plus_right) are driven
' by the buttons create_gc_sheet places on the sheet. They are Public only so the
' buttons can call them, but each takes an Optional argument, which keeps them
' OUT of the Alt+F8 list (Excel hides any sub that takes arguments). Everything
' else is Private.
'==============================================================================
Option Explicit

' ---- columns on the generated guess-and-check sheet ----
Private Const COL_GAME As Long = 1          'A  game numbers
Private Const COL_CORRECT As Long = 2       'B  confirmed answers (blank = unsolved)
Private Const COL_GUESS As Long = 3         'C  guesses to submit this round
Private Const COL_ELIM_MIN As Long = 4      'D  low edge of the eliminated block
Private Const COL_ELIM_MAX As Long = 5      'E  high edge of the eliminated block
Private Const COL_POSS_A As Long = 6        'F  held-out bucket A (<= 1 right)
Private Const COL_POSS_B As Long = 7        'G  held-out bucket B (2+ right)
Private Const COL_INIT_GUESS As Long = 8    'H  seed guess per game

Private Const SIG_CELL As String = "B6"     ' Level Significance on the GC sheet
Private Const HEADER_TEXT As String = "Game Numbers"


'==============================================================================
' PUBLIC entry point
'==============================================================================
Public Sub create_gc_sheet()
    Dim src As Worksheet
    Set src = ActiveSheet

    Dim lvl As Long
    lvl = gc_level_from_name(src.Name)
    If lvl = 0 Then
        MsgBox "Run this from a level tab named _L1 .. _L7." & vbLf & _
               "The active sheet is '" & src.Name & "'.", vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- locate the level's data table header ("Game #" in column B) ----
    Dim hdrRow As Long
    hdrRow = gc_find_text_row(src, 2, "Game #", 1, 200)
    If hdrRow = 0 Then
        MsgBox "Could not find the 'Game #' table header on '" & src.Name & "'.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- example answer(s): rows just under the header whose Game # is ExampleN ----
    Dim exVals() As Double, exN As Long
    ReDim exVals(1 To 16): exN = 0
    Dim r As Long, gnum As String, ans As Variant
    For r = hdrRow + 1 To hdrRow + 20
        gnum = UCase$(Trim$(CStr(src.Cells(r, 2).Value)))       ' col B = Game #
        If Left$(gnum, 7) = "EXAMPLE" Then
            ans = src.Cells(r, 5).Value                          ' col E = Answer
            If Not IsEmpty(ans) And Trim$(CStr(ans)) <> "" Then
                ' PRE-FLIGHT GUARD: guess-and-check only works on numeric answers
                If Not gc_is_plain_number(ans) Then
                    MsgBox "Guess and Check only works on numeric answers.", _
                           vbExclamation, "Guess and Check"
                    Exit Sub
                End If
                exN = exN + 1
                exVals(exN) = CDbl(ans)
            End If
        End If
    Next r
    If exN = 0 Then
        MsgBox "Could not read a numeric example answer for level " & lvl & ".", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- game numbers (column A, below the table) ----
    Dim firstGame As Long, lastGame As Long
    If Not gc_game_range(src, firstGame, lastGame) Then
        MsgBox "Could not find the game numbers (column A) for level " & lvl & ".", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- hints: parse the "Game #g Hint: ... between X and Y" lines above the table.
    '      Assumption (long-stable): hints are exactly the level's first three
    '      games, or none. Each hint is placed against its own game's slot; a hint
    '      that is not in "between X and Y" form is ignored (that slot stays blank).
    Dim hMin(1 To 3) As Variant, hMax(1 To 3) As Variant
    Dim mn As Double, mx As Double, hg As Long, pos As Long
    For r = 1 To hdrRow - 1
        Dim ctext As String
        ctext = CStr(src.Cells(r, 3).Value)                     ' col C = hint text
        If InStr(1, ctext, "Hint:", vbTextCompare) > 0 Then
            hg = gc_parse_gamenum(ctext)
            If hg > 0 And gc_parse_between(ctext, mn, mx) Then
                pos = hg - firstGame + 1
                If pos >= 1 And pos <= 3 Then
                    hMin(pos) = mn: hMax(pos) = mx
                End If
            End If
        End If
    Next r

    ' ---- inferences over every numeric sample (examples + usable hint bounds) ----
    Dim samples() As Double, sN As Long
    ReDim samples(1 To exN + 6): sN = 0
    Dim i As Long
    For i = 1 To exN
        sN = sN + 1: samples(sN) = exVals(i)
    Next i
    For i = 1 To 3
        If Not IsEmpty(hMin(i)) Then sN = sN + 1: samples(sN) = CDbl(hMin(i))
        If Not IsEmpty(hMax(i)) Then sN = sN + 1: samples(sN) = CDbl(hMax(i))
    Next i

    Dim sig As Double: sig = gc_significance(samples, sN)
    Dim exMean As Double: exMean = gc_mean(exVals, exN)
    Dim negAllowed As Long: negAllowed = 0
    For i = 1 To sN
        If samples(i) < 0 Then negAllowed = 1: Exit For
    Next i

    ' ---- points per game (display only, for the button captions) ----
    Dim pts As Double: pts = gc_points_for_level(src.Parent, lvl)

    ' ==========================================================================
    ' Build the sheet
    ' ==========================================================================
    Application.ScreenUpdating = False

    Dim gc As Worksheet
    Set gc = src.Parent.Worksheets.Add(After:=src)
    gc.Name = gc_unique_name(src.Parent, "_GC" & lvl)

    ' ---- setup block ----
    gc.Range("A2").Value = "Level N":            gc.Range("B2").Value = lvl
    gc.Range("A3").Value = "Points Per Game":    gc.Range("B3").Value = pts
    gc.Range("A4").Value = "Game Ns Range":      gc.Range("B4").Value = firstGame: gc.Range("C4").Value = lastGame
    gc.Range("A5").Value = "Example Answer":     gc.Range("B5").Value = exMean
    gc.Range("A6").Value = "Level Significance": gc.Range("B6").Value = sig
    gc.Range("A7").Value = "Hint 1 Range"
    gc.Range("A8").Value = "Hint 2 Range"
    gc.Range("A9").Value = "Hint 3 Range"
    For i = 1 To 3
        If Not IsEmpty(hMin(i)) Then gc.Cells(6 + i, 2).Value = hMin(i)
        If Not IsEmpty(hMax(i)) Then gc.Cells(6 + i, 3).Value = hMax(i)
    Next i
    gc.Range("A10").Value = "Guess Center"
    gc.Range("B10").Formula = "=CEILING.MATH(IF(COUNT(B7:C9)=0,B5,0.5*B5+0.5*AVERAGE(B7:C9)),B6)"
    gc.Range("A11").Value = "Negative Allowed":  gc.Range("B11").Value = negAllowed

    ' ---- feedback caption cells ----
    gc.Range("E6").Value = "Feedback Buttons"
    gc.Range("E8").Formula = "=""0 / ""&B3*((C4-B4)+1)&"" Points"""
    gc.Range("F8").Formula = "=B3&"" / ""&B3*((C4-B4)+1)&"" Points"""
    gc.Range("G8").Formula = "=2*B3&""+ / ""&B3*((C4-B4)+1)&"" Points"""

    ' ---- diagnostics row (live counts) ----
    Dim col As Long
    For col = 1 To 8
        gc.Cells(12, col).Formula = "=COUNT(" & Cells(14, col).Address(False, False) & ":" & _
                                     Cells(1000, col).Address(False, False) & ")"
    Next col

    ' ---- table header ----
    gc.Range("A13").Value = HEADER_TEXT
    gc.Range("B13").Value = "Correct Answers"
    gc.Range("C13").Value = "Guesses"
    gc.Range("D13").Value = "Eliminated Min"
    gc.Range("E13").Value = "Eliminated Max"
    gc.Range("F13").Value = "Possible A"
    gc.Range("G13").Value = "Possible B"
    gc.Range("H13").Value = "Initial Guesses"

    ' ---- game numbers + per-game initial-guess formulas ----
    Dim nGames As Long: nGames = lastGame - firstGame + 1
    Dim rr As Long
    For i = 1 To nGames
        rr = 13 + i
        gc.Cells(rr, COL_GAME).Value = firstGame + i - 1
        Select Case i
            Case 1: gc.Cells(rr, COL_INIT_GUESS).Formula = "=IF(B7="""",$B$10,CEILING.MATH(AVERAGE(B7:C7),$B$6))"
            Case 2: gc.Cells(rr, COL_INIT_GUESS).Formula = "=IF(B8="""",$B$10,CEILING.MATH(AVERAGE(B8:C8),$B$6))"
            Case 3: gc.Cells(rr, COL_INIT_GUESS).Formula = "=IF(B9="""",$B$10,CEILING.MATH(AVERAGE(B9:C9),$B$6))"
            Case Else: gc.Cells(rr, COL_INIT_GUESS).Formula = "=IF(A" & rr & "="""","""",$B$10)"
        End Select
    Next i

    gc.Columns("A:H").AutoFit
    gc.Calculate

    ' ---- seed the first round of guesses from the initial guesses, and copy ----
    Dim firstRow As Long, lastRow As Long
    If gc_locate(gc, firstRow, lastRow) Then
        gc_generate_guesses gc, firstRow, lastRow, sig
        gc_place_buttons gc
        gc_copy_guesses gc, firstRow, lastRow
    End If

    Application.ScreenUpdating = True
    gc.Activate
End Sub


'==============================================================================
' Feedback subs  (button-driven; Optional arg keeps them out of Alt+F8)
'==============================================================================

' 0 correct this round.
Public Sub gc_zero_right(Optional ByVal ignore As Variant)
    Dim ws As Worksheet, firstRow As Long, lastRow As Long, sig As Double
    If Not gc_prep(ws, firstRow, lastRow, sig) Then Exit Sub
    Application.ScreenUpdating = False
    Randomize

    Dim r As Long
    Dim cntA0 As Long, cntB0 As Long, idxSingleA0 As Long
    Dim g As Variant, pa As String, pb As String, gn As Double
    Dim emn As Variant, emx As Variant, mv As Long

    ' snapshot the possibles before touching anything
    For r = firstRow To lastRow
        If Trim$(CStr(ws.Cells(r, COL_POSS_A).Value)) <> "" Then
            cntA0 = cntA0 + 1
            If idxSingleA0 = 0 Then idxSingleA0 = r
        End If
        If Trim$(CStr(ws.Cells(r, COL_POSS_B).Value)) <> "" Then cntB0 = cntB0 + 1
    Next r

    ' every guess was wrong: extend the eliminated block, drop matching possibles
    For r = firstRow To lastRow
        If Trim$(CStr(ws.Cells(r, COL_GUESS).Value)) <> "" Then
            gn = CDbl(ws.Cells(r, COL_GUESS).Value)
            emn = ws.Cells(r, COL_ELIM_MIN).Value
            emx = ws.Cells(r, COL_ELIM_MAX).Value
            If IsEmpty(emn) And IsEmpty(emx) Then
                ws.Cells(r, COL_ELIM_MIN).Value = gn
                ws.Cells(r, COL_ELIM_MAX).Value = gn
            Else
                If Not IsEmpty(emn) Then If gn = CDbl(emn) - sig Then ws.Cells(r, COL_ELIM_MIN).Value = gn
                If Not IsEmpty(emx) Then If gn = CDbl(emx) + sig Then ws.Cells(r, COL_ELIM_MAX).Value = gn
            End If
            g = CStr(ws.Cells(r, COL_GUESS).Value)
            If CStr(ws.Cells(r, COL_POSS_A).Value) = g Then ws.Cells(r, COL_POSS_A).ClearContents
            If CStr(ws.Cells(r, COL_POSS_B).Value) = g Then ws.Cells(r, COL_POSS_B).ClearContents
        End If
        ws.Cells(r, COL_GUESS).ClearContents
    Next r

    If cntA0 = 1 And cntB0 = 0 Then
        ' the lone survivor in A is the answer -> promote, then fresh guesses
        ws.Cells(idxSingleA0, COL_CORRECT).Value = ws.Cells(idxSingleA0, COL_POSS_A).Value
        ws.Cells(idxSingleA0, COL_POSS_A).ClearContents
        gc_generate_guesses ws, firstRow, lastRow, sig
    ElseIf cntA0 >= 2 Or cntB0 >= 1 Then
        ' mid-bisection: retest half (rounded up) of each held-out bucket, no new guesses
        gc_move_half ws, firstRow, lastRow, COL_POSS_A, COL_GUESS, (cntA0 + 1) \ 2
        gc_move_half ws, firstRow, lastRow, COL_POSS_B, COL_GUESS, (cntB0 + 1) \ 2
    Else
        ' ordinary elimination round -> fresh guesses
        gc_generate_guesses ws, firstRow, lastRow, sig
    End If

    gc_copy_guesses ws, firstRow, lastRow
    Application.ScreenUpdating = True
End Sub

' exactly 1 correct this round.
Public Sub gc_one_right(Optional ByVal ignore As Variant)
    Dim ws As Worksheet, firstRow As Long, lastRow As Long, sig As Double
    If Not gc_prep(ws, firstRow, lastRow, sig) Then Exit Sub
    Application.ScreenUpdating = False
    Randomize

    ' one_right ignores bucket B
    ws.Range(ws.Cells(firstRow, COL_POSS_B), ws.Cells(lastRow, COL_POSS_B)).ClearContents

    Dim r As Long, cntGuess As Long, firstGuessRow As Long
    For r = firstRow To lastRow
        If Trim$(CStr(ws.Cells(r, COL_GUESS).Value)) <> "" Then
            cntGuess = cntGuess + 1
            If firstGuessRow = 0 Then firstGuessRow = r
        End If
    Next r

    If cntGuess = 1 Then
        ' the single guess is the answer -> promote, fresh guesses
        ws.Cells(firstGuessRow, COL_CORRECT).Value = ws.Cells(firstGuessRow, COL_GUESS).Value
        ws.Range(ws.Cells(firstRow, COL_GUESS), ws.Cells(lastRow, COL_GUESS)).ClearContents
        ws.Range(ws.Cells(firstRow, COL_POSS_A), ws.Cells(lastRow, COL_POSS_A)).ClearContents
        gc_generate_guesses ws, firstRow, lastRow, sig
    Else
        ' fold bucket A into the eliminated ranges, then bisect the guesses
        Dim pa As Variant, emn As Variant, emx As Variant, pn As Double
        For r = firstRow To lastRow
            If Trim$(CStr(ws.Cells(r, COL_POSS_A).Value)) <> "" Then
                pn = CDbl(ws.Cells(r, COL_POSS_A).Value)
                emn = ws.Cells(r, COL_ELIM_MIN).Value
                emx = ws.Cells(r, COL_ELIM_MAX).Value
                If Not IsEmpty(emn) Then If pn = CDbl(emn) - sig Then ws.Cells(r, COL_ELIM_MIN).Value = pn
                If Not IsEmpty(emx) Then If pn = CDbl(emx) + sig Then ws.Cells(r, COL_ELIM_MAX).Value = pn
            End If
        Next r
        ws.Range(ws.Cells(firstRow, COL_POSS_A), ws.Cells(lastRow, COL_POSS_A)).ClearContents

        cntGuess = gc_count(ws, firstRow, lastRow, COL_GUESS)
        gc_move_half ws, firstRow, lastRow, COL_GUESS, COL_POSS_A, cntGuess \ 2
    End If

    gc_copy_guesses ws, firstRow, lastRow
    Application.ScreenUpdating = True
End Sub

' 2 or more correct this round.
Public Sub gc_two_plus_right(Optional ByVal ignore As Variant)
    Dim ws As Worksheet, firstRow As Long, lastRow As Long, sig As Double
    If Not gc_prep(ws, firstRow, lastRow, sig) Then Exit Sub
    Application.ScreenUpdating = False

    ' unused here
    ws.Range(ws.Cells(firstRow, COL_POSS_A), ws.Cells(lastRow, COL_POSS_A)).ClearContents

    ' hold half (rounded down) of the guesses in bucket B, retest the rest; no new guesses
    Dim cntGuess As Long: cntGuess = gc_count(ws, firstRow, lastRow, COL_GUESS)
    gc_move_half ws, firstRow, lastRow, COL_GUESS, COL_POSS_B, cntGuess \ 2

    gc_copy_guesses ws, firstRow, lastRow
    Application.ScreenUpdating = True
End Sub


'==============================================================================
' Shared engine (Private)
'==============================================================================

' Locate the working table on a GC sheet and read the significance. Returns False
' (with a message) if the sheet is not a guess-and-check sheet.
Private Function gc_prep(ByRef ws As Worksheet, ByRef firstRow As Long, _
                         ByRef lastRow As Long, ByRef sig As Double) As Boolean
    Set ws = ActiveSheet
    If Not gc_locate(ws, firstRow, lastRow) Then
        MsgBox "This doesn't look like a guess-and-check sheet.", vbExclamation, "Guess and Check"
        gc_prep = False: Exit Function
    End If
    sig = CDbl(ws.Range(SIG_CELL).Value)
    gc_prep = True
End Function

Private Function gc_locate(ByVal ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long) As Boolean
    firstRow = 0
    Dim r As Long
    For r = 1 To 100
        If LCase$(Trim$(CStr(ws.Cells(r, COL_GAME).Value))) = LCase$(HEADER_TEXT) Then
            firstRow = r + 1: Exit For
        End If
    Next r
    If firstRow = 0 Then gc_locate = False: Exit Function
    lastRow = ws.Cells(ws.Rows.Count, COL_GAME).End(xlUp).Row
    gc_locate = (lastRow >= firstRow)
End Function

' Build the next Guesses column: solved games stay blank; a game with no eliminated
' block yet uses its Initial Guess; otherwise probe one step past a random edge.
Private Sub gc_generate_guesses(ByVal ws As Worksheet, ByVal firstRow As Long, _
                                ByVal lastRow As Long, ByVal sig As Double)
    Dim r As Long, cv As Variant, emn As Variant, emx As Variant
    For r = firstRow To lastRow
        cv = ws.Cells(r, COL_CORRECT).Value
        If Not IsEmpty(cv) And IsNumeric(cv) Then
            ' solved (0 counts) -> no guess
        Else
            emn = ws.Cells(r, COL_ELIM_MIN).Value
            emx = ws.Cells(r, COL_ELIM_MAX).Value
            If IsEmpty(emn) Then
                ws.Cells(r, COL_GUESS).Value = ws.Cells(r, COL_INIT_GUESS).Value
            ElseIf Rnd < 0.5 Then
                ws.Cells(r, COL_GUESS).Value = CDbl(emn) - sig
            Else
                ws.Cells(r, COL_GUESS).Value = CDbl(emx) + sig
            End If
        End If
    Next r
End Sub

' Move the first `howMany` non-empty cells from column `fromCol` into `toCol`.
Private Sub gc_move_half(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                         ByVal fromCol As Long, ByVal toCol As Long, ByVal howMany As Long)
    If howMany <= 0 Then Exit Sub
    Dim r As Long
    For r = firstRow To lastRow
        If howMany = 0 Then Exit For
        If Trim$(CStr(ws.Cells(r, fromCol).Value)) <> "" Then
            ws.Cells(r, toCol).Value = ws.Cells(r, fromCol).Value
            ws.Cells(r, fromCol).ClearContents
            howMany = howMany - 1
        End If
    Next r
End Sub

Private Function gc_count(ByVal ws As Worksheet, ByVal firstRow As Long, _
                          ByVal lastRow As Long, ByVal col As Long) As Long
    Dim r As Long
    For r = firstRow To lastRow
        If Trim$(CStr(ws.Cells(r, col).Value)) <> "" Then gc_count = gc_count + 1
    Next r
End Function

' Copy the Guesses column to the clipboard so the operator can paste-submit,
' then click a feedback button, then paste again next round.
Private Sub gc_copy_guesses(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    ws.Range(ws.Cells(firstRow, COL_GUESS), ws.Cells(lastRow, COL_GUESS)).Copy
End Sub

Private Sub gc_place_buttons(ByVal ws As Worksheet)
    gc_add_button ws, ws.Range("E9"), "gc_zero_right", "0 right"
    gc_add_button ws, ws.Range("F9"), "gc_one_right", "1 right"
    gc_add_button ws, ws.Range("G9"), "gc_two_plus_right", "2+ right"
End Sub

Private Sub gc_add_button(ByVal ws As Worksheet, ByVal cel As Range, _
                          ByVal macro As String, ByVal caption As String)
    Dim b As Object
    Set b = ws.Buttons.Add(cel.Left, cel.Top, cel.Width, cel.Height * 2)
    b.OnAction = macro
    b.Caption = caption
End Sub


'==============================================================================
' Parsing / inference helpers (Private)
'==============================================================================

' "_L6" -> 6 ; anything else -> 0
Private Function gc_level_from_name(ByVal nm As String) As Long
    nm = Trim$(nm)
    If UCase$(Left$(nm, 2)) <> "_L" Then Exit Function
    Dim tail As String: tail = Mid$(nm, 3)
    If Len(tail) = 0 Then Exit Function
    Dim i As Long
    For i = 1 To Len(tail)
        If Mid$(tail, i, 1) < "0" Or Mid$(tail, i, 1) > "9" Then Exit Function
    Next i
    gc_level_from_name = CLng(tail)
End Function

' First row at/after startRow (in the given column) whose text = `text`, else 0.
Private Function gc_find_text_row(ByVal ws As Worksheet, ByVal col As Long, ByVal text As String, _
                                  ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If LCase$(Trim$(CStr(ws.Cells(r, col).Value))) = LCase$(text) Then
            gc_find_text_row = r: Exit Function
        End If
    Next r
End Function

' Min/max game number from the numeric cells in column A.
Private Function gc_game_range(ByVal ws As Worksheet, ByRef firstGame As Long, ByRef lastGame As Long) As Boolean
    Dim r As Long, lastR As Long, v As Variant, found As Boolean
    lastR = ws.Cells(ws.Rows.Count, COL_GAME).End(xlUp).Row
    For r = 1 To lastR
        v = ws.Cells(r, COL_GAME).Value
        If IsNumeric(v) And Not IsEmpty(v) Then
            If Not found Then
                firstGame = CLng(v): lastGame = CLng(v): found = True
            Else
                If CLng(v) < firstGame Then firstGame = CLng(v)
                If CLng(v) > lastGame Then lastGame = CLng(v)
            End If
        End If
    Next r
    gc_game_range = found
End Function

' Game number out of "Game #91 Hint: ..." -> 91 (0 if none).
Private Function gc_parse_gamenum(ByVal s As String) As Long
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "Game\s*#\s*(\d+)"
    End If
    If re.Test(s) Then gc_parse_gamenum = CLng(re.Execute(s)(0).SubMatches(0))
End Function

' Extract the two numbers from "... between X and Y ..." -> True on success.
Private Function gc_parse_between(ByVal s As String, ByRef mn As Double, ByRef mx As Double) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Pattern = "between\s+(-?\d+(?:\.\d+)?)\s+and\s+(-?\d+(?:\.\d+)?)"
    End If
    If re.Test(s) Then
        Dim m As Object: Set m = re.Execute(s)(0)
        mn = CDbl(m.SubMatches(0)): mx = CDbl(m.SubMatches(1))
        gc_parse_between = True
    End If
End Function

' A single plain number: rejects text, cell refs, and delimited series like "3;7".
Private Function gc_is_plain_number(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        ' IsNumeric accepts things like "1,000" / "$5"; require a clean number too
        Dim s As String: s = Trim$(CStr(v))
        If InStr(s, ",") > 0 Or InStr(s, ";") > 0 Or InStr(s, " ") > 0 Then Exit Function
        gc_is_plain_number = True
    End If
End Function

Private Function gc_mean(ByRef vals() As Double, ByVal n As Long) As Double
    Dim i As Long, t As Double
    For i = 1 To n: t = t + vals(i): Next i
    If n > 0 Then gc_mean = t / n
End Function

' Coarsest power of ten (1e-6 .. 1e6) that divides EVERY sample exactly.
Private Function gc_significance(ByRef vals() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double, best As Double
    For i = 1 To n
        s = gc_indiv_sig(vals(i))
        If i = 1 Or s < best Then best = s
    Next i
    If best = 0 Then best = 1
    gc_significance = best
End Function

Private Function gc_indiv_sig(ByVal x As Double) As Double
    Dim ax As Double: ax = Abs(x)
    If ax = 0 Then gc_indiv_sig = 1: Exit Function
    Dim k As Long, s As Double, q As Double
    For k = 6 To -6 Step -1
        s = 10 ^ k
        q = ax / s
        If Abs(q - CDbl(CLng(q))) < 0.0000001 Then gc_indiv_sig = s: Exit Function
    Next k
    gc_indiv_sig = 0.000001
End Function

' Best-effort points-per-game for a level, read from the imported case table.
' Display only (drives the button captions); returns 0 if not found.
Private Function gc_points_for_level(ByVal wb As Workbook, ByVal lvl As Long) As Double
    Dim cand As Variant, nm As Variant, ws As Worksheet, r As Long, lastR As Long
    cand = Array("Case", "case copy", "Answers", "case data")
    For Each nm In cand
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row     ' col B = Game #
            For r = 1 To lastR
                If IsNumeric(ws.Cells(r, 2).Value) And ws.Cells(r, 3).Value = lvl Then
                    If IsNumeric(ws.Cells(r, 4).Value) Then
                        If CDbl(ws.Cells(r, 4).Value) > 0 Then
                            gc_points_for_level = CDbl(ws.Cells(r, 4).Value): Exit Function
                        End If
                    End If
                End If
            Next r
        End If
    Next nm
End Function

Private Function gc_sheet_exists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    gc_sheet_exists = Not ws Is Nothing
End Function

' "_GC6" if free, else "_GC6(2)", "_GC6(3)", ...
Private Function gc_unique_name(ByVal wb As Workbook, ByVal base As String) As String
    If Not gc_sheet_exists(wb, base) Then gc_unique_name = base: Exit Function
    Dim k As Long
    For k = 2 To 999
        If Not gc_sheet_exists(wb, base & "(" & k & ")") Then
            gc_unique_name = base & "(" & k & ")": Exit Function
        End If
    Next k
    gc_unique_name = base & "(x)"
End Function
