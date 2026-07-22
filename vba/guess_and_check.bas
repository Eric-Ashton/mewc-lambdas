Attribute VB_Name = "guess_and_check"
' deploy: template
'==============================================================================
' guess_and_check  -  build and drive a guess-and-check sheet for a level.
'
' Strategy: the live MEWC platform scores a whole level at once, telling you HOW
' MANY of your answers are right but not WHICH. For levels whose answers are
' small numbers it can be faster to guess and check against that feedback than
' to solve the case. See docs/guess-and-check-business-rules.md for the full
' spec; this module implements it.
'
'   create_gc_sheet   PUBLIC. The only sub shown in Alt+F8. Run it FROM a level
'                     tab (_L1.._L7): it reads the active sheet's name to pick
'                     the level, parses that level's example / hints / points /
'                     game numbers, infers significance and whether negatives
'                     are allowed, and builds a new "_GCn" sheet (or _GCn(2),
'                     _GCn(3), ... if one already exists) wired up with feedback
'                     buttons.
'
' Two searches run at once, both driven only by the per-round COUNT of correct:
'
'   * Value search   - each unsolved game marches its guess outward from a
'                      centre in steps of the answer's significance, staying
'                      inside a feasible range [Range Low, Range High] and never
'                      going below 0 unless negatives are allowed. Un-hinted
'                      games have their range and centre re-derived from the
'                      answers already solved (adaptive).
'
'   * Attribution    - when a batch scores 0 < k < (games submitted), we know k
'                      of the submitted guesses are right but not which. The
'                      resolver bisects the submitted set: a sub-set that scores
'                      0 is eliminated wholesale, one that scores its own size is
'                      solved wholesale, otherwise it is split again. This uses
'                      the EXACT count, so zero halves are pruned immediately.
'
' The feedback buttons (0..7 points, plus an 8+ fallback) are driven by wrapper
' subs (gc_fb0..gc_fb7, gc_fbN) that live in the companion module gc_buttons; that
' module is Option Private Module, which keeps them OUT of the Alt+F8 list while
' still letting their buttons call them. They forward to gc_feedback here.
' create_gc_sheet is the one Alt+F8-visible entry point; everything else is Private
' or (like gc_feedback, which has a required argument) otherwise hidden.
'==============================================================================
Option Explicit

' ---- columns on the generated guess-and-check sheet ----
Private Const COL_GAME As Long = 1          'A  game numbers
Private Const COL_CORRECT As Long = 2       'B  confirmed answers (blank = unsolved)
Private Const COL_GUESS As Long = 3         'C  guesses to submit this round
Private Const COL_ELIM_MIN As Long = 4      'D  low edge of the eliminated block
Private Const COL_ELIM_MAX As Long = 5      'E  high edge of the eliminated block
Private Const COL_RANGE_LO As Long = 6      'F  feasible range - low bound  (blank = open)
Private Const COL_RANGE_HI As Long = 7      'G  feasible range - high bound (blank = open)
Private Const COL_INIT As Long = 8          'H  seed guess per game
Private Const COL_ATTR As Long = 9          'I  candidate value held during attribution

Private Const SIG_CELL As String = "B6"     ' Level Significance
Private Const CENTER_CELL As String = "B10" ' Guess Centre
Private Const NEG_CELL As String = "B11"    ' Negative Allowed (0/1)
Public Const FB_CELL As String = "N8"       ' operator types the POINTS here for the 8+ fallback button (used by gc_buttons)
Private Const HEADER_TEXT As String = "Game Numbers"

Private Const OPEN_LO As Double = -1E+300   ' sentinel: unbounded below
Private Const OPEN_HI As Double = 1E+300    ' sentinel: unbounded above


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

    ' ---- points per game (display only) ----
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

    ' ---- feedback header + 8+ points-entry cells ----
    '      One button per "number correct" 0..7, captioned with the POINTS the
    '      platform shows for that many (count x points-per-game). For 8+ the
    '      operator types the points into N8 and clicks the fallback button.
    gc.Range("E6").Value = "Number of points:"
    gc.Range("N6").Value = "8 or more?"
    gc.Range("N7").Value = "type points:"
    ' (the 0..7 point buttons + the N8 entry cell are placed / styled below)

    ' ---- diagnostics row (live counts) ----
    Dim col As Long
    For col = 1 To 9
        gc.Cells(12, col).Formula = "=COUNT(" & Cells(14, col).Address(False, False) & ":" & _
                                     Cells(1000, col).Address(False, False) & ")"
    Next col

    ' ---- table header ----
    gc.Range("A13").Value = HEADER_TEXT
    gc.Range("B13").Value = "Correct Answers"
    gc.Range("C13").Value = "Guesses"
    gc.Range("D13").Value = "Eliminated Min"
    gc.Range("E13").Value = "Eliminated Max"
    gc.Range("F13").Value = "Range Low"
    gc.Range("G13").Value = "Range High"
    gc.Range("H13").Value = "Initial Guesses"
    gc.Range("I13").Value = "Attribution"

    ' ---- game numbers + per-game initial-guess formulas + hint ranges ----
    Dim nGames As Long: nGames = lastGame - firstGame + 1
    Dim rr As Long
    For i = 1 To nGames
        rr = 13 + i
        gc.Cells(rr, COL_GAME).Value = firstGame + i - 1
        Select Case i
            Case 1: gc.Cells(rr, COL_INIT).Formula = "=IF(B7="""",$B$10,CEILING.MATH(AVERAGE(B7:C7),$B$6))"
            Case 2: gc.Cells(rr, COL_INIT).Formula = "=IF(B8="""",$B$10,CEILING.MATH(AVERAGE(B8:C8),$B$6))"
            Case 3: gc.Cells(rr, COL_INIT).Formula = "=IF(B9="""",$B$10,CEILING.MATH(AVERAGE(B9:C9),$B$6))"
            Case Else: gc.Cells(rr, COL_INIT).Formula = "=IF(A" & rr & "="""","""",$B$10)"
        End Select
        ' hinted games get a fixed feasible range from their hint
        If i <= 3 Then
            If Not IsEmpty(hMin(i)) Then gc.Cells(rr, COL_RANGE_LO).Value = CDbl(hMin(i))
            If Not IsEmpty(hMax(i)) Then gc.Cells(rr, COL_RANGE_HI).Value = CDbl(hMax(i))
        End If
    Next i

    gc.Columns("A:I").AutoFit
    gc.Calculate

    ' ---- seed the first round of guesses, wire up buttons, format, copy ----
    Dim firstRow As Long, lastRow As Long
    If gc_locate(gc, firstRow, lastRow) Then
        gc_scan_regenerate gc, firstRow, lastRow, sig, (negAllowed <> 0)
        gc_format_sheet gc, firstRow, lastRow
        gc_place_buttons gc                 ' after formatting, so column widths are final
        gc_copy_guesses gc, firstRow, lastRow
    End If

    Application.ScreenUpdating = True
    gc.Activate
    On Error Resume Next
    ActiveWindow.Zoom = 130          ' default zoom for the GC sheet
    On Error GoTo 0
End Sub


' The button wrappers (gc_fb0..gc_fb7, gc_fbN) live in module gc_buttons, which is
' Option Private Module so they stay OUT of the Alt+F8 list while remaining callable
' by their buttons. They forward the exact count to gc_feedback below.


'==============================================================================
' Unified feedback handler
'   k = the EXACT number of currently-submitted guesses that were correct.
'   Public so gc_buttons can call it; the required argument keeps it out of Alt+F8.
'==============================================================================
Public Sub gc_feedback(ByVal k As Long)
    Dim ws As Worksheet, firstRow As Long, lastRow As Long, sig As Double, neg As Boolean
    If Not gc_prep(ws, firstRow, lastRow, sig, neg) Then Exit Sub
    Application.ScreenUpdating = False

    If gc_count(ws, firstRow, lastRow, COL_ATTR) > 0 Then
        gc_attrib_step ws, firstRow, lastRow, sig, neg, k
    Else
        gc_scan_step ws, firstRow, lastRow, sig, neg, k
    End If

    gc_copy_guesses ws, firstRow, lastRow
    Application.ScreenUpdating = True
End Sub

' A plain scanning round: the whole Guesses column was submitted.
Private Sub gc_scan_step(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                         ByVal sig As Double, ByVal neg As Boolean, ByVal k As Long)
    Dim m As Long: m = gc_count(ws, firstRow, lastRow, COL_GUESS)
    If m = 0 Then
        MsgBox "There are no guesses to score - the Guesses column is empty.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    If k > m Then
        MsgBox "You entered " & k & " correct, but only " & m & " games were submitted this round.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    Dim r As Long
    If k = 0 Then
        ' every guess wrong -> eliminate them all, then scan again
        For r = firstRow To lastRow
            If gc_has(ws, r, COL_GUESS) Then
                gc_eliminate ws, r, CDbl(ws.Cells(r, COL_GUESS).Value), sig
                ws.Cells(r, COL_GUESS).ClearContents
            End If
        Next r
        gc_scan_regenerate ws, firstRow, lastRow, sig, neg

    ElseIf k = m Then
        ' every guess right -> solve them all, then scan again (probably done)
        For r = firstRow To lastRow
            If gc_has(ws, r, COL_GUESS) Then
                ws.Cells(r, COL_CORRECT).Value = ws.Cells(r, COL_GUESS).Value
                ws.Cells(r, COL_GUESS).ClearContents
            End If
        Next r
        gc_scan_regenerate ws, firstRow, lastRow, sig, neg

    Else
        ' 0 < k < m -> ambiguous: begin attribution on the submitted set
        For r = firstRow To lastRow
            If gc_has(ws, r, COL_GUESS) Then ws.Cells(r, COL_ATTR).Value = ws.Cells(r, COL_GUESS).Value
        Next r
        gc_hold_half ws, firstRow, lastRow          ' keep first half active, hold the rest
    End If
End Sub

' An attribution round: only the current "window" (the submitted subset) was sent.
Private Sub gc_attrib_step(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                           ByVal sig As Double, ByVal neg As Boolean, ByVal k As Long)
    Dim w As Long: w = gc_count(ws, firstRow, lastRow, COL_GUESS)
    If w = 0 Then
        ' nothing in the window (shouldn't happen): re-gather all candidates
        gc_regather ws, firstRow, lastRow
        Exit Sub
    End If
    If k > w Then
        MsgBox "You entered " & k & " correct, but only " & w & " guesses were submitted this round.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    Dim r As Long, resolved As Boolean
    If k = 0 Then
        ' the whole window is wrong -> eliminate each, release them back to scanning
        For r = firstRow To lastRow
            If gc_has(ws, r, COL_GUESS) Then
                gc_eliminate ws, r, CDbl(ws.Cells(r, COL_GUESS).Value), sig
                ws.Cells(r, COL_GUESS).ClearContents
                ws.Cells(r, COL_ATTR).ClearContents
            End If
        Next r
        resolved = True

    ElseIf k = w Then
        ' the whole window is right -> solve each
        For r = firstRow To lastRow
            If gc_has(ws, r, COL_GUESS) Then
                ws.Cells(r, COL_CORRECT).Value = ws.Cells(r, COL_GUESS).Value
                ws.Cells(r, COL_GUESS).ClearContents
                ws.Cells(r, COL_ATTR).ClearContents
            End If
        Next r
        resolved = True

    Else
        ' 0 < k < w -> split: keep the first half active, hold the rest as candidates
        gc_hold_half ws, firstRow, lastRow
        resolved = False
    End If

    If resolved Then
        If gc_count(ws, firstRow, lastRow, COL_ATTR) > 0 Then
            gc_regather ws, firstRow, lastRow       ' re-test everything still unresolved
        Else
            gc_scan_regenerate ws, firstRow, lastRow, sig, neg   ' attribution done -> back to scanning
        End If
    End If
End Sub


'==============================================================================
' Value scan (Private)
'==============================================================================

' Rebuild the Guesses column for a fresh scanning round: every unsolved game gets
' its next candidate value; solved games (and any whose range is exhausted) stay
' blank. Also refreshes the adaptive range/centre for un-hinted games.
Private Sub gc_scan_regenerate(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                               ByVal sig As Double, ByVal neg As Boolean)
    Dim nSolved As Long, meanS As Double, mnS As Double, mxS As Double
    gc_solved_stats ws, firstRow, lastRow, nSolved, meanS, mnS, mxS
    gc_update_adaptive ws, firstRow, lastRow, sig, neg, nSolved, mnS, mxS

    Dim r As Long
    For r = firstRow To lastRow
        ws.Cells(r, COL_GUESS).ClearContents
        ws.Cells(r, COL_ATTR).ClearContents
    Next r

    Dim nUnsolved As Long, nGuess As Long, center As Double, baseC As Double
    Dim hv As Variant, g As Variant
    For r = firstRow To lastRow
        If Not gc_is_solved(ws, r) Then
            nUnsolved = nUnsolved + 1
            hv = ws.Cells(r, COL_INIT).Value
            If IsNumeric(hv) And Trim$(CStr(hv)) <> "" Then
                baseC = CDbl(hv)
            Else
                baseC = CDbl(ws.Range(CENTER_CELL).Value)
            End If
            If (Not gc_is_hinted(ws, firstRow, r)) And nSolved >= 2 Then
                center = meanS                          ' adaptive re-centre for un-hinted games
            Else
                center = baseC
            End If
            g = gc_next_guess(ws, r, sig, neg, center)
            If Not IsEmpty(g) Then
                ws.Cells(r, COL_GUESS).Value = CDbl(g)
                nGuess = nGuess + 1
            End If
        End If
    Next r

    If nUnsolved > 0 And nGuess = 0 Then
        MsgBox "Every remaining game has exhausted its search range." & vbLf & vbLf & _
               "Widen 'Range Low' / 'Range High' (columns F/G) for those games, or check " & _
               "'Negative Allowed' (" & NEG_CELL & "), then click a feedback button again.", _
               vbExclamation, "Guess and Check"
    End If
End Sub

' The next value to try for game `r`: its initial guess when nothing is eliminated
' yet, otherwise one step past whichever edge of the eliminated block is closer to
' the centre - staying inside [lo, hi]. Returns Empty when the range is exhausted.
Private Function gc_next_guess(ByVal ws As Worksheet, ByVal r As Long, ByVal sig As Double, _
                               ByVal neg As Boolean, ByVal center As Double) As Variant
    Dim lo As Double, hi As Double
    lo = gc_bound(ws, r, COL_RANGE_LO, IIf(neg, OPEN_LO, 0#))
    hi = gc_bound(ws, r, COL_RANGE_HI, OPEN_HI)

    Dim c As Double: c = center
    If c < lo Then c = lo
    If c > hi Then c = hi

    Dim emn As Variant: emn = ws.Cells(r, COL_ELIM_MIN).Value
    If Not (IsNumeric(emn) And Trim$(CStr(emn)) <> "") Then
        ' nothing eliminated yet -> the (clamped) centre is the first guess
        If c < lo - sig * 0.000001 Or c > hi + sig * 0.000001 Then
            gc_next_guess = Empty
        Else
            gc_next_guess = c
        End If
        Exit Function
    End If

    Dim emx As Double: emx = CDbl(ws.Cells(r, COL_ELIM_MAX).Value)
    Dim nl As Double, nh As Double, vl As Boolean, vh As Boolean, eps As Double
    eps = sig * 0.000001
    nl = CDbl(emn) - sig
    nh = emx + sig
    vl = (nl >= lo - eps)
    vh = (nh <= hi + eps)

    If vl And vh Then
        If (c - nl) <= (nh - c) Then gc_next_guess = nl Else gc_next_guess = nh
    ElseIf vl Then
        gc_next_guess = nl
    ElseIf vh Then
        gc_next_guess = nh
    Else
        gc_next_guess = Empty       ' both edges out of range -> exhausted
    End If
End Function

' Fold a proven-wrong value into game r's eliminated block (contiguous extension).
Private Sub gc_eliminate(ByVal ws As Worksheet, ByVal r As Long, ByVal val As Double, ByVal sig As Double)
    Dim emn As Variant, emx As Variant, eps As Double
    eps = sig * 0.000001
    emn = ws.Cells(r, COL_ELIM_MIN).Value
    emx = ws.Cells(r, COL_ELIM_MAX).Value
    If Not (IsNumeric(emn) And Trim$(CStr(emn)) <> "") Then
        ws.Cells(r, COL_ELIM_MIN).Value = val
        ws.Cells(r, COL_ELIM_MAX).Value = val
        Exit Sub
    End If
    If Abs(val - (CDbl(emn) - sig)) < eps Then
        ws.Cells(r, COL_ELIM_MIN).Value = val
    ElseIf Abs(val - (CDbl(emx) + sig)) < eps Then
        ws.Cells(r, COL_ELIM_MAX).Value = val
    Else
        ' non-adjacent (defensive): widen the block to include it
        If val < CDbl(emn) Then ws.Cells(r, COL_ELIM_MIN).Value = val
        If val > CDbl(emx) Then ws.Cells(r, COL_ELIM_MAX).Value = val
    End If
End Sub

' For un-hinted, unsolved games, re-derive a generous feasible range from the
' answers already solved. Bounds runaway scanning without cutting off good values.
Private Sub gc_update_adaptive(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                               ByVal sig As Double, ByVal neg As Boolean, _
                               ByVal nSolved As Long, ByVal mnS As Double, ByVal mxS As Double)
    If nSolved < 2 Then Exit Sub
    Dim spread As Double: spread = mxS - mnS
    Dim hw As Double: hw = 3 * spread
    If hw < 20 * sig Then hw = 20 * sig
    Dim lo As Double, hi As Double
    lo = mnS - hw
    hi = mxS + hw
    If (Not neg) And lo < 0 Then lo = 0

    Dim r As Long
    For r = firstRow To lastRow
        If (Not gc_is_solved(ws, r)) And (Not gc_is_hinted(ws, firstRow, r)) Then
            ws.Cells(r, COL_RANGE_LO).Value = lo
            ws.Cells(r, COL_RANGE_HI).Value = hi
        End If
    Next r
End Sub


'==============================================================================
' Attribution helpers (Private)
'==============================================================================

' Keep the first half (rounded up) of the current window in Guesses; blank the
' Guesses of the rest so they are held out (their Attribution value is preserved).
Private Sub gc_hold_half(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim rows() As Long, n As Long
    gc_collect ws, firstRow, lastRow, COL_GUESS, rows, n
    If n <= 1 Then Exit Sub
    Dim keep As Long: keep = (n + 1) \ 2
    Dim i As Long
    For i = keep + 1 To n
        ws.Cells(rows(i), COL_GUESS).ClearContents
    Next i
End Sub

' Re-submit every still-unresolved candidate (Attribution set) as one new window,
' so the next count re-derives how many of them remain correct.
Private Sub gc_regather(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    For r = firstRow To lastRow
        If gc_has(ws, r, COL_ATTR) Then ws.Cells(r, COL_GUESS).Value = ws.Cells(r, COL_ATTR).Value
    Next r
End Sub


'==============================================================================
' Small shared utilities (Private)
'==============================================================================

' Locate the working table + read sig / negative-allowed. False (with a message)
' if the active sheet isn't a guess-and-check sheet.
Private Function gc_prep(ByRef ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long, _
                         ByRef sig As Double, ByRef neg As Boolean) As Boolean
    Set ws = ActiveSheet
    If Not gc_locate(ws, firstRow, lastRow) Then
        MsgBox "This doesn't look like a guess-and-check sheet.", vbExclamation, "Guess and Check"
        gc_prep = False: Exit Function
    End If
    sig = CDbl(ws.Range(SIG_CELL).Value)
    Dim nv As Variant: nv = ws.Range(NEG_CELL).Value
    neg = (IsNumeric(nv) And Val(CStr(nv)) <> 0)
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

' A cell in `col` on row r holds something non-blank.
Private Function gc_has(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long) As Boolean
    gc_has = (Trim$(CStr(ws.Cells(r, col).Value)) <> "")
End Function

' A game is solved iff its Correct Answers cell is numeric-and-present (0 counts).
Private Function gc_is_solved(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim v As Variant: v = ws.Cells(r, COL_CORRECT).Value
    gc_is_solved = (IsNumeric(v) And Trim$(CStr(v)) <> "")
End Function

' A game has its own parsed hint (fixed feasible range) iff it is one of the first
' three games and its Hint N Range min cell (B7/B8/B9) is numeric.
Private Function gc_is_hinted(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal r As Long) As Boolean
    Dim rel As Long: rel = r - firstRow + 1
    If rel < 1 Or rel > 3 Then Exit Function
    Dim v As Variant: v = ws.Cells(6 + rel, 2).Value       ' B7 / B8 / B9
    gc_is_hinted = (IsNumeric(v) And Trim$(CStr(v)) <> "")
End Function

' Feasible bound from col (COL_RANGE_LO/HI), or `dflt` when the cell is blank.
Private Function gc_bound(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal dflt As Double) As Double
    Dim v As Variant: v = ws.Cells(r, col).Value
    If IsNumeric(v) And Trim$(CStr(v)) <> "" Then gc_bound = CDbl(v) Else gc_bound = dflt
End Function

' Rows (in order) whose cell in `col` is non-blank.
Private Sub gc_collect(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                       ByVal col As Long, ByRef rows() As Long, ByRef n As Long)
    ReDim rows(1 To (lastRow - firstRow + 1))
    n = 0
    Dim r As Long
    For r = firstRow To lastRow
        If gc_has(ws, r, col) Then n = n + 1: rows(n) = r
    Next r
End Sub

Private Function gc_count(ByVal ws As Worksheet, ByVal firstRow As Long, _
                          ByVal lastRow As Long, ByVal col As Long) As Long
    Dim r As Long
    For r = firstRow To lastRow
        If gc_has(ws, r, col) Then gc_count = gc_count + 1
    Next r
End Function

Private Sub gc_solved_stats(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, _
                            ByRef n As Long, ByRef meanV As Double, ByRef mnV As Double, ByRef mxV As Double)
    Dim r As Long, s As Double, v As Double, seen As Boolean
    n = 0: s = 0
    For r = firstRow To lastRow
        If gc_is_solved(ws, r) Then
            v = CDbl(ws.Cells(r, COL_CORRECT).Value)
            n = n + 1: s = s + v
            If Not seen Then
                mnV = v: mxV = v: seen = True
            Else
                If v < mnV Then mnV = v
                If v > mxV Then mxV = v
            End If
        End If
    Next r
    If n > 0 Then meanV = s / n Else meanV = 0
End Sub

' Copy the Guesses column to the clipboard so the operator can paste-submit,
' click a feedback button, and paste again next round.
Private Sub gc_copy_guesses(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    ws.Range(ws.Cells(firstRow, COL_GUESS), ws.Cells(lastRow, COL_GUESS)).Copy
End Sub


'==============================================================================
' Buttons + formatting (Private)
'==============================================================================
' A 4x2 grid of buttons for 0..7 correct, each captioned with the POINTS the
' platform shows for that many (count x points-per-game), plus an 8+ fallback.
Private Sub gc_place_buttons(ByVal ws As Worksheet)
    Dim pts As Double
    If IsNumeric(ws.Range("B3").Value) Then pts = CDbl(ws.Range("B3").Value)

    gc_add_button ws, ws.Range("E8:F9"),   "gc_fb0", gc_pts_caption(0, pts)
    gc_add_button ws, ws.Range("G8:H9"),   "gc_fb1", gc_pts_caption(1, pts)
    gc_add_button ws, ws.Range("I8:J9"),   "gc_fb2", gc_pts_caption(2, pts)
    gc_add_button ws, ws.Range("K8:L9"),   "gc_fb3", gc_pts_caption(3, pts)
    gc_add_button ws, ws.Range("E10:F11"), "gc_fb4", gc_pts_caption(4, pts)
    gc_add_button ws, ws.Range("G10:H11"), "gc_fb5", gc_pts_caption(5, pts)
    gc_add_button ws, ws.Range("I10:J11"), "gc_fb6", gc_pts_caption(6, pts)
    gc_add_button ws, ws.Range("K10:L11"), "gc_fb7", gc_pts_caption(7, pts)

    gc_add_button ws, ws.Range("O8:P9"),   "gc_fbN", "8+ pts"
End Sub

' Caption for the "k correct" button: the points the platform shows for k right,
' i.e. k x points-per-game. Falls back to the bare count when points are unknown.
Private Function gc_pts_caption(ByVal k As Long, ByVal pts As Double) As String
    If pts > 0 Then gc_pts_caption = CStr(k * pts) Else gc_pts_caption = CStr(k)
End Function

Private Sub gc_add_button(ByVal ws As Worksheet, ByVal rng As Range, _
                          ByVal macro As String, ByVal caption As String)
    Dim b As Object
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.OnAction = macro
    b.Caption = caption
End Sub

' Cosmetic pass - fills, borders, widths. Purely visual; no state lives in styling.
Private Sub gc_format_sheet(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim hdr As Long: hdr = firstRow - 1
    Dim cNavy As Long, cLabel As Long, cGreen As Long, cGuessH As Long, cGuess As Long
    Dim cHead As Long, cInput As Long, cGrid As Long
    cNavy = RGB(31, 78, 120)
    cLabel = RGB(221, 235, 247)
    cGreen = RGB(226, 239, 218)
    cGuessH = RGB(191, 143, 0)
    cGuess = RGB(255, 242, 204)
    cHead = RGB(47, 117, 181)
    cInput = RGB(255, 255, 204)
    cGrid = RGB(200, 200, 200)

    With ws
        ' title
        .Range("A1").Value = "Guess & Check  -  Level " & .Range("B2").Value
        With .Range("A1")
            .Font.Bold = True: .Font.Size = 14: .Font.Color = cNavy
        End With

        ' setup block
        .Range("A2:A11").Font.Bold = True
        .Range("A2:A11").Interior.Color = cLabel
        With .Range("B2:C11")
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = cGrid
        End With

        ' feedback area
        .Range("E6").Font.Bold = True
        .Range("E6").Font.Size = 12
        .Range("N6").Font.Bold = True
        .Range("N7").Font.Size = 9
        With .Range(FB_CELL)
            .Interior.Color = cInput
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
            .HorizontalAlignment = xlCenter
        End With

        ' table header
        With .Range(.Cells(hdr, COL_GAME), .Cells(hdr, COL_ATTR))
            .Interior.Color = cHead
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
        .Cells(hdr, COL_GUESS).Interior.Color = cGuessH   ' spotlight the column you copy

        ' data grid
        If lastRow >= firstRow Then
            With .Range(.Cells(firstRow, COL_GAME), .Cells(lastRow, COL_ATTR))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = cGrid
                .HorizontalAlignment = xlCenter
            End With
            .Range(.Cells(firstRow, COL_CORRECT), .Cells(lastRow, COL_CORRECT)).Interior.Color = cGreen
            .Range(.Cells(firstRow, COL_GUESS), .Cells(lastRow, COL_GUESS)).Interior.Color = cGuess
        End If

        ' widths (A:I is the table; J:L back the 0..7 button grid so it reads evenly)
        .Columns("A:L").ColumnWidth = 13
        .Columns("A").ColumnWidth = 14
        .Rows(hdr).RowHeight = 28
    End With
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
        Dim mch As Object: Set mch = re.Execute(s)(0)
        mn = CDbl(mch.SubMatches(0)): mx = CDbl(mch.SubMatches(1))
        gc_parse_between = True
    End If
End Function

' A single plain number: rejects text, cell refs, and delimited series like "3;7".
Private Function gc_is_plain_number(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
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
' Display only; returns 0 if not found.
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
