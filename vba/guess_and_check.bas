Attribute VB_Name = "guess_and_check"
' deploy: template
'==============================================================================
' guess_and_check  -  build and drive a guess-and-check sheet for a level.
'
' The live MEWC platform scores a whole level at once (points = correct x P),
' telling you HOW MANY answers are right, not WHICH. For levels whose answers are
' small numbers it can be faster to guess and check against that feedback than to
' solve the case. See docs/guess-and-check-business-rules.md for the full spec.
'
' create_gc_sheet is the one Alt+F8-visible entry point. Run it FROM a level tab
' (_L1.._L7); it builds a "_GCn" sheet wired to feedback buttons (in module
' gc_buttons) and drives two interleaved searches, both from the per-round count:
'
'   Value search  - each unsolved game tries its next candidate. Order: values
'                   already CONFIRMED elsewhere on the level (most frequent first),
'                   then the example answer, then an outward scan from a centre
'                   that re-derives from solved answers (mode/median). Hinted games
'                   are walled to their hint; un-hinted games are unbounded (only
'                   floored at 0 unless negatives are allowed) - the search never
'                   declares a value permanently impossible.
'
'   Attribution   - when a submission scores 0 < k < (active games), the exact
'                   count is CARRIED through a bisection: a parent group of known
'                   count k is split, one half tested, and the sibling's count is
'                   inferred as k - (tested count) for free. A sheet-persisted
'                   stack of (group, known-count) means all-right / all-wrong
'                   subsets resolve entirely inside Excel with no re-test.
'
' Confirmed answers stay in the submitted column, so the leaderboard banks points
' continuously; the buttons show the ABSOLUTE score the platform will display.
' Every feedback snapshots state first, so one bad click can be undone.
'==============================================================================
Option Explicit

' ---- working-table columns ----
Private Const COL_GAME As Long = 1        'A  game numbers
Private Const COL_CORRECT As Long = 2     'B  confirmed answers (blank = unsolved)
Private Const COL_GUESS As Long = 3       'C  active candidate this round (unsolved only)
Private Const COL_SUBMIT As Long = 4      'D  =IF(ISNUMBER(B),B,C) - THE column to copy
Private Const COL_EMIN As Long = 5        'E  low edge of the contiguous eliminated block
Private Const COL_EMAX As Long = 6        'F  high edge of that block
Private Const COL_TRIED As Long = 7       'G  extra (non-contiguous) tried values, comma list
Private Const COL_HLO As Long = 8         'H  hard lower bound (hint) - blank if none
Private Const COL_HHI As Long = 9         'I  hard upper bound (hint) - blank if none
Private Const COL_INIT As Long = 10       'J  seed guess per game
Private Const COL_ATTR As Long = 11       'K  parked candidate value during attribution
Private Const COL_GRP As Long = 12        'L  attribution group tag (see below)
Private Const LAST_COL As Long = 12

' Group tags in COL_GRP: blank = not a candidate; >0 = pending stacked group id;
' GRP_ACTIVE = the half submitted now; GRP_SIB = its held sibling.
Private Const GRP_ACTIVE As Long = -1
Private Const GRP_SIB As Long = -2

' ---- setup-block / feedback cells ----
Private Const PTS_CELL As String = "B3"
Private Const GLO_CELL As String = "B4"
Private Const GHI_CELL As String = "C4"
Private Const EX_CELL As String = "B5"
Private Const SIG_CELL As String = "B6"
Private Const CENTER_CELL As String = "B10"
Private Const NEG_CELL As String = "B11"
Public Const FB_CELL As String = "N8"     ' operator types the platform POINTS here for the "8+" fallback

' ---- resolver state (scalars + LIFO stack), kept out to the right ----
Private Const ST_DEPTH As String = "T1"   ' stack depth
Private Const ST_SEQ As String = "T2"     ' next group id
Private Const ST_PARENTK As String = "T3" ' known count of the parent of the current split
Private Const ST_ROUND As String = "T4"   ' round counter
Private Const STK_ID_COL As Long = 21     'U  stack: group id per level
Private Const STK_CNT_COL As Long = 22    'V  stack: known count per level

' ---- undo backup (values only; laid out to the far right) ----
Private Const BAK_DATA_COL As Long = 40   ' backup of B..L starts here (col 40 = B, 41 = C, ...)
Private Const BAK_STATE_COL As Long = 52  ' backup of T..V starts here

Private Const HEADER_TEXT As String = "Game Numbers"
Private Const OPEN_HI As Double = 1E+300


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

    Dim hdrRow As Long
    hdrRow = gc_find_text_row(src, 2, "Game #", 1, 200)
    If hdrRow = 0 Then
        MsgBox "Could not find the 'Game #' table header on '" & src.Name & "'.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- example answer(s) ----
    Dim exVals() As Double, exN As Long
    ReDim exVals(1 To 16): exN = 0
    Dim r As Long, gnum As String, ans As Variant
    For r = hdrRow + 1 To hdrRow + 20
        gnum = UCase$(Trim$(CStr(src.Cells(r, 2).Value)))
        If Left$(gnum, 7) = "EXAMPLE" Then
            ans = src.Cells(r, 5).Value
            If Not IsEmpty(ans) And Trim$(CStr(ans)) <> "" Then
                If Not gc_is_plain_number(ans) Then
                    MsgBox "Guess and Check only works on numeric answers.", _
                           vbExclamation, "Guess and Check"
                    Exit Sub
                End If
                exN = exN + 1: exVals(exN) = CDbl(ans)
            End If
        End If
    Next r
    If exN = 0 Then
        MsgBox "Could not read a numeric example answer for level " & lvl & ".", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    Dim firstGame As Long, lastGame As Long
    If Not gc_game_range(src, firstGame, lastGame) Then
        MsgBox "Could not find the game numbers (column A) for level " & lvl & ".", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    ' ---- hints (top three games, or none): "Game #g ... between X and Y" ----
    Dim hMin(1 To 3) As Variant, hMax(1 To 3) As Variant
    Dim mn As Double, mx As Double, hg As Long, pos As Long, ctext As String
    For r = 1 To hdrRow - 1
        ctext = CStr(src.Cells(r, 3).Value)
        If InStr(1, ctext, "Hint:", vbTextCompare) > 0 Then
            hg = gc_parse_gamenum(ctext)
            If hg > 0 And gc_parse_between(ctext, mn, mx) Then
                pos = hg - firstGame + 1
                If pos >= 1 And pos <= 3 Then hMin(pos) = mn: hMax(pos) = mx
            End If
        End If
    Next r

    ' ---- inferences ----
    Dim samples() As Double, sN As Long, i As Long
    ReDim samples(1 To exN + 6): sN = 0
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
    Dim pts As Double: pts = gc_points_for_level(src.Parent, lvl)

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

    ' ---- feedback labels ----
    gc.Range("E6").Value = "Points on the platform (banked answers included):"
    gc.Range("N6").Value = "platform points"
    gc.Range("N7").Value = "(8+ / re-eval):"

    ' ---- diagnostics ----
    gc.Range("A12").Formula = "=COUNT(A14:A1000)"
    gc.Range("B12").Formula = "=COUNT(B14:B1000)"
    gc.Range("C12").Formula = "=COUNT(C14:C1000)"
    gc.Range("K12").Formula = "=COUNT(K14:K1000)"

    ' ---- table header ----
    gc.Range("A13").Value = HEADER_TEXT
    gc.Range("B13").Value = "Correct Answers"
    gc.Range("C13").Value = "Guess"
    gc.Range("D13").Value = "Submit  (copy this)"
    gc.Range("E13").Value = "Elim Min"
    gc.Range("F13").Value = "Elim Max"
    gc.Range("G13").Value = "Tried Extras"
    gc.Range("H13").Value = "Hint Lo"
    gc.Range("I13").Value = "Hint Hi"
    gc.Range("J13").Value = "Initial Guess"
    gc.Range("K13").Value = "Attribution"
    gc.Range("L13").Value = "Grp"

    ' ---- one row per game ----
    Dim nGames As Long, rr As Long
    nGames = lastGame - firstGame + 1
    For i = 1 To nGames
        rr = 13 + i
        gc.Cells(rr, COL_GAME).Value = firstGame + i - 1
        ' solved -> banked answer; else the active guess; else BLANK (not 0 - blank C
        ' coerces to 0, and 0 is a valid answer we must not submit for a held game)
        gc.Cells(rr, COL_SUBMIT).Formula = _
            "=IF(ISNUMBER(B" & rr & "),B" & rr & ",IF(C" & rr & "="""","""",C" & rr & "))"
        Select Case i
            Case 1: gc.Cells(rr, COL_INIT).Formula = "=IF(B7="""",$B$10,CEILING.MATH(AVERAGE(B7:C7),$B$6))"
            Case 2: gc.Cells(rr, COL_INIT).Formula = "=IF(B8="""",$B$10,CEILING.MATH(AVERAGE(B8:C8),$B$6))"
            Case 3: gc.Cells(rr, COL_INIT).Formula = "=IF(B9="""",$B$10,CEILING.MATH(AVERAGE(B9:C9),$B$6))"
            Case Else: gc.Cells(rr, COL_INIT).Formula = "=IF(A" & rr & "="""","""",$B$10)"
        End Select
        If i <= 3 Then
            If Not IsEmpty(hMin(i)) Then gc.Cells(rr, COL_HLO).Value = CDbl(hMin(i))
            If Not IsEmpty(hMax(i)) Then gc.Cells(rr, COL_HHI).Value = CDbl(hMax(i))
        End If
    Next i

    ' ---- resolver state = empty ----
    gc.Range(ST_DEPTH).Value = 0
    gc.Range(ST_SEQ).Value = 1
    gc.Range(ST_PARENTK).Value = 0
    gc.Range(ST_ROUND).Value = 0

    gc.Columns("A:L").AutoFit
    gc.Calculate

    Dim firstRow As Long, lastRow As Long
    If gc_locate(gc, firstRow, lastRow) Then
        gc_scan_regenerate gc, firstRow, lastRow, sig, (negAllowed <> 0)
        gc_format_sheet gc, firstRow, lastRow
        gc_place_buttons gc
        gc_recaption gc
        gc_copy_submit gc, firstRow, lastRow
    End If

    Application.ScreenUpdating = True
    gc.Activate
    On Error Resume Next
    ActiveWindow.Zoom = 130
    On Error GoTo 0
End Sub


'==============================================================================
' PUBLIC feedback handler (called by gc_buttons; the required argument keeps it
' out of Alt+F8). activeHits = number of the currently-SUBMITTED active guesses
' that were correct (confirmed answers are excluded - the caller already did
' points/P - confirmed, or a fixed 0..7 button).
'==============================================================================
Public Sub gc_feedback(ByVal activeHits As Long)
    Dim ws As Worksheet, firstRow As Long, lastRow As Long, sig As Double, neg As Boolean
    If Not gc_prep(ws, firstRow, lastRow, sig, neg) Then Exit Sub

    Dim attrib As Boolean: attrib = (gc_count(ws, firstRow, lastRow, COL_ATTR) > 0)
    Dim subCount As Long: subCount = gc_count(ws, firstRow, lastRow, COL_GUESS)

    ' ---- validation (protect state from a mis-read) ----
    If subCount = 0 Then
        MsgBox "There is nothing submitted to score right now.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    If activeHits < 0 Or activeHits > subCount Then
        MsgBox "That score implies " & activeHits & " correct among " & subCount & _
               " submitted guesses, which is impossible. Nothing was changed.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If
    If attrib Then
        Dim aSize As Long, bSize As Long, parentK As Long
        aSize = gc_count_tag(ws, firstRow, lastRow, GRP_ACTIVE)
        bSize = gc_count_tag(ws, firstRow, lastRow, GRP_SIB)
        parentK = CLng(ws.Range(ST_PARENTK).Value)
        If activeHits > parentK Or (parentK - activeHits) > bSize Then
            MsgBox "That count is inconsistent with the group being resolved " & _
                   "(parent had " & parentK & " correct). Nothing was changed.", _
                   vbExclamation, "Guess and Check"
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    gc_snapshot ws, firstRow, lastRow                 ' undo point BEFORE any change

    If attrib Then
        gc_attrib_step ws, firstRow, lastRow, sig, neg, activeHits
    Else
        gc_scan_step ws, firstRow, lastRow, sig, neg, activeHits
    End If

    ws.Range(ST_ROUND).Value = CLng(ws.Range(ST_ROUND).Value) + 1
    gc_recaption ws
    gc_copy_submit ws, firstRow, lastRow
    Application.ScreenUpdating = True
End Sub

' PUBLIC undo (arg keeps it out of Alt+F8) - restore the pre-feedback snapshot.
Public Sub gc_apply_undo(ByVal ws As Worksheet)
    Dim firstRow As Long, lastRow As Long
    If Not gc_locate(ws, firstRow, lastRow) Then
        MsgBox "This doesn't look like a guess-and-check sheet.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    If Trim$(CStr(ws.Cells(13, BAK_DATA_COL).Value)) = "" And _
       Trim$(CStr(ws.Cells(14, BAK_DATA_COL).Value)) = "" Then
        MsgBox "There is no snapshot to undo yet.", vbInformation, "Guess and Check"
        Exit Sub
    End If
    Application.ScreenUpdating = False

    ' restore B,C (cols 2,3) and E..L (cols 5..12); D is a formula, leave it
    ws.Range(ws.Cells(13, 2), ws.Cells(lastRow, 3)).Value = _
        ws.Range(ws.Cells(13, BAK_DATA_COL), ws.Cells(lastRow, BAK_DATA_COL + 1)).Value
    ws.Range(ws.Cells(13, 5), ws.Cells(lastRow, LAST_COL)).Value = _
        ws.Range(ws.Cells(13, BAK_DATA_COL + 3), ws.Cells(lastRow, BAK_DATA_COL + LAST_COL - 2)).Value
    ' restore scalars + stack (T..V)
    ws.Range(ws.Cells(1, 20), ws.Cells(60, 22)).Value = _
        ws.Range(ws.Cells(1, BAK_STATE_COL), ws.Cells(60, BAK_STATE_COL + 2)).Value

    gc_recaption ws
    gc_copy_submit ws, firstRow, lastRow
    Application.ScreenUpdating = True
    MsgBox "Reverted to before the last feedback.", vbInformation, "Guess and Check"
End Sub

' PUBLIC recovery from operator error (arg keeps it out of Alt+F8). Forgets which
' submitted answers were "confirmed" vs merely guessed, and re-derives purely from
' the fact that the CURRENT Submit column scored `truePoints`. It re-parks the whole
' current submission as one attribution group of known count truePoints/P, so a
' wrongly-confirmed answer is found and dropped. Solved-value / elimination history
' is kept, so the re-derivation is cheap.
Public Sub gc_do_reeval(ByVal ws As Worksheet, ByVal truePoints As Double)
    Dim fr As Long, lr As Long, sig As Double
    If Not gc_locate(ws, fr, lr) Then
        MsgBox "This doesn't look like a guess-and-check sheet.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    sig = CDbl(ws.Range(SIG_CELL).Value)
    Dim nv As Variant: nv = ws.Range(NEG_CELL).Value
    Dim neg As Boolean: neg = (IsNumeric(nv) And Val(CStr(nv)) <> 0)
    Dim p As Double
    If IsNumeric(ws.Range(PTS_CELL).Value) Then p = CDbl(ws.Range(PTS_CELL).Value)
    If p <= 0 Then
        MsgBox "Points Per Game (B3) is not set.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim tc As Double: tc = truePoints / p
    If Abs(tc - CLng(tc)) > 0.000001 Then
        MsgBox truePoints & " points is not a whole multiple of " & p & ".", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    Dim trueCorrect As Long: trueCorrect = CLng(tc)

    ' count the current submission (non-blank Submit) before touching anything
    Dim r As Long, subN As Long
    For r = fr To lr
        If gc_num(ws, r, COL_SUBMIT) Then subN = subN + 1
    Next r
    If subN = 0 Then
        MsgBox "Nothing is currently in the Submit column to re-evaluate.", vbExclamation, "Guess and Check"
        Exit Sub
    End If
    If trueCorrect < 0 Or trueCorrect > subN Then
        MsgBox trueCorrect & " correct is impossible for the " & subN & " values now submitted.", _
               vbExclamation, "Guess and Check"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    gc_snapshot ws, fr, lr                       ' undo point

    ' abandon any in-flight attribution stack, then re-park the whole submission
    Dim gid As Long: gid = gc_new_gid(ws)
    ws.Range(ws.Cells(1, STK_ID_COL), ws.Cells(1000, STK_CNT_COL)).ClearContents
    ws.Range(ST_DEPTH).Value = 0
    ws.Range(ST_PARENTK).Value = 0
    Dim dv As Variant
    For r = fr To lr
        dv = ws.Cells(r, COL_SUBMIT).Value       ' read D BEFORE clearing B/C
        ws.Cells(r, COL_CORRECT).ClearContents   ' un-confirm
        ws.Cells(r, COL_GUESS).ClearContents
        If IsNumeric(dv) And Trim$(CStr(dv)) <> "" Then
            ws.Cells(r, COL_ATTR).Value = CDbl(dv)
            ws.Cells(r, COL_GRP).Value = gid
        Else
            ws.Cells(r, COL_ATTR).ClearContents
            ws.Cells(r, COL_GRP).ClearContents
        End If
    Next r

    gc_push ws, gid, trueCorrect
    gc_pump ws, fr, lr, sig, neg                 ' handles all/none/mixed

    ws.Range(ST_ROUND).Value = CLng(ws.Range(ST_ROUND).Value) + 1
    gc_recaption ws
    gc_copy_submit ws, fr, lr
    Application.ScreenUpdating = True
    MsgBox "Re-evaluated from " & truePoints & " points (" & trueCorrect & " of " & subN & _
           " correct). Paste the Submit column and carry on.", vbInformation, "Guess and Check"
End Sub


'==============================================================================
' Scan step / attribution step (Private)
'==============================================================================
Private Sub gc_scan_step(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                         ByVal sig As Double, ByVal neg As Boolean, ByVal k As Long)
    Dim m As Long: m = gc_count(ws, fr, lr, COL_GUESS)
    Dim r As Long
    If k = 0 Then
        For r = fr To lr
            If gc_has(ws, r, COL_GUESS) Then
                gc_eliminate ws, r, CDbl(ws.Cells(r, COL_GUESS).Value), sig
                ws.Cells(r, COL_GUESS).ClearContents
            End If
        Next r
        gc_scan_regenerate ws, fr, lr, sig, neg
    ElseIf k = m Then
        For r = fr To lr
            If gc_has(ws, r, COL_GUESS) Then
                ws.Cells(r, COL_CORRECT).Value = ws.Cells(r, COL_GUESS).Value
                ws.Cells(r, COL_GUESS).ClearContents
            End If
        Next r
        gc_scan_regenerate ws, fr, lr, sig, neg
    Else
        ' park the whole submitted set as one group of known count k, then resolve
        Dim gid As Long: gid = gc_new_gid(ws)
        For r = fr To lr
            If gc_has(ws, r, COL_GUESS) Then
                ws.Cells(r, COL_ATTR).Value = ws.Cells(r, COL_GUESS).Value
                ws.Cells(r, COL_GRP).Value = gid
                ws.Cells(r, COL_GUESS).ClearContents
            End If
        Next r
        gc_push ws, gid, k
        gc_pump ws, fr, lr, sig, neg
    End If
End Sub

Private Sub gc_attrib_step(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                           ByVal sig As Double, ByVal neg As Boolean, ByVal a As Long)
    Dim parentK As Long: parentK = CLng(ws.Range(ST_PARENTK).Value)
    gc_absorb ws, fr, lr, GRP_ACTIVE, a, sig          ' tested half
    gc_absorb ws, fr, lr, GRP_SIB, parentK - a, sig   ' sibling, count carried for free
    gc_pump ws, fr, lr, sig, neg
End Sub

' Resolve every row tagged `tag`, whose known correct-count is `cnt`:
'   0        -> all wrong  (eliminate, free)
'   size     -> all right  (solve)
'   mixed    -> becomes a fresh pending group on the stack
Private Sub gc_absorb(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                      ByVal tag As Long, ByVal cnt As Long, ByVal sig As Double)
    Dim rows() As Long, n As Long
    gc_collect_tag ws, fr, lr, tag, rows, n
    If n = 0 Then Exit Sub
    Dim i As Long, r As Long
    If cnt <= 0 Then
        For i = 1 To n
            r = rows(i)
            gc_eliminate ws, r, CDbl(ws.Cells(r, COL_ATTR).Value), sig
            ws.Cells(r, COL_ATTR).ClearContents
            ws.Cells(r, COL_GRP).ClearContents
            ws.Cells(r, COL_GUESS).ClearContents
        Next i
    ElseIf cnt >= n Then
        For i = 1 To n
            r = rows(i)
            ws.Cells(r, COL_CORRECT).Value = ws.Cells(r, COL_ATTR).Value
            ws.Cells(r, COL_ATTR).ClearContents
            ws.Cells(r, COL_GRP).ClearContents
            ws.Cells(r, COL_GUESS).ClearContents
        Next i
    Else
        Dim gid As Long: gid = gc_new_gid(ws)
        For i = 1 To n
            r = rows(i)
            ws.Cells(r, COL_GRP).Value = gid
            ws.Cells(r, COL_GUESS).ClearContents
        Next i
        gc_push ws, gid, cnt
    End If
End Sub

' Process the stack until a group must be tested (submission set) or it is empty
' (attribution done -> back to scanning).
Private Sub gc_pump(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                    ByVal sig As Double, ByVal neg As Boolean)
    Do
        If gc_depth(ws) = 0 Then
            gc_scan_regenerate ws, fr, lr, sig, neg
            Exit Sub
        End If
        Dim gid As Long, cnt As Long
        gc_pop ws, gid, cnt
        Dim rows() As Long, n As Long
        gc_collect_tag ws, fr, lr, gid, rows, n
        Dim i As Long, r As Long
        If cnt <= 0 Or n = 0 Then
            For i = 1 To n
                r = rows(i)
                gc_eliminate ws, r, CDbl(ws.Cells(r, COL_ATTR).Value), sig
                ws.Cells(r, COL_ATTR).ClearContents: ws.Cells(r, COL_GRP).ClearContents
            Next i
        ElseIf cnt >= n Then
            For i = 1 To n
                r = rows(i)
                ws.Cells(r, COL_CORRECT).Value = ws.Cells(r, COL_ATTR).Value
                ws.Cells(r, COL_ATTR).ClearContents: ws.Cells(r, COL_GRP).ClearContents
            Next i
        Else
            ' split: first half active (submit), rest sibling (held)
            Dim keep As Long: keep = (n + 1) \ 2
            For i = 1 To n
                r = rows(i)
                If i <= keep Then
                    ws.Cells(r, COL_GRP).Value = GRP_ACTIVE
                    ws.Cells(r, COL_GUESS).Value = ws.Cells(r, COL_ATTR).Value
                Else
                    ws.Cells(r, COL_GRP).Value = GRP_SIB
                    ws.Cells(r, COL_GUESS).ClearContents
                End If
            Next i
            ws.Range(ST_PARENTK).Value = cnt
            Exit Sub                                   ' wait for the operator to report A's count
        End If
    Loop
End Sub


'==============================================================================
' Value scan (Private)
'==============================================================================

' Rebuild the Guess column for a fresh scanning round. Solved games stay blank;
' every unsolved game gets its next candidate (priority values, then outward
' scan). Also clears any attribution scratch (we are back in scan mode).
Private Sub gc_scan_regenerate(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                               ByVal sig As Double, ByVal neg As Boolean)
    Dim r As Long
    For r = fr To lr
        ws.Cells(r, COL_GUESS).ClearContents
        ws.Cells(r, COL_ATTR).ClearContents
        ws.Cells(r, COL_GRP).ClearContents
    Next r

    ' solved values -> centre + priority list
    Dim sv() As Double, ns As Long
    gc_gather_solved ws, fr, lr, sv, ns
    Dim baseCenter As Double: baseCenter = CDbl(ws.Range(CENTER_CELL).Value)
    Dim center As Double: center = gc_center(sv, ns, baseCenter)
    Dim prio() As Double, np As Long
    gc_priority_values sv, ns, sig, prio, np

    Dim nUnsolved As Long, nGuess As Long, g As Variant
    For r = fr To lr
        If Not gc_is_solved(ws, r) Then
            nUnsolved = nUnsolved + 1
            g = gc_next_guess(ws, r, sig, neg, center, prio, np)
            If Not IsEmpty(g) Then ws.Cells(r, COL_GUESS).Value = CDbl(g): nGuess = nGuess + 1
        End If
    Next r

    If nUnsolved > 0 And nGuess = 0 Then
        MsgBox "Every remaining game is a hinted game whose hint range is fully " & _
               "eliminated. Check the hints / significance (" & SIG_CELL & ").", _
               vbExclamation, "Guess and Check"
    End If
End Sub

' Next value to try for game r. Priority values first (confirmed elsewhere, then
' example) if untried & in-bounds, else the nearest-to-centre untried value.
' Returns Empty only when a HINTED game's hard range is exhausted.
Private Function gc_next_guess(ByVal ws As Worksheet, ByVal r As Long, ByVal sig As Double, _
                               ByVal neg As Boolean, ByVal center As Double, _
                               ByRef prio() As Double, ByVal np As Long) As Variant
    Dim hardLo As Boolean, hardHi As Boolean, lo As Double, hi As Double
    hardLo = gc_num(ws, r, COL_HLO): If hardLo Then lo = CDbl(ws.Cells(r, COL_HLO).Value)
    hardHi = gc_num(ws, r, COL_HHI): If hardHi Then hi = CDbl(ws.Cells(r, COL_HHI).Value)
    Dim eps As Double: eps = sig * 0.000001

    Dim i As Long, v As Double
    For i = 1 To np
        v = prio(i)
        If gc_in_bounds(v, hardLo, lo, hardHi, hi, neg, eps) Then
            If Not gc_tried(ws, r, v, sig) Then gc_next_guess = v: Exit Function
        End If
    Next i

    ' outward scan from the (grid-snapped) centre. A candidate is only returned if
    ' it is inside BOTH bounds (gc_in_bounds); the single-sided loPast/hiPast tests
    ' are only for deciding a hinted range is exhausted (never fires when unbounded).
    Dim c As Double: c = gc_snap(center, sig)
    Dim k As Long, vlo As Double, vhi As Double, loPast As Boolean, hiPast As Boolean
    For k = 0 To 200000
        vlo = c - k * sig: vhi = c + k * sig
        If gc_in_bounds(vlo, hardLo, lo, hardHi, hi, neg, eps) Then
            If Not gc_tried(ws, r, vlo, sig) Then gc_next_guess = vlo: Exit Function
        End If
        If k > 0 Then
            If gc_in_bounds(vhi, hardLo, lo, hardHi, hi, neg, eps) Then
                If Not gc_tried(ws, r, vhi, sig) Then gc_next_guess = vhi: Exit Function
            End If
        End If
        loPast = (hardLo And vlo < lo - eps) Or ((Not neg) And vlo < -eps)
        hiPast = (hardHi And vhi > hi + eps)
        If loPast And hiPast Then Exit For         ' hinted game: whole range tried -> exhausted
    Next k
    gc_next_guess = Empty
End Function

' Fold a proven-wrong value into game r's tried state: extend the contiguous
' block when adjacent, otherwise record it in the Tried Extras list.
Private Sub gc_eliminate(ByVal ws As Worksheet, ByVal r As Long, ByVal v As Double, ByVal sig As Double)
    Dim eps As Double: eps = sig * 0.000001
    If Not gc_num(ws, r, COL_EMIN) Then
        ws.Cells(r, COL_EMIN).Value = v: ws.Cells(r, COL_EMAX).Value = v
        Exit Sub
    End If
    Dim emin As Double, emax As Double
    emin = CDbl(ws.Cells(r, COL_EMIN).Value): emax = CDbl(ws.Cells(r, COL_EMAX).Value)
    If v >= emin - eps And v <= emax + eps Then Exit Sub          ' already in block
    If Abs(v - (emin - sig)) < eps Then ws.Cells(r, COL_EMIN).Value = v: Exit Sub
    If Abs(v - (emax + sig)) < eps Then ws.Cells(r, COL_EMAX).Value = v: Exit Sub
    gc_tried_add ws, r, v, sig
End Sub

' Has game r already tried value v? (contiguous block OR the extras list)
Private Function gc_tried(ByVal ws As Worksheet, ByVal r As Long, ByVal v As Double, ByVal sig As Double) As Boolean
    Dim eps As Double: eps = sig * 0.000001
    If gc_num(ws, r, COL_EMIN) Then
        If v >= CDbl(ws.Cells(r, COL_EMIN).Value) - eps And _
           v <= CDbl(ws.Cells(r, COL_EMAX).Value) + eps Then gc_tried = True: Exit Function
    End If
    Dim s As String: s = CStr(ws.Cells(r, COL_TRIED).Value)
    If Len(s) = 0 Then Exit Function
    Dim parts() As String, i As Long
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            If Abs(CDbl(parts(i)) - v) < eps Then gc_tried = True: Exit Function
        End If
    Next i
End Function

Private Sub gc_tried_add(ByVal ws As Worksheet, ByVal r As Long, ByVal v As Double, ByVal sig As Double)
    If gc_tried(ws, r, v, sig) Then Exit Sub
    Dim s As String: s = CStr(ws.Cells(r, COL_TRIED).Value)
    If Len(s) = 0 Then s = CStr(v) Else s = s & "," & CStr(v)
    ws.Cells(r, COL_TRIED).Value = s
End Sub

Private Function gc_in_bounds(ByVal v As Double, ByVal hardLo As Boolean, ByVal lo As Double, _
                              ByVal hardHi As Boolean, ByVal hi As Double, _
                              ByVal neg As Boolean, ByVal eps As Double) As Boolean
    If hardLo And v < lo - eps Then Exit Function
    If hardHi And v > hi + eps Then Exit Function
    If (Not neg) And v < -eps Then Exit Function
    gc_in_bounds = True
End Function

Private Function gc_snap(ByVal x As Double, ByVal sig As Double) As Double
    gc_snap = CDbl(CLng(x / sig)) * sig
End Function


'==============================================================================
' Centre + priority values (Private)
'==============================================================================
Private Sub gc_gather_solved(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                             ByRef sv() As Double, ByRef n As Long)
    ReDim sv(1 To (lr - fr + 1)): n = 0
    Dim r As Long, v As Variant
    For r = fr To lr
        v = ws.Cells(r, COL_CORRECT).Value
        If IsNumeric(v) And Trim$(CStr(v)) <> "" Then n = n + 1: sv(n) = CDbl(v)
    Next r
End Sub

' Search centre: mode of solved answers if one repeats, else their median, else
' the (example/hint-blended) build centre. Solved data beats the seed once it exists.
Private Function gc_center(ByRef sv() As Double, ByVal n As Long, ByVal fallback As Double) As Double
    If n = 0 Then gc_center = fallback: Exit Function
    Dim a() As Double: a = gc_sorted(sv, n)
    Dim bestV As Double, bestF As Long, curV As Double, curF As Long, i As Long
    bestF = 0: curF = 0
    For i = 1 To n
        If i = 1 Or a(i) <> curV Then curV = a(i): curF = 1 Else curF = curF + 1
        If curF > bestF Then bestF = curF: bestV = curV
    Next i
    If bestF > 1 Then gc_center = bestV Else gc_center = a((n + 1) \ 2)   ' mode, else median
End Function

' Priority list: the distinct CONFIRMED values on this level, most frequent first
' (ties low). These are the values worth re-testing across the other games first,
' because a level's answers tend to repeat. The example answer is deliberately NOT
' here - it is a different game's answer and already feeds the centre (B10); round 1
' (nothing solved) therefore probes the centre, not the example.
Private Sub gc_priority_values(ByRef sv() As Double, ByVal n As Long, _
                               ByVal sig As Double, ByRef prio() As Double, ByRef np As Long)
    np = 0
    If n = 0 Then ReDim prio(1 To 1): Exit Sub
    ReDim prio(1 To n)
    Dim eps As Double: eps = sig * 0.000001
    Dim a() As Double: a = gc_sorted(sv, n)
    ' distinct values + frequency
    Dim vals() As Double, fq() As Long, d As Long, i As Long, isNew As Boolean
    ReDim vals(1 To n): ReDim fq(1 To n): d = 0
    For i = 1 To n
        If d = 0 Then                          ' VBA has no short-circuit: guard vals(d) separately
            isNew = True
        Else
            isNew = (Abs(a(i) - vals(d)) > eps)
        End If
        If isNew Then
            d = d + 1: vals(d) = a(i): fq(d) = 1
        Else
            fq(d) = fq(d) + 1
        End If
    Next i
    ' selection-sort distinct by freq desc (ties keep ascending value)
    Dim j As Long, bi As Long
    For i = 1 To d
        bi = i
        For j = i + 1 To d
            If fq(j) > fq(bi) Then bi = j
        Next j
        If bi <> i Then
            Dim tv As Double, tf As Long
            tv = vals(i): vals(i) = vals(bi): vals(bi) = tv
            tf = fq(i): fq(i) = fq(bi): fq(bi) = tf
        End If
        np = np + 1: prio(np) = vals(i)
    Next i
End Sub

Private Function gc_sorted(ByRef src() As Double, ByVal n As Long) As Double()
    Dim a() As Double: ReDim a(1 To n)
    Dim i As Long: For i = 1 To n: a(i) = src(i): Next i
    Dim j As Long, key As Double
    For i = 2 To n
        key = a(i): j = i - 1
        Do While j >= 1
            If a(j) <= key Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = key
    Next i
    gc_sorted = a
End Function


'==============================================================================
' Resolver stack + small utilities (Private)
'==============================================================================
Private Function gc_depth(ByVal ws As Worksheet) As Long
    gc_depth = CLng(ws.Range(ST_DEPTH).Value)
End Function
Private Function gc_new_gid(ByVal ws As Worksheet) As Long
    gc_new_gid = CLng(ws.Range(ST_SEQ).Value)
    ws.Range(ST_SEQ).Value = gc_new_gid + 1
End Function
Private Sub gc_push(ByVal ws As Worksheet, ByVal gid As Long, ByVal cnt As Long)
    Dim d As Long: d = gc_depth(ws) + 1
    ws.Cells(d, STK_ID_COL).Value = gid
    ws.Cells(d, STK_CNT_COL).Value = cnt
    ws.Range(ST_DEPTH).Value = d
End Sub
Private Sub gc_pop(ByVal ws As Worksheet, ByRef gid As Long, ByRef cnt As Long)
    Dim d As Long: d = gc_depth(ws)
    gid = CLng(ws.Cells(d, STK_ID_COL).Value)
    cnt = CLng(ws.Cells(d, STK_CNT_COL).Value)
    ws.Cells(d, STK_ID_COL).ClearContents: ws.Cells(d, STK_CNT_COL).ClearContents
    ws.Range(ST_DEPTH).Value = d - 1
End Sub

Private Function gc_prep(ByRef ws As Worksheet, ByRef fr As Long, ByRef lr As Long, _
                         ByRef sig As Double, ByRef neg As Boolean) As Boolean
    Set ws = ActiveSheet
    If Not gc_locate(ws, fr, lr) Then
        MsgBox "This doesn't look like a guess-and-check sheet.", vbExclamation, "Guess and Check"
        Exit Function
    End If
    sig = CDbl(ws.Range(SIG_CELL).Value)
    Dim nv As Variant: nv = ws.Range(NEG_CELL).Value
    neg = (IsNumeric(nv) And Val(CStr(nv)) <> 0)
    gc_prep = True
End Function

Private Function gc_locate(ByVal ws As Worksheet, ByRef fr As Long, ByRef lr As Long) As Boolean
    fr = 0
    Dim r As Long
    For r = 1 To 100
        If LCase$(Trim$(CStr(ws.Cells(r, COL_GAME).Value))) = LCase$(HEADER_TEXT) Then fr = r + 1: Exit For
    Next r
    If fr = 0 Then Exit Function
    lr = ws.Cells(ws.Rows.Count, COL_GAME).End(xlUp).Row
    gc_locate = (lr >= fr)
End Function

Private Function gc_has(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long) As Boolean
    gc_has = (Trim$(CStr(ws.Cells(r, col).Value)) <> "")
End Function
Private Function gc_num(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long) As Boolean
    Dim v As Variant: v = ws.Cells(r, col).Value
    gc_num = (IsNumeric(v) And Trim$(CStr(v)) <> "")
End Function
Private Function gc_is_solved(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    gc_is_solved = gc_num(ws, r, COL_CORRECT)
End Function
Private Function gc_count(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, ByVal col As Long) As Long
    Dim r As Long
    For r = fr To lr
        If gc_has(ws, r, col) Then gc_count = gc_count + 1
    Next r
End Function
Private Function gc_count_tag(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, ByVal tag As Long) As Long
    Dim r As Long
    For r = fr To lr
        If gc_num(ws, r, COL_GRP) Then If CLng(ws.Cells(r, COL_GRP).Value) = tag Then gc_count_tag = gc_count_tag + 1
    Next r
End Function
Private Sub gc_collect_tag(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long, _
                           ByVal tag As Long, ByRef rows() As Long, ByRef n As Long)
    ReDim rows(1 To (lr - fr + 1)): n = 0
    Dim r As Long
    For r = fr To lr
        If gc_num(ws, r, COL_GRP) Then If CLng(ws.Cells(r, COL_GRP).Value) = tag Then n = n + 1: rows(n) = r
    Next r
End Sub

' Copy the Submit column (confirmed answers + active guesses) for the operator.
Private Sub gc_copy_submit(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long)
    ws.Range(ws.Cells(fr, COL_SUBMIT), ws.Cells(lr, COL_SUBMIT)).Copy
End Sub

' Snapshot mutable state (values only, no clipboard) for one-level undo.
Private Sub gc_snapshot(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long)
    ws.Range(ws.Cells(13, BAK_DATA_COL), ws.Cells(lr, BAK_DATA_COL + LAST_COL - 2)).Value = _
        ws.Range(ws.Cells(13, 2), ws.Cells(lr, LAST_COL)).Value
    ws.Range(ws.Cells(1, BAK_STATE_COL), ws.Cells(60, BAK_STATE_COL + 2)).Value = _
        ws.Range(ws.Cells(1, 20), ws.Cells(60, 22)).Value
End Sub


'==============================================================================
' Buttons + formatting (Private)
'==============================================================================
Private Sub gc_place_buttons(ByVal ws As Worksheet)
    gc_add_button ws, ws.Range("E8:F9"), "gc_fb0", "gc_btn0"
    gc_add_button ws, ws.Range("G8:H9"), "gc_fb1", "gc_btn1"
    gc_add_button ws, ws.Range("I8:J9"), "gc_fb2", "gc_btn2"
    gc_add_button ws, ws.Range("K8:L9"), "gc_fb3", "gc_btn3"
    gc_add_button ws, ws.Range("E10:F11"), "gc_fb4", "gc_btn4"
    gc_add_button ws, ws.Range("G10:H11"), "gc_fb5", "gc_btn5"
    gc_add_button ws, ws.Range("I10:J11"), "gc_fb6", "gc_btn6"
    gc_add_button ws, ws.Range("K10:L11"), "gc_fb7", "gc_btn7"
    gc_add_button ws, ws.Range("O8:P9"), "gc_fbN", "gc_btnN"
    gc_add_button ws, ws.Range("O10:P11"), "gc_undo", "gc_btnUndo"
    gc_add_button ws, ws.Range("Q8:R11"), "gc_reeval", "gc_btnReeval"
End Sub

Private Sub gc_add_button(ByVal ws As Worksheet, ByVal rng As Range, ByVal macro As String, ByVal nm As String)
    Dim b As Object
    Set b = ws.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    b.Name = nm
    b.OnAction = macro
End Sub

' Re-label the buttons to the ABSOLUTE score the platform will show: a button for
' "j more correct" reads (confirmed + j) * P. Refreshed after every feedback.
Private Sub gc_recaption(ByVal ws As Worksheet)
    Dim p As Double, confirmed As Long
    If IsNumeric(ws.Range(PTS_CELL).Value) Then p = CDbl(ws.Range(PTS_CELL).Value)
    confirmed = Application.WorksheetFunction.Count(ws.Range("B14:B100000"))
    Dim j As Long
    On Error Resume Next
    For j = 0 To 7
        ws.Buttons("gc_btn" & j).Caption = gc_pts_caption(confirmed + j, p)
    Next j
    ws.Buttons("gc_btnN").Caption = "8+ pts"
    ws.Buttons("gc_btnUndo").Caption = "Undo"
    ws.Buttons("gc_btnReeval").Caption = "Re-evaluate" & vbLf & "(from N8 points)"
    On Error GoTo 0
End Sub

Private Function gc_pts_caption(ByVal correctTotal As Long, ByVal p As Double) As String
    If p > 0 Then gc_pts_caption = CStr(correctTotal * p) Else gc_pts_caption = CStr(correctTotal)
End Function

Private Sub gc_format_sheet(ByVal ws As Worksheet, ByVal fr As Long, ByVal lr As Long)
    Dim hdr As Long: hdr = fr - 1
    Dim cNavy As Long, cLabel As Long, cGreen As Long, cSubmitH As Long, cSubmit As Long, cHead As Long, cInput As Long, cGrid As Long
    cNavy = RGB(31, 78, 120): cLabel = RGB(221, 235, 247): cGreen = RGB(226, 239, 218)
    cSubmitH = RGB(191, 143, 0): cSubmit = RGB(255, 242, 204): cHead = RGB(47, 117, 181)
    cInput = RGB(255, 255, 204): cGrid = RGB(200, 200, 200)
    With ws
        .Range("A1").Value = "Guess & Check  -  Level " & .Range("B2").Value
        .Range("A1").Font.Bold = True: .Range("A1").Font.Size = 14: .Range("A1").Font.Color = cNavy
        .Range("A2:A11").Font.Bold = True: .Range("A2:A11").Interior.Color = cLabel
        With .Range("B2:C11")
            .Interior.Color = RGB(255, 255, 255): .Borders.LineStyle = xlContinuous: .Borders.Color = cGrid
        End With
        .Range("E6").Font.Bold = True: .Range("N6").Font.Bold = True: .Range("N7").Font.Size = 9
        With .Range(FB_CELL)
            .Interior.Color = cInput: .Borders.LineStyle = xlContinuous: .Borders.Weight = xlMedium: .HorizontalAlignment = xlCenter
        End With
        With .Range(.Cells(hdr, COL_GAME), .Cells(hdr, LAST_COL))
            .Interior.Color = cHead: .Font.Color = RGB(255, 255, 255): .Font.Bold = True
            .HorizontalAlignment = xlCenter: .WrapText = True
        End With
        .Cells(hdr, COL_SUBMIT).Interior.Color = cSubmitH
        If lr >= fr Then
            With .Range(.Cells(fr, COL_GAME), .Cells(lr, LAST_COL)).Borders
                .LineStyle = xlContinuous: .Color = cGrid
            End With
            .Range(.Cells(fr, COL_GAME), .Cells(lr, LAST_COL)).HorizontalAlignment = xlCenter
            .Range(.Cells(fr, COL_CORRECT), .Cells(lr, COL_CORRECT)).Interior.Color = cGreen
            .Range(.Cells(fr, COL_SUBMIT), .Cells(lr, COL_SUBMIT)).Interior.Color = cSubmit
        End If
        .Columns("A:L").ColumnWidth = 11
        .Columns("A").ColumnWidth = 13
        .Rows(hdr).RowHeight = 28
        .Columns("T:BB").Hidden = True                ' resolver state + undo backup
    End With
End Sub


'==============================================================================
' Parsing / inference helpers (Private)
'==============================================================================
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

Private Function gc_find_text_row(ByVal ws As Worksheet, ByVal col As Long, ByVal text As String, _
                                  ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If LCase$(Trim$(CStr(ws.Cells(r, col).Value))) = LCase$(text) Then gc_find_text_row = r: Exit Function
    Next r
End Function

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

Private Function gc_parse_gamenum(ByVal s As String) As Long
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp"): re.IgnoreCase = True: re.Pattern = "Game\s*#\s*(\d+)"
    End If
    If re.Test(s) Then gc_parse_gamenum = CLng(re.Execute(s)(0).SubMatches(0))
End Function

Private Function gc_parse_between(ByVal s As String, ByRef mn As Double, ByRef mx As Double) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp"): re.IgnoreCase = True
        re.Pattern = "between\s+(-?\d+(?:\.\d+)?)\s+and\s+(-?\d+(?:\.\d+)?)"
    End If
    If re.Test(s) Then
        Dim mch As Object: Set mch = re.Execute(s)(0)
        mn = CDbl(mch.SubMatches(0)): mx = CDbl(mch.SubMatches(1)): gc_parse_between = True
    End If
End Function

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

' Significance = the FINEST step the samples justify, never coarser than 1.
' All-integer samples -> 1 (do NOT infer a coarse step just because 20 divides by
' 10). Decimal samples -> the smallest decimal unit that divides them exactly.
Private Function gc_significance(ByRef vals() As Double, ByVal n As Long) As Double
    Dim i As Long, s As Double, best As Double: best = 1
    For i = 1 To n
        s = gc_decimal_sig(vals(i))
        If s < best Then best = s
    Next i
    gc_significance = best
End Function

Private Function gc_decimal_sig(ByVal x As Double) As Double
    Dim ax As Double: ax = Abs(x)
    If ax = 0 Then gc_decimal_sig = 1: Exit Function
    Dim k As Long, s As Double, q As Double
    For k = 0 To -6 Step -1                    ' 1, 0.1, 0.01, ... (never > 1)
        s = 10 ^ k: q = ax / s
        If Abs(q - CDbl(CLng(q))) < 0.0000001 Then gc_decimal_sig = s: Exit Function
    Next k
    gc_decimal_sig = 0.000001
End Function

Private Function gc_points_for_level(ByVal wb As Workbook, ByVal lvl As Long) As Double
    Dim cand As Variant, nm As Variant, ws As Worksheet, r As Long, lastR As Long
    cand = Array("Case", "case copy", "Answers", "case data")
    For Each nm In cand
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For r = 1 To lastR
                If IsNumeric(ws.Cells(r, 2).Value) And ws.Cells(r, 3).Value = lvl Then
                    If IsNumeric(ws.Cells(r, 4).Value) Then
                        If CDbl(ws.Cells(r, 4).Value) > 0 Then gc_points_for_level = CDbl(ws.Cells(r, 4).Value): Exit Function
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

Private Function gc_unique_name(ByVal wb As Workbook, ByVal base As String) As String
    If Not gc_sheet_exists(wb, base) Then gc_unique_name = base: Exit Function
    Dim k As Long
    For k = 2 To 999
        If Not gc_sheet_exists(wb, base & "(" & k & ")") Then gc_unique_name = base & "(" & k & ")": Exit Function
    Next k
    gc_unique_name = base & "(x)"
End Function
