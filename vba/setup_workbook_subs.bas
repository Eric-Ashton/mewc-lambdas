Attribute VB_Name = "setup_workbook_subs"
' deploy: shared
' Module-level declarations

' Workbooks
Public attempt_workbook As Workbook
Public case_workbook As Workbook

' Worksheets
Public prep_worksheet As Worksheet
Public case_worksheet As Worksheet
Public case_copy As Worksheet
Public all_games_worksheet As Worksheet

' Strings
Public attempt_filename As String
Public case_path As String
Public case_filename As String

' Integers and Arrays
Public last_row_case As Long
Public last_col_case As Long
Public number_of_levels As Long
Public last_col_by_level() As Long
Public last_col_all_levels As Long
Public example_row_numbers() As Long
Public header_row_numbers() As Long
Public yellow_cell_rows() As Long
Public offset_rows_needed() As Long
Public case_link_rows() As Long
Public level_link_rows() As Long
    
Sub setup_workbook()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set attempt_workbook = ThisWorkbook
    ThisWorkbook.Save

    ' --- Best-effort pipeline -------------------------------------------------
    ' Tuned for the MEWC competition use-case: there's no time to debug a
    ' stage failure mid-run, and a partially-completed template is more useful
    ' than nothing. Each stage Sub still logs its own failures to ErrorLog
    ' (via its internal ErrorHandler -> LogError -> Err.Raise). We absorb the
    ' re-raise here with On Error Resume Next + Err.Clear so later stages
    ' still get a shot at running. Stages whose preconditions weren't met
    ' will likely log secondary failures of their own - the FIRST entry in
    ' ErrorLog is the real cause; the rest are cascades.
    On Error Resume Next

    Call import_case
    Err.Clear
    Call classify_rows
    Err.Clear
    Call create_level_worksheets
    Err.Clear
    Call align_level_worksheets
    Err.Clear
    Call create_all_games_worksheet
    Err.Clear
    Call zoom_all_worksheets
    Err.Clear
    Call create_internal_links
    Err.Clear
    Call create_hints_sheet
    Err.Clear

    ' Final UI nicety - guarded the same way (no _L1 sheet means earlier
    ' stages failed; we don't want that to abort the cleanup).
    Worksheets("_L1").Select
    Err.Clear

    On Error GoTo ErrorHandler
    ' --- end best-effort pipeline --------------------------------------------

Cleanup:
    ' Hide the ErrorLog sheet if it exists. Use SheetExists rather than
    ' exception-based lookup so the VBA debugger's "Break on All Errors"
    ' setting can't trip on a missing sheet.
    Dim logSheet As Worksheet
    If SheetExists(ThisWorkbook, "ErrorLog") Then
        Set logSheet = ThisWorkbook.Sheets("ErrorLog")
        logSheet.Visible = xlSheetHidden
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Call LogError(Err.Number, Err.Description, "setup_workbook")
    Resume Cleanup

End Sub


' === Import Case Subroutine ===
Private Sub import_case()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Dim full_case_path As String
    
    
    Set prep_worksheet = attempt_workbook.Sheets("Prep")
    attempt_filename = prep_worksheet.Range("B3").Value
    case_path = prep_worksheet.Range("B1").Value
    case_filename = prep_worksheet.Range("B2").Value
    
    'saves workbook as new filename given in B3 of prep
   If attempt_filename <> "" Then
        Dim full_attempt_path As String
        full_attempt_path = case_path & "\" & attempt_filename

        ' Check if file already exists
        If Dir(full_attempt_path) <> "" Then
            overwrite = MsgBox("The target filename already exists." & vbCrLf & _
                           "Do you want to overwrite it?", vbYesNo + vbQuestion)
            If overwrite <> vbYes Then Exit Sub
            ' Back up the file we're about to overwrite so a crashed or
            ' mid-run-failed setup doesn't destroy a previously-good attempt.
            Call BackupExistingFile(full_attempt_path)
        End If
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = False
       attempt_workbook.SaveAs filename:=full_attempt_path
      Application.DisplayAlerts = True
      Application.Calculation = xlCalculationManual
   End If

    
    'opens case workbook and copies over worksheets
    full_case_path = case_path & Application.PathSeparator & case_filename
    If Dir(full_case_path) = "" Then
        MsgBox "Case file not found: " & full_case_path, vbCritical
        Exit Sub
    End If

    ' Open the case workbook defensively:
    '   UpdateLinks:=0                  -> don't refresh external links on open
    '   ReadOnly:=True                  -> guarantee we can't modify the author's file
    '   IgnoreReadOnlyRecommended:=True -> skip the "open as read-only?" prompt
    ' Also disable auto-running of any macros embedded in the case workbook for
    ' the duration of the open. AutomationSecurity is restored in the cleanup
    ' path above setup_workbook via a module-level save/restore. Belt-and-
    ' suspenders: also disable the link-update prompt at the Application level.
    Dim prev_auto_sec As Long
    Dim prev_ask_update As Boolean
    prev_auto_sec = Application.AutomationSecurity
    prev_ask_update = Application.AskToUpdateLinks
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.AskToUpdateLinks = False
    Set case_workbook = Workbooks.Open(filename:=full_case_path, _
                                       UpdateLinks:=0, _
                                       ReadOnly:=True, _
                                       IgnoreReadOnlyRecommended:=True)
    Application.AskToUpdateLinks = prev_ask_update
    Application.AutomationSecurity = prev_auto_sec

    ' --- Candidate names for the case-content sheet ---
    ' Add new candidates to this array as future cases introduce new names.
    Dim candidate_names As Variant
    candidate_names = Array("Case", "Challenge", "Task", _
                            "Question", "Problem", "Main", _
                            "Puzzle", "Round")
    Dim cand As Variant

    ' --- Remove any leftover candidate sheets in the destination workbook ---
    ' If we don't do this, Excel auto-renames the newly-copied sheet to
    ' "Case (2)" etc. and the lookup binds to the stale template copy.
    ' We also remove any previously-produced "case copy" so MakeCaseCopy
    ' starts clean.
    Dim alerts_prev As Boolean
    alerts_prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each cand In candidate_names
        If SheetExists(attempt_workbook, CStr(cand)) Then
            attempt_workbook.Worksheets(CStr(cand)).Delete
        End If
    Next cand
    If SheetExists(attempt_workbook, "case copy") Then
        attempt_workbook.Worksheets("case copy").Delete
    End If
    Application.DisplayAlerts = alerts_prev

    ' Copy all sheets at the same time to preserve internal links.
    ' Worksheets.Copy requires every source sheet to be xlSheetVisible,
    ' so temporarily unhide anything that isn't and restore state afterward.
    Dim srcSheetCount As Long
    srcSheetCount = case_workbook.Worksheets.count

    Dim origVisibility() As Long
    ReDim origVisibility(1 To srcSheetCount)
    Dim k As Long
    For k = 1 To srcSheetCount
        origVisibility(k) = case_workbook.Worksheets(k).Visible
        case_workbook.Worksheets(k).Visible = xlSheetVisible
    Next k

    ' Anchor the copy on a (now-visible) sheet
    case_workbook.Worksheets(1).Activate

    Dim destBase As Long
    destBase = attempt_workbook.Sheets.count
    case_workbook.Worksheets.Copy _
        After:=attempt_workbook.Sheets(destBase)

    ' Move focus off the just-copied sheets so we can re-hide any of them
    prep_worksheet.Activate

    ' Restore each copied sheet's original visibility on the destination
    For k = 1 To srcSheetCount
        On Error Resume Next
        attempt_workbook.Sheets(destBase + k).Visible = origVisibility(k)
        On Error GoTo ErrorHandler
    Next k

    case_workbook.Close SaveChanges:=False

    ' --- Locate the case-content sheet among the sheets we just copied ---
    ' Only inspect the freshly-imported range (destBase+1 .. destBase+srcSheetCount)
    ' so we cannot accidentally bind to a stale template sheet. We also accept
    ' names that Excel has auto-suffixed with " (2)", " (3)", etc. in case a
    ' collision still slipped through.
    Set case_worksheet = Nothing
    Dim ws_i As Long
    Dim cand_ws As Worksheet
    Dim cand_name As String, base_name As String
    Dim paren_pos As Long
    For ws_i = destBase + 1 To destBase + srcSheetCount
        Set cand_ws = attempt_workbook.Sheets(ws_i)
        cand_name = cand_ws.Name
        ' Strip any trailing " (N)" Excel appended for name-collision resolution
        base_name = cand_name
        paren_pos = InStrRev(base_name, " (")
        If paren_pos > 0 Then
            If Right$(base_name, 1) = ")" Then
                Dim inner As String
                inner = Mid$(base_name, paren_pos + 2, Len(base_name) - paren_pos - 2)
                If IsNumeric(inner) Then base_name = Left$(base_name, paren_pos - 1)
            End If
        End If
        For Each cand In candidate_names
            If StrComp(base_name, CStr(cand), vbTextCompare) = 0 Then
                Set case_worksheet = cand_ws
                Exit For
            End If
        Next cand
        If Not case_worksheet Is Nothing Then Exit For
    Next ws_i

    ' --- Fallback: no alias matched, pick the visible imported sheet with ---
    ' --- the most used rows and ask the user to confirm. Gives us a manual ---
    ' --- override for cases whose author used a completely novel sheet name. ---
    If case_worksheet Is Nothing Then
        Dim best_ws As Worksheet
        Dim best_rows As Long
        Dim cand_rows As Long
        best_rows = 0
        For ws_i = destBase + 1 To destBase + srcSheetCount
            Set cand_ws = attempt_workbook.Sheets(ws_i)
            If cand_ws.Visible = xlSheetVisible Then
                cand_rows = GetLastUsedRow(cand_ws)
                If cand_rows > best_rows Then
                    best_rows = cand_rows
                    Set best_ws = cand_ws
                End If
            End If
        Next ws_i
        If Not best_ws Is Nothing Then
            If MsgBox("No imported sheet matched the expected names (" & _
                      Join(candidate_names, ", ") & ")." & vbCrLf & vbCrLf & _
                      "Use """ & best_ws.Name & """ as the case sheet?", _
                      vbYesNo + vbQuestion, "Case sheet not found") = vbYes Then
                Set case_worksheet = best_ws
            End If
        End If
    End If

    ' Standardize the name to "Case" so the rest of the pipeline keeps working
    If Not case_worksheet Is Nothing Then
        If case_worksheet.Name <> "Case" Then
            On Error Resume Next
            case_worksheet.Name = "Case"
            On Error GoTo ErrorHandler
        End If
    Else
        MsgBox "Error: none of " & Join(candidate_names, ", ") & _
               " worksheets were found in the imported case.", vbCritical
        Exit Sub
    End If

    ' --- Remove third-party add-in tracking sheets (done AFTER case lookup) ---
    ' Some cases are saved by authors who have Capital IQ / AlphaSense / similar
    ' add-ins installed. Those add-ins silently inject veryHidden tracking
    ' sheets like "_avcts_version_hash". Besides being dead weight, their
    ' presence at the tail of the Worksheets collection breaks Copy After:=
    ' position indexing (the new copy ends up BEFORE the veryHidden tail
    ' rather than after it). Delete them so downstream code (MakeCaseCopy,
    ' create_level_worksheets, etc.) operates on a clean workbook.
    alerts_prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim junk_i As Long
    Dim junk_ws As Worksheet
    Dim junk_orig_name As String
    Dim junk_del_err As Long
    For junk_i = attempt_workbook.Worksheets.count To 1 Step -1
        Set junk_ws = attempt_workbook.Worksheets(junk_i)
        If junk_ws.Visible = xlSheetVeryHidden And Left$(junk_ws.Name, 1) = "_" Then
            ' Some add-in-injected sheets (Capital IQ, AlphaSense, etc.) carry
            ' sheet protection or COM-add-in hooks that make .Delete raise
            ' 1004. Treat junk-sheet removal as best-effort: try delete, then
            ' try unprotect+delete, then fall back to rename + veryHidden so
            ' the sheet can't shadow legitimate "_"-prefixed names downstream.
            junk_orig_name = junk_ws.Name
            On Error Resume Next
            Err.Clear
            junk_ws.Delete
            junk_del_err = Err.Number
            On Error GoTo ErrorHandler
            If junk_del_err <> 0 Then
                On Error Resume Next
                junk_ws.Unprotect
                Err.Clear
                junk_ws.Delete
                junk_del_err = Err.Number
                On Error GoTo ErrorHandler
            End If
            If junk_del_err <> 0 Then
                On Error Resume Next
                junk_ws.Name = "_x_junk_" & junk_i
                junk_ws.Visible = xlSheetVeryHidden
                On Error GoTo ErrorHandler
                Call LogError(junk_del_err, _
                    "Could not delete add-in tracking sheet '" & junk_orig_name & _
                    "' - renamed to '_x_junk_" & junk_i & "' and hidden instead.", _
                    "import_case")
            End If
        End If
    Next junk_i
    Application.DisplayAlerts = alerts_prev

    case_worksheet.Tab.Color = RGB(255, 0, 0)
    Exit Sub
    
ErrorHandler:
    ' Re-raise so setup_workbook's top-level Cleanup restores ScreenUpdating,
    ' Calculation and Events. Silent Resume Next here would leave Excel frozen.
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "import_case")
    Err.Raise e_num, "import_case", e_desc
    
    
End Sub



' === Classify Rows Subroutine ===
' Parse case workbook and classify which rows are questions, instructions etc...
Private Sub classify_rows()
    On Error GoTo ErrorHandler
    Dim i As Long, header_count As Long
    Dim colB As String, colDStr As String, colE As String, colD As Variant
    
    On Error GoTo ErrorHandler
    
    Set case_copy = MakeCaseCopy(attempt_workbook, case_worksheet)

    case_copy.Cells.UnMerge
    ' Clear any pre-existing content and conditional formatting in col A before
    ' we start writing row classifications into it. Some authors put text in
    ' col A of their Case sheet (sidebars, notes, column headers) which would
    ' otherwise collide with our classification labels and mislead the
    ' downstream pattern matching.
    case_copy.Columns("A").ClearContents
    On Error Resume Next
    case_copy.Columns("A").FormatConditions.Delete
    On Error GoTo ErrorHandler
    case_copy.Columns("A").ColumnWidth = 8.14
    
    last_row_case = GetLastUsedRow(case_copy)
    last_col_case = GetLastUsedCol(case_copy)

    ' Pre-flight: case files in competition often contain author formulas
    ' that reference user-defined lambdas not present in our template, so
    ' they surface as #NAME? / #VALUE! cells. Log a single notice to
    ' ErrorLog if any are present, so the operator can tell at a glance
    ' whether errored cells played a role in any later anomaly. The
    ' SanitizeRow call inside the per-row loop below ensures the
    ' classification logic itself can't be derailed by them.
    Dim err_count As Long
    err_count = CountErrorCells(case_copy)
    If err_count > 0 Then
        Call LogError(0, "Pre-flight: " & err_count & _
                      " error cell(s) found in case sheet; will be treated as empty.", _
                      "classify_rows")
    End If

    Dim rowValues As Variant
    For i = 1 To last_row_case
        rowValues = case_copy.Range(case_copy.Cells(i, 1), case_copy.Cells(i, last_col_case)).Value
        ' Convert any error variants to "" so the string/numeric ops below
        ' don't throw Err 13 (Type mismatch) on rows that contain #VALUE!,
        ' #NAME?, etc.
        Call SanitizeRow(rowValues)
        colB = LCase(CStr(rowValues(1, 2)))
        colD = rowValues(1, 4)
        colDStr = LCase(CStr(rowValues(1, 4)))
        colE = LCase(CStr(rowValues(1, 5)))

        Select Case True
            Case InStr(colB, "level code") > 0
                case_copy.Cells(i, 1).Value = "Level Code"
            Case Application.WorksheetFunction.CountA(Application.Index(rowValues, 1, 0)) = 0
                case_copy.Cells(i, 1).Value = "Spacing Row"
            Case IsHeaderRow(colDStr, colE)
                case_copy.Cells(i, 1).Value = "Header"
            Case IsNumeric(colD) And colD <> 0 And InStr(colB, "bonus") > 0
                case_copy.Cells(i, 1).Value = "Bonus Question"
            Case IsNumeric(colD) And colD = 0 And InStr(colB, "example") > 0
                case_copy.Cells(i, 1).Value = "Example Question"
            Case IsNumeric(rowValues(1, 2)) And rowValues(1, 2) <> 0 And IsNumeric(colD) And colD <> 0
                case_copy.Cells(i, 1).Value = "Level Question"
            'instruction is default classificaiton if none of the other criteria are met
            Case Else
                case_copy.Cells(i, 1).Value = "Instructions"
        End Select
    Next i

    ' --- Cross-validate each Header row ---
    ' A genuine header is always followed within a short window by a Bonus,
    ' Example, or Level Question row. A row that merely mentions "points" and
    ' "answer" in an instructions paragraph would be a false positive. Demote
    ' any Header that isn't backed up by at least one question row within the
    ' lookahead window; this prevents phantom levels.
    Const HEADER_LOOKAHEAD As Long = 15
    Dim j As Long, look_end As Long, has_followup As Boolean
    Dim nextClass As String
    For i = 1 To last_row_case
        If case_copy.Cells(i, 1).Value = "Header" Then
            has_followup = False
            look_end = i + HEADER_LOOKAHEAD
            If look_end > last_row_case Then look_end = last_row_case
            For j = i + 1 To look_end
                nextClass = CStr(case_copy.Cells(j, 1).Value)
                If nextClass = "Bonus Question" Or _
                   nextClass = "Level Question" Or _
                   nextClass = "Example Question" Then
                    has_followup = True
                    Exit For
                End If
                ' Another header in the window means no questions live between
                ' them -- this one is suspicious.
                If nextClass = "Header" Then Exit For
            Next j
            If Not has_followup Then
                case_copy.Cells(i, 1).Value = "Instructions"
            End If
        End If
    Next i


    'to assign level numbers, find all headers. If a bonus section exists, the first
    'header is the Bonus Header and the remaining headers are Level 1..N. If no bonus
    'section exists (e.g., Othello), every header is a Level header starting at Level 1.
    ReDim header_row_numbers(1 To last_row_case)
    header_count = 0
    For i = 1 To last_row_case
        If case_copy.Cells(i, 1).Value = "Header" Then
            header_count = header_count + 1
            header_row_numbers(header_count) = i
        End If
    Next i
    ReDim Preserve header_row_numbers(1 To header_count)

    ' Detect whether a bonus section exists by scanning for any row already
    ' classified as "Bonus Question" above.
    Dim has_bonus As Boolean
    has_bonus = False
    For i = 1 To last_row_case
        If case_copy.Cells(i, 1).Value = "Bonus Question" Then
            has_bonus = True
            Exit For
        End If
    Next i

    ' header_offset = how many of the leading headers are NOT level headers.
    ' With a bonus section the first header is the Bonus Header (offset = 1).
    ' Without a bonus section every header is a level header (offset = 0).
    Dim header_offset As Long
    If has_bonus Then
        header_offset = 1
    Else
        header_offset = 0
    End If
    number_of_levels = header_count - header_offset

    If has_bonus Then
        If header_count >= 1 Then case_copy.Cells(header_row_numbers(1), 1).Value = "Bonus Header"
        For i = 2 To header_count
            case_copy.Cells(header_row_numbers(i), 1).Value = "Level " & (i - 1) & " Header"
        Next i
    Else
        For i = 1 To header_count
            case_copy.Cells(header_row_numbers(i), 1).Value = "Level " & i & " Header"
        Next i
    End If

    Dim lvl As Long
    For i = 1 To last_row_case
        If case_copy.Cells(i, 1).Value = "Instructions" Then
            If has_bonus Then
                ' With bonus: Instructions before header(1) = General; between
                ' header(k-1) and header(k) (for k >= 2) = Level (k-1) Instructions.
                For lvl = header_count To 2 Step -1
                    If i > header_row_numbers(lvl - 1) And i < header_row_numbers(lvl) Then
                        case_copy.Cells(i, 1).Value = "Level " & (lvl - 1) & " Instructions"
                        Exit For
                    End If
                Next lvl
                If i < header_row_numbers(1) Then
                    case_copy.Cells(i, 1).Value = "General Instructions"
                End If
            Else
                ' No bonus: Instructions before header(1) = Level 1 Instructions;
                ' between header(k-1) and header(k) (for k >= 2) = Level k Instructions.
                For lvl = header_count To 2 Step -1
                    If i > header_row_numbers(lvl - 1) And i < header_row_numbers(lvl) Then
                        case_copy.Cells(i, 1).Value = "Level " & lvl & " Instructions"
                        Exit For
                    End If
                Next lvl
                If header_count >= 1 Then
                    If i < header_row_numbers(1) Then
                        case_copy.Cells(i, 1).Value = "Level 1 Instructions"
                    End If
                End If
            End If
        End If
    Next i

    Dim level_index As Long
    For i = 1 To last_row_case
        Select Case case_copy.Cells(i, 1).Value
            Case "Example Question", "Level Question"
                If has_bonus Then
                    ' With bonus: header(level_index) is the header for Level (level_index - 1).
                    For level_index = header_count To 2 Step -1
                        If i > header_row_numbers(level_index) Then
                            case_copy.Cells(i, 1).Value = "Level " & (level_index - 1) & " " & case_copy.Cells(i, 1).Value
                            Exit For
                        End If
                    Next level_index
                Else
                    ' No bonus: header(level_index) is the header for Level level_index.
                    For level_index = header_count To 1 Step -1
                        If i > header_row_numbers(level_index) Then
                            case_copy.Cells(i, 1).Value = "Level " & level_index & " " & case_copy.Cells(i, 1).Value
                            Exit For
                        End If
                    Next level_index
                End If
        End Select
    Next i
    
    
    
        ' === Find and store the row number for each level's Example Question ===
    ReDim example_row_numbers(1 To number_of_levels)
    Dim p1 As Long, p2 As Long, numTxt As String, lbl As String
    
    For i = 1 To last_row_case
        lbl = CStr(case_copy.Cells(i, 1).Value) ' Get the classification text
    
        ' Check if the row is an "Example Question"
        If InStr(1, lbl, " Example Question", vbTextCompare) > 0 Then
            ' Parse the level number from the classification text (e.g., "Level 1 Example Question")
            p1 = InStr(1, lbl, "Level ", vbTextCompare)
            p2 = InStr(p1, lbl, " Example Question", vbTextCompare)
            If p1 > 0 And p2 > p1 Then
                numTxt = Mid$(lbl, p1 + 6, p2 - (p1 + 6))
                If IsNumeric(numTxt) Then
                    lvl = CLng(numTxt)
                    If lvl >= 1 And lvl <= number_of_levels Then
                        ' Store the row number in our new array
                        example_row_numbers(lvl) = i
                    End If
                End If
            End If
        End If
    Next i

    
    ' === Calculate the last used column for each level's data range ===
    ReDim last_col_by_level(1 To number_of_levels)
    Dim level_num As Long
    Dim start_row As Long, end_row As Long
    Dim max_col As Long
    Dim r As Long, c As Long
    
    For level_num = 1 To number_of_levels
        ' Start scanning from this level's header. With a bonus section the
        ' first header is the Bonus Header, so Level N's header is at index
        ' N + 1; without a bonus section Level N's header is at index N.
        start_row = header_row_numbers(level_num + header_offset)

        ' Scan all the way down to just before the NEXT level's header
        ' (or to the very end of the sheet if it's the last level)
        If level_num < number_of_levels Then
            end_row = header_row_numbers(level_num + header_offset + 1) - 1
        Else
            end_row = last_row_case
        End If
    
        max_col = 0
        ' Loop through each row in the level's range
        For r = start_row To end_row
            ' Loop backwards from the overall last column to find the true last cell with data
            For c = last_col_case To 1 Step -1
                ' Check if the cell's value, AFTER trimming spaces, is not empty
                If Trim(CStr(case_copy.Cells(r, c).Value)) <> "" Then
                    ' If we found a cell with real data, check if its column is the new max
                    If c > max_col Then
                        max_col = c
                    End If
                    ' Break the inner loop (c) and move to the next row (r)
                    Exit For
                End If
            Next c
        Next r
        
        ' Store the result. Default to a safe value if nothing was found.
        If max_col > 0 Then
            last_col_by_level(level_num) = max_col
        Else
            last_col_by_level(level_num) = 7 ' Default to G if level is empty
        End If
    Next level_num
    
    last_col_all_levels = Application.WorksheetFunction.Max(last_col_by_level)
        
    Exit Sub
    
ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "classify_rows")
    Err.Raise e_num, "classify_rows", e_desc
    
    
End Sub

' === Create Level Worksheets Subroutine ===
Private Sub create_level_worksheets()
    On Error GoTo ErrorHandler

    ' Precondition: classify_rows must have established the level header
    ' map. If it didn't (e.g. it died on an error cell), bail cleanly with
    ' a single notice in ErrorLog rather than crashing with Err 9
    ' (Subscript out of range) on an unallocated header_row_numbers().
    If number_of_levels <= 0 Then
        Call LogError(0, _
            "Skipped: number_of_levels is 0; classify_rows did not populate level headers.", _
            "create_level_worksheets")
        Exit Sub
    End If

    ' ---- Declarations (grouped & explicit) ----
    Dim level_index As Long, i As Long, j As Long
    Dim dest_row As Long, header_row As Long
    Dim first_game_number As Long, last_game_number As Long
    Dim difficulty_header As Long, last_row_inst As Long
    Dim last_instr_nonblank As Long, rScan As Long, target_header_row As Long
    Dim endCol As Long
    Dim level_ws As Worksheet
    Dim header_range As Range, boxRange As Range
    Dim cf As FormatCondition
    Dim row_type As String
    Dim first_example_game_name As String
    Dim delete_end As Long
    Dim col_letter As String
    Dim xlookup_formula_part As String

    ' --- Level-header detection helpers (robust: do not rely on a fixed
    ' difficulty vocabulary like easy/medium/hard; authors sometimes use
    ' "possible", "impossible", "extreme", or leave it blank entirely.)
    Dim level_label As String       ' e.g. "LEVEL 1", "LEVEL 12"
    Dim b_text As String            ' trimmed/upper value from col B of the scan row
    Dim next_ch As String           ' character after level_label in b_text (used to avoid
                                    ' matching "LEVEL 10" when we want "LEVEL 1")

    ReDim yellow_cell_rows(1 To number_of_levels)

    For level_index = 1 To number_of_levels

        ' Create & name the level sheet
        Set level_ws = attempt_workbook.Sheets.Add(After:=attempt_workbook.Sheets(attempt_workbook.Sheets.count))
        level_ws.Name = "_L" & level_index
        If level_index = 1 Then
            level_ws.Tab.Color = RGB(255, 255, 0)
        Else
            level_ws.Tab.Color = RGB(0, 100, 0)
        End If

        ' Base layout
        With level_ws.Cells
            .UnMerge
            .RowHeight = 14.3
            .ColumnWidth = 5.23
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Font.Name = "Aptos Narrow"
            .Font.size = 11
        End With

        ' Paste "Level n Instructions"
        dest_row = 1
        For i = 1 To last_row_case
            row_type = CStr(case_copy.Cells(i, 1).Value)
            If row_type = "Level " & level_index & " Instructions" Then
                level_ws.Rows(dest_row).Value = case_copy.Rows(i).Value
                dest_row = dest_row + 1
            End If
        Next i

        ' Locate difficulty header inside pasted instructions.
        '
        ' We identify the header solely by column B containing "Level N" where N is
        ' the current level_index. Whatever text is in column C is the "difficulty"
        ' label ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ï¿½ we preserve it as-is and do NOT require any specific vocabulary.
        ' This tolerates "easy/medium/hard", "possible", "impossible", "extreme",
        ' a blank cell, or anything else an author chooses to write.
        '
        ' Matching rules for col B (case-insensitive, trimmed):
        '   1. Exact match: "LEVEL N"
        '   2. Starts-with match: "LEVEL N<non-digit>..." (e.g. "Level 1:", "Level 1 -")
        '      We reject "LEVEL 10" when searching for "LEVEL 1" by checking the next
        '      character is not a digit.
        last_row_inst = GetLastUsedRow(level_ws)
        difficulty_header = 0
        level_label = "LEVEL " & level_index
        For i = 1 To last_row_inst
            b_text = Trim$(UCase$(CStr(level_ws.Cells(i, 2).Value)))
            If b_text = level_label Then
                difficulty_header = i
                Exit For
            ElseIf Len(b_text) > Len(level_label) Then
                If Left$(b_text, Len(level_label)) = level_label Then
                    next_ch = Mid$(b_text, Len(level_label) + 1, 1)
                    If next_ch < "0" Or next_ch > "9" Then
                        difficulty_header = i
                        Exit For
                    End If
                End If
            End If
        Next i
        If difficulty_header = 0 Then difficulty_header = 2

        ' Keep exactly one row above the difficulty header
        delete_end = difficulty_header - 2
        If delete_end >= 1 Then level_ws.Rows("1:" & delete_end).Delete

        ' Freeze panes below row 3
        level_ws.Rows("1").RowHeight = 14.3
        level_ws.Rows("2").RowHeight = 24.4
        level_ws.Rows("3").RowHeight = 14.3
        level_ws.Range("A4").Select
        level_ws.Application.ActiveWindow.FreezePanes = True

        ' Normalize a single blank line before the header we will paste
        last_instr_nonblank = 0
        For rScan = dest_row - 1 To 1 Step -1
            If Application.WorksheetFunction.CountA(level_ws.Rows(rScan)) > 0 Then
                last_instr_nonblank = rScan
                Exit For
            End If
        Next rScan
        If last_instr_nonblank = 0 Then last_instr_nonblank = 1
        target_header_row = last_instr_nonblank + 2   ' one blank row after instructions

        If dest_row > target_header_row Then
            level_ws.Rows(target_header_row & ":" & (dest_row - 1)).Delete
            dest_row = target_header_row
        ElseIf dest_row < target_header_row Then
            dest_row = target_header_row
        End If

        ' What-If framing
        With level_ws.Range("I1:K3").Interior: .Color = RGB(34, 138, 184): End With
        With level_ws.Range("M1:O1").Interior: .Color = RGB(0, 123, 0): End With
        With level_ws.Range("M2:O3").Interior: .Color = RGB(218, 242, 208): End With
        With level_ws.Range("N1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Value = "What-If Data Table"
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
        End With

        ' Buttons/toggles
        With level_ws.Buttons.Add(217.875, 9.75, 92.25, 32.625): .OnAction = "done": .Text = "Done": End With
        With level_ws.Buttons.Add(410.625, 11.25, 95.25, 30.75): .OnAction = "copy_previous": .Text = "Copy Previous": End With
        With level_ws.OptionButtons.Add(586.125, 21.75, 58.875, 18.375): .OnAction = "what_if_on": .Text = "What If On": End With
        With level_ws.OptionButtons.Add(653.625, 22.5, 61.125, 18.375): .OnAction = "what_if_off": .Text = "What If Off": End With

        ' Copy "Level n Header" row
        For i = 1 To last_row_case
            If CStr(case_copy.Cells(i, 1).Value) = "Level " & level_index & " Header" Then
                level_ws.Rows(dest_row).Value = case_copy.Rows(i).Value
                dest_row = dest_row + 1
                Exit For
            End If
        Next i

        ' Explicitly format the header row (no Selection reliance)
        header_row = dest_row - 1
        Set header_range = level_ws.Range(level_ws.Cells(header_row, 1), level_ws.Cells(header_row, last_col_all_levels))
        With header_range.EntireRow
            .RowHeight = 42.3
            .WrapText = True
        End With
        With header_range
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlNone
            .FormatConditions.Delete
            Set cf = .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=" & .Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "<>""""")
            With cf
                .SetFirstPriority
                .Interior.Color = RGB(34, 138, 184)
                With .Font
                    .Bold = True
                    .Color = vbWhite
                End With
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = vbBlack
                End With
            End With
        End With

        level_ws.Cells.UnMerge

        ' Example anchor (game name + unified XLOOKUP)
        For i = 1 To last_row_case
            If CStr(case_copy.Cells(i, 1).Value) = "Level " & level_index & " Example Question" Then
                first_example_game_name = CStr(case_copy.Cells(i, 2).Value)
                level_ws.Cells(dest_row, 2).Value = first_example_game_name

        If Len(first_example_game_name) > 0 Then
                    col_letter = ColNumToLetter(last_col_all_levels)
                    xlookup_formula_part = "XLOOKUP(B" & dest_row & ",Case!B1:B" & last_row_case & ",Case!C1:" & col_letter & last_row_case & ","""")"
                    ' MAP/LAMBDA per-element pattern: ISBLANK over the XLOOKUP
                    ' result array misbehaves (silently nukes some cells - hit
                    ' on the Order-Logs column in April 2026 RTLV). MAP applies
                    ' the blank-check element-wise, which keeps strings intact
                    ' while still converting retrieved blanks (which XLOOKUP
                    ' surfaces as 0) into "" instead of zeros.
                    level_ws.Cells(dest_row, 3).Formula2 = "=LET(ans, " & xlookup_formula_part & ", MAP(ans, LAMBDA(item, IFERROR(IF(item="""", """", item), """"))))"
                End If

                Exit For
            End If
        Next i

        ' === Box exactly the header row and the XLOOKUP row (columns B .. last_col_all_levels) ===
        endCol = IIf(last_col_all_levels >= 2, last_col_all_levels, 2) ' at least column B
        Set boxRange = level_ws.Range(level_ws.Cells(header_row, 2), level_ws.Cells(dest_row, endCol))
        With boxRange.Borders
            .LineStyle = xlNone  ' clear existing
        End With
        With boxRange
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
        ' === End box ===

        ' Clear column A content (keep formats)
        level_ws.Columns("A").ClearContents

        ' First/last game numbers for the level
        For i = 1 To last_row_case
            If CStr(case_copy.Cells(i, 1).Value) = "Level " & level_index & " Level Question" Then
                first_game_number = CLng(case_copy.Cells(i, 2).Value)
                Exit For
            End If
        Next i
        For i = last_row_case To 1 Step -1
            If CStr(case_copy.Cells(i, 1).Value) = "Level " & level_index & " Level Question" Then
                last_game_number = CLng(case_copy.Cells(i, 2).Value)
                Exit For
            End If
        Next i

        ' Yellow anchor and game list
        yellow_cell_rows(level_index) = dest_row + 5
        level_ws.Cells(yellow_cell_rows(level_index), 2).Interior.Color = RGB(255, 255, 0)
        level_ws.Cells(yellow_cell_rows(level_index) + 1, 1).Value = first_game_number
        For i = first_game_number + 1 To last_game_number
            level_ws.Cells(yellow_cell_rows(level_index) + i - first_game_number + 1, 1).Value = i
        Next i

        ' Conditional formatting near "Done"
        With level_ws.Range("E1:G3").FormatConditions
            .Delete
            .Add Type:=xlExpression, _
                 Formula1:="=$B$" & yellow_cell_rows(level_index) & "=$E$" & (yellow_cell_rows(level_index) - 5)
            .item(.count).SetFirstPriority
            With .item(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 12093986
                .TintAndShade = 0
            End With
        End With
        
        ' Format B2 to match M1 (Dark Green, Bold White Text)
        With level_ws.Range("B2")
            .Interior.Color = RGB(0, 123, 0)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter ' Change to xlLeft if you prefer it uncentered
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
        End With
        
        ' Format C2 to match M2 (Light Green background, standard text)
        With level_ws.Range("C2")
            .Interior.Color = RGB(218, 242, 208)
            .Font.Bold = False
            .Font.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter ' Change to xlLeft if you prefer it uncentered
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
        End With
        
    
    Next level_index
    Exit Sub

ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "create_level_worksheets")
    Err.Raise e_num, "create_level_worksheets", e_desc
End Sub

' === Align Level Worksheets Subroutine ===
' Pads each L# so yellow anchors align, creates a solution start strip on each sheet,
' and stores the header-start address (max_yellow_row-6, last_col_all_levels+1) in case_copy!Z1.
Private Sub align_level_worksheets()
    On Error GoTo ErrorHandler

    ' Precondition: needs the level sheets created by create_level_worksheets.
    ' If number_of_levels was never set (classify_rows failed), there are no
    ' _L* sheets to align - skip cleanly instead of cascading to Err 5
    ' (Invalid procedure call) on uninitialized state.
    If number_of_levels <= 0 Then
        Call LogError(0, _
            "Skipped: number_of_levels is 0; no level worksheets to align.", _
            "align_level_worksheets")
        Exit Sub
    End If

    Dim i As Long
    Dim ws As Worksheet
    Dim max_yellow_row As Long
    Dim solRow As Long, solCol As Long
    Dim solution_start_cell As Range
    Dim solutionHeaderRange As Range
    Dim fc As FormatCondition
    Dim fAddr As String
    Dim lastColOnSheet As Long

    Dim headerStartRow As Long, headerStartCol As Long
    Dim headerStartAddr As String

    ' 1) Align yellow anchors
    max_yellow_row = Application.Max(yellow_cell_rows)
    ReDim offset_rows_needed(1 To number_of_levels)

    For i = 1 To number_of_levels
        offset_rows_needed(i) = max_yellow_row - yellow_cell_rows(i)
        If offset_rows_needed(i) > 0 Then
            Set ws = attempt_workbook.Sheets("_L" & i)
            With ws
                .Rows("1:" & offset_rows_needed(i)).Insert Shift:=xlDown
                .Rows("1:" & offset_rows_needed(i)).Hidden = True
            End With
        End If
    Next i

    ' 2) Compute & save the canonical "solution header start" address (always header row)
    headerStartRow = max_yellow_row - 6
    If headerStartRow < 1 Then headerStartRow = 1

    headerStartCol = last_col_all_levels + 1
    If headerStartCol < 1 Then headerStartCol = 1

    headerStartAddr = case_copy.Cells(headerStartRow, headerStartCol).Address(False, False)
    case_copy.Range("Z1").Value = headerStartAddr  ' e.g., "H36"

    ' 3) Create solution strip on each level sheet
    For i = 1 To number_of_levels
        Set ws = attempt_workbook.Sheets("_L" & i)

        ' Placement rule for the LIVE solution strip:
        '   - If lots of columns (>=21), put strip on yellow row, col G
        '   - Otherwise put strip in header row at col (last_col_all_levels+1)
        If last_col_all_levels >= 21 Then
            solRow = max_yellow_row
            solCol = 7                         ' column G
        Else
            solRow = headerStartRow
            solCol = headerStartCol
        End If
        If solCol < 1 Then solCol = 1

        lastColOnSheet = ws.Columns.count
        If solCol + 100 > lastColOnSheet Then
            Set solutionHeaderRange = ws.Range(ws.Cells(solRow, solCol), ws.Cells(solRow, lastColOnSheet))
        Else
            Set solutionHeaderRange = ws.Range(ws.Cells(solRow, solCol), ws.Cells(solRow, solCol + 100))
        End If
        Set solution_start_cell = ws.Cells(solRow, solCol)

        ' Row height + seed value
        ws.Rows(solRow).RowHeight = 42
        solution_start_cell.Value = " "

        ' Clear formats and reset style on the strip
        On Error Resume Next
        solutionHeaderRange.Style = "Normal"
        On Error GoTo 0
        solutionHeaderRange.ClearFormats
        solutionHeaderRange.FormatConditions.Delete

        ' Base formatting (Roboto 11, centered)
        With solutionHeaderRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            With .Font
                .Name = "Roboto"
                .size = 11
                .Bold = False
                .Color = vbBlack
            End With
        End With

        ' CF: non-empty cells -> #007B00 fill + bold white text
        fAddr = solutionHeaderRange.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Set fc = solutionHeaderRange.FormatConditions.Add(Type:=xlExpression, _
                   Formula1:="=" & fAddr & "<>""""")
        With fc
            .SetFirstPriority
            .StopIfTrue = False
            .Interior.Color = RGB(0, 123, 0)   ' #007B00
            With .Font
                .Bold = True
                .Color = vbWhite
            End With
        End With

        ' Select the first cell of the solution range on this sheet
        ws.Activate
        solution_start_cell.Select
    Next i

    Exit Sub

ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "align_level_worksheets")
    Err.Raise e_num, "align_level_worksheets", e_desc
End Sub



' === Create All Games Worksheet (AG) ===
' Produces AG as a copy of case_copy, keeps only "Level * Level Question" rows,
' removes the label column, and writes an answer formula in Column D.
'
' Row-span detection for each level ("first" and "last" row where the label in case_copy
' equals "Level X Level Question") is computed in VBA. A game-name ? level map is also
' built in VBA. The D-column formula for each AG row uses those constants to XLOOKUP
' from the Case sheet within the correct bounded range.
'
' Inputs:
'   - case_copy (Worksheet): annotated copy of Case; Column A contains labels;
'     Column B contains game names; row numbers correspond 1:1 with Case.
'   - case_worksheet (Worksheet): the original Case sheet.
'   - number_of_levels, last_row_case (Long): established earlier in the pipeline.
'
' Outputs:
'   - all_games_worksheet (Worksheet) named "AG".
'   - Column A on AG: game names.
'   - Column D on AG: answer formula returning "" when not found/blank.
'
' Assumptions:
'   - Labels for playable rows in case_copy Column A are in the form "Level X Level Question".
'   - Game names used for lookups are in Column B in both case_copy and Case.
'   - Answers to retrieve are in Case Column E.
Private Sub create_all_games_worksheet()
    On Error GoTo ErrorHandler

    ' Precondition: case_copy is set by classify_rows. If classify_rows
    ' died before assigning it, bail cleanly rather than crashing with
    ' Err 424 (Object required) when we try to use case_copy below.
    If case_copy Is Nothing Then
        Call LogError(0, _
            "Skipped: case_copy is Nothing; classify_rows did not complete.", _
            "create_all_games_worksheet")
        Exit Sub
    End If

    Dim wsAG As Worksheet
    Dim lastRowAG As Long, lastColAG As Long
    Dim rng As Range

    Dim firstRowByLevel() As Long
    Dim lastRowByLevel() As Long
    Dim gameToLevel As Object

    Dim i As Long, r As Long
    Dim lbl As String, lvl As Long, p1 As Long, p2 As Long, numTxt As String
    Dim gameName As String
    Dim topRow As Long, botRow As Long
    Dim f As String
    Dim caseName As String

    ' Delete prior AG if present
    On Error Resume Next
    Application.DisplayAlerts = False
    attempt_workbook.Worksheets("AG").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Copy case_copy to AG
    case_copy.Copy After:=attempt_workbook.Sheets(attempt_workbook.Sheets.count)
    Set all_games_worksheet = attempt_workbook.Sheets(attempt_workbook.Sheets.count)
    all_games_worksheet.Name = "AG"
    Set wsAG = all_games_worksheet

    ' Prepare a view of AG's used range
    lastRowAG = wsAG.Cells(wsAG.Rows.count, 1).End(xlUp).Row
    lastColAG = wsAG.Cells(1, wsAG.Columns.count).End(xlToLeft).Column
    Set rng = wsAG.Range(wsAG.Cells(1, 1), wsAG.Cells(lastRowAG, lastColAG))

    ' Keep only rows where Column A is "Level * Level Question"
    With rng
        .AutoFilter Field:=1, Criteria1:="<>*Level Question*"
        On Error Resume Next
        wsAG.Range(wsAG.Rows(1), wsAG.Rows(lastRowAG)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo ErrorHandler
        .AutoFilter
    End With

    ' Recompute last row after deletions
    lastRowAG = wsAG.Cells(wsAG.Rows.count, 1).End(xlUp).Row

    ' Remove label column so Column A holds the game name
    wsAG.Columns(1).Delete

    ' Remove pictures copied over with the sheet (if any)
    On Error Resume Next
    wsAG.Pictures.Delete
    On Error GoTo ErrorHandler

    ' -------------------------------------------------------------------------
    ' Build level ? [firstRow, lastRow] bounds and gameName ? level map (from case_copy)
    ' -------------------------------------------------------------------------
    ReDim firstRowByLevel(1 To number_of_levels)
    ReDim lastRowByLevel(1 To number_of_levels)
    Set gameToLevel = CreateObject("Scripting.Dictionary")

    For i = 1 To last_row_case
        lbl = CStr(case_copy.Cells(i, 1).Value)

        ' Identify "Level X Level Question" and parse X
        If InStr(1, lbl, "Level ", vbTextCompare) > 0 And _
           InStr(1, lbl, " Level Question", vbTextCompare) > 0 Then

            p1 = InStr(1, lbl, "Level ", vbTextCompare)                  ' start of "Level "
            p2 = InStr(p1 + 6, lbl, " Level Question", vbTextCompare)    ' start of " Level Question"
            If p1 > 0 And p2 > p1 Then
                numTxt = Mid$(lbl, p1 + 6, p2 - (p1 + 6))                ' numeric portion between markers
                If IsNumeric(numTxt) Then
                    lvl = CLng(numTxt)
                    If lvl >= 1 And lvl <= number_of_levels Then
                        If firstRowByLevel(lvl) = 0 Then firstRowByLevel(lvl) = i
                        lastRowByLevel(lvl) = i

                        ' Map game name in Column B to its level
                        gameName = CStr(case_copy.Cells(i, 2).Value)
                        If Len(gameName) > 0 Then
                            If Not gameToLevel.Exists(gameName) Then
                                gameToLevel.Add gameName, lvl
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' -------------------------------------------------------------------------
    ' Write the answer formula into AG Column D using precomputed row spans
    ' -------------------------------------------------------------------------
    caseName = case_worksheet.Name

    For r = 1 To lastRowAG
        gameName = CStr(wsAG.Cells(r, 1).Value) ' AG Column A = game name

        If Len(gameName) > 0 And gameToLevel.Exists(gameName) Then
            lvl = CLng(gameToLevel(gameName))
            topRow = firstRowByLevel(lvl)
            botRow = lastRowByLevel(lvl)

            If topRow > 0 And botRow >= topRow Then
                ' D(r): answer lookup limited to Case!B[topRow:botRow] ? Case!E[topRow:botRow]
                f = "=LET(ans," & _
                    "XLOOKUP(A" & r & "," & _
                    "'" & caseName & "'!B$" & topRow & ":B$" & botRow & "," & _
                    "'" & caseName & "'!E$" & topRow & ":E$" & botRow & ","""")," & _
                    "IF(ans="""","""",ans))"
                wsAG.Cells(r, 4).Formula2 = f
            Else
                wsAG.Cells(r, 4).Value = ""   ' No valid span for this level
            End If
        Else
            wsAG.Cells(r, 4).Value = ""       ' Game not mapped to a level
        End If
    Next r
    
    'clear column E which should just be checks and Xs
    all_games_worksheet.Columns("E").ClearContents
    
    ' Hide the working copy to keep the surface simple
    case_copy.Visible = xlSheetHidden
       
    ' Work within the used range of the sheet
    Set rng = wsAG.UsedRange
    
    For r = rng.Row To rng.Row + rng.Rows.count - 1
        With wsAG
            ' Check if column B has a value and column A is empty
            If Len(Trim(.Cells(r, 2).Value)) > 0 And Len(Trim(.Cells(r, 1).Value)) = 0 Then
                .Cells(r, 1).Value = r
            End If
        End With
    Next r
    
    Exit Sub
    
ErrorHandler:
    ' NOTE: this handler is inside create_all_games_worksheet; the previous
    ' label read "align_level_worksheets" by mistake.
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "create_all_games_worksheet")
    Err.Raise e_num, "create_all_games_worksheet", e_desc


End Sub

' Creates bi-directional internal hyperlinks between the "Level n" cell on the Case sheet
' (Column B) and the corresponding "Level n" cell on each level sheet Ln (Column B).
' For each level 1..number_of_levels, the first exact match of "Level n" in Column B
' is recorded. If not found on a given sheet, 9999 is recorded for that level.
' Existing hyperlinks in the target cells are replaced. The original text formatting
' (font name, size, bold, italic, underline, and color) is preserved.
Private Sub create_internal_links()
On Error GoTo ErrorHandler
    Dim lvl As Long
    Dim found As Range
    Dim wsLevel As Worksheet
    Dim caseCell As Range, levelCell As Range

    ' Precondition: needs level worksheets to link to. Log + bail
    ' rather than silently skip, so ErrorLog shows the operator that
    ' the link stage was reached but no work was possible.
    If number_of_levels <= 0 Then
        Call LogError(0, _
            "Skipped: number_of_levels is 0; no level worksheets to link.", _
            "create_internal_links")
        Exit Sub
    End If

    ' Size and initialize the link-row arrays
    ReDim case_link_rows(1 To number_of_levels)
    ReDim level_link_rows(1 To number_of_levels)

    For lvl = 1 To number_of_levels
        case_link_rows(lvl) = 9999
        level_link_rows(lvl) = 9999
    Next lvl

    ' --- locate "Level n" on Case sheet, Column B (exact text match)
    For lvl = 1 To number_of_levels
        With case_worksheet.Columns(2) ' Column B
            Set found = .Find(What:="Level " & lvl, _
                              After:=.Cells(.Rows.count), _
                              LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                              MatchCase:=True)
            If Not found Is Nothing Then
                case_link_rows(lvl) = found.Row
            End If
        End With
    Next lvl

    ' --- locate "Level n" on each L# sheet, Column B (exact text match)
    For lvl = 1 To number_of_levels
        Set wsLevel = Nothing
        On Error Resume Next
        Set wsLevel = attempt_workbook.Worksheets("_L" & lvl)
        On Error GoTo ErrorHandler

        If Not wsLevel Is Nothing Then
            With wsLevel.Columns(2) ' Column B
                Set found = .Find(What:="Level " & lvl, _
                                  After:=.Cells(.Rows.count), _
                                  LookIn:=xlValues, LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=True)
                If Not found Is Nothing Then
                    level_link_rows(lvl) = found.Row
                End If
            End With
        End If
    Next lvl

    ' --- create reciprocal links where both endpoints exist
    For lvl = 1 To number_of_levels
        If case_link_rows(lvl) <> 9999 And level_link_rows(lvl) <> 9999 Then
            Set wsLevel = attempt_workbook.Worksheets("_L" & lvl)
            Set caseCell = case_worksheet.Cells(case_link_rows(lvl), 2)   ' B[row] on Case
            Set levelCell = wsLevel.Cells(level_link_rows(lvl), 2)        ' B[row] on L#

            ' Case -> L#
            AddInternalLinkPreserveFont targetCell:=caseCell, _
                destSheetName:=wsLevel.Name, destCell:=levelCell

            ' L# -> Case
            AddInternalLinkPreserveFont targetCell:=levelCell, _
                destSheetName:=case_worksheet.Name, destCell:=caseCell
        End If
    Next lvl
    Exit Sub
    
ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "create_internal_links")
    Err.Raise e_num, "create_internal_links", e_desc
    
    
End Sub

' Adds an internal hyperlink to targetCell that navigates to destSheetName!destCell,
' replacing any existing hyperlink on targetCell. The cell's text formatting is preserved.
Private Sub AddInternalLinkPreserveFont(targetCell As Range, _
                                        destSheetName As String, _
                                        destCell As Range)
    Dim txt As String
    Dim fName As String, fSize As Double, fBold As Boolean, fItalic As Boolean
    Dim fUnderline As Long, fColor As Long
On Error GoTo ErrorHandler

    ' Capture current text and formatting
    txt = CStr(targetCell.Value)
    With targetCell.Font
        fName = .Name
        fSize = .size
        fBold = .Bold
        fItalic = .Italic
        fUnderline = .Underline
        fColor = .Color
    End With

    ' Replace any existing hyperlink
    If targetCell.Hyperlinks.count > 0 Then
        targetCell.Hyperlinks.Delete
    End If

    ' Create the hyperlink to the destination cell
    targetCell.Parent.Hyperlinks.Add anchor:=targetCell, Address:="", _
        SubAddress:="'" & destSheetName & "'!" & destCell.Address(False, False), _
        TextToDisplay:=txt

    ' Restore original formatting
    With targetCell.Font
        .Name = fName
        .size = fSize
        .Bold = fBold
        .Italic = fItalic
        .Underline = fUnderline
        .Color = fColor
    End With
    Exit Sub
    
ErrorHandler:
    Call LogError(Err.Number, Err.Description, "AddInternalLinkPreserveFont")
    Resume Next ' Continues execution on the line after the error
    
    
End Sub



' === Create Hints Worksheet Subroutine ===
Private Sub create_hints_sheet()
    On Error GoTo ErrorHandler
    
    Dim wsHints As Worksheet
    Dim wsAG As Worksheet
    Dim lastHintRow As Long
    Dim safe_last_row As Long
    
    ' Ensure we have a valid last row for the Case sheet lookup
    If last_row_case > 0 Then
        safe_last_row = last_row_case
    Else
        safe_last_row = 1000
    End If
    
    ' Delete old Hints sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    attempt_workbook.Worksheets("Hints").Delete
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = True
    
    ' Add new sheet at the end
    Set wsHints = attempt_workbook.Worksheets.Add(After:=attempt_workbook.Sheets(attempt_workbook.Sheets.count))
    wsHints.Name = "Hints"
    
    ' Match AG tab color
    On Error Resume Next
    Set wsAG = attempt_workbook.Worksheets("AG")
    If Not wsAG Is Nothing Then
        wsHints.Tab.Color = wsAG.Tab.Color
    End If
    On Error GoTo ErrorHandler
    
    ' Turn off gridlines
    wsHints.Activate
    ActiveWindow.DisplayGridlines = False
    
    ' Setup Headers
    wsHints.Range("A1:F1").Value = Array("Hint", "Game Number", "Lower Bound", "Upper Bound", "My Answer", "Correct?")
    
    ' Format Headers
    With wsHints.Range("A1:F1")
        .Interior.Color = RGB(34, 139, 34) ' Dark Green
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
    
    ' Set Column Widths
    wsHints.Columns("A").ColumnWidth = 70.5
    wsHints.Columns("B").ColumnWidth = 12
    wsHints.Columns("C").ColumnWidth = 12
    wsHints.Columns("D").ColumnWidth = 12
    wsHints.Columns("E").ColumnWidth = 12
    wsHints.Columns("F").ColumnWidth = 10
    
    ' Base alignment for all columns
    With wsHints.Columns("A:F")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    
    ' Insert Spilling FILTER formula in A2
    ' (Assumes Case information is in column C)
    wsHints.Range("A2").Formula2 = "=FILTER(Case!$C$1:$C$" & safe_last_row & ", ISNUMBER(SEARCH(""Hint"", Case!$C$1:$C$" & safe_last_row & ")), ""No hints found"")"
    
    ' Force calculation to allow the dynamic array to spill so we can count the rows
    Application.Calculate
    
    ' Find the last row of the spilled hints array
    lastHintRow = wsHints.Cells(wsHints.Rows.count, "A").End(xlUp).Row
    
    ' Only add and pull down formulas if valid hints were found
    If lastHintRow >= 2 And wsHints.Range("A2").Value <> "No hints found" Then
        
        ' Insert standard formulas in row 2
        wsHints.Range("B2").Formula2 = "=TEXTBETWEEN(A2, ""game #"", ""hint"")"
        wsHints.Range("C2").Formula2 = "=TEXTBETWEEN(A2, ""between"", ""and"")"
        wsHints.Range("D2").Formula2 = "=TEXTAFTER(A2, ""and"")*1"
        wsHints.Range("E2").Formula2 = "=LET(a, XLOOKUP(B2*1, Case!$B$1:$B$" & safe_last_row & ", Case!$E$1:$E$" & safe_last_row & ", """"), IF(a="""", """", a))"
        wsHints.Range("F2").Formula2 = "=IFS(COUNTA(B2:E2)<4, """", AND(E2*1>=C2*1, E2*1<=D2*1), 1, TRUE, -1)"
        
        ' Code-based equivalent of double-clicking the autofill handle
        If lastHintRow > 2 Then
            wsHints.Range("B2:F2").AutoFill Destination:=wsHints.Range("B2:F" & lastHintRow)
        End If
        
        ' Format Column F - Hide numbers and apply Icon Sets
        With wsHints.Range("F2:F" & lastHintRow)
            ' Make text white so you only see the icons
            .Font.Color = vbWhite
            
            ' Add Icon Set Conditional Formatting
            .FormatConditions.Delete
            Dim iconSetCF As IconSetCondition
            Set iconSetCF = .FormatConditions.AddIconSetCondition
            
            With iconSetCF
                .IconSet = ActiveWorkbook.IconSets(xl3Symbols)
                .ReverseOrder = False
                .ShowIconOnly = False ' Leaving false because font is explicitly white
                
                With .IconCriteria(2)
                    .Type = xlConditionValueNumber
                    .Value = 0
                    .Operator = xlGreaterEqual
                End With
                
                With .IconCriteria(3)
                    .Type = xlConditionValueNumber
                    .Value = 1
                    .Operator = xlGreaterEqual
                End With
            End With
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "create_hints_sheet")
    Err.Raise e_num, "create_hints_sheet", e_desc
End Sub


' === Global Error Handling Subroutine ===
' Writes error details to the "ErrorLog" worksheet.
Public Sub LogError(errNum As Long, errDesc As String, procedureName As String)
    Dim logSheet As Worksheet
    Dim nextRow As Long

    ' Use SheetExists for the lookup so we don't rely on exception-based
    ' control flow (and so the debugger's "Break on All Errors" setting
    ' can't trip on a missing sheet).
    If SheetExists(ThisWorkbook, "ErrorLog") Then
        Set logSheet = ThisWorkbook.Sheets("ErrorLog")
    End If

    ' From here on, guard with a real handler. The previous version left
    ' On Error Resume Next in effect for the entire body, which silently
    ' swallowed any structural failure (protected workbook, invalid cell
    ' write, etc.) and made logging look like it had succeeded.
    On Error GoTo ErrorHandler

    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        logSheet.Name = "ErrorLog"
        logSheet.Range("A1:D1").Value = Array("Timestamp", "Procedure", "Error Number", "Description")
        logSheet.Range("A1:D1").Font.Bold = True
    End If

    nextRow = logSheet.Cells(logSheet.Rows.count, "A").End(xlUp).Row + 1
    logSheet.Cells(nextRow, "A").Value = Now()
    logSheet.Cells(nextRow, "B").Value = procedureName
    logSheet.Cells(nextRow, "C").Value = errNum
    logSheet.Cells(nextRow, "D").Value = errDesc
    logSheet.Columns("A:D").AutoFit

    Exit Sub

ErrorHandler:
    ' Logging itself failed for some structural reason (workbook is closing,
    ' structure protected, etc.). Do NOT recurse into LogError - that risks
    ' an infinite loop / stack overflow. Surface to the Immediate window
    ' instead so a developer running interactively can still see what was
    ' lost. In production runs this is a silent best-effort degrade.
    Debug.Print "LogError failed: [" & Err.Number & "] " & Err.Description & _
                " | while logging: [" & errNum & "] " & errDesc & _
                " | from: " & procedureName
End Sub



' === Helper Functions ===
Function GetLastUsedRow(ws As Worksheet) As Long
    GetLastUsedRow = ws.Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function
Function GetLastUsedCol(ws As Worksheet) As Long
    GetLastUsedCol = ws.Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function
Function ColNumToLetter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ColNumToLetter = vArr(0)
End Function
' Public so utility_subs (and any other module) can share it. Two-arg
' signature is unambiguous about WHICH workbook is being searched,
' unlike a bare Sheets() call that silently binds to ActiveWorkbook.
Public Function SheetExists(wb As Workbook, SHEET_NAME As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_NAME)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' === SanitizeRow ===
' Replace any error variants in a row read from a Range.Value with "".
' Error cells cause CStr() and most numeric comparisons to throw Err 13
' (Type mismatch); coercing them to "" lets downstream classification
' logic treat them as plain blank cells, which is the right default.
' Accepts both the 2D Variant array shape (multi-column row read) and
' the scalar shape (single-cell read).
Public Sub SanitizeRow(ByRef rowValues As Variant)
    Dim col_idx As Long
    If IsArray(rowValues) Then
        For col_idx = LBound(rowValues, 2) To UBound(rowValues, 2)
            If IsError(rowValues(1, col_idx)) Then rowValues(1, col_idx) = ""
        Next col_idx
    Else
        If IsError(rowValues) Then rowValues = ""
    End If
End Sub

' === CountErrorCells ===
' Returns the count of cells in ws.UsedRange whose value is an Excel
' error (#VALUE!, #NAME?, #REF!, etc.). Used by classify_rows as a
' pre-flight check so the operator can see at-a-glance from ErrorLog
' whether the imported case sheet has error cells that the pipeline
' will treat as empty.
Public Function CountErrorCells(ws As Worksheet) As Long
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim count As Long

    arr = ws.UsedRange.Value
    If IsArray(arr) Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If IsError(arr(i, j)) Then count = count + 1
            Next j
        Next i
    Else
        If IsError(arr) Then count = 1
    End If
    CountErrorCells = count
End Function

' === IsHeaderRow ===
' Decide whether a row is a Level/Bonus header based on col D ("points") and
' col E ("answer"). Intentionally tolerant of author variations:
'   col D -- accepts "points", "point", "pts", "score" (plus "marks" for UK
'            authors).
'   col E -- accepts "answer", "answers", "response", "your answer" (anything
'            containing "answer" or "response"); also accepts a BLANK col E
'            when col D clearly indicates a scoring row, since a few authors
'            put the "answer" label one column further right.
' Both arguments are expected to be LCased already by the caller.
Private Function IsHeaderRow(colD As String, colE As String) As Boolean
    Dim hasPoints As Boolean, hasAnswer As Boolean
    hasPoints = (InStr(colD, "point") > 0) Or _
                (InStr(colD, "pts") > 0) Or _
                (InStr(colD, "score") > 0) Or _
                (InStr(colD, "marks") > 0)
    hasAnswer = (InStr(colE, "answer") > 0) Or _
                (InStr(colE, "response") > 0)
    IsHeaderRow = hasPoints And (hasAnswer Or Len(Trim$(colE)) = 0)
End Function

' === BackupExistingFile ===
' Copy the file at filePath to a timestamped backup under a "backups" sibling
' folder. Best-effort: any failure (missing dir rights, locked file, etc.)
' is silently ignored so setup isn't blocked by backup failures.
Private Sub BackupExistingFile(filePath As String)
    On Error GoTo failed
    If Len(Dir(filePath)) = 0 Then Exit Sub

    Dim sep As String
    sep = Application.PathSeparator

    Dim slashPos As Long
    slashPos = InStrRev(filePath, sep)
    If slashPos = 0 Then Exit Sub

    Dim folder As String, filename As String
    folder = Left$(filePath, slashPos - 1)
    filename = Mid$(filePath, slashPos + 1)

    Dim dotPos As Long, stem As String, ext As String
    dotPos = InStrRev(filename, ".")
    If dotPos > 0 Then
        stem = Left$(filename, dotPos - 1)
        ext = Mid$(filename, dotPos)
    Else
        stem = filename
        ext = ""
    End If

    Dim backupDir As String
    backupDir = folder & sep & "backups"
    If Len(Dir(backupDir, vbDirectory)) = 0 Then MkDir backupDir

    Dim backupPath As String
    backupPath = backupDir & sep & stem & "_backup_" & _
                 Format$(Now, "yyyymmdd_hhnnss") & ext
    FileCopy filePath, backupPath
    Exit Sub
failed:
    ' swallow: backups are best-effort
End Sub

Private Function MakeCaseCopy(attempt_workbook As Workbook, case_worksheet As Worksheet) As Worksheet
    Const TARGET_NAME As String = "case copy"
    Dim alerts_prev As Boolean
    Dim new_ws As Worksheet
    Dim k As Long

    If attempt_workbook.ProtectStructure Then
        Err.Raise vbObjectError + 513, , "Workbook structure is protected; cannot delete or rename sheets."
    End If

    If Not case_worksheet.Parent Is attempt_workbook Then
        Err.Raise vbObjectError + 514, , "'case_worksheet' is not in attempt_workbook."
    End If

    If SheetExists(attempt_workbook, TARGET_NAME) Then
        alerts_prev = Application.DisplayAlerts
        Application.DisplayAlerts = False
        attempt_workbook.Worksheets(TARGET_NAME).Delete
        Application.DisplayAlerts = alerts_prev
    End If

    ' --- Snapshot sheet names BEFORE the copy ---
    ' We identify the newly-copied sheet by name diff rather than by position.
    ' Position indexing is unreliable when the workbook contains veryHidden
    ' sheets (e.g. _avcts_version_hash from Capital IQ/AlphaSense add-ins):
    ' Copy After:= a veryHidden sheet does NOT place the new sheet at the
    ' tail of the Worksheets collection, it places it after the last VISIBLE
    ' sheet. So Worksheets(Worksheets.Count) still returns the veryHidden
    ' junk sheet, not the fresh copy.
    Dim pre_names As Object
    Set pre_names = CreateObject("Scripting.Dictionary")
    For k = 1 To attempt_workbook.Worksheets.count
        pre_names(attempt_workbook.Worksheets(k).Name) = True
    Next k

    ' --- Anchor the Copy on a VISIBLE sheet ---
    ' This avoids the Copy-After-veryHidden quirk described above. Any visible
    ' sheet works; we just need to guarantee the anchor isn't veryHidden.
    Dim anchor As Worksheet
    Set anchor = attempt_workbook.Worksheets(1)
    For k = 1 To attempt_workbook.Worksheets.count
        If attempt_workbook.Worksheets(k).Visible = xlSheetVisible Then
            Set anchor = attempt_workbook.Worksheets(k)
            Exit For
        End If
    Next k

    case_worksheet.Copy After:=anchor

    ' --- Find the new sheet by name diff ---
    Set new_ws = Nothing
    For k = 1 To attempt_workbook.Worksheets.count
        If Not pre_names.Exists(attempt_workbook.Worksheets(k).Name) Then
            Set new_ws = attempt_workbook.Worksheets(k)
            Exit For
        End If
    Next k

    If new_ws Is Nothing Then
        Err.Raise vbObjectError + 515, , _
            "MakeCaseCopy: could not locate the newly copied sheet. " & _
            "Is the source workbook protected or was the Copy blocked?"
    End If

    ' Defensive: ensure the new sheet is visible (veryHidden copies can happen
    ' when the source sheet was veryHidden; Worksheets.Copy preserves state).
    If new_ws.Visible <> xlSheetVisible Then new_ws.Visible = xlSheetVisible

    new_ws.Name = TARGET_NAME

    ' --- Sheet protection handling ---
    ' If the source sheet was password-protected, the copy inherited that
    ' protection and downstream UnMerge/ClearContents in classify_rows will
    ' fail. Try Unprotect with empty password first; if that fails, fall
    ' back to deleting the protected copy and reconstituting a fresh
    ' unprotected sheet by reading from the (still-protected) source --
    ' reads are always allowed on protected sheets.
    If new_ws.ProtectContents Then
        On Error Resume Next
        new_ws.Unprotect ""
        On Error GoTo 0

        If new_ws.ProtectContents Then
            ' Tier 3: reconstitute via fresh sheet
            Dim alerts_prev2 As Boolean
            alerts_prev2 = Application.DisplayAlerts
            Application.DisplayAlerts = False
            new_ws.Delete
            Application.DisplayAlerts = alerts_prev2

            ' Find a visible anchor (same approach as the copy anchor above)
            Dim anchor2 As Worksheet
            Set anchor2 = attempt_workbook.Worksheets(1)
            Dim k2 As Long
            For k2 = 1 To attempt_workbook.Worksheets.count
                If attempt_workbook.Worksheets(k2).Visible = xlSheetVisible Then
                    Set anchor2 = attempt_workbook.Worksheets(k2)
                    Exit For
                End If
            Next k2

            Set new_ws = attempt_workbook.Worksheets.Add(After:=anchor2)
            new_ws.Name = TARGET_NAME

            ' Copy data/formulas/formatting from the protected source.
            ' Reads are always allowed on protected sheets, so this works
            ' without the password. Loses sheet-level items (shapes, pictures,
            ' row heights, column widths) but preserves all cell-level content.
            case_worksheet.UsedRange.Copy Destination:=new_ws.Range("A1")
            Application.CutCopyMode = False
        End If
    End If

    Set MakeCaseCopy = new_ws
End Function

Public Sub zoom_all_worksheets()
On Error GoTo ErrorHandler
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
        ws.Select
        ActiveWindow.Zoom = 130
        End If
    Next ws
    Exit Sub
    
ErrorHandler:
    Dim e_num As Long, e_desc As String
    e_num = Err.Number: e_desc = Err.Description
    Call LogError(e_num, e_desc, "zoom_all_worksheets")
    Err.Raise e_num, "zoom_all_worksheets", e_desc
    
    
End Sub

Sub safe_setup()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set attempt_workbook = ThisWorkbook
    ThisWorkbook.Save
    
    Call import_case
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Call LogError(Err.Number, Err.Description, "safe_setup")
    Resume Cleanup

End Sub
