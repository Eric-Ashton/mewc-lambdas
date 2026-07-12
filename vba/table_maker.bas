Attribute VB_Name = "table_maker"
Option Explicit

'==============================================================================
' table_maker  -  scaffolds a "what-if data table" style sheet for a level's
'                 game data.
'
' Workflow:
'   1. On the case sheet, select the input array for a level (header row plus
'      the data rows underneath, leftmost column = game number).
'   2. Alt+F8 -> make_level_data_table (or bind to a button).
'   3. Sheet _T1 (or next free _Tn) is added to the workbook and populated:
'        A1..    header row from the input (values + formatting)
'        A2      first game number (as value)
'        B2      =XLOOKUP(A2, <game# column on case>, <full input on case>)
'                so it spills the matching row's data horizontally
'        B4      empty, yellow fill - drop your target formula here
'        A5..    every game number from the leftmost column (blanks skipped)
'   4. Fill B4 with your target formula (referring back to B2#), select
'      A4:B<last> (or wider), Data > What-If Analysis > Data Table, use A2
'      as the Column input cell.
'==============================================================================

Public Sub make_level_data_table()
    Dim case_sheet As Worksheet
    Dim table_sheet As Worksheet
    Dim input_range As Range
    Dim input_data_range As Range          ' input_range minus the header row
    Dim first_game As Variant
    Dim have_first As Boolean
    Dim src_val As Variant
    Dim dest_row As Long
    Dim r As Long
    Dim next_num As Long
    Dim sheet_prefix As String
    Dim lookup_addr As String
    Dim return_addr As String
    Dim prev_screen As Boolean
    Dim prev_calc As XlCalculation
    Dim prev_events As Boolean

    On Error GoTo CleanFail

    ' --- Validate the selection ---------------------------------------------
    If TypeName(Selection) <> "Range" Then
        MsgBox "make_level_data_table: please select the level's data range " & _
               "(header row plus one or more rows underneath) first.", _
               vbExclamation
        Exit Sub
    End If
    Set input_range = Selection
    If input_range.Cells.CountLarge < 2 Then
        MsgBox "make_level_data_table: only one cell is selected. Please " & _
               "select the array of data for the level (header row plus " & _
               "the rows underneath).", vbExclamation
        Exit Sub
    End If
    If input_range.Areas.count > 1 Then
        MsgBox "make_level_data_table: please select a single rectangular " & _
               "range, not multiple separate areas.", vbExclamation
        Exit Sub
    End If
    If input_range.Columns.count < 2 Then
        MsgBox "make_level_data_table: the selection needs at least two " & _
               "columns - a game-number column plus one or more data columns.", _
               vbExclamation
        Exit Sub
    End If

    ' Note where we started - the case_sheet is the sheet the selection is on.
    Set case_sheet = input_range.Worksheet

    ' --- Snapshot Application state so a mid-run failure can't strand it ----
    prev_screen = Application.ScreenUpdating
    prev_calc = Application.Calculation
    prev_events = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- Create the _Tn sheet in the same workbook as the case sheet --------
    next_num = 1
    Do While SheetExists(case_sheet.Parent, "_T" & next_num)
        next_num = next_num + 1
    Loop
    Set table_sheet = case_sheet.Parent.Sheets.Add( _
        After:=case_sheet.Parent.Sheets(case_sheet.Parent.Sheets.count))
    table_sheet.Name = "_T" & next_num

    ' --- Copy header row (values + formatting) into A1 of table_sheet -------
    ' Two-pass PasteSpecial (Values then Formats) so we get the visible values
    ' and cell formatting but not any formulas that might have been in the
    ' header row referring back to case_sheet cells.
    input_range.Rows(1).Copy
    table_sheet.Range("A1").PasteSpecial Paste:=xlPasteValues
    table_sheet.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' --- Write game numbers from leftmost column of input_range -------------
    ' Skip the top cell (header) and any truly blank cells. Keep non-numeric
    ' entries like "Example4a" verbatim - they're valid game identifiers for
    ' the XLOOKUP.
    Set input_data_range = input_range.Offset(1, 0).Resize( _
        input_range.Rows.count - 1, input_range.Columns.count)

    dest_row = 5
    have_first = False
    For r = 1 To input_data_range.Rows.count
        src_val = input_data_range.Cells(r, 1).Value
        If Not IsEmpty(src_val) Then
            ' Also treat cells that hold "" (via a formula or entered
            ' apostrophe) as blank for this purpose.
            If VarType(src_val) <> vbString Or Len(Trim$(CStr(src_val))) > 0 Then
                table_sheet.Cells(dest_row, 1).Value = src_val
                If Not have_first Then
                    first_game = src_val
                    have_first = True
                End If
                dest_row = dest_row + 1
            End If
        End If
    Next r

    ' --- A2: first game number as a value; user will vary this cell ---------
    If have_first Then
        table_sheet.Range("A2").Value = first_game
    End If

    ' --- B4: yellow fill; user drops target formula here --------------------
    table_sheet.Range("B4").Interior.Color = RGB(255, 255, 0)

    ' --- B2: XLOOKUP that spills the matching row's data --------------------
    ' Return columns 2..N of input_data_range (skip the game-number column
    ' itself, since A2 already holds it). That way the spilled row starting
    ' at B2 lines up column-for-column with the header row starting at B1.
    ' Sheet name wrapped in single quotes (with any embedded quotes doubled)
    ' so sheets containing spaces or special characters work.
    Dim return_data_range As Range
    Set return_data_range = input_data_range.Offset(0, 1).Resize( _
        input_data_range.Rows.count, input_data_range.Columns.count - 1)
    sheet_prefix = "'" & Replace(case_sheet.Name, "'", "''") & "'!"
    lookup_addr = input_data_range.Columns(1).Address( _
                      RowAbsolute:=True, ColumnAbsolute:=True)
    return_addr = return_data_range.Address( _
                      RowAbsolute:=True, ColumnAbsolute:=True)
    table_sheet.Range("B2").Formula2 = _
        "=XLOOKUP(A2," & sheet_prefix & lookup_addr & _
        "," & sheet_prefix & return_addr & ")"

    ' --- Sheet-wide formatting ----------------------------------------------
    ' Left-align every cell. Applying to Cells (the whole sheet) is a single
    ' property assignment, so it doesn't materialise 17 billion cells.
    table_sheet.Cells.HorizontalAlignment = xlLeft

    ' Set every column that carries data to a uniform 8.00-unit width. The
    ' used range spans A..(input_range.Columns.Count) - both the header row
    ' and the spilled B2 data end at that column.
    table_sheet.Range( _
        table_sheet.Columns(1), _
        table_sheet.Columns(input_range.Columns.count)).ColumnWidth = 8

    ' --- Land the user on the yellow cell -----------------------------------
    table_sheet.Activate
    table_sheet.Range("B4").Select

CleanExit:
    Application.CutCopyMode = False
    Application.ScreenUpdating = prev_screen
    Application.Calculation = prev_calc
    Application.EnableEvents = prev_events
    Exit Sub

CleanFail:
    MsgBox "make_level_data_table stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
