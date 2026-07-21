Attribute VB_Name = "utility_subs"
' deploy: shared
'===================================
' Sub: write_background_color
'===================================
' Replaces the contents of each cell in the selected range with a string representing its background color.
'
' - If the cell has a solid fill, the color is written as a hex RGB string (e.g., "#FF0000").
' - If the cell has a two-stop gradient, both RGB colors are written, separated by "/".
'
' SELECTION REQUIRED:
'   A rectangular range of cells whose background colors you want to extract.
'
' Output:
'   The original values in the cells will be overwritten with the color information.

Sub write_background_color()
    Dim selectedRange As Range
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation
    Dim priorEnableEvents As Boolean

    On Error GoTo CleanFail

    Set selectedRange = Selection

    ' Snapshot Application state before changing it so the cleanup path
    ' restores whatever was there - we don't want to silently flip the
    ' workbook back to xlCalculationAutomatic if it was on Manual.
    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation
    priorEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ReDim arr(1 To selectedRange.Rows.count, 1 To selectedRange.Columns.count)

    ' Delegate color extraction to get_color() so no-fill and
    ' black-filled cells are distinguished. Output formats:
    '   no fill        -> "#FFFFFF"
    '   solid black    -> "#000000"
    '   solid <other>  -> "#RRGGBB"
    '   gradient (N)   -> "#RRGGBB/#RRGGBB/..." for all N stops
    For i = 1 To selectedRange.Rows.count
        For j = 1 To selectedRange.Columns.count
            arr(i, j) = get_color(selectedRange.Cells(i, j))
        Next j
    Next i

    selectedRange.Value = arr

CleanExit:
    Application.ScreenUpdating = priorScreenUpdating
    Application.Calculation = priorCalculation
    Application.EnableEvents = priorEnableEvents
    Exit Sub

CleanFail:
    MsgBox "write_background_color stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


'===================================
' Sub: frequency_table
'===================================
' Generates a new sheet showing the frequency of each unique value in
' the selected range. Comparison is case-sensitive ("A" and "a" are
' distinct rows).
'
' Output sheet layout:
'   Column A: Unique Values
'   Column B: Counts
'   AutoFilter is enabled. Rows are sorted by Counts descending.
'   The row(s) with the max count are highlighted blue, with the min
'   count orange (white text in both cases).
'
' SELECTION REQUIRED:
'   Any range of cells containing the values to analyze.
'
' Output:
'   A new worksheet named "Freqs1", "Freqs2", etc.

Sub frequency_table()
    Dim selectedRange As Range
    Dim cell As Range
    Dim dataDict As Object
    Dim nextSheetNumber As Integer
    Dim newSheet As Worksheet
    Dim uniqueKey As Variant
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation
    Dim priorEnableEvents As Boolean

    On Error GoTo CleanFail

    ' Intersect the selection with UsedRange so a Ctrl-A collapses to
    ' the data region instead of walking ~17 billion cells.
    Set selectedRange = Intersect(Selection, ActiveSheet.UsedRange)
    If selectedRange Is Nothing Then
        MsgBox "frequency_table: the selection contains no used cells.", vbExclamation
        Exit Sub
    End If

    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation
    priorEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set dataDict = CreateObject("Scripting.Dictionary")
    dataDict.CompareMode = 0    ' vbBinaryCompare - case-sensitive

    ' Tally each non-blank cell value.
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            If dataDict.Exists(cell.Value) Then
                dataDict(cell.Value) = dataDict(cell.Value) + 1
            Else
                dataDict.Add cell.Value, 1
            End If
        End If
    Next cell

    ' Allocate the output sheet.
    nextSheetNumber = 1
    Do While SheetExists(ActiveWorkbook, "Freqs" & nextSheetNumber)
        nextSheetNumber = nextSheetNumber + 1
    Loop
    Set newSheet = ActiveWorkbook.Sheets.Add
    newSheet.Name = "Freqs" & nextSheetNumber

    ' Headers.
    newSheet.Cells(1, 1).Value = "Unique Values"
    newSheet.Cells(1, 2).Value = "Counts"

    ' Write rows.
    rowIndex = 2
    For Each uniqueKey In dataDict.Keys
        newSheet.Cells(rowIndex, 1).Value = uniqueKey
        newSheet.Cells(rowIndex, 2).Value = dataDict(uniqueKey)
        rowIndex = rowIndex + 1
    Next uniqueKey
    lastRow = rowIndex - 1

    ' Sort by Counts descending, then enable AutoFilter, then color.
    If lastRow >= 2 Then
        newSheet.Range("A1:B" & lastRow).Sort _
            Key1:=newSheet.Range("B1"), Order1:=xlDescending, Header:=xlYes
        newSheet.Range("A1:B" & lastRow).AutoFilter
        Call apply_freq_colors(newSheet, lastRow)
    End If

    newSheet.Columns("A:B").AutoFit

CleanExit:
    Application.ScreenUpdating = priorScreenUpdating
    Application.Calculation = priorCalculation
    Application.EnableEvents = priorEnableEvents
    Exit Sub

CleanFail:
    MsgBox "frequency_table stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


'===================================
' Sub: frequency_table_by_char
'===================================
' Generates a new sheet showing the frequency of each unique CHARACTER
' across every cell in the selection. Every character is counted,
' including spaces, tabs, and any other whitespace. Comparison is
' case-sensitive ("A" and "a" are distinct rows).
'
' Output sheet layout:
'   Column A: Unique Values (the character, formatted as text)
'   Column B: Counts
'   Column C: Unicode (the UTF-16 code unit, via AscW)
'   AutoFilter is enabled. Rows are sorted by Counts descending.
'   Same blue (max) / orange (min) coloring as frequency_table,
'   applied to columns A and B (not C).
'
' SELECTION REQUIRED:
'   Any range of cells whose characters you want to tally.
'
' Output:
'   A new worksheet named "FreqsChar1", "FreqsChar2", etc.

Sub frequency_table_by_char()
    Dim selectedRange As Range
    Dim cell As Range
    Dim dataDict As Object
    Dim nextSheetNumber As Integer
    Dim newSheet As Worksheet
    Dim uniqueKey As Variant
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim txt As String
    Dim k As Long
    Dim ch As String
    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation
    Dim priorEnableEvents As Boolean

    On Error GoTo CleanFail

    Set selectedRange = Intersect(Selection, ActiveSheet.UsedRange)
    If selectedRange Is Nothing Then
        MsgBox "frequency_table_by_char: the selection contains no used cells.", vbExclamation
        Exit Sub
    End If

    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation
    priorEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set dataDict = CreateObject("Scripting.Dictionary")
    dataDict.CompareMode = 0    ' vbBinaryCompare - case-sensitive

    ' Walk each cell, then each character. Mid$ returns one UTF-16
    ' code unit at a time; whitespace and any other character is
    ' counted as its own key.
    For Each cell In selectedRange
        If Not IsEmpty(cell.Value) Then
            txt = CStr(cell.Value)
            For k = 1 To Len(txt)
                ch = Mid$(txt, k, 1)
                If dataDict.Exists(ch) Then
                    dataDict(ch) = dataDict(ch) + 1
                Else
                    dataDict.Add ch, 1
                End If
            Next k
        End If
    Next cell

    ' Allocate the output sheet.
    nextSheetNumber = 1
    Do While SheetExists(ActiveWorkbook, "FreqsChar" & nextSheetNumber)
        nextSheetNumber = nextSheetNumber + 1
    Loop
    Set newSheet = ActiveWorkbook.Sheets.Add
    newSheet.Name = "FreqsChar" & nextSheetNumber

    ' Headers.
    newSheet.Cells(1, 1).Value = "Unique Values"
    newSheet.Cells(1, 2).Value = "Counts"
    newSheet.Cells(1, 3).Value = "Unicode"

    ' Force column A to Text format BEFORE writing - prevents Excel
    ' from interpreting a leading "=" as a formula or coercing numeric-
    ' looking characters into numbers.
    newSheet.Columns("A").NumberFormat = "@"

    ' Write rows.
    rowIndex = 2
    For Each uniqueKey In dataDict.Keys
        newSheet.Cells(rowIndex, 1).Value = uniqueKey
        newSheet.Cells(rowIndex, 2).Value = dataDict(uniqueKey)
        ' Mask AscW into the unsigned 16-bit range so BMP code points
        ' above 32767 don't come back as negative numbers on hosts
        ' where AscW is signed Integer.
        newSheet.Cells(rowIndex, 3).Value = AscW(uniqueKey) And &HFFFF&
        rowIndex = rowIndex + 1
    Next uniqueKey
    lastRow = rowIndex - 1

    ' Sort by Counts descending, enable AutoFilter on A:C, color A:B.
    If lastRow >= 2 Then
        newSheet.Range("A1:C" & lastRow).Sort _
            Key1:=newSheet.Range("B1"), Order1:=xlDescending, Header:=xlYes
        newSheet.Range("A1:C" & lastRow).AutoFilter
        Call apply_freq_colors(newSheet, lastRow)
    End If

    newSheet.Columns("A:C").AutoFit

CleanExit:
    Application.ScreenUpdating = priorScreenUpdating
    Application.Calculation = priorCalculation
    Application.EnableEvents = priorEnableEvents
    Exit Sub

CleanFail:
    MsgBox "frequency_table_by_char stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


' Private helper shared by frequency_table and frequency_table_by_char.
' Colors columns A and B of ws:
'   - Rows whose count (column B) equals the maximum get blue / white text.
'   - Rows whose count equals the minimum get orange / white text.
' When max = min (only one distinct count), every row gets blue.
Private Sub apply_freq_colors(ws As Worksheet, lastRow As Long)
    Dim r As Long
    Dim currentVal As Long, maxVal As Long, minVal As Long

    If lastRow < 2 Then Exit Sub

    maxVal = ws.Cells(2, 2).Value
    minVal = ws.Cells(2, 2).Value
    For r = 3 To lastRow
        currentVal = ws.Cells(r, 2).Value
        If currentVal > maxVal Then maxVal = currentVal
        If currentVal < minVal Then minVal = currentVal
    Next r

    For r = 2 To lastRow
        currentVal = ws.Cells(r, 2).Value
        If currentVal = maxVal Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(0, 0, 255)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Color = RGB(255, 255, 255)
        ElseIf currentVal = minVal Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 165, 0)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Color = RGB(255, 255, 255)
        End If
    Next r
End Sub

'===================================
' Sub: tall_board
'===================================
' Creates a new sheet listing each cell in the selected range, including:
'   - The cell's address
'   - Its value
'   - Its background color in hex RGB format
'   - The font color of the first character, when available
'   - The font name of the first character, when available
'
' SELECTION REQUIRED:
'   A single rectangular range of cells to document.
'
' Output:
'   A new sheet named "tall_board", "tall_board1", etc., with flattened cell info.

Sub tall_board()
    Dim selectedRange As Range
    Dim newSheet As Worksheet
    Dim cell As Range
    Dim outputData() As Variant
    Dim rowIndex As Long
    Dim itemCount As Long
    Dim sheetName As String
    Dim suffix As Long
    Dim priorScreenUpdating As Boolean
    Dim priorEnableEvents As Boolean
    Dim priorCalculation As XlCalculation
    Dim fontColorHex As String
    Dim fontName As String
    
    On Error GoTo CleanFail
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a rectangular range of cells first.", vbExclamation
        Exit Sub
    End If
    
    Set selectedRange = Selection
    
    If selectedRange.Areas.count > 1 Then
        MsgBox "Please select a single rectangular range, not multiple separate areas.", vbExclamation
        Exit Sub
    End If
    
    priorScreenUpdating = Application.ScreenUpdating
    priorEnableEvents = Application.EnableEvents
    priorCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    sheetName = "tall_board"
    suffix = 0
    
    Do While SheetExists(ActiveWorkbook, sheetName & IIf(suffix = 0, "", suffix))
        suffix = suffix + 1
    Loop
    
    sheetName = sheetName & IIf(suffix = 0, "", suffix)
    
    Set newSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
    newSheet.Name = sheetName
    
    With newSheet
        .Cells(1, 1).Value = "Address"
        .Cells(1, 2).Value = "Value"
        .Cells(1, 3).Value = "Background Color (RGB)"
        .Cells(1, 4).Value = "Font Color (RGB, 1st char)"
        .Cells(1, 5).Value = "Font Name (1st char)"
    End With
    
    itemCount = selectedRange.Cells.CountLarge
    ReDim outputData(1 To itemCount, 1 To 5)
    
    rowIndex = 1
    
    For Each cell In selectedRange.Cells
        GetFirstCharacterFont cell, fontColorHex, fontName
        
        outputData(rowIndex, 1) = cell.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False)
        
        If IsError(cell.Value2) Then
            outputData(rowIndex, 2) = cell.Text
        ElseIf IsEmpty(cell.Value2) Then
            outputData(rowIndex, 2) = vbNullString
        Else
            outputData(rowIndex, 2) = cell.Value2
        End If
        
        ' get_color() reports no-fill as "#FFFFFF" and solid-black as
        ' "#000000" unambiguously by inspecting .Pattern before reading
        ' .Color. Raw Interior.Color can confuse those two states.
        outputData(rowIndex, 3) = get_color(cell)
        outputData(rowIndex, 4) = fontColorHex
        outputData(rowIndex, 5) = fontName
        
        rowIndex = rowIndex + 1
    Next cell
    
    newSheet.Range("A2").Resize(itemCount, 5).Value = outputData
    newSheet.Columns("A:E").AutoFit

CleanExit:
    Application.ScreenUpdating = priorScreenUpdating
    Application.EnableEvents = priorEnableEvents
    Application.Calculation = priorCalculation
    Exit Sub

CleanFail:
    MsgBox "tall_board stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub GetFirstCharacterFont(ByVal cell As Range, ByRef fontColorHex As String, ByRef fontName As String)
    Dim cellText As String
    Dim fontColor As Long
    
    fontColorHex = vbNullString
    fontName = vbNullString
    
    If IsError(cell.Value2) Then Exit Sub
    
    cellText = CStr(cell.Value2)
    If Len(cellText) = 0 Then Exit Sub
    
    On Error Resume Next
    fontColor = cell.Characters(1, 1).Font.Color
    fontName = cell.Characters(1, 1).Font.Name
    
    If Err.Number <> 0 Then
        Err.Clear
        fontColor = cell.Font.Color
        fontName = cell.Font.Name
    End If
    On Error GoTo 0
    
    fontColorHex = ColorLongToHex(fontColor)
End Sub

Private Function ColorLongToHex(ByVal colorValue As Long) As String
    ColorLongToHex = "#" & _
        Right$("0" & Hex$(colorValue Mod 256), 2) & _
        Right$("0" & Hex$((colorValue \ 256) Mod 256), 2) & _
        Right$("0" & Hex$((colorValue \ 65536) Mod 256), 2)
End Function

'===================================
' subs to select first and last sheet
'===================================
Sub select_first_sheet()
    Worksheets(1).Select
End Sub

Sub select_last_sheet()
    Worksheets(Worksheets.count).Select
End Sub

'===================================
' Sub: fill_color
'===================================
' Fills all blank cells ("" or empty) in the used range of the active sheet
' with the value of each selected (non-blank) cell, *but only for cells that share its background color*.
'
' SELECTION REQUIRED:
'   You may select one or more colored cells with values to propagate.
'   Each source cell must have a unique background color.
'
' Output:
'   Modifies only those blank cells within the UsedRange that match the background color of a selected cell.

Sub fill_color()
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim data As Variant, colors() As Variant
    Dim i As Long, j As Long
    Dim srcCell As Range
    Dim colorMap As Object
    Dim key As String

    ' Optimize performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set ws = ActiveSheet
    Set usedRng = ws.UsedRange
    Set colorMap = CreateObject("Scripting.Dictionary")

    ' Build color -> value map from selected cells.
    ' Use get_color() rather than raw Interior.Color so that gradient
    ' cells produce a distinct key per gradient (e.g. "#FF0909/#4774C5")
    ' rather than all collapsing to "0". Solid-black cells are also
    ' distinguished from gradients and from unfilled cells.
    '
    ' We REJECT "#FFFFFF" as a source key. get_color() returns "#FFFFFF"
    ' for both no-fill cells and explicitly white-filled cells, and the
    ' default blank cell on every worksheet is also no-fill. If a no-fill
    ' cell with a value sneaks into the selection, treating it as a
    ' source would cause us to write its value into ~every blank cell in
    ' UsedRange. That's never the user's intent and would be a painful
    ' undo if it slipped through.
    For Each srcCell In Selection
        If Not IsEmpty(srcCell.Value) Then
            key = get_color(srcCell)
            If key <> "#FFFFFF" Then
                If Not colorMap.Exists(key) Then
                    colorMap.Add key, srcCell.Value
                End If
            End If
        End If
    Next srcCell

    ' Nothing usable in the selection - bail out before scanning UsedRange.
    If colorMap.count = 0 Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        MsgBox "fill_color: no usable source colors in selection." & vbCrLf & _
               "(No-fill and white-fill cells are ignored - they would " & _
               "match every blank cell on the sheet.)", vbExclamation
        Exit Sub
    End If

    ' Load values and background colors (as get_color signatures)
    data = usedRng.Value
    ReDim colors(1 To usedRng.Rows.count, 1 To usedRng.Columns.count)
    For i = 1 To usedRng.Rows.count
        For j = 1 To usedRng.Columns.count
            colors(i, j) = get_color(usedRng.Cells(i, j))
        Next j
    Next i

    ' Write each matched cell individually so untouched cells in
    ' usedRng are preserved exactly - formulas stay as formulas, data
    ' validation stays, literals stay byte-for-byte the same. Skip
    ' cells that contain a formula (even one that evaluates to "")
    ' so we don't replace a deliberate formula with a fill value.
    For i = 1 To UBound(data, 1)
        For j = 1 To UBound(data, 2)
            If Trim(CStr(data(i, j))) = "" Then
                If Not usedRng.Cells(i, j).HasFormula Then
                    key = CStr(colors(i, j))
                    If colorMap.Exists(key) Then
                        usedRng.Cells(i, j).Value = colorMap(key)
                    End If
                End If
            End If
        Next j
    Next i

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


'===================================
' Sub: flood_fill
'===================================
' Performs a flood-fill operation starting from each non-empty cell in the selected range.
' Fills all connected blank cells with the starting cell's value and fill color.
'
' Flooding continues outward (up/down/left/right) but *stops* when encountering:
'   - A visible border between cells (used as a wall/barrier)
'   - A background color mismatch between adjacent cells
'   - A non-blank cell
'
' SELECTION REQUIRED:
'   Select the range from which flood fill should start.
'   Only non-empty cells within this range will be used as flood-fill seeds.
'
' Output:
'   Updates values and fill colors of connected blank cells that match the region.
'   Ignores cells outside the selected range.
'
' Notes:
'   - Flooding is constrained to contiguous regions that do not cross walls or color zones.

Sub flood_fill()
    Dim ws As Worksheet
    Dim visited As Object
    Dim cell As Range, rng As Range, nbr As Range
    Dim fillValue As Variant
    Dim fillColorKey As String, nbrColorKey As String
    Dim key As String
    Dim i As Long
    Dim hasBorder As Boolean
    Dim rOffset As Long, cOffset As Long
    Dim nbrRow As Long, nbrCol As Long
    Dim startRange As Range

    ' Bounds of the (single-area) start range, precomputed so the inner
    ' loop doesn't have to call Intersect for every neighbor.
    Dim startTop As Long, startLeft As Long, startBot As Long, startRight As Long

    ' Array-based LIFO stack of (row, col) Long pairs. O(1) push/pop,
    ' and storing primitives instead of Range objects avoids pinning
    ' COM references for every queued cell.
    Dim stackR() As Long, stackC() As Long
    Dim stackSize As Long, stackTop As Long

    ' Row/col offsets for L, U, R, D - row delta in dRow, col delta in dCol.
    Dim dRow As Variant, dCol As Variant
    dRow = Array(0, -1, 0, 1)
    dCol = Array(-1, 0, 1, 0)

    Dim priorScreenUpdating As Boolean
    Dim priorCalculation As XlCalculation
    Dim priorEnableEvents As Boolean

    On Error GoTo CleanFail

    Set ws = ActiveSheet
    Set startRange = Selection

    ' Restrict to a single rectangular area. The bounds-check optimization
    ' below assumes one contiguous rectangle; multi-area selections would
    ' need a per-area loop. Easier to ask the user to re-select than to
    ' over-engineer this.
    If startRange.Areas.count > 1 Then
        MsgBox "flood_fill: please select a single rectangular range, " & _
               "not multiple separate areas.", vbExclamation
        Exit Sub
    End If

    priorScreenUpdating = Application.ScreenUpdating
    priorCalculation = Application.Calculation
    priorEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set visited = CreateObject("Scripting.Dictionary")

    startTop = startRange.Row
    startLeft = startRange.Column
    startBot = startTop + startRange.Rows.count - 1
    startRight = startLeft + startRange.Columns.count - 1

    ' Initial stack capacity. Grows by doubling on push when full.
    stackSize = 256
    ReDim stackR(1 To stackSize)
    ReDim stackC(1 To stackSize)
    stackTop = 0

    ' Walk each potential seed in the start range.
    For Each cell In startRange.Cells
        key = cell.Address
        If Not visited.Exists(key) And Not IsEmpty(cell.Value) Then
            fillValue = cell.Value
            ' Color comparison via get_color() string keys so a no-fill
            ' seed and an explicit-color seed can't compare equal by
            ' accident (raw Interior.Color comparison was the latent
            ' bug the user flagged elsewhere in this module).
            fillColorKey = get_color(cell)

            ' Push the seed
            stackTop = 1
            stackR(stackTop) = cell.Row
            stackC(stackTop) = cell.Column

            Do While stackTop > 0
                ' Pop (LIFO)
                rOffset = stackR(stackTop)
                cOffset = stackC(stackTop)
                stackTop = stackTop - 1

                Set rng = ws.Cells(rOffset, cOffset)
                key = rng.Address
                If visited.Exists(key) Then GoTo NextPop
                visited.Add key, True

                If IsEmpty(rng.Value) Then rng.Value = fillValue
                ' Paint via the seed's actual numeric color so empties
                ' end up uniformly colored even when the seed was a
                ' gradient (paint with the gradient's primary color).
                If rng.Interior.Color <> cell.Interior.Color Then _
                    rng.Interior.Color = cell.Interior.Color

                ' Visit four neighbors
                For i = 0 To 3
                    nbrRow = rOffset + dRow(i)
                    nbrCol = cOffset + dCol(i)

                    ' Inline bounds check against startRange - avoids
                    ' allocating a Range object per neighbor.
                    If nbrRow >= startTop And nbrRow <= startBot _
                       And nbrCol >= startLeft And nbrCol <= startRight Then

                        Set nbr = ws.Cells(nbrRow, nbrCol)
                        key = nbr.Address
                        If Not visited.Exists(key) Then
                            ' Stop if not blank (formulas count as not-blank
                            ' since IsEmpty is False for any formula cell).
                            If IsEmpty(nbr.Value) Then

                                ' Border check between rng and nbr
                                Select Case i
                                    Case 0 ' Left
                                        hasBorder = (rng.Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone) Or _
                                                    (nbr.Borders(xlEdgeRight).LineStyle <> xlLineStyleNone)
                                    Case 1 ' Up
                                        hasBorder = (rng.Borders(xlEdgeTop).LineStyle <> xlLineStyleNone) Or _
                                                    (nbr.Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone)
                                    Case 2 ' Right
                                        hasBorder = (rng.Borders(xlEdgeRight).LineStyle <> xlLineStyleNone) Or _
                                                    (nbr.Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone)
                                    Case 3 ' Down
                                        hasBorder = (rng.Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone) Or _
                                                    (nbr.Borders(xlEdgeTop).LineStyle <> xlLineStyleNone)
                                End Select

                                If Not hasBorder Then
                                    ' Color match against the seed (via get_color
                                    ' strings, not raw .Color, so no-fill vs.
                                    ' explicit colors can't collide).
                                    nbrColorKey = get_color(nbr)
                                    If nbrColorKey = fillColorKey Then
                                        ' Push neighbor onto stack, growing
                                        ' the array if needed (amortized O(1)).
                                        stackTop = stackTop + 1
                                        If stackTop > stackSize Then
                                            stackSize = stackSize * 2
                                            ReDim Preserve stackR(1 To stackSize)
                                            ReDim Preserve stackC(1 To stackSize)
                                        End If
                                        stackR(stackTop) = nbrRow
                                        stackC(stackTop) = nbrCol
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
NextPop:
            Loop
        End If
    Next cell

CleanExit:
    Application.ScreenUpdating = priorScreenUpdating
    Application.Calculation = priorCalculation
    Application.EnableEvents = priorEnableEvents
    Exit Sub

CleanFail:
    MsgBox "flood_fill stopped because of this error:" & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


'==============================================================================
' Formula cleanup  (merged in from the old utils module)
'==============================================================================
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


' Delete every defined name whose definition is a LAMBDA. The sync path calls
' this before re-adding the current lambdas from the Lamb sheet, so a lambda
' that was renamed or removed does not linger as an orphan (the old per-row
' delete only removed names that were still in Lamb). Non-lambda names -
' repo_path, plain named ranges, add-in names such as IQ_* - are left alone,
' since a blanket clear would break repo_export's repo_path. Returns the count
' deleted.
Public Function delete_all_lambda_names(ByVal wb As Workbook) As Long
    Dim nm As Name, kill As Collection, s As Variant
    Dim rt As String, removed As Long

    Set kill = New Collection
    For Each nm In wb.Names
        rt = ""
        On Error Resume Next
        rt = nm.RefersTo            ' broken names can raise; treat as non-lambda
        On Error GoTo 0
        If InStr(1, rt, "LAMBDA(", vbTextCompare) > 0 Then
            kill.Add nm.Name        ' collect first; don't delete mid-iteration
        End If
    Next nm

    For Each s In kill
        On Error Resume Next
        Err.Clear
        wb.Names(s).Delete
        If Err.Number = 0 Then removed = removed + 1
        On Error GoTo 0
    Next s

    delete_all_lambda_names = removed
End Function


' One-shot cleanup of Excel's hidden LAMBDA-internal reserved names
' (_xleta.*, _xlop.*, _xlpm.*) that accumulate from editing lambdas. They are
' invisible in the Name Manager, show #NAME?, and only bloat the file. Excel-
' protected names that refuse deletion are skipped. This is NOT part of sync -
' run it deliberately, and if anything looks wrong afterwards just close
' without saving.
Public Sub sweep_name_cruft()
    Dim nm As Name, kill As Collection, s As Variant, removed As Long

    Set kill = New Collection
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.Name, "_xleta.", vbTextCompare) = 1 _
        Or InStr(1, nm.Name, "_xlop.", vbTextCompare) = 1 _
        Or InStr(1, nm.Name, "_xlpm.", vbTextCompare) = 1 Then
            kill.Add nm.Name
        End If
    Next nm

    For Each s In kill
        On Error Resume Next
        Err.Clear
        ThisWorkbook.Names(s).Delete
        If Err.Number = 0 Then removed = removed + 1
        On Error GoTo 0
    Next s

    MsgBox "Removed " & removed & " of " & kill.count & _
        " hidden reserved name(s) (_xleta. / _xlop. / _xlpm.).", _
        vbInformation, "sweep_name_cruft"
End Sub
