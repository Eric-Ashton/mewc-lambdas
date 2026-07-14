Attribute VB_Name = "utility_functions"
' deploy: shared


'========================
' Function get_color
'========================
' Returns a string representing the RGB color of a cell.
' - For gradients (2+ stops): returns "#RRGGBB/#RRGGBB/..." for each stop in order.
' - For solid fills: returns "#RRGGBB".
' - For truly unfilled cells (Pattern = xlPatternNone): returns "#FFFFFF".
'   Note: a deliberately-filled-black cell has Pattern = xlPatternSolid and
'   Color = 0, and correctly returns "#000000" via the solid branch. Only
'   cells with no fill at all are treated as white.
'
' Parameters:
'   cell - The target cell to extract color from.

Function get_color(cell As Range) As String
    Dim rgb_color As String
    Dim r As Long, g As Long, b As Long
    Dim grad As Variant
    Dim n As Long, i As Long, c As Long
    Dim parts() As String

    With cell.Interior
        If .Pattern = xlPatternLinearGradient Or .Pattern = xlPatternRectangularGradient Then
            Set grad = .gradient
            n = grad.ColorStops.count
            If n >= 2 Then
                ReDim parts(0 To n - 1)
                For i = 1 To n
                    c = grad.ColorStops(i).Color
                    r = c Mod 256
                    g = (c \ 256) Mod 256
                    b = (c \ 65536) Mod 256
                    parts(i - 1) = "#" & Hex2(r) & Hex2(g) & Hex2(b)
                Next i
                rgb_color = Join(parts, "/")
            Else
                ' Fallback (unlikely): treat as solid
                c = .Color
                r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
                rgb_color = "#" & Hex2(r) & Hex2(g) & Hex2(b)
            End If

        ElseIf .Pattern = xlPatternNone Then
            ' Truly unfilled cell -> treat as white. Distinct from a
            ' deliberately-filled-black cell (which has Pattern=xlPatternSolid
            ' and Color=0, and falls through to the solid branch below).
            rgb_color = "#FFFFFF"

        Else
            c = .Color
            r = c Mod 256
            g = (c \ 256) Mod 256
            b = (c \ 65536) Mod 256
            rgb_color = "#" & Hex2(r) & Hex2(g) & Hex2(b)
        End If
    End With

    get_color = rgb_color
End Function
'Helper function for get_color
Private Function Hex2(n As Long) As String
    ' Returns a two-digit uppercase hex string for 0â€“255
    Hex2 = Right$("0" & Hex$(n And &HFF), 2)
End Function


' ============
' sheet_names
' ============
' Returns a vertical array of all sheet names in the workbook, in order.
'
' Usage:
'   =sheet_names()
'
' Returns:
'   A single-column array with the names of all worksheets in the workbook.
'
Function sheet_names() As Variant
    Dim ws As Worksheet
    Dim sheetNames() As String
    Dim i As Long

    ReDim sheetNames(1 To ThisWorkbook.Worksheets.count)

    For i = 1 To ThisWorkbook.Worksheets.count
        sheetNames(i) = ThisWorkbook.Worksheets(i).Name
    Next i

    sheet_names = Application.Transpose(sheetNames)
End Function


' ============
' sheet_data
' ============
' Returns vertically stacked or flattened data from the same range across multiple sheets.
'
' Parameters:
'   sheet_names (Range) - An array of worksheet names (one per cell).
'                         Each must match a valid worksheet in the workbook.
'
'   range_ref (String)  - An A1-style reference (e.g., "B4:C9") indicating the cell range
'                         to extract from each sheet.
'
'   flatten (Integer) [Optional] - If 0 (default), returns a stacked table:
'       - First row contains headers: Sheet | Row | Column1 | Column2 | ...
'       - Each sheet's data appears as a block of rows.
'     If 1, returns a flattened table in row-major order (like TOCOL):
'       - First column is A1-style addresses (B4, C4, B5, etc.)
'       - First row contains: [blank] | Sheet1 | Sheet2 | ...
'       - Each sheetâ€™s data appears in a single column.
'
' Returns:
'   Variant 2D array:
'     - If flatten = 0, returns (n * sheet_count) + 1 rows by (columns + 2)
'     - If flatten = 1, returns (n * columns) + 1 rows by (sheet_count + 1)
'     - If a sheet name is invalid or range is incorrect, returns #REF! error
'
' Example usage:
'   =sheet_data(A1:A3, "B2:D4")           ' Stacked table from 3 sheets
'   =sheet_data(A1:A3, "B2:D4", 1)        ' Flattened TOCOL-style view
'
Function sheet_data(sheet_names As Variant, range_ref As String, Optional flatten As Integer = 0) As Variant
    Dim nameList() As String
    Dim numSheets As Long, i As Long
    Dim refRange As Range
    Dim rowCount As Long, colCount As Long
    Dim startRow As Long, startCol As Long
    Dim out(), arr()
    Dim wsName As String
    Dim r As Long, c As Long, k As Long, idx As Long

    ' Normalize sheet_names to 1D row-major list
    nameList = RowMajorNames(sheet_names)
    If (Not Not nameList) = 0 Then
        sheet_data = CVErr(xlErrRef)
        Exit Function
    End If
    numSheets = UBound(nameList)

    ' Validate first sheet/range and size
    On Error GoTo InvalidRange
    Set refRange = ThisWorkbook.Sheets(nameList(1)).Range(range_ref)
    rowCount = refRange.Rows.count
    colCount = refRange.Columns.count
    startRow = refRange.Row
    startCol = refRange.Column
    On Error GoTo 0

    If flatten = 0 Then
        ' === Stacked ===
        ReDim out(0 To numSheets * rowCount, 1 To 2 + colCount)
        out(0, 1) = "Sheet"
        out(0, 2) = "Row"
        For c = 1 To colCount
            out(0, 2 + c) = Split(Cells(1, startCol + c - 1).Address(False, False), "1")(0)
        Next c

        k = 1
        For i = 1 To numSheets
            wsName = nameList(i)
            On Error Resume Next
            Set refRange = ThisWorkbook.Sheets(wsName).Range(range_ref)
            If Err.Number <> 0 Then
                sheet_data = CVErr(xlErrRef): Exit Function
            End If
            On Error GoTo 0

            arr = refRange.Value
            For r = 1 To rowCount
                out(k, 1) = wsName
                out(k, 2) = startRow + r - 1
                For c = 1 To colCount
                    If IsEmpty(arr(r, c)) Or arr(r, c) = "" Then
                        out(k, 2 + c) = ""
                    Else
                        out(k, 2 + c) = arr(r, c)
                    End If
                Next c
                k = k + 1
            Next r
        Next i

    Else
        ' === Flatten (TOCOL-style) ===
        Dim flatCount As Long
        flatCount = rowCount * colCount
        ReDim out(0 To flatCount, 0 To numSheets)

        out(0, 0) = ""
        For i = 1 To numSheets
            out(0, i) = nameList(i)
        Next i

        idx = 1
        For r = 1 To rowCount
            For c = 1 To colCount
                out(idx, 0) = Cells(startRow + r - 1, startCol + c - 1).Address(False, False)
                idx = idx + 1
            Next c
        Next r

        For i = 1 To numSheets
            wsName = nameList(i)
            On Error Resume Next
            Set refRange = ThisWorkbook.Sheets(wsName).Range(range_ref)
            If Err.Number <> 0 Then
                sheet_data = CVErr(xlErrRef): Exit Function
            End If
            On Error GoTo 0

            arr = refRange.Value
            idx = 1
            For r = 1 To rowCount
                For c = 1 To colCount
                    If IsEmpty(arr(r, c)) Or arr(r, c) = "" Then
                        out(idx, i) = ""
                    Else
                        out(idx, i) = arr(r, c)
                    End If
                    idx = idx + 1
                Next c
            Next r
        Next i
    End If

    sheet_data = out
    Exit Function

InvalidRange:
    sheet_data = CVErr(xlErrRef)
End Function


' --- Helper for sheet_data: normalize sheet_names (Range | 1D/2D array | String) to 1D row-major String() ---
Private Function RowMajorNames(ByVal sheet_names As Variant) As String()
    Dim names() As String
    Dim r As Long, c As Long, i As Long
    Dim n As Long

    If IsObject(sheet_names) Then
        ' Range (any shape) ? row-major
        Dim rng As Range
        Set rng = sheet_names
        ReDim names(1 To rng.Rows.count * rng.Columns.count)
        For r = 1 To rng.Rows.count
            For c = 1 To rng.Columns.count
                n = n + 1
                names(n) = CStr(rng.Cells(r, c).Value)
            Next c
        Next r

    ElseIf IsArray(sheet_names) Then
        ' VBA array (1D or 2D) ? row-major
        Dim rL As Long, rU As Long, cL As Long, cU As Long
        Dim is2D As Boolean

        On Error Resume Next
        rL = LBound(sheet_names, 1): rU = UBound(sheet_names, 1)
        cL = LBound(sheet_names, 2)
        is2D = (Err.Number = 0)
        If is2D Then cU = UBound(sheet_names, 2)
        Err.Clear
        On Error GoTo 0

        If is2D Then
            ReDim names(1 To (rU - rL + 1) * (cU - cL + 1))
            For r = rL To rU
                For c = cL To cU
                    n = n + 1
                    names(n) = CStr(sheet_names(r, c))
                Next c
            Next r
        Else
            ReDim names(1 To (rU - rL + 1))
            For i = rL To rU
                n = n + 1
                names(n) = CStr(sheet_names(i))
            Next i
        End If

    Else
        ' Single value ? 1 element
        ReDim names(1 To 1)
        names(1) = CStr(sheet_names)
    End If

    RowMajorNames = names
End Function

Function UNICODE_SPLIT(textToSplit As String) As Variant
    Dim result() As String
    Dim charCount As Long
    Dim i As Long
    Dim currentChar As String
    Dim lastItem As String
    
    If Len(textToSplit) = 0 Then
        UNICODE_SPLIT = ""
        Exit Function
    End If
    
    ReDim result(0 To 0)
    result(0) = Mid(textToSplit, 1, 1)
    charCount = 0
    
    For i = 2 To Len(textToSplit)
        currentChar = Mid(textToSplit, i, 1)
        lastItem = result(charCount)
        
        'Check if the new character merges with the last one (e.g., emoji + modifier)
        If Len(lastItem & currentChar) = 1 Then
            result(charCount) = lastItem & currentChar 'Merge by replacing the last item
        Else
            charCount = charCount + 1
            ReDim Preserve result(0 To charCount)
            result(charCount) = currentChar 'Add as a new item
        End If
    Next i
    
    'Transpose the array to make it spill vertically in the worksheet
    UNICODE_SPLIT = Application.Transpose(result)
End Function


'==============================================================================
' Unit conversion  (merged in from the old conversions module)
'==============================================================================
' =========
' xconvert
' =========
' Converts a quantity from one unit to another using a user-supplied conversion table.
'
' Parameters:
'   conversion_table (Variant)
'       A 3-column table supplied either as:
'         • A Range (e.g., A2:C20), or
'         • An inline dynamic array from LET/VSTACK/HSTACK, or
'         • A literal array constant.
'       Columns (1-based):
'         [1] From_Unit  (String)
'         [2] Factor     (Numeric)  ; amount_in_to_unit = amount_in_from_unit * Factor
'         [3] To_Unit    (String)
'       Example rows:
'         "USD", 18.2, "ZAR"
'         "m",   100,  "cm"
'
'   unit_from (String)
'       The source unit symbol (case-insensitive).
'
'   quantity (Double)
'       The numeric value to convert.
'
'   unit_to (String)
'       The target unit symbol (case-insensitive).
'
' Returns:
'   Variant
'       Converted numeric value if a path exists; otherwise a worksheet error:
'         • #N/A if no conversion path is found
'         • #VALUE! if the table is malformed
'
' Notes:
'   • Works directly with dynamic arrays and spilled ranges.
'   • Builds a directed graph of conversions. Reverse links are added automatically
'     with reciprocal factors, so you only need to enter one direction in the table.
'   • Finds a path using BFS and composes multiplicative factors along the path.
Public Function xconvert(conversion_table As Variant, _
                         unit_from As String, _
                         quantity As Double, _
                         unit_to As String) As Variant
    On Error GoTo FailHard

    Dim tbl As Variant
    tbl = CoerceTo2D(conversion_table)

    Dim r As Long, rLo As Long, rHi As Long
    Dim cLo As Long, cHi As Long
    rLo = LBound(tbl, 1): rHi = UBound(tbl, 1)
    cLo = LBound(tbl, 2): cHi = UBound(tbl, 2)

    If (cHi - cLo + 1) < 3 Then
        xconvert = CVErr(xlErrValue)
        Exit Function
    End If

    Dim fromU As String, toU As String
    Dim factor As Double
    Dim g As Object ' Dictionary: key=unit (String), item=Collection of Edge objects [neighbor|factor]
    Set g = CreateObject("Scripting.Dictionary")

    ' Build graph with forward and reverse edges
    For r = rLo To rHi
        fromU = CStr(tbl(r, cLo))
        toU = CStr(tbl(r, cLo + 2))

        If Not IsBlankLike(fromU) And Not IsBlankLike(toU) Then
            If IsNumeric(tbl(r, cLo + 1)) Then
                factor = CDbl(tbl(r, cLo + 1))
                If factor <> 0 Then
                    AddEdge g, LCase$(fromU), LCase$(toU), factor
                    AddEdge g, LCase$(toU), LCase$(fromU), 1# / factor
                End If
            End If
        End If
    Next r

    Dim src As String, dst As String
    src = LCase$(unit_from)
    dst = LCase$(unit_to)

    ' Trivial case
    If src = dst Then
        xconvert = quantity
        Exit Function
    End If

    ' No nodes?
    If Not g.Exists(src) And Not g.Exists(dst) Then
        xconvert = CVErr(xlErrNA)
        Exit Function
    End If

    ' BFS over units, accumulating product of factors
    Dim q As Collection: Set q = New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary") ' unit -> cumulative factor
    seen.Add src, 1#
    q.Add src

    Dim u As String, v As String, i As Long
    Dim cum As Double, newCum As Double

    Do While q.count > 0
        u = q(1): q.Remove 1
        cum = seen(u)

        If g.Exists(u) Then
            Dim edges As Collection
            Set edges = g(u)
            For i = 1 To edges.count
                v = edges(i)(0)
                newCum = cum * edges(i)(1)
                If Not seen.Exists(v) Then
                    seen.Add v, newCum
                    If v = dst Then
                        xconvert = quantity * newCum
                        Exit Function
                    End If
                    q.Add v
                End If
            Next i
        End If
    Loop

    xconvert = CVErr(xlErrNA)
    Exit Function

FailHard:
    xconvert = CVErr(xlErrValue)
End Function

' ---- Helpers ---------------------------------------------------------------

' Adds a directed edge u -> v with multiplicative factor f
Private Sub AddEdge(ByRef g As Object, ByVal u As String, ByVal v As String, ByVal f As Double)
    Dim edges As Collection
    If Not g.Exists(u) Then
        Set edges = New Collection
        g.Add u, edges
    Else
        Set edges = g(u)
    End If
    Dim e(1) As Variant
    e(0) = v
    e(1) = f
    edges.Add e
End Sub

' Coerces a Range, 1D array, 2D array, or scalar into a 2D Variant array
Private Function CoerceTo2D(ByVal v As Variant) As Variant
    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            CoerceTo2D = v.Value2
            Exit Function
        End If
    End If

    If IsArray(v) Then
        On Error GoTo Make2D
        Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
        lb1 = LBound(v, 1): ub1 = UBound(v, 1)
        lb2 = LBound(v, 2): ub2 = UBound(v, 2)
        CoerceTo2D = v
        Exit Function
Make2D:
        Dim n As Long, i As Long
        n = UBound(v) - LBound(v) + 1
        Dim a(): ReDim a(1 To n, 1 To 1)
        For i = 1 To n
            a(i, 1) = v(LBound(v) + i - 1)
        Next
        CoerceTo2D = a
        Exit Function
    End If

    Dim s(1 To 1, 1 To 1) As Variant
    s(1, 1) = v
    CoerceTo2D = s
End Function

' Treats Empty, zero-length, or whitespace-only strings as blank
Private Function IsBlankLike(ByVal x As Variant) As Boolean
    If IsEmpty(x) Then
        IsBlankLike = True
    ElseIf VarType(x) = vbString Then
        IsBlankLike = (Len(Trim$(x)) = 0)
    Else
        IsBlankLike = False
    End If
End Function
