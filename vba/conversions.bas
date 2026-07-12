Attribute VB_Name = "conversions"
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

