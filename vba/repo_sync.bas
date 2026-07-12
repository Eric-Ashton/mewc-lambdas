Attribute VB_Name = "repo_sync"
'==============================================================================
' repo_sync  -  sync the lambda library between this workbook and a local Git
'               repo of plain-text ".lambda" files (the source of truth).
'
'   import_lambdas   Read every lambdas\*.lambda file from the repo, rewrite the
'                    Lamb sheet (alphabetically) as quote-prefixed text, then run
'                    lambda_update to push them into the Name Manager.
'   export_lambdas   Write the current Lamb sheet back out to lambdas\*.lambda
'                    (one file per lambda) so in-Excel edits can be committed.
'   set_repo_path    Prompt for and remember the repo root folder.
'
' The repo root is stored in the workbook defined name "repo_path". The lambda
' files live in  <repo_path>\lambdas .  File format (one file per lambda):
'
'     === SIGNATURE ===
'     name(args)
'
'     === COMMENT ===
'     <=255 char comment
'
'     === CODE ===
'     =LAMBDA(...)
'
'     === DESCRIPTION ===
'     <full description>
'
' (A blank line precedes every header except the first; the parser tolerates it
'  either way, but export_lambdas writes it so files stay in this style.)
'==============================================================================
Option Explicit

Private Const LAMB_SHEET As String = "Lamb"

' ---- repo path (remembered in defined name "repo_path") ---------------------
Public Sub set_repo_path()
    Dim p As String
    p = InputBox("Path to the repo root folder (the one containing 'lambdas'):", _
                 "set_repo_path", repo_root_raw())
    If Len(p) = 0 Then Exit Sub
    Do While Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    On Error Resume Next
    ThisWorkbook.Names("repo_path").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="repo_path", _
        RefersTo:="=""" & p & """", Visible:=True
    MsgBox "Repo path set to:" & vbLf & p, vbInformation, "set_repo_path"
End Sub

Private Function repo_root_raw() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names("repo_path").RefersTo   ' looks like  ="C:\...\repo"
    On Error GoTo 0
    s = Replace(s, "=", "")
    s = Replace(s, Chr$(34), "")
    repo_root_raw = s
End Function

Private Function repo_root() As String
    Dim s As String
    s = repo_root_raw()
    If Len(s) = 0 Then
        set_repo_path
        s = repo_root_raw()
    End If
    repo_root = s
End Function

Private Function lambdas_folder() As String
    Dim r As String
    r = repo_root()
    If Len(r) = 0 Then
        lambdas_folder = ""
    Else
        lambdas_folder = r & "\lambdas"
    End If
End Function

' ---- import: repo -> Lamb sheet -> Name Manager -----------------------------
Public Sub import_lambdas()
    Dim folder As String
    folder = lambdas_folder()
    If Len(folder) = 0 Then Exit Sub                       ' user cancelled path prompt
    If Len(Dir$(folder, vbDirectory)) = 0 Then
        MsgBox "Lambdas folder not found:" & vbLf & folder & vbLf & vbLf & _
               "Run set_repo_path or check the repo location.", _
               vbExclamation, "import_lambdas"
        Exit Sub
    End If

    Dim names() As String, sigs() As String, coms() As String, cds() As String, dss() As String
    ReDim names(1 To 1000): ReDim sigs(1 To 1000)
    ReDim coms(1 To 1000): ReDim cds(1 To 1000): ReDim dss(1 To 1000)

    Dim f As String, content As String, n As Long, problems As String
    f = Dir$(folder & "\*.lambda")
    Do While Len(f) > 0
        content = read_file(folder & "\" & f)
        Dim sg As String, cd As String
        sg = section(content, "SIGNATURE")
        cd = section(content, "CODE")
        If Len(sg) = 0 Or Len(cd) = 0 Then
            problems = problems & vbLf & "  " & f & " (missing SIGNATURE or CODE)"
        Else
            n = n + 1
            sigs(n) = sg
            coms(n) = section(content, "COMMENT")
            cds(n) = cd
            dss(n) = section(content, "DESCRIPTION")
            names(n) = name_from_sig(sg)
        End If
        f = Dir$()
    Loop

    If n = 0 Then
        MsgBox "No usable .lambda files found in:" & vbLf & folder & _
               IIf(Len(problems) > 0, vbLf & vbLf & "Problems:" & problems, ""), _
               vbExclamation, "import_lambdas"
        Exit Sub
    End If

    sort_parallel names, sigs, coms, cds, dss, n

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LAMB_SHEET)
    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Range("A2:D" & lastRow).ClearContents

    Dim i As Long, rr As Long
    For i = 1 To n
        rr = i + 1
        ws.Cells(rr, 1).Value = sigs(i)
        ws.Cells(rr, 2).Value = coms(i)
        ws.Cells(rr, 3).Value = "'" & cds(i)   ' leading apostrophe -> stored as text
        ws.Cells(rr, 4).Value = dss(i)
        With ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 4)).Font
            .Name = "Aptos Narrow": .Size = 14
        End With
        ws.Cells(rr, 1).Font.Bold = True
        ws.Rows(rr).RowHeight = 19.25
    Next i

    Application.ScreenUpdating = True

    lambda_update        ' push the freshly-loaded Lamb sheet into the Name Manager

    MsgBox "Imported " & n & " lambda(s) from the repo and updated the Name Manager." & _
           IIf(Len(problems) > 0, vbLf & vbLf & "Skipped:" & problems, ""), _
           vbInformation, "import_lambdas"
End Sub

' ---- export: Lamb sheet -> repo ---------------------------------------------
Public Sub export_lambdas()
    Dim folder As String
    folder = lambdas_folder()
    If Len(folder) = 0 Then Exit Sub
    If Len(Dir$(folder, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
        If Len(Dir$(folder, vbDirectory)) = 0 Then
            MsgBox "Could not find or create:" & vbLf & folder, vbExclamation, "export_lambdas"
            Exit Sub
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LAMB_SHEET)
    Dim lastRow As Long, r As Long, cnt As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        Dim sig As String
        sig = ws.Cells(r, 1).Value
        If Len(sig) > 0 Then
            Dim body As String
            ' blank line before every header except the first (SIGNATURE)
            body = "=== SIGNATURE ===" & vbLf & sig & vbLf & _
                   vbLf & "=== COMMENT ===" & vbLf & CStr(ws.Cells(r, 2).Value) & vbLf & _
                   vbLf & "=== CODE ===" & vbLf & CStr(ws.Cells(r, 3).Value) & vbLf & _
                   vbLf & "=== DESCRIPTION ===" & vbLf & CStr(ws.Cells(r, 4).Value) & vbLf
            write_file folder & "\" & name_from_sig(sig) & ".lambda", body
            cnt = cnt + 1
        End If
    Next r
    MsgBox "Exported " & cnt & " lambda(s) to:" & vbLf & folder, vbInformation, "export_lambdas"
End Sub

' ---- helpers ---------------------------------------------------------------
Private Function name_from_sig(ByVal sig As String) As String
    Dim p As Long
    p = InStr(sig, "(")
    If p > 0 Then
        name_from_sig = Trim$(Left$(sig, p - 1))
    Else
        name_from_sig = Trim$(sig)
    End If
End Function

Private Function section(ByVal content As String, ByVal secName As String) As String
    Dim mark As String, p As Long, afterP As Long, q As Long, body As String
    mark = "=== " & secName & " ==="
    p = InStr(1, content, mark, vbTextCompare)
    If p = 0 Then section = "": Exit Function
    afterP = p + Len(mark)
    q = InStr(afterP, content, vbLf & "=== ")
    If q = 0 Then
        body = Mid$(content, afterP)
    Else
        body = Mid$(content, afterP, q - afterP)
    End If
    section = trim_ws(body)
End Function

Private Function trim_ws(ByVal s As String) As String
    Do While Len(s) > 0 And (Left$(s, 1) = vbCr Or Left$(s, 1) = vbLf Or Left$(s, 1) = " " Or Left$(s, 1) = vbTab)
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0 And (Right$(s, 1) = vbCr Or Right$(s, 1) = vbLf Or Right$(s, 1) = " " Or Right$(s, 1) = vbTab)
        s = Left$(s, Len(s) - 1)
    Loop
    trim_ws = s
End Function

Private Function read_file(ByVal path As String) As String
    ' UTF-8 read (ADODB.Stream) so Unicode in the .lambda files survives
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    read_file = st.ReadText(-1)
    st.Close
    read_file = Replace(read_file, vbCrLf, vbLf)
End Function

Private Sub write_file(ByVal path As String, ByVal content As String)
    ' UTF-8 write, no BOM
    Dim st As Object, bt As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText content
    st.Position = 3                       ' skip the 3-byte UTF-8 BOM
    Set bt = CreateObject("ADODB.Stream")
    bt.Type = 1: bt.Open
    st.CopyTo bt
    bt.SaveToFile path, 2                 ' 2 = overwrite
    st.Close: bt.Close
End Sub

Private Sub sort_parallel(names() As String, sigs() As String, coms() As String, _
                          cds() As String, dss() As String, ByVal n As Long)
    Dim i As Long, j As Long
    For i = 2 To n
        Dim kn As String, ks As String, kc As String, kd As String, ke As String
        kn = names(i): ks = sigs(i): kc = coms(i): kd = cds(i): ke = dss(i)
        j = i - 1
        Do While j >= 1
            If StrComp(names(j), kn, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j): sigs(j + 1) = sigs(j)
            coms(j + 1) = coms(j): cds(j + 1) = cds(j): dss(j + 1) = dss(j)
            j = j - 1
        Loop
        names(j + 1) = kn: sigs(j + 1) = ks
        coms(j + 1) = kc: cds(j + 1) = kd: dss(j + 1) = ke
    Next i
End Sub
