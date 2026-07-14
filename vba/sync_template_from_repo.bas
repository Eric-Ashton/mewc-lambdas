Attribute VB_Name = "sync_template_from_repo"
' deploy: template
Option Private Module
'==============================================================================
' sync_template_from_repo  -  pull the lambda library and VBA from the repo into
'   the DOWNSTREAM competition template. Consumer side: repo is the source of
'   truth; this only writes INTO the template.
'
'   sync_lambdas   Rewrite the Lamb sheet + (re)create the Name Manager names.
'   sync_vba       Import every repo vba\*.bas tagged 'deploy: shared' or
'                  'template', remove any stale standard module not in that set
'                  (with a confirmation), and skip this module itself.
'   sync_all       Run both (lambdas first, VBA last).
'
' Option Private Module keeps sync_* out of the Alt+F8 macro list during
' competition. Run these from the VBE (F5) or a button when updating the template.
'
' Self-contained on purpose: it can bootstrap a bare template. Place this one
' module in the template, set REPO_DIR, and run sync_all; the rest comes from the
' repo. It does NOT hide z_ helpers (they stay visible for competition).
'
' Requirements for sync_vba / sync_all:
'   Trust Center > Macro Settings > "Trust access to the VBA project object model".
'==============================================================================
Option Explicit

' ---- configuration (edit REPO_DIR for this machine) ------------------------
Private Const REPO_DIR   As String = "C:\MEWC Dev Cowork\mewc-lambdas"
Private Const MY_ROLE    As String = "template"
Private Const MODULE_ID  As String = "sync_template_from_repo"
Private Const LAMB_SHEET As String = "Lamb"

'==============================================================================
' Public entry points (hidden from Alt+F8 by Option Private Module)
'==============================================================================
Public Sub sync_lambdas()
    On Error GoTo fail
    MsgBox do_sync_lambdas(ThisWorkbook), vbInformation, MODULE_ID
    Exit Sub
fail:
    Dim em As String: em = "Error " & Err.Number & ": " & Err.Description
    restore_app
    MsgBox "sync_lambdas failed:" & vbLf & vbLf & em, vbExclamation, MODULE_ID
End Sub

Public Sub sync_vba()
    On Error GoTo fail
    MsgBox do_sync_vba(ThisWorkbook), vbInformation, MODULE_ID
    Exit Sub
fail:
    Dim em As String: em = "Error " & Err.Number & ": " & Err.Description
    restore_app
    MsgBox "sync_vba failed:" & vbLf & vbLf & em, vbExclamation, MODULE_ID
End Sub

Public Sub sync_all()
    On Error GoTo fail
    Dim wb As Workbook, lam As String, vb As String
    Set wb = ThisWorkbook
    lam = do_sync_lambdas(wb)          ' Lamb sheet + Name Manager (safe) first
    vb = do_sync_vba(wb)               ' modify the VBA project last
    MsgBox lam & vbLf & vbLf & vb, vbInformation, MODULE_ID
    Exit Sub
fail:
    Dim em As String: em = "Error " & Err.Number & ": " & Err.Description
    restore_app
    MsgBox "sync_all failed:" & vbLf & vbLf & em, vbExclamation, MODULE_ID
End Sub

'==============================================================================
' Workers
'==============================================================================
Private Function do_sync_lambdas(ByVal wb As Workbook) As String
    Dim folder As String
    folder = REPO_DIR & "\lambdas"
    If Len(Dir$(folder, vbDirectory)) = 0 Then _
        Err.Raise vbObjectError + 1, MODULE_ID, "Lambdas folder not found:" & vbLf & folder

    Dim nms() As String, sigs() As String, coms() As String, cds() As String, dss() As String
    ReDim nms(1 To 2000): ReDim sigs(1 To 2000): ReDim coms(1 To 2000)
    ReDim cds(1 To 2000): ReDim dss(1 To 2000)

    Dim n As Long, f As String, content As String, sg As String, cd As String
    f = Dir$(folder & "\*.lambda")
    Do While Len(f) > 0
        content = read_file(folder & "\" & f)
        sg = section(content, "SIGNATURE")
        cd = section(content, "CODE")
        If Len(sg) > 0 And Len(cd) > 0 Then
            n = n + 1
            sigs(n) = sg
            coms(n) = section(content, "COMMENT")
            cds(n) = cd
            dss(n) = section(content, "DESCRIPTION")
            nms(n) = name_from_sig(sg)
        End If
        f = Dir$()
    Loop
    If n = 0 Then Err.Raise vbObjectError + 2, MODULE_ID, "No usable .lambda files in:" & vbLf & folder

    sort_parallel nms, sigs, coms, cds, dss, n

    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    ' --- rewrite the Lamb sheet ---
    Dim ws As Worksheet
    Set ws = ensure_lamb_sheet(wb)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Range("A2:D" & lastRow).ClearContents

    Dim i As Long, rr As Long
    For i = 1 To n
        rr = i + 1
        ws.Cells(rr, 1).Value = sigs(i)
        ws.Cells(rr, 2).Value = coms(i)
        ws.Cells(rr, 3).Value = "'" & cds(i)          ' leading apostrophe -> stored as text
        ws.Cells(rr, 4).Value = dss(i)
    Next i

    ' --- push into the Name Manager (multi-pass so interdependent lambdas resolve) ---
    Dim i2 As Long
    For i2 = 1 To n
        On Error Resume Next
        wb.names(nms(i2)).Delete
        On Error GoTo 0
    Next i2

    Dim added() As Boolean: ReDim added(1 To n)
    Dim pass As Long, progress As Boolean, addedCount As Long
    For pass = 1 To 25
        progress = False
        For i = 1 To n
            If Not added(i) Then
                On Error Resume Next
                Err.Clear
                wb.names.Add Name:=nms(i), RefersTo:=cds(i)
                If Err.Number = 0 Then added(i) = True: progress = True: addedCount = addedCount + 1
                On Error GoTo 0
            End If
        Next i
        If Not progress Then Exit For
    Next pass

    For i = 1 To n
        If added(i) And Len(coms(i)) > 0 Then
            On Error Resume Next
            wb.names(nms(i)).Comment = Left$(coms(i), 255)
            On Error GoTo 0
        End If
    Next i

    Dim failed As String
    For i = 1 To n
        If Not added(i) Then failed = failed & vbLf & "  " & nms(i)
    Next i

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    do_sync_lambdas = "Lambdas: " & addedCount & " of " & n & " loaded into the Name Manager." & _
        IIf(Len(failed) > 0, vbLf & "  Could not add (check the Code):" & failed, "")
End Function

Private Function do_sync_vba(ByVal wb As Workbook) As String
    Dim proj As Object
    On Error Resume Next
    Set proj = wb.VBProject
    On Error GoTo 0
    If proj Is Nothing Then _
        Err.Raise vbObjectError + 3, MODULE_ID, _
            "Can't reach the VBA project. Turn on Trust Center > Macro Settings > " & _
            "'Trust access to the VBA project object model', then retry."

    Dim folder As String
    folder = REPO_DIR & "\vba"
    If Len(Dir$(folder, vbDirectory)) = 0 Then _
        Err.Raise vbObjectError + 4, MODULE_ID, "VBA folder not found:" & vbLf & folder

    Dim keep As Object
    Set keep = CreateObject("Scripting.Dictionary")
    keep.Add LCase$(MODULE_ID), True                   ' never prune self

    Dim f As String, base As String, role As String, comp As Object
    Dim cnt As Long, skipped As String
    f = Dir$(folder & "\*.bas")
    Do While Len(f) > 0
        base = Left$(f, Len(f) - 4)
        role = deploy_role(folder & "\" & f)
        If LCase$(base) <> LCase$(MODULE_ID) And (role = "shared" Or role = MY_ROLE) Then
            If Not keep.Exists(LCase$(base)) Then keep.Add LCase$(base), True
            Set comp = Nothing
            On Error Resume Next
            Set comp = proj.VBComponents(base)
            On Error GoTo 0
            If Not comp Is Nothing Then proj.VBComponents.Remove comp
            proj.VBComponents.Import folder & "\" & f
            cnt = cnt + 1
        Else
            skipped = skipped & " " & base
        End If
        f = Dir$()
    Loop

    Dim pruned As String
    pruned = prune_modules(proj, keep)

    do_sync_vba = "VBA: imported " & cnt & " module(s)." & _
        IIf(Len(skipped) > 0, vbLf & "  Skipped (other role):" & skipped, "") & _
        IIf(Len(pruned) > 0, vbLf & "  " & pruned, "")
End Function

'==============================================================================
' Helpers
'==============================================================================
Private Function deploy_role(ByVal basPath As String) As String
    Dim txt As String, p As Long, rest As String, tok As String, i As Long, ch As String
    txt = read_file(basPath)
    p = InStr(1, txt, "' deploy:", vbTextCompare)
    If p = 0 Then deploy_role = "shared": Exit Function
    rest = LTrim$(Mid$(txt, p + Len("' deploy:")))
    For i = 1 To Len(rest)
        ch = Mid$(rest, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit For
        tok = tok & ch
    Next i
    deploy_role = LCase$(Trim$(tok))
    If Len(deploy_role) = 0 Then deploy_role = "shared"
End Function

Private Function prune_modules(ByVal proj As Object, ByVal keep As Object) As String
    Dim comp As Object, victims As Collection, nm As String, v As Variant
    Set victims = New Collection
    For Each comp In proj.VBComponents
        If comp.Type = 1 Then                          ' vbext_ct_StdModule
            nm = comp.Name
            If Not keep.Exists(LCase$(nm)) Then victims.Add nm
        End If
    Next comp
    If victims.count = 0 Then prune_modules = "": Exit Function

    Dim lst As String
    For Each v In victims: lst = lst & vbLf & "  " & v: Next v
    If MsgBox("These standard modules are in the template but not in the repo for this " & _
              "role. Remove them?" & vbLf & lst, vbYesNo + vbQuestion, MODULE_ID) <> vbYes Then
        prune_modules = "prune skipped by user (" & victims.count & " stale module(s) left)."
        Exit Function
    End If
    Dim removed As Long
    For Each v In victims
        On Error Resume Next
        proj.VBComponents.Remove proj.VBComponents(CStr(v))
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
        On Error GoTo 0
    Next v
    prune_modules = "pruned " & removed & " stale module(s)."
End Function

Private Function ensure_lamb_sheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(LAMB_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = LAMB_SHEET
        ws.Range("A1").Value = "Signature"
        ws.Range("B1").Value = "Comment"
        ws.Range("C1").Value = "Code"
        ws.Range("D1").Value = "Description"
        ws.Range("A1:D1").Font.Bold = True
    End If
    Set ensure_lamb_sheet = ws
End Function

Private Sub restore_app()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Function name_from_sig(ByVal sig As String) As String
    Dim p As Long
    sig = Trim$(sig)
    p = InStr(sig, "(")
    If p > 0 Then sig = Left$(sig, p - 1)
    name_from_sig = Trim$(sig)
End Function

Private Function section(ByVal content As String, ByVal secName As String) As String
    Dim mark As String, p As Long, afterP As Long, q As Long, body As String
    mark = "=== " & secName & " ==="
    p = InStr(1, content, mark, vbTextCompare)
    If p = 0 Then section = "": Exit Function
    afterP = p + Len(mark)
    q = InStr(afterP, content, vbLf & "=== ")
    If q = 0 Then body = Mid$(content, afterP) Else body = Mid$(content, afterP, q - afterP)
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
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    read_file = st.ReadText(-1)
    st.Close
    read_file = Replace(read_file, vbCrLf, vbLf)
End Function

Private Sub sort_parallel(nms() As String, sigs() As String, coms() As String, _
                          cds() As String, dss() As String, ByVal n As Long)
    Dim i As Long, j As Long
    For i = 2 To n
        Dim kn As String, ks As String, kc As String, kd As String, ke As String
        kn = nms(i): ks = sigs(i): kc = coms(i): kd = cds(i): ke = dss(i)
        j = i - 1
        Do While j >= 1
            If StrComp(nms(j), kn, vbTextCompare) <= 0 Then Exit Do
            nms(j + 1) = nms(j): sigs(j + 1) = sigs(j)
            coms(j + 1) = coms(j): cds(j + 1) = cds(j): dss(j + 1) = dss(j)
            j = j - 1
        Loop
        nms(j + 1) = kn: sigs(j + 1) = ks
        coms(j + 1) = kc: cds(j + 1) = kd: dss(j + 1) = ke
    Next i
End Sub
