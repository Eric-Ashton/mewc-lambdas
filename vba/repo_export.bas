Attribute VB_Name = "repo_export"
' deploy: test
'==============================================================================
' repo_export  -  the write-BACK half of the old repo_sync: push in-Excel edits
'                 of the Lamb sheet out to the repo's lambdas\*.lambda files.
'                 (The read side - repo -> workbook - now lives in
'                 sync_test_workbook_from_repo.) Test-workbook only.
'
'   set_repo_path   Prompt for and remember the repo root (defined name repo_path).
'   export_lambdas  Write the current Lamb sheet back out to lambdas\*.lambda,
'                   one file per lambda, so a Lane-B (Excel-first) edit can be
'                   committed. Writes the blank-line-before-headers format.
'==============================================================================
Option Explicit

Private Const LAMB_SHEET As String = "Lamb"

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
    ThisWorkbook.Names.Add Name:="repo_path", RefersTo:="=""" & p & """", Visible:=True
    MsgBox "Repo path set to:" & vbLf & p, vbInformation, "set_repo_path"
End Sub

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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
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
Private Function repo_root_raw() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names("repo_path").RefersTo
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
    If Len(r) = 0 Then lambdas_folder = "" Else lambdas_folder = r & "\lambdas"
End Function

Private Function name_from_sig(ByVal sig As String) As String
    Dim p As Long
    sig = Trim$(sig)
    p = InStr(sig, "(")
    If p > 0 Then sig = Left$(sig, p - 1)
    name_from_sig = Trim$(sig)
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
