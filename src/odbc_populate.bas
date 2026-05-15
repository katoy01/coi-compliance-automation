' Module 3

Option Explicit

' === CONFIGURATION CONSTANTS ===
Private Const CONST_ODBC_DSN As String = "YOUR_DSN_HERE"
Private Const CONST_ODBC_UID As String = "YOUR_UID_HERE"
Private Const CONST_ODBC_PWD As String = "YOUR_PWD_HERE"
Private Const CONST_COMPANY_A As String = "CompanyA"

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo1 As ListObject, lo2 As ListObject, rngBody1 As Range, rngBody2 As Range

    If Target.Count > 1 Then Exit Sub
    If Target.Column <> 1 Or Target.row < 2 Then Exit Sub
    If Trim(CStr(Target.Value)) = "" Then Exit Sub

    On Error Resume Next
    Set lo1 = Me.ListObjects("COIRequest")
    On Error GoTo 0
    If lo1 Is Nothing Then Exit Sub
    Set rngBody1 = lo1.DataBodyRange
    If rngBody1 Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo2 = Me.ListObjects("FindWO")
    On Error GoTo 0
    If lo2 Is Nothing Then Exit Sub
    Set rngBody2 = lo2.DataBodyRange
    If rngBody2 Is Nothing Then Exit Sub

    If Not Intersect(Target, rngBody1) Is Nothing Then Call PopulateFromWO(Target.row)
    If Not Intersect(Target, rngBody2) Is Nothing Then Call MarkFoundStatus_Match(CStr(Target.Value), Target.row)
End Sub

Sub MarkFoundStatus_Match(woNum As String, rowNum As Long)
    Dim lo As ListObject
    Dim m As Variant

    Set lo = Me.ListObjects("COIRequest")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        Me.Cells(rowNum, 2).Value = "Not Found"
        Exit Sub
    End If

    m = Application.Match(Val(Trim(woNum)), lo.ListColumns(1).DataBodyRange, 0)
    If IsError(m) Then
        Me.Cells(rowNum, 2).Value = "Not Found"
    Else
        Me.Cells(rowNum, 2).Value = "Found"
    End If
End Sub

Sub Click_PopulateFromWO()
    Dim row As ListRow
    Dim response As Integer
    Dim send As Integer
    For Each row In Sheet1.ListObjects("COIRequest").ListRows
        If row.Range.Cells(1, 15).Value = True Then
            Call PopulateFromWO(row.Range.row)
            row.Range.Cells(1, 15).Value = False
        End If
    Next row
End Sub

Sub PopulateFromWO(rowNum As Long)
    Dim woNum As String
    Dim ws As Worksheet
    
    Set ws = Sheet1
    
    woNum = ws.Cells(rowNum, 1).Value
    
    If woNum = "" Then Exit Sub
    
    Dim conn As Object
    Dim rs As Object
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    On Error GoTo CleanUp
    
    conn.Open "DSN=" & CONST_ODBC_DSN & ";UID=" & CONST_ODBC_UID & ";PWD=" & CONST_ODBC_PWD & ";"
    
    Dim sql As String
    sql = "SELECT TOP 1 " & _
              "w.wrkordr_id, " & _
              "cs.clntste_nme, " & _
              "jcc.jbcstcde_nme, " & _
              "j.jb_nme, " & _
              "c.clnt_eml, " & _
              "w.wrkordr_entrd_by, " & _
              "wt.wrkordrtchncn_dte_schdld " & _
          "FROM wrkordr w " & _
          "JOIN clntste cs ON w.clntste_rn = cs.clntste_rn " & _
          "JOIN jbcstcde jcc ON w.jbcstcde_rn = jcc.jbcstcde_rn " & _
          "LEFT JOIN jbbllngitm jbi ON w.jbbllngitm_rn = jbi.jbbllngitm_rn " & _
          "LEFT JOIN jb j ON jbi.jb_rn = j.jb_rn " & _
          "LEFT JOIN clnt c ON j.clnt_rn = c.clnt_rn " & _
          "LEFT JOIN wrkordrtchncn wt ON w.wrkordr_rn = wt.wrkordr_rn " & _
              "AND wt.wrkordrlgtype_id IN ('Waiting Service', 'Pending', 'Scheduled') " & _
          "WHERE w.wrkordr_id = '" & woNum & "' " & _
          "ORDER BY wt.wrkordrtchncn_dte_schdld DESC"
    
    rs.Open sql, conn
    
    If Not rs.EOF Then
        ' Date of Appt - column 2
        If IsEmpty(ws.Cells(rowNum, 2)) Then
            If Not IsNull(rs.Fields("wrkordrtchncn_dte_schdld").Value) Then
                ws.Cells(rowNum, 2).Value = CDate(rs.Fields("wrkordrtchncn_dte_schdld").Value)
                ws.Cells(rowNum, 2).NumberFormat = "[$-en-US]dddd, mmmm d, yyyy"
            End If
        End If
        If IsEmpty(ws.Cells(rowNum, 3)) Then ws.Cells(rowNum, 3).Value = rs.Fields("clntste_nme").Value
        ' Map ERP cost code to internal company identifier
        If IsEmpty(ws.Cells(rowNum, 4)) Then
            If InStr(rs.Fields("jbcstcde_nme").Value, CONST_COMPANY_A) > 0 Then
                ws.Cells(rowNum, 4).Value = "CompanyA"
            Else
                ws.Cells(rowNum, 4).Value = "CompanyB"
            End If
        End If
        If IsEmpty(ws.Cells(rowNum, 6)) Then ws.Cells(rowNum, 6).Value = rs.Fields("jb_nme").Value
        If IsEmpty(ws.Cells(rowNum, 8)) Then ws.Cells(rowNum, 8).Value = rs.Fields("clnt_eml").Value
        If IsEmpty(ws.Cells(rowNum, 10)) Then ws.Cells(rowNum, 10).Value = rs.Fields("wrkordr_entrd_by").Value
    Else
        MsgBox "No results found for WO# " & woNum
    End If
    
    Exit Sub
    
CleanUp:
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
End Sub