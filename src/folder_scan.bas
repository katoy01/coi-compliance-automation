' Module 2

Option Explicit

' === CONFIGURATION CONSTANTS ===
Private Const CONST_USER_PATH As String = "C:\Users\username\"

'This Module is for COI Tracking

Sub ScanCOIFolder()
'Declare Dim variables
    Dim coiTable As ListObject
    Dim mainTable As ListObject
    Dim newRow As ListRow
    Dim matchresult As Variant
    
    Dim folderPath As String
    Dim filename As String
    Dim woNum As String
    Dim response As Integer
    
'User can cancel if desired
    response = MsgBox("Scan COI Folder?", vbOKCancel)
    If response = vbCancel Then Exit Sub
'Populate variables
    Set mainTable = Sheet1.ListObjects("COIRequest")
    Set coiTable = Sheet4.ListObjects("tableCOI")
    folderPath = CONST_USER_PATH & "Downloads\1_COI_WO\"
'Loop: iterates through every file in folder and adds new files to excel table
    filename = Dir(folderPath & "*.pdf")
    Do While filename <> ""
        ' Parse WO# from filename
        woNum = Split(Split(filename, ".")(0), "_")(1)
        
        ' Check if WO already exists in excel WO tracker
        ' If WO doesn't exist then add new line:
        If coiTable.ListRows.Count = 0 Then
            matchresult = CVErr(xlErrNA)
        Else
            matchresult = Application.Match(filename, coiTable.ListColumns(2).DataBodyRange, 0)
        End If
        If IsError(matchresult) Then
            Set newRow = coiTable.ListRows.Add()
            newRow.Range.Cells(1, 1).Value = woNum
            newRow.Range.Cells(1, 2).Value = filename
            newRow.Range.Cells(1, 3).Value = Format(FileDateTime(folderPath & filename), "mm/dd/yyyy")
            newRow.Range.Cells(1, 4).Value = "Received"
        End If
        
        filename = Dir()
    Loop
    
    Call UpdateCOITable
End Sub

Sub ScanWCFolder()
'Declare Dim variables
    Dim wcTable As ListObject
    Dim mainTable As ListObject
    Dim newRow As ListRow
    Dim matchresult As Variant
    
    Dim folderPath As String
    Dim filename As String
    Dim certHolder As String
    Dim company As String
    Dim response As Integer
    
'User can cancel if desired
    response = MsgBox("Scan WC Folder?", vbOKCancel)
    If response = vbCancel Then Exit Sub
'Populate variables
    Set mainTable = Sheet1.ListObjects("COIRequest")
    Set wcTable = Sheet5.ListObjects("tableWC")
    folderPath = CONST_USER_PATH & "Downloads\1_WC_Cert\"
'Loop: iterates through every file in folder and adds new files to excel table
    filename = Dir(folderPath & "*.pdf")
    Do While filename <> ""
        ' Parse WC Certholder & Company from filename
        certHolder = Split(Split(filename, ".")(0), "_")(1)
        company = Split(Split(filename, ".")(0), "_")(2)
        
        ' Check if WC already exists in WC tracker
        ' If WC doesn't exist then add new row:
        If wcTable.ListRows.Count = 0 Then
            matchresult = CVErr(xlErrNA)
        Else
            matchresult = Application.Match(filename, wcTable.ListColumns(3).DataBodyRange, 0)
        End If
        If IsError(matchresult) Then
            Set newRow = wcTable.ListRows.Add()
            newRow.Range.Cells(1, 1).Value = certHolder
            newRow.Range.Cells(1, 2).Value = company
            newRow.Range.Cells(1, 3).Value = filename
            newRow.Range.Cells(1, 4).Value = Format(FileDateTime(folderPath & filename), "mm/dd/yyyy")
            newRow.Range.Cells(1, 5).Value = "Received"
        End If
        
        filename = Dir()
    Loop
    
    Call UpdateWCTable
    
End Sub

Sub ScanHHFolder()

    Dim mainTable As ListObject
    Dim hhTable As ListObject
    Dim newRow As ListRow
    Dim matchresult As Variant
    
    Dim folderPath As String
    Dim filename As String
    Dim clientName As String
    Dim response As Integer
    
'User can cancel if desired
    response = MsgBox("Scan HH Folder?", vbOKCancel)
    If response = vbCancel Then Exit Sub
'Populate variables
    Set mainTable = Sheet1.ListObjects("COIRequest")
    Set hhTable = Sheet6.ListObjects("tableHH")
    folderPath = CONST_USER_PATH & "Downloads\1_Indemnity-HH\"
    filename = Dir(folderPath & "*.pdf")
'Loop: iterates through every file in folder and adds new files to excel table
    Do While filename <> ""
        ' Parse client name / building name from filename
        clientName = Split(Split(filename, "-")(1), ".")(0)
        
        ' Check if HH already exists in tracker
        ' If HH doesn't exist then add new row:
        If hhTable.ListRows.Count = 0 Then
            matchresult = CVErr(xlErrNA)
        Else
            matchresult = Application.Match(filename, hhTable.ListColumns(2).DataBodyRange, 0)
        End If
        If IsError(matchresult) Then
            Set newRow = hhTable.ListRows.Add()
            newRow.Range.Cells(1, 1).Value = clientName
            newRow.Range.Cells(1, 2).Value = filename
            newRow.Range.Cells(1, 3).Value = Format(FileDateTime(folderPath & filename), "mm/dd/yyyy")
            newRow.Range.Cells(1, 4).Value = "Received"
        End If
        
        filename = Dir()
    Loop
    
    Call UpdateHHTable
    
End Sub

' Helper Functions

Sub UpdateCOITable()
    Dim mainTable As ListObject
    Dim row As ListRow
    Dim folderPath As String
    Dim filename As String
    Dim woNum As String
    
    Set mainTable = Sheet1.ListObjects("COIRequest")
    folderPath = CONST_USER_PATH & "Downloads\1_COI_WO\"
    ' Looks through folder for a file with matching work order if:
    ' (1) wo is active (2) coi status is "req"
    For Each row In mainTable.ListRows
        If Not row.Range.Cells(1, 17).Value And row.Range.Cells(1, 11).Value Like "*req*" Then
            woNum = row.Range.Cells(1, 1).Value
            ' Look through folder for matching coi file
            filename = Dir(folderPath & "*" & woNum & "*")
            If filename <> "" Then
                row.Range.Cells(1, 11).Value = "rec" & vbNewLine & Format(FileDateTime(folderPath & filename), "mm/dd")
            End If
        End If
    Next row
End Sub

Sub UpdateWCTable()
    Dim mainTable As ListObject
    Dim row As ListRow
    Dim certHolder As String
    Dim company As String
    Dim folderPath As String
    Dim filename As String
    
    Set mainTable = Sheet1.ListObjects("COIRequest")
    folderPath = CONST_USER_PATH & "Downloads\1_WC_Cert\"
    ' Looks through folder for a file with matching cert holder if:
    ' (1) wo is active (2) wc status is "req"
    For Each row In mainTable.ListRows
        If Not row.Range.Cells(1, 17).Value And row.Range.Cells(1, 12).Value Like "*req*" Then
            company = row.Range.Cells(1, 4).Value
            certHolder = row.Range.Cells(1, 14).Text
            If certHolder Like "*INDIVIDUAL*" Then certHolder = row.Range.Cells(1, 6).Value
            ' Look through folder for matching wc file
            filename = Dir(folderPath & "*" & certHolder & "*" & company & "*")
            If filename <> "" Then
                row.Range.Cells(1, 12).Value = "rec" & vbNewLine & Format(FileDateTime(folderPath & filename), "mm/dd")
            End If
        End If
    Next row
End Sub

Sub UpdateHHTable()
    Dim mainTable As ListObject
    Dim row As ListRow
    Dim clientName As String
    Dim address As String
    Dim folderPath As String
    Dim filename As String
    
    Set mainTable = Sheet1.ListObjects("COIRequest")
    folderPath = CONST_USER_PATH & "Downloads\1_Indemnity-HH\"
    ' Looks through folder for a file if:
    ' (1) wo is active (2) hh status is not "no"
    For Each row In mainTable.ListRows
        If Not row.Range.Cells(1, 17).Value And InStr(1, row.Range.Cells(1, 13).Value, "no") = 0 And InStr(1, row.Range.Cells(1, 13).Value, "sent") = 0 And Not IsError(row.Range.Cells(1, 14).Value) Then
            clientName = row.Range.Cells(1, 6).Value
            address = row.Range.Cells(1, 3).Value
            filename = Dir(folderPath & "*" & Replace(clientName, "/", "") & "*")
            If filename = "" Then filename = Dir(folderPath & "Indemnity-" & address & "*")
            If filename <> "" Then
                row.Range.Cells(1, 13).Value = "signed" & vbNewLine & Format(FileDateTime(folderPath & filename), "mm/dd")
            End If
        End If
    Next row
End Sub