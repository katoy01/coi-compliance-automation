' Module 1

Option Explicit

' === CONFIGURATION CONSTANTS ===
Private Const CONST_EMAIL_MGMT_1 As String = "manager1@company.com"
Private Const CONST_EMAIL_MGMT_2 As String = "manager2@company.com"
Private Const CONST_EMAIL_BROKER_COI As String = "coi.request@broker1.com"
Private Const CONST_EMAIL_BROKER_WC As String = "wc.request@broker2.com"
Private Const CONST_EMAIL_SIGNATURE As String = "user@company.com"
Private Const CONST_USER_PATH As String = "C:\Users\username\"

' Requests insurance for WO#s that are selected on Spreadsheet
Sub RequestSelected()
'' Set up Dim variables
Dim row As ListRow
Dim response As Integer
Dim send As Integer

'' Code starts here
' Morning or Afternoon
response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
If response = vbCancel Then Exit Sub
' Send or Draft
send = MsgBox("Send (yes) or Draft (no)?", vbYesNoCancel)
If send = vbCancel Then Exit Sub


' Iterate through each row
For Each row In Sheet1.ListObjects("COIRequest").ListRows
    If row.Range.Cells(1, 15).Value = True Then
        ' Calls a helper function
        Call DraftRequest(row.Index, response, send)
        row.Range.Cells(1, 15).Value = False
    End If
Next row

End Sub

' Sends insurance gathered to Clients+Mgmt+FigliaTeam for WO#s selected on Spreadsheet
Sub DraftSelected()
'' Set up Dim variables
Dim row As ListRow
Dim response As Integer
Dim blank As Integer
Dim send As Integer

'' Code starts here
' Morning or Afternoon
response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
If response = vbCancel Then Exit Sub
' Blank or Attachments
blank = MsgBox("Blank attachments?", vbYesNoCancel)
If blank = vbCancel Then Exit Sub
' Send or Draft
If blank = vbNo Then
    send = MsgBox("Send (yes) or Draft (no)?", vbYesNoCancel)
    If send = vbCancel Then Exit Sub
Else
    send = vbNo
End If

' Iterate through each row
For Each row In Sheet1.ListObjects("COIRequest").ListRows
    If row.Range.Cells(1, 15).Value = True Then
        ' Call a helper function
        Call DraftToClient(row.Index, response, blank, send)
        row.Range.Cells(1, 15).Value = False
    End If
Next row

End Sub

' Internal helper functions called by above
' Made private to cause less confusion
Private Sub DraftRequest(rowNum As Long, response As Integer, send As Integer)
' Set up Dim variables
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Range
    
    Dim emailSubject As String
    Dim emailBody As String
    Dim recipient As String
    Dim closedCopy As String
    Dim company As String
    Dim companyTag1 As String
    Dim companyTag2 As String
    Dim clientName As String
    Dim woNum As String
    Dim address As String
    Dim aptNum As String
    Dim WCFilePath As String
    Dim certHolder As String
    
    ' Set up active row
    Set activeRow = Sheet1.ListObjects("COIRequest").DataBodyRange.Rows(rowNum)
    
    ' Set up Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Set up constants
    recipient = CONST_EMAIL_BROKER_COI & ";"
    
    If activeRow.Cells(1, 14).Text Like "*INDIVIDUAL*" Then
        certHolder = activeRow.Cells(1, 6).Value
    Else
        certHolder = activeRow.Cells(1, 14).Text
    End If
    
    WCFilePath = Dir(CONST_USER_PATH & "Downloads\1_WC_Cert\WC_" & "*" & certHolder & "*" & company & "*")
    If WCFilePath = "" Then recipient = recipient & CONST_EMAIL_BROKER_WC & ";"
    
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    ' Get data from the active row
    company = activeRow.Cells(1, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = activeRow.Cells(1, 6).Value
    woNum = activeRow.Cells(1, 1).Value
    address = activeRow.Cells(1, 3).Value
    If activeRow.Cells(1, 5).Value <> "" Then aptNum = " Apt " & activeRow.Cells(1, 5).Value
    

    ' Compose the email subject
    emailSubject = "" & companyTag1 & " Priority - " & clientName

    ' Compose the email body
    Dim timeOfDay As String
    If response = vbYes Then
        timeOfDay = "morning,"
    Else
        If response = vbNo Then
            timeOfDay = "afternoon,"
        Else
            Exit Sub
        End If
    End If
     
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body> Good " & timeOfDay & "<br><br>" & _
                "Please provide " & companyTag2 & " Certificates for the following:" & "<br><br>" & _
                "-  WO# " & woNum & "<br>" & _
                "-  " & clientName & "<br>" & _
                "-  " & address & " " & aptNum & "<br><br>" & _
                "Thank you!" & "<br><br><br>" & S & _
                "</body></html>"
    
    ' Set email parameters
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\2_COI_Sample_Reqs\COI Reqs - " & address & ".pdf"
    End With
    
    If send = vbYes Then
        OutMail.send
    Else
        OutMail.Display
    End If

    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    activeRow.Cells(1, 11) = "req" & vbNewLine & Format(Date, "mm/dd")
    activeRow.Cells(1, 12) = "req" & vbNewLine & Format(Date, "mm/dd")
    
End Sub

Private Sub DraftToClient(rowNum As Long, response As Integer, blank As Integer, send As Integer)
' Section#1 Set Variables
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Range
    
    Dim emailSubject As String
    Dim emailBody As String
    Dim recipient As String
    Dim closedCopy As String
    
    Dim company As String
    Dim companyTag1 As String
    Dim companyTag2 As String
    Dim clientName As String
    Dim woNum As String
    Dim address As String
    Dim aptNum As String
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    Dim timeOfDay As String
    Dim S As String
    Dim mgmtCopy As String
    
    Dim WOFilePath As String
    Dim WCFilePath As String
    Dim HHFilePath As String
    Dim certHolder As String
    Dim companyCheck As Integer
    Dim wcCheck As Integer
    
    ' Set up active row
    Set activeRow = Sheet1.ListObjects("COIRequest").DataBodyRange.Rows(rowNum)
    
    ' Set up Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Set up constants
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    ' Get data from the active row
    company = activeRow.Cells(1, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = activeRow.Cells(1, 6).Value
    woNum = activeRow.Cells(1, 1).Value
    address = activeRow.Cells(1, 3).Value
    If activeRow.Cells(1, 5).Value <> "" Then aptNum = " Apt " & activeRow.Cells(1, 5).Value
    scopeWork = activeRow.Cells(1, 7).Value
    clientEmail = activeRow.Cells(1, 8).Value
    mgmtEmail = activeRow.Cells(1, 9).Value
    apptDate = activeRow.Cells(1, 2).Value
    
    If activeRow.Cells(1, 14).Text Like "*INDIVIDUAL*" Then
        certHolder = clientName
    Else
        certHolder = activeRow.Cells(1, 14).Text
    End If
    
    ' If SalesRep isn't included
    If Not IsEmpty(activeRow.Cells(1, 10).Value) Then
        salesRep = activeRow.Cells(1, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").ListObjects("SalesContact").DataBodyRange, 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
' Section#2 Draft the email
    ' Set recipient and cc
    recipient = clientEmail
    
    If mgmtEmail <> "0" And mgmtEmail <> "?" And Not IsEmpty(mgmtEmail) Then
        closedCopy = closedCopy & mgmtEmail
        mgmtCopy = "I have copied your building management on this email." & "<br>"
    Else
        mgmtCopy = "Kindly share these documents with any appropriate contacts for review." & "<br>"
    End If
    
    If InStr(1, activeRow.Cells(1, 13).Value, "no") = 0 And InStr(1, activeRow.Cells(1, 13).Value, "sent") = 0 Then
        mgmtCopy = "If it is appropriate to do so, please countersign and send back the attached indemnity." & "<br>" & "<br>" & mgmtCopy
    End If
    
    ' Compose the email subject
    emailSubject = "" & companyTag2 & " - Insurance for " & address & aptNum & " (WO# " & woNum & ")"

    ' Compose the email body
    If response = vbYes Then
        timeOfDay = "morning"
    Else
        If response = vbNo Then
            timeOfDay = "afternoon"
        Else
            Exit Sub
        End If
    End If
    
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & aptNum & "." & "<br>" & _
                companyTag2 & " is scheduled to perform " & scopeWork & " for " & address & aptNum & " on " & Format(apptDate, "dddd, mmmm d, yyyy") & ", between 9AM-4PM." & "<br>" & _
                mgmtCopy & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
    
    ' Set email parameters
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
    End With
    
' Section#3 Attach documents

    If blank = vbNo Then
        With OutMail
            ' COI
            WOFilePath = Dir(CONST_USER_PATH & "Downloads\1_COI_WO\COI_" & "*" & woNum & "*")
            If WOFilePath <> "" Then
                .Attachments.Add CONST_USER_PATH & "Downloads\1_COI_WO\" & WOFilePath
                activeRow.Cells(1, 11) = "sent" & vbNewLine & Format(Date, "mm/dd")
            Else
                MsgBox "No COI Found for WO# " & woNum
                Exit Sub
            End If
            
            ' WC
            WCFilePath = Dir(CONST_USER_PATH & "Downloads\1_WC_Cert\WC_" & "*" & certHolder & "*" & company & "*")
            If WCFilePath <> "" Then
                .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Cert\" & WCFilePath
                activeRow.Cells(1, 12) = "sent" & vbNewLine & Format(Date, "mm/dd")
            Else
                WCFilePath = Dir(CONST_USER_PATH & "Downloads\1_WC_Cert\WC_" & "*" & certHolder & "*")
                If WCFilePath <> "" Then
                    companyCheck = MsgBox("Wrong company OK?", vbYesNo)
                    If companyCheck = vbYes Then
                        .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Cert\" & WCFilePath
                        activeRow.Cells(1, 12) = "sent" & vbNewLine & Format(Date, "mm/dd")
                    Else
                        MsgBox ("No WC Found for WO# " & woNum & " with company " & company)
                        Exit Sub
                    End If
                Else
                    wcCheck = MsgBox("No WC Found for WO# " & woNum & " with company " & company, vbOKCancel)
                    If wcCheck = vbCancel Then Exit Sub
                End If
            End If
            
            ' If HH
            If InStr(1, activeRow.Cells(1, 13).Value, "no") = 0 And InStr(1, activeRow.Cells(1, 13).Value, "sent") = 0 Then
                HHFilePath = Dir(CONST_USER_PATH & "Downloads\1_Indemnity-HH\Indemnity-" & "*" & Replace(clientName, "/", "") & "*")
                If HHFilePath = "" Then
                    HHFilePath = Dir(CONST_USER_PATH & "Downloads\1_Indemnity-HH\Indemnity-" & "*" & address & "*")
                End If
                If HHFilePath <> "" Then
                    .Attachments.Add CONST_USER_PATH & "Downloads\1_Indemnity-HH\" & HHFilePath
                    activeRow.Cells(1, 13) = "sent" & vbNewLine & Format(Date, "mm/dd")
                Else
                    MsgBox "No HH Found for WO# " & woNum
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If send = vbYes Then
        OutMail.send
    Else
        OutMail.Display
    End If

    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub