' Module 4 — Archived / Deprecated
' Preserved to document pipeline evolution.
' These functions represent earlier iterations before
' the batch-selection and ODBC auto-populate patterns were introduced.

Option Explicit

' === CONFIGURATION CONSTANTS ===
Private Const CONST_EMAIL_MGMT_1 As String = "manager1@company.com"
Private Const CONST_EMAIL_MGMT_2 As String = "manager2@company.com"
Private Const CONST_EMAIL_BROKER_COI As String = "coi.request@broker1.com"
Private Const CONST_EMAIL_BROKER_WC As String = "wc.request@broker2.com"
Private Const CONST_EMAIL_SIGNATURE As String = "user@company.com"
Private Const CONST_USER_PATH As String = "C:\Users\username\"

Private Sub DraftType2_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim activeRow As Long
    
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
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    ' Set up Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Set up constants
    activeRow = ActiveCell.row
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    ' Get data from the active row
    woNum = Cells(activeRow, 1).Value
    apptDate = Format(Cells(activeRow, 2).Value, "dddd, mmmm dd, yyyy")
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value
    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    scopeWork = Cells(activeRow, 7).Value
    clientEmail = Cells(activeRow, 8).Value
    mgmtEmail = Cells(activeRow, 9).Value
    
    If Not IsEmpty(Cells(activeRow, 10).Value) Then
        salesRep = Cells(activeRow, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").Range("A2:B8"), 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
    recipient = clientEmail
    closedCopy = closedCopy & mgmtEmail

    emailSubject = "" & companyTag2 & " - Insurance for " & address & " (WO# " & woNum & ")"

    Dim response As Integer
    Dim timeOfDay As String
    response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
    If response = vbYes Then
        timeOfDay = "morning"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & "." & "<br>" & _
                "Your appointment - " & scopeWork & " - is scheduled for " & apptDate & ", between 9AM-4PM." & "<br>" & _
                "I have copied your building management on this email." & "<br>" & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
    
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub DraftType2wCOIandWC_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Long
    
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
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    activeRow = ActiveCell.row
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    woNum = Cells(activeRow, 1).Value
    apptDate = Format(Cells(activeRow, 2).Value, "dddd, mmmm dd, yyyy")
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value
    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    scopeWork = Cells(activeRow, 7).Value
    clientEmail = Cells(activeRow, 8).Value
    mgmtEmail = Cells(activeRow, 9).Value
    
    If Not IsEmpty(Cells(activeRow, 10).Value) Then
        salesRep = Cells(activeRow, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").Range("A2:B8"), 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
    recipient = clientEmail
    closedCopy = closedCopy & mgmtEmail

    emailSubject = "" & companyTag2 & " - Insurance for " & address & " (WO# " & woNum & ")"

    Dim response As Integer
    Dim timeOfDay As String
    response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
    If response = vbYes Then
        timeOfDay = "morning"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & "." & "<br>" & _
                "Your appointment - " & scopeWork & " - is scheduled for " & apptDate & ", between 9AM-4PM." & "<br>" & _
                "I have copied your building management on this email." & "<br>" & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
                
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\1_COI_Name_Address_Company\COI-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Holder_(Company)\WC-" & Cells(activeRow, 3).Value & " (" & company & ")" & ".pdf"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub DraftType2wCOIandWC_Click1(response)
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Long
    
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
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    activeRow = ActiveCell.row
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    woNum = Cells(activeRow, 1).Value
    apptDate = Format(Cells(activeRow, 2).Value, "dddd, mmmm dd, yyyy")
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value
    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    scopeWork = Cells(activeRow, 7).Value
    clientEmail = Cells(activeRow, 8).Value
    mgmtEmail = Cells(activeRow, 9).Value
    
    If Not IsEmpty(Cells(activeRow, 10).Value) Then
        salesRep = Cells(activeRow, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").Range("A2:B8"), 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
    recipient = clientEmail
    closedCopy = closedCopy & mgmtEmail

    emailSubject = "" & companyTag2 & " - Insurance for " & address & " (WO# " & woNum & ")"

    Dim timeOfDay As String
    If response = vbYes Then
        timeOfDay = "morning"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & "." & "<br>" & _
                "Your appointment - " & scopeWork & " - is scheduled for " & apptDate & ", between 9AM-4PM." & "<br>" & _
                "I have copied your building management on this email." & "<br>" & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
                
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\1_COI_Name_Address_Company\COI-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Holder_(Company)\WC-" & Cells(activeRow, 3).Value & " (" & company & ")" & ".pdf"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Cells(activeRow, 11) = "sent" & vbNewLine & Format(Date, "mm/dd")
    Cells(activeRow, 12) = "sent" & vbNewLine & Format(Date, "mm/dd")
End Sub

Private Sub DraftType2wCOIandWC_Individual_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Long
    
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
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    activeRow = ActiveCell.row
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    woNum = Cells(activeRow, 1).Value
    apptDate = Format(Cells(activeRow, 2).Value, "dddd, mmmm dd, yyyy")
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value
    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    scopeWork = Cells(activeRow, 7).Value
    clientEmail = Cells(activeRow, 8).Value
    mgmtEmail = Cells(activeRow, 9).Value
    
    If Not IsEmpty(Cells(activeRow, 10).Value) Then
        salesRep = Cells(activeRow, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").Range("A2:B8"), 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
    recipient = clientEmail
    closedCopy = closedCopy & mgmtEmail

    emailSubject = "" & companyTag2 & " - Insurance for " & address & " (WO# " & woNum & ")"

    Dim response As Integer
    Dim timeOfDay As String
    response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
    If response = vbYes Then
        timeOfDay = "morning"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & "." & "<br>" & _
                "Your appointment - " & scopeWork & " - is scheduled for " & apptDate & ", between 9AM-4PM." & "<br>" & _
                "I have copied your building management on this email." & "<br>" & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
                
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\1_COI_Name_Address_Company\COI-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Holder_(Company)\WC-I-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub DraftType2wCOIandWC_Individual_Click1(response)
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Long
    
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
    Dim apptDate As String
    Dim scopeWork As String
    Dim clientEmail As String
    Dim mgmtEmail As String
    Dim salesRep As String
    Dim salesRepEmail As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    activeRow = ActiveCell.row
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    woNum = Cells(activeRow, 1).Value
    apptDate = Format(Cells(activeRow, 2).Value, "dddd, mmmm dd, yyyy")
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value
    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    scopeWork = Cells(activeRow, 7).Value
    clientEmail = Cells(activeRow, 8).Value
    mgmtEmail = Cells(activeRow, 9).Value
    
    If Not IsEmpty(Cells(activeRow, 10).Value) Then
        salesRep = Cells(activeRow, 10).Value
        salesRepEmail = Application.VLookup(salesRep, Sheets("Static Info").Range("A2:B8"), 2, False)
        closedCopy = closedCopy & salesRepEmail
    End If
    
    recipient = clientEmail
    closedCopy = closedCopy & mgmtEmail

    emailSubject = "" & companyTag2 & " - Insurance for " & address & " (WO# " & woNum & ")"

    Dim timeOfDay As String
    If response = vbYes Then
        timeOfDay = "morning"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body>" & _
                "Good " & timeOfDay & " " & clientName & "," & "<br>" & "<br>" & _
                "Attached are the insurance certificates for " & address & "." & "<br>" & _
                "Your appointment - " & scopeWork & " - is scheduled for " & apptDate & ", between 9AM-4PM." & "<br>" & _
                "I have copied your building management on this email." & "<br>" & "<br>" & _
                "Thank you & have a nice day!" & "<br><br>" & S & _
                "</body></html>"
                
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\1_COI_Name_Address_Company\COI-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Attachments.Add CONST_USER_PATH & "Downloads\1_WC_Holder_(Company)\WC-I-" & clientName & " " & address & " (" & company & ")" & ".pdf"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub DraftSelected_Individual()
    Dim table As Range, row As Range
    Dim response As Integer

    Set table = ActiveSheet.ListObjects("COIRequest").Range

    response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
    If response = vbCancel Then Exit Sub
    
    For Each row In table.Rows
        If row.Cells(1, 15).Value = "True" Then
            row.Cells(1, 1).Activate
            Call DraftType2wCOIandWC_Individual_Click1(response)
            row.Cells(1, 15).Value = "False"
        End If
    Next row
End Sub

Private Sub DraftType1_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim activeRow As Long
    
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
    Dim apptDate As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    activeRow = ActiveCell.row
    recipient = CONST_EMAIL_BROKER_COI & ";" & CONST_EMAIL_BROKER_WC & ";"
    closedCopy = CONST_EMAIL_MGMT_1 & ";" & CONST_EMAIL_MGMT_2 & ";"

    company = Cells(activeRow, 4).Value
    companyTag1 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 2, False)
    companyTag2 = Application.VLookup(company, Sheets("Static Info").Range("D2:F3"), 3, False)
    clientName = Cells(activeRow, 6).Value
    woNum = Cells(activeRow, 1).Value
    address = Cells(activeRow, 3).Value & " Apt " & Cells(activeRow, 5).Value

    emailSubject = "" & companyTag1 & " Priority - " & clientName

    Dim response As Integer
    Dim timeOfDay As String
    response = MsgBox("Good morning (yes) or afternoon (no)?", vbYesNoCancel)
    If response = vbYes Then
        timeOfDay = "morning,"
    ElseIf response = vbNo Then
        timeOfDay = "afternoon,"
    Else
        Exit Sub
    End If
    
    Dim S As String
    S = CONST_USER_PATH & "AppData\Roaming\Microsoft\Signatures\Official (" & CONST_EMAIL_SIGNATURE & ").htm"
    S = CreateObject("Scripting.FileSystemObject").GetFile(S).OpenAsTextStream(1, -2).ReadAll
    
    emailBody = "<html><body> Good " & timeOfDay & "<br><br>" & _
                "Please provide " & companyTag2 & " Certificates for the following:" & "<br><br>" & _
                "-  WO# " & woNum & "<br>" & _
                "-  " & clientName & "<br>" & _
                "-  " & address & "<br><br>" & _
                "Thank you!" & "<br><br><br>" & S & _
                "</body></html>"
    
    With OutMail
        .To = recipient
        .cc = closedCopy
        .subject = emailSubject
        .HTMLbody = emailBody
        .Attachments.Add CONST_USER_PATH & "Downloads\2_COI_Sample_Reqs\COI Reqs - " & Cells(activeRow, 3).Value & ".pdf"
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Cells(activeRow, 11) = "req" & vbNewLine & Format(Date, "mm/dd")
    Cells(activeRow, 12) = "req" & vbNewLine & Format(Date, "mm/dd")
End Sub