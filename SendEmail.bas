Attribute VB_Name = "SendEmail"
Sub SendEmailWithBCC_v3()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim mailRecip As String
    Dim bccList As String
    Dim olApp As Object
    Dim olMail As Object
    Dim emailSubject As String
    Dim hasNone As Boolean
    Dim hasTentative As Boolean
    
    ' Assume you're using the first sheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Dynamically find the last row in column D
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    ' Get the subject from K2
    emailSubject = "Awaiting Your Feedback on " & ws.Range("K2").Value

    ' Initialize boolean variables
    hasNone = False
    hasTentative = False

    For i = 2 To lastRow
        ' Check if there's an email address in column D
        If ws.Cells(i, 4).Value <> "" Then
            Dim responseValue As String
            responseValue = ws.Cells(i, 3).Value

            ' Check for None and Tentative responses
            If responseValue = "None" Then hasNone = True
            If responseValue = "Tentative" Then hasTentative = True

            ' Collect emails for now, will filter later
            mailRecip = ws.Cells(i, 4).Value
            If bccList = "" Then
                bccList = mailRecip
            Else
                bccList = bccList & ";" & mailRecip
            End If
        End If
    Next i
    
    ' Check if there are any email addresses in bccList
    If bccList = "" Then
        MsgBox "No email addresses found in the table. Please ensure there are valid email addresses and try again.", vbExclamation, "No Email Addresses"
        Exit Sub
    End If

    ' If there are None responses, ask user about them
    If hasNone Then
        Dim includeNone As VbMsgBoxResult
        includeNone = MsgBox("Do you want to include emails with 'None' response?", vbYesNoCancel, "Include 'None' Response?")
        If includeNone = vbCancel Then Exit Sub
    End If

    ' If there are Tentative responses, ask user about them
    If hasTentative Then
        Dim includeTentative As VbMsgBoxResult
        includeTentative = MsgBox("Do you want to include emails with 'Tentative' response?", vbYesNoCancel, "Include 'Tentative' Response?")
        If includeTentative = vbCancel Then Exit Sub
    End If
    
' Initialize the Outlook application
Set olApp = CreateObject("Outlook.Application")

Dim OApp As Object, OMail As Object, signature As String, emailBody As String

' Get the email body from K3
emailBody = Replace(ws.Range("K3").Value, Chr(10), "<br>")

Set OApp = CreateObject("Outlook.Application")
Set OMail = OApp.CreateItem(0)

With OMail
    .Display
End With

signature = OMail.HTMLBody

With OMail
    .BCC = bccList
    .Subject = emailSubject
    .HTMLBody = emailBody & "<br><br>" & "<br>" & signature
    .Display ' This will display the email. Replace with .Send to send directly
End With

Set OMail = Nothing
Set OApp = Nothing
End Sub




