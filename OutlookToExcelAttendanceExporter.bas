Attribute VB_Name = "V01DataFOutlook"
Option Explicit

Sub EksporterDeltakereFraOutlookMote()

    ' Define variables
    Dim olApp As Object, olNamespace As Object, olFolder As Object
    Dim olItems As Object, olAppointment As Object
    Dim ws As Worksheet, iRow As Integer
    Dim olRecipient As Object
    Dim attendanceStatus As String, responseStatus As String
    Dim organizerName As String
    Dim requiredAttendees() As String, optionalAttendees() As String
    Dim tbl As ListObject, rng As Range
    Dim meetingTitle As String
    Dim foundMeeting As Boolean
    Dim organizerCount As Integer
    Dim alreadyNotified As Boolean
    Dim countAccepted As Integer, countTentative As Integer
    Dim countDeclined As Integer, countNone As Integer
    Dim olExchangeUser As Object


    countAccepted = 0
    countTentative = 0
    countDeclined = 0
    countNone = 0
    
    ' Use existing worksheetA
    Set ws = ThisWorkbook.Sheets(1)
       
   ' Connect to Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(9) ' 9 = olFolderCalendar
    Set olItems = olFolder.Items

    ' Get meeting title
     meetingTitle = ws.Cells(2, 11).Value
    foundMeeting = False
    
   If meetingTitle = "" Then
    MsgBox "Please enter a meeting title in K2 and try again.", vbCritical, "Error: Meeting Title Missing"
    ClearTableData ws, "OutlookData"
    Exit Sub
End If


    ' Check if meeting with specified title exists in Outlook
    
    foundMeeting = False
    
    For Each olAppointment In olItems
     If olAppointment.Subject = meetingTitle Then
    foundMeeting = True
       If DataIsUnchanged(ws, olAppointment) Then
           MsgBox "The data in the worksheet matches the data in Outlook. No updates are required.", vbInformation, "No Update Needed"
           Exit Sub
       End If
              
          ' Before defining the range for the new table, delete the existing table if it exists
    If TableExists(ws, "OutlookData") Then
    ws.ListObjects("OutlookData").Delete
    End If
             
       ' Clear sheet before adding new data
       
       ws.Range("E2:H2").Clear
       
        ' Clear data, but not headers
        ws.Range("A2:C" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row).Clear


       ' Get the organizer's name
       organizerName = olAppointment.GetOrganizer.Name

       organizerCount = 0

       ' Split the required and optional attendees
       requiredAttendees = Split(olAppointment.requiredAttendees, ";")
       optionalAttendees = Split(olAppointment.optionalAttendees, ";")

       Dim organizerProcessed As Boolean
       organizerProcessed = False

        Dim iRowForSummary As Integer
        iRowForSummary = iRow
        
        ' Reset iRow
        iRow = 2

       For Each olRecipient In olAppointment.Recipients
           If olRecipient.Name = organizerName Then
               organizerCount = organizerCount + 1
           End If
       Next olRecipient

       If organizerCount > 1 And Not alreadyNotified Then
          MsgBox "The name of the meeting organizer, " & organizerName & ", appears " & organizerCount & " times. For summary purposes, only the instance that does not have the status 'Organizer' will be included.", vbInformation, "Multiple Entries Noted"


           alreadyNotified = True
       End If

       ' Reset iRow
         iRow = 2

       For Each olRecipient In olAppointment.Recipients
           If olRecipient.Name = organizerName And Not organizerProcessed Then
               attendanceStatus = "Meeting Organizer"
               organizerProcessed = True
           ElseIf IsInArray(olRecipient.Name, requiredAttendees) Then
               attendanceStatus = "Required Attendee"
           ElseIf IsInArray(olRecipient.Name, optionalAttendees) Then
               attendanceStatus = "Optional Attendee"
           Else
               attendanceStatus = "Unknown"
           End If
            
                ' Populate the sheet
                ws.Cells(iRow, 1).Value = olRecipient.Name
                ws.Cells(iRow, 2).Value = attendanceStatus
                
                If olRecipient.AddressEntry.Type = "EX" Then
                    Set olExchangeUser = olRecipient.AddressEntry.GetExchangeUser()
                    If Not olExchangeUser Is Nothing Then
                        ws.Cells(iRow, 4).Value = olExchangeUser.PrimarySmtpAddress
                    Else
                        ws.Cells(iRow, 4).Value = olRecipient.Address
                    End If
                Else
                    ws.Cells(iRow, 4).Value = olRecipient.Address
                End If


                Dim colorCode As Long
                Select Case olRecipient.MeetingResponseStatus
                    Case 0:
                        ws.Cells(iRow, 3).Value = "None"
                        colorCode = RGB(208, 206, 206) ' #D0CECE
                    Case 1:
                        ws.Cells(iRow, 3).Value = "Organizer"
                        colorCode = RGB(208, 206, 206) ' #D0CECE
                    Case 2:
                        ws.Cells(iRow, 3).Value = "Tentative"
                        colorCode = RGB(255, 242, 204) ' #FFF2CC
                    Case 3:
                        ws.Cells(iRow, 3).Value = "Accepted"
                        colorCode = RGB(226, 239, 218) ' #E2EFDA
                    Case 4:
                        ws.Cells(iRow, 3).Value = "Declined"
                        colorCode = RGB(252, 228, 214) ' #FCE4D6
                End Select
                
                
                If Not (organizerCount > 1 And olRecipient.Name = organizerName) Then
                   Select Case olRecipient.MeetingResponseStatus
                       Case 0:
                           countNone = countNone + 1
                       Case 2:
                           countTentative = countTentative + 1
                       Case 3:
                           countAccepted = countAccepted + 1
                       Case 4:
                           countDeclined = countDeclined + 1
                   End Select
                End If

                ' Color the whole row based on the response
                ws.Range("A" & iRow & ":D" & iRow).Interior.Color = colorCode

                iRow = iRow + 1
            Next olRecipient
            iRowForSummary = iRow
            
                End If
            Next olAppointment
            
       If Not foundMeeting Then
    MsgBox "The specified meeting title in K2 does not exist in Outlook. Please ensure the title is accurate and try again.", vbCritical, "Error: Invalid Meeting Title"
    
    ' Clear table data and formatting
    ClearTableData ws, "OutlookData"
    
    Exit Sub
End If

    'Define the range for the table
    Set rng = ws.Range("A1:D" & iRow - 1)

    'Convert the range to a table
   Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "OutlookData"
   tbl.TableStyle = "TableStyleMedium9" ' This is optional and can be adjusted to your preference

    ' Set headers
    ws.Cells(1, 1).Value = "Name"
    ws.Cells(1, 2).Value = "Attendance"
    ws.Cells(1, 3).Value = "Response"
    ws.Cells(1, 10).Value = "Input parameters"
    ws.Cells(1, 11).Value = "Value"
    ws.Cells(1, 4).Value = "Email"
    ws.Cells(2, 10).Value = "Meeting Title"
    ws.Cells(3, 10).Value = "Email Text"
   

    iRow = 2

    ' Add summary columns with formulas
  
      Application.AutoCorrect.AutoFillFormulasInLists = False
      Application.AutoCorrect.AutoExpandListRange = False

      
      With ws
            
      
        .Cells(1, 5).Value = "Accepted"
        If organizerCount > 1 Then
            .Cells(2, 5).Formula = "=COUNTIFS(C2:C" & iRowForSummary - 1 & ",""Accepted"",B2:B" & iRowForSummary - 1 & ",""<>Meeting Organizer"")"
        Else
            .Cells(2, 5).Formula = "=COUNTIF(C2:C" & iRowForSummary - 1 & ",""Accepted"")"
        End If
        .Range("E1:E2").Interior.Color = RGB(226, 239, 218) ' #E2EFDA
        
        .Cells(1, 6).Value = "Tentative"
        If organizerCount > 1 Then
            .Cells(2, 6).Formula = "=COUNTIFS(C2:C" & iRowForSummary - 1 & ",""Tentative"",B2:B" & iRowForSummary - 1 & ",""<>Meeting Organizer"")"
        Else
            .Cells(2, 6).Formula = "=COUNTIF(C2:C" & iRowForSummary - 1 & ",""Tentative"")"
        End If
        .Range("F1:F2").Interior.Color = RGB(255, 242, 204) ' #FFF2CC
        
        .Cells(1, 7).Value = "Declined"
        If organizerCount > 1 Then
            .Cells(2, 7).Formula = "=COUNTIFS(C2:C" & iRowForSummary - 1 & ",""Declined"",B2:B" & iRowForSummary - 1 & ",""<>Meeting Organizer"")"
        Else
            .Cells(2, 7).Formula = "=COUNTIF(C2:C" & iRowForSummary - 1 & ",""Declined"")"
        End If
        .Range("G1:G2").Interior.Color = RGB(252, 228, 214) ' #FCE4D6
        
        .Cells(1, 8).Value = "None"
        If organizerCount > 1 Then
            .Cells(2, 8).Formula = "=COUNTIFS(C2:C" & iRowForSummary - 1 & ",""None"",B2:B" & iRowForSummary - 1 & ",""<>Meeting Organizer"")"
        Else
            .Cells(2, 8).Formula = "=COUNTIF(C2:C" & iRowForSummary - 1 & ",""None"")"
        End If
        .Range("H1:H2").Interior.Color = RGB(208, 206, 206) ' #D0CECE
        
    .Cells(1, 9).Value = "Total"
    .Cells(2, 9).Formula = "=SUM(E2:H2)"
    .Range("I1:I2").Interior.Color = RGB(255, 255, 255)
        
        
    End With
    
        Application.AutoCorrect.AutoFillFormulasInLists = True
        Application.AutoCorrect.AutoExpandListRange = True

        ' Colour cells and set formatting
        With ws
        
            ' First set of cells
            Dim hdrCells As Range
            Set hdrCells = .Range("E1:Q1")
            
            hdrCells.Font.Bold = True
            hdrCells.Interior.Color = RGB(68, 114, 196) ' #4472C4
            hdrCells.Font.Color = RGB(255, 255, 255)
        
            ' Second set of cells
            Set hdrCells = Union(.Range("J2:J2"), .Range("J2:J3"))
            
            hdrCells.Font.Bold = True
            hdrCells.Interior.Color = RGB(180, 198, 231)
            hdrCells.Font.Color = RGB(255, 255, 255)
            
        End With
       

    ' Autofit the summary columns
        ws.Range("F1:K2").Columns.AutoFit
        ws.Columns("A:A").AutoFit
        ws.Columns("B:B").AutoFit
        ws.Columns("C:C").AutoFit
        ws.Columns("D:D").AutoFit
    
    'Merge and center
        With ws.Range("K3:Q11")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        End With
        
        With ws.Range("J3:J11")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        End With
        
          With ws.Range("K1:Q1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        End With
        
          With ws.Range("K2:Q2")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        End With
        
        
    
    ' Borders: xlHairline, xlThin, xlMedium, xlThick
    
        'Columns
            With ws.Columns("D:D")
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
            End With
            
            With ws.Columns("I:I")
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
            End With
            
        'Vertical cell range
         With Union(ws.Range("K1:k2"), ws.Range("J1:J11"), ws.Range("Q1:Q11")).Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
         End With
         
         'Horizontal cell range
          With Union(ws.Range("A1:Q1"), ws.Range("E2:Q2"), ws.Range("J11:Q11")).Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
         End With

  ' Align an wrap text
     ws.Range("K3").WrapText = True
     
     With ws.Range("K3")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlTop
    End With
    
    With ws.Range("J:J", "K1:K2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Cleanup
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, valToBeFound)) > -1)
End Function

Function DataIsUnchanged(ws As Worksheet, olAppointment As Object) As Boolean
    Dim olRecipient As Object
    Dim iRow As Integer
    Dim organizerName As String
    Dim requiredAttendees() As String, optionalAttendees() As String
    organizerName = olAppointment.GetOrganizer.Name
    requiredAttendees = Split(olAppointment.requiredAttendees, ";")
    optionalAttendees = Split(olAppointment.optionalAttendees, ";")
    Dim organizerProcessed As Boolean
    organizerProcessed = False
    
    
    iRow = 2

    For Each olRecipient In olAppointment.Recipients
        If ws.Cells(iRow, 1).Value <> olRecipient.Name Then
            DataIsUnchanged = False
            Exit Function
        End If

        If olRecipient.Name = organizerName And Not organizerProcessed Then
            If ws.Cells(iRow, 2).Value <> "Meeting Organizer" Then
                DataIsUnchanged = False
                Exit Function
            End If
            organizerProcessed = True
        ElseIf IsInArray(olRecipient.Name, requiredAttendees) Then
            If ws.Cells(iRow, 2).Value <> "Required Attendee" Then
                DataIsUnchanged = False
                Exit Function
            End If
        ElseIf IsInArray(olRecipient.Name, optionalAttendees) Then
            If ws.Cells(iRow, 2).Value <> "Optional Attendee" Then
                DataIsUnchanged = False
                Exit Function
            End If
        Else
            If ws.Cells(iRow, 2).Value <> "Unknown" Then
                DataIsUnchanged = False
                Exit Function
            End If
        End If

        Select Case olRecipient.MeetingResponseStatus
            Case 0:
                If ws.Cells(iRow, 3).Value <> "None" Then
                    DataIsUnchanged = False
                    Exit Function
                End If
            Case 1:
                If ws.Cells(iRow, 3).Value <> "Organizer" Then
                    DataIsUnchanged = False
                    Exit Function
                End If
            Case 2:
                If ws.Cells(iRow, 3).Value <> "Tentative" Then
                    DataIsUnchanged = False
                    Exit Function
                End If
            Case 3:
                If ws.Cells(iRow, 3).Value <> "Accepted" Then
                    DataIsUnchanged = False
                    Exit Function
                End If
            Case 4:
                If ws.Cells(iRow, 3).Value <> "Declined" Then
                    DataIsUnchanged = False
                    Exit Function
                End If
        End Select
        iRow = iRow + 1
    Next olRecipient

    DataIsUnchanged = True
End Function

Function TableExists(ws As Worksheet, tableName As String) As Boolean
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not tbl Is Nothing Then TableExists = True
End Function

Sub ClearTableData(ByRef ws As Worksheet, ByVal tableName As String)
    Dim tbl As ListObject
    Dim rng As Range
    Dim lastRow As Long
    
    ' Check if table exists
    If TableExists(ws, tableName) Then
        Set tbl = ws.ListObjects(tableName)
        
        ' Clear the data rows without affecting the headers
        If tbl.ListRows.Count > 0 Then
            tbl.DataBodyRange.ClearContents
            tbl.DataBodyRange.Interior.ColorIndex = xlNone
        End If
        
        ' Resize the table to fit the headers only
        lastRow = tbl.HeaderRowRange.Row
        Set rng = ws.Range(tbl.HeaderRowRange.Cells(1, 1).Address, tbl.HeaderRowRange.Cells(1, tbl.ListColumns.Count).Offset(1, 0).Address)
        tbl.Resize rng
    End If
End Sub









