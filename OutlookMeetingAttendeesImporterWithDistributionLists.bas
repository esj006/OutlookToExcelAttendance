Attribute VB_Name = "Module1"
Sub ImportMeetingAttendeesToExcel()
    Dim olApp As Object, olNamespace As Object, olFolder As Object
    Dim olItems As Object, olAppt As Object, olRecipient As Object
    Dim ws As Worksheet, r As Long
    Dim Found As Range

    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(9) ' olFolderCalendar
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Endre til riktig arknavn
    r = 1 ' Starter på første rad

    Dim MeetingSubject As String
    MeetingSubject = "GGP1 gathering at Stord 16-17. November 2023"

    Set olItems = olFolder.Items
    Set olAppt = olItems.Find("[Subject] = '" & MeetingSubject & "'")

    If Not olAppt Is Nothing Then
        For Each olRecipient In olAppt.Recipients
            olRecipient.Resolve
            If olRecipient.AddressEntry.GetExchangeUser Is Nothing And _
               Not olRecipient.AddressEntry.GetExchangeDistributionList Is Nothing Then
                ProcessDistributionList olRecipient.AddressEntry, ws, r
            Else
                Set Found = ws.Range("A:A").Find(What:=olRecipient.Name, LookIn:=xlValues, LookAt:=xlWhole)
                If Found Is Nothing Then
                    Debug.Print "Behandler enkeltbruker: " & olRecipient.Name
                    ws.Cells(r, 1).Value = olRecipient.Name
                    r = r + 1
                End If
            End If
        Next olRecipient
    Else
        MsgBox "Møte med tittelen '" & MeetingSubject & "' ble ikke funnet."
    End If

    Set olApp = Nothing
    Set olNamespace = Nothing
    Set olFolder = Nothing
    Set olAppt = Nothing
End Sub

Private Sub ProcessDistributionList(oDL As Object, ws As Worksheet, ByRef r As Long)
    Dim members As Object, oMember As Object
    Dim Found As Range

    On Error Resume Next
    Set members = oDL.members
    If Err.Number <> 0 Then
        Debug.Print "Feil ved å hente medlemmer fra distribusjonslisten: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If Not members Is Nothing Then
        Debug.Print "Fant medlemmer i distribusjonslisten: " & oDL.Name
        For Each oMember In members
            On Error Resume Next
            Set Found = ws.Range("A:A").Find(What:=oMember.Name, LookIn:=xlValues, LookAt:=xlWhole)
            If Found Is Nothing And Err.Number = 0 Then
                ws.Cells(r, 1).Value = oMember.Name
                Debug.Print "Medlem: " & oMember.Name
                r = r + 1
            End If
            On Error GoTo 0
        Next oMember
    Else
        Debug.Print "Ingen medlemmer funnet i distribusjonslisten: " & oDL.Name
    End If
End Sub



