Attribute VB_Name = "Module1"
Sub GetMembersOfGALDistributionList()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim oGAL As Outlook.AddressList
    Dim oDistList As Outlook.AddressEntry
    Dim oExDistList As Outlook.ExchangeDistributionList
    Dim members As Outlook.AddressEntries
    Dim oMember As Outlook.AddressEntry
    Dim i As Long
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set oGAL = olNamespace.AddressLists("Global Address List")
    
    On Error Resume Next
    Set oDistList = oGAL.AddressEntries("F.CCN GGP GGP1")
    On Error GoTo 0
    
    If Not oDistList Is Nothing Then
        If oDistList.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
            Set oExDistList = oDistList.GetExchangeDistributionList
            If Not oExDistList Is Nothing Then
                Set members = oExDistList.members
                If Not members Is Nothing Then
                    For Each oMember In members
                        Debug.Print oMember.Name
                    Next oMember
                End If
            End If
        End If
    End If
    
    Set olApp = Nothing
    Set olNamespace = Nothing
    Set oGAL = Nothing
    Set oDistList = Nothing
    Set oExDistList = Nothing
End Sub

